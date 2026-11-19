#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# TICK tests. Every registry, service, CIM, native and console call is mocked.

Describe 'TICK entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'TICK.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    }
    It 'defines Invoke-TickMain and Exit-KillerScript' {
        $names = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-TickMain'
        $names | Should -Contain 'Exit-KillerScript'
    }
    It 'ends with the test-mode guard and keeps only definitions and constants at the top level' {
        $stmts = @($script:Ast.EndBlock.Statements)
        $stmts[-1].Extent.Text | Should -Match 'KillerScriptsTestMode'
        $stmts[-1].Extent.Text | Should -Match 'Invoke-TickMain'
        $bad = @($stmts | Select-Object -SkipLast 1 | Where-Object {
            $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $_ -isnot [System.Management.Automation.Language.AssignmentStatementAst]
        })
        ($bad | ForEach-Object { $_.Extent.Text.Split("`n")[0] }) -join "`n" | Should -BeNullOrEmpty
    }
}

Describe 'TICK' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'TICK')

        function Get-Run {
            # Runs Main and returns the exit code plus everything written to the output stream.
            param([switch]$Silent, [string]$NtpServer)
            $code = -1
            $lines = New-Object System.Collections.ArrayList
            try { Invoke-TickMain -Silent:$Silent -NtpServer $NtpServer *>&1 | ForEach-Object { [void]$lines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($lines -join "`n") }
        }
        function New-W32Result {
            param([int]$Code = 0, [string[]]$Lines = @())
            [pscustomobject]@{ ExitCode = $Code; Output = $Lines }
        }
        $script:EnglishStatus = @(
            'Leap Indicator: 0(no warning)',
            'Stratum: 3 (secondary reference - syncd by (S)NTP)',
            'Precision: -23 (119.209ns per tick)',
            'Root Delay: 0.0312500s',
            'Root Dispersion: 0.0500000s',
            'ReferenceId: 0x0A0A0A01 (source IP:  10.10.10.1)',
            'Last Successful Sync Time: 11/19/2026 9:00:00 AM',
            'Source: time.google.com,0x8',
            'Poll Interval: 6 (64s)'
        )
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:W32Calls = New-Object System.Collections.ArrayList
        $script:W32Results = @{}
        $script:Keys = New-Object System.Collections.Queue
        $script:Lines = New-Object System.Collections.Queue
        $script:SourceReply = 'time.google.com,0x8'
        $script:StatusLines = $script:EnglishStatus
        # Console boundaries
        Mock Write-Host {}
        Mock Clear-Host {}
        Mock Write-HLine {}
        Mock Write-Banner {}
        Mock Test-TickElevated { $true }
        Mock Test-TickInteractive { $true }
        Mock Read-TickKey { if ($script:Keys.Count -gt 0) { $script:Keys.Dequeue() } else { $null } }
        Mock Read-TickLine { if ($script:Lines.Count -gt 0) { $script:Lines.Dequeue() } else { $null } }
        # Machine boundaries
        Mock Get-TickComputerSystem { [pscustomobject]@{ PartOfDomain = $false; DomainRole = 0; Domain = 'WORKGROUP' } }
        Mock Get-TickPdcName { $null }
        Mock Stop-Service {}
        Mock Start-Service {}
        Mock Get-Service { [pscustomobject]@{ Name = 'w32time'; Status = 'Running' } }
        Mock Invoke-TickW32tm {
            $key = ($Arguments -join ' ')
            [void]$script:W32Calls.Add($key)
            if ($script:W32Results.ContainsKey($key)) { return $script:W32Results[$key] }
            if ($key -eq '/query /status') { return (New-W32Result 0 $script:StatusLines) }
            if ($key -eq '/query /source') { return (New-W32Result 0 @($script:SourceReply)) }
            New-W32Result 0 @('The command completed successfully.')
        }
    }

    Context 'entry-point convention' {
        It 'defines Invoke-TickMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-TickMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Invoke-TickW32tm -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'registers, configures, restarts and resyncs against the named server and exits 0' {
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 0
            $script:W32Calls | Should -Contain '/register'
            $script:W32Calls | Should -Contain '/config /manualpeerlist:time.google.com,0x8 /syncfromflags:manual /reliable:yes /update'
            $script:W32Calls | Should -Contain '/resync /force'
            Should -Invoke Stop-Service -Times 1 -Exactly
            Should -Invoke Start-Service -Times 1 -Exactly
        }
        It 'uses the first server in the list when running silent with no server' {
            $script:SourceReply = 'time.cloudflare.com,0x8'
            $null = Get-Run -Silent
            $script:W32Calls | Should -Contain '/config /manualpeerlist:time.cloudflare.com,0x8 /syncfromflags:manual /reliable:yes /update'
        }
    }
}
