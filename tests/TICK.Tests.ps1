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
            [pscustomobject]@{ Code = $code; Text = ($lines -join "`n"); HostText = ($script:HostLines -join "`n") }
        }
        function New-W32Result {
            param([int]$Code = 0, [string[]]$Lines = @())
            [pscustomobject]@{ ExitCode = $Code; Output = $Lines }
        }
        function Get-EnglishStatus {
            param([datetime]$LastSync = (Get-Date), [string]$Source = 'time.google.com,0x8')
            @(
            'Leap Indicator: 0(no warning)',
            'Stratum: 3 (secondary reference - syncd by (S)NTP)',
            'Precision: -23 (119.209ns per tick)',
            'Root Delay: 0.0312500s',
            'Root Dispersion: 0.0500000s',
            'ReferenceId: 0x0A0A0A01 (source IP:  10.10.10.1)',
            ('Last Successful Sync Time: ' + $LastSync.ToString()),
            ('Source: ' + $Source),
            'Poll Interval: 6 (64s)'
            )
        }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:W32Calls = New-Object System.Collections.ArrayList
        $script:W32Results = @{}
        $script:Keys = New-Object System.Collections.Queue
        $script:Lines = New-Object System.Collections.Queue
        $script:SourceReply = 'time.google.com,0x8'
        $script:StatusLines = Get-EnglishStatus
        $script:HostLines = New-Object System.Collections.ArrayList
        # Console boundaries
        Mock Write-Host { [void]$script:HostLines.Add([string]$Object) }
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
    Context 'finding 69: not elevated' {
        It 'exits 2 with a FAIL line and touches nothing' {
            Mock Test-TickElevated { $false }
            $r = Get-Run -Silent
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
            Should -Invoke Invoke-TickW32tm -Times 0 -Exactly
            Should -Invoke Stop-Service -Times 0 -Exactly
            Should -Invoke Start-Service -Times 0 -Exactly
        }
    }

    Context 'finding 26: every step is checked and the result is verified' {
        It 'reports a failed /register as a failure, not a success' {
            $script:W32Results['/register'] = New-W32Result 1 @('Access is denied.')
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\]'
            $r.Text | Should -Not -Match '\[PASS\]'
        }
        It 'reports a failed /config as a failure and still starts the time service again' {
            $script:W32Results['/config /manualpeerlist:time.google.com,0x8 /syncfromflags:manual /reliable:yes /update'] = New-W32Result 1 @('error')
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 1
            Should -Invoke Start-Service -Times 1 -Exactly
            $r.HostText | Should -Match '\[FAIL\]'
        }
        It 'exits 1 when the time service is not running after the restart' {
            Mock Get-Service { [pscustomobject]@{ Name = 'w32time'; Status = 'Stopped' } }
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\]'
        }
        It 'exits 3 with a not-confirmed warning when the resync fails' {
            $script:W32Results['/resync /force'] = New-W32Result 1 @('The computer did not resync because no time data was available.')
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\].*not confirmed.*exit 3'
        }
        It 'exits 3 when the reported source is not the requested peer' {
            $script:SourceReply = 'Local CMOS Clock'
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'Local CMOS Clock'
            $r.Text | Should -Not -Match '\[PASS\]'
        }
        It 'exits 3 when the last successful sync is not recent' {
            $script:StatusLines = Get-EnglishStatus -LastSync (Get-Date).AddHours(-3)
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'last successful sync'
        }
        It 'does not throw on non-English status output and reports the sync as unknown with exit 3' {
            $script:StatusLines = @('Indicateur de saut: 0(aucun avertissement)', 'Strate: 3', 'Precision: -23', 'Dispersion de la racine: 0,0500000s', 'Source: time.google.com,0x8')
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'UNKNOWN'
            $r.HostText | Should -Match 'UNKNOWN'
        }
        It 'exits 0 with a PASS line naming the verified source on a fully confirmed sync' {
            $r = Get-Run -Silent -NtpServer 'time.google.com'
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\].*time\.google\.com.*exit 0'
        }
        It 'says the machine is synced to the domain controller only when the source is a domain controller' {
            Mock Get-TickComputerSystem { [pscustomobject]@{ PartOfDomain = $true; DomainRole = 1; Domain = 'corp.local' } }
            Mock Get-TickPdcName { 'dc01.corp.local' }
            $script:SourceReply = 'dc01.corp.local'
            $script:StatusLines = Get-EnglishStatus -Source 'dc01.corp.local'
            $r = Get-Run -Silent
            $r.Code | Should -Be 0
            $script:W32Calls | Should -Contain '/config /syncfromflags:domhier /update'
            $r.HostText | Should -Match 'synced to the Domain Controller'
        }
        It 'does not claim a domain sync when the source is the local clock' {
            Mock Get-TickComputerSystem { [pscustomobject]@{ PartOfDomain = $true; DomainRole = 1; Domain = 'corp.local' } }
            Mock Get-TickPdcName { 'dc01.corp.local' }
            $script:SourceReply = 'Local CMOS Clock'
            $r = Get-Run -Silent
            $r.Code | Should -Be 3
            $r.HostText | Should -Not -Match 'synced to the Domain Controller'
        }
    }

    Context 'finding 26: no console behaves like -Silent and never hangs' {
        It 'does not read a key or a line when there is no console and no parameters' {
            Mock Test-TickInteractive { $false }
            $script:SourceReply = 'time.cloudflare.com,0x8'
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Read-TickKey -Times 0 -Exactly
            Should -Invoke Read-TickLine -Times 0 -Exactly
        }
        It 'exits 2 and changes nothing when Other is chosen and no address is typed' {
            $script:Keys.Enqueue((New-Object System.ConsoleKeyInfo ([char]'8'), ([ConsoleKey]::D8), $false, $false, $false))
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
            Should -Invoke Stop-Service -Times 0 -Exactly
        }
        It 'exits 2 when the given NTP server is not a plausible host name' {
            $r = Get-Run -Silent -NtpServer 'time.google.com 0x8 evil'
            $r.Code | Should -Be 2
            Should -Invoke Invoke-TickW32tm -Times 0 -Exactly
        }
    }
}
