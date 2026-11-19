#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# FACTS tests. Every registry, service, scheduler and console call is mocked.

BeforeDiscovery {
    $script:FactsPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'FACTS.ps1'
}

Describe 'FACTS entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'FACTS.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    }
    It 'defines Invoke-FactsMain and Exit-KillerScript' {
        $names = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-FactsMain'
        $names | Should -Contain 'Exit-KillerScript'
    }
    It 'ends with the test-mode guard and keeps only definitions and constants at the top level' {
        $stmts = @($script:Ast.EndBlock.Statements)
        $stmts[-1].Extent.Text | Should -Match 'KillerScriptsTestMode'
        $stmts[-1].Extent.Text | Should -Match 'Invoke-FactsMain'
        $bad = @($stmts | Select-Object -SkipLast 1 | Where-Object {
            $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $_ -isnot [System.Management.Automation.Language.AssignmentStatementAst]
        })
        ($bad | ForEach-Object { $_.Extent.Text.Split("`n")[0] }) -join "`n" | Should -BeNullOrEmpty
    }
}

Describe 'FACTS' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'FACTS')

        function Get-Run {
            # Runs Main and returns the exit code plus everything written to the output stream.
            $code = -1
            $lines = New-Object System.Collections.ArrayList
            try { Invoke-FactsMain *>&1 | ForEach-Object { [void]$lines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($lines -join "`n") }
        }
        function New-FakeTask {
            param([string]$Name, [string]$Path = '\', [string]$State = 'Ready')
            [pscustomobject]@{ TaskName = $Name; TaskPath = $Path; State = $State }
        }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:Registered = New-Object System.Collections.ArrayList
        $script:Disabled   = New-Object System.Collections.ArrayList
        $script:SetCalls   = New-Object System.Collections.ArrayList
        # Console boundaries
        Mock Write-Host {}
        Mock Clear-Host {}
        Mock Get-FactsCursorTop { -1 }
        Mock Set-FactsCursor { $false }
        Mock Test-FactsElevated { $true }
        # Registry boundaries (64-bit view helpers)
        Mock Get-FactsFoxitApps { [pscustomobject]@{ Ok = $true; Names = @('Foxit PDF Reader') } }
        Mock Get-FactsUpdaterKeys { @('SOFTWARE\Foxit Software\Foxit Reader\Updater', 'SOFTWARE\WOW6432Node\Foxit Software\Foxit PDF Editor\Updater') }
        Mock Set-FactsUpdaterValues { [void]$script:SetCalls.Add($Path); $true }
        # Service and scheduler boundaries
        Mock Get-Service { @([pscustomobject]@{ Name = 'FoxitReaderUpdateService'; StartType = 'Disabled' }) }
        Mock Set-Service {}
        Mock Stop-Service {}
        Mock Get-ScheduledTask {
            if ($TaskName -eq 'FACTS_Universal_Enforcement') { return (New-FakeTask 'FACTS_Universal_Enforcement' '\' 'Ready') }
            @((New-FakeTask 'FoxitUpdateTask' '\' 'Disabled'))
        }
        Mock Stop-ScheduledTask {}
        Mock Disable-ScheduledTask { [void]$script:Disabled.Add(@{ Path = $TaskPath; Name = $TaskName }) }
        Mock Register-ScheduledTask { [void]$script:Registered.Add($TaskName); $script:RegisteredAction = $Action; $null }
        Mock Unregister-ScheduledTask {}
    }

    Context 'entry-point convention' {
        It 'defines Invoke-FactsMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-FactsMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Register-ScheduledTask -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'exits 2 and changes nothing when not elevated' {
            Mock Test-FactsElevated { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
            Should -Invoke Set-FactsUpdaterValues -Times 0 -Exactly
            Should -Invoke Register-ScheduledTask -Times 0 -Exactly
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
        }
        It 'exits 0 with a PASS line on a healthy run with Foxit present' {
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\].*exit 0'
            $script:SetCalls.Count | Should -Be 2
            $script:Registered | Should -Contain 'FACTS_Universal_Enforcement'
        }
        It 'removes the maintenance task and exits 0 when no Foxit is installed' {
            Mock Get-FactsFoxitApps { [pscustomobject]@{ Ok = $true; Names = @() } }
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'FACTS_Universal_Enforcement' }
            Should -Invoke Register-ScheduledTask -Times 0 -Exactly
        }
        It 'exits 1 and leaves the maintenance task alone when the software list cannot be read' {
            Mock Get-FactsFoxitApps { [pscustomobject]@{ Ok = $false; Names = @() } }
            $r = Get-Run
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\]'
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
        }
    }
}
