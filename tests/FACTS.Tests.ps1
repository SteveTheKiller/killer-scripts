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
        function Get-MaintenanceText {
            # The hourly task body, decoded from the -EncodedCommand argument handed to Register-ScheduledTask.
            $m = [regex]::Match([string]$script:RegisteredAction.Arguments, '-EncodedCommand (\S+)')
            if (-not $m.Success) { return $null }
            [System.Text.Encoding]::Unicode.GetString([Convert]::FromBase64String($m.Groups[1].Value))
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
    Context 'finding 15: maintenance task enforces the full settings set' {
        BeforeEach {
            $script:MaintText = $null
        }
        It 'puts the registry enforcement into the hourly task body' {
            $null = Get-Run
            $script:RegisteredAction | Should -Not -BeNullOrEmpty
            $arg = [string]$script:RegisteredAction.Arguments
            $text = Get-MaintenanceText
            $tokens = $null; $errors = $null
            $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$tokens, [ref]$errors)
            @($errors).Count | Should -Be 0
            $calls = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.CommandAst] }, $true) | ForEach-Object { $_.GetCommandName() })
            $calls | Should -Contain 'Get-FactsUpdaterKeys'
            $calls | Should -Contain 'Set-FactsUpdaterValues'
            $calls | Should -Contain 'Disable-ScheduledTask'
        }
        It 'exits nonzero from the hourly task body when a value does not stick' {
            $null = Get-Run
            $text = Get-MaintenanceText
            $text | Should -Match 'exit 3'
        }
        It 'keeps the task folder when it stops and disables a Foxit update task in a subfolder' {
            Mock Get-ScheduledTask {
                if ($TaskName -eq 'FACTS_Universal_Enforcement') { return (New-FakeTask 'FACTS_Universal_Enforcement' '\' 'Ready') }
                if ($TaskName -eq 'FoxitUpdateTask') { return (New-FakeTask 'FoxitUpdateTask' '\Foxit Software\' 'Disabled') }
                @((New-FakeTask 'FoxitUpdateTask' '\Foxit Software\' 'Ready'))
            }
            $r = Get-Run
            Should -Invoke Disable-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'FoxitUpdateTask' -and $TaskPath -eq '\Foxit Software\' }
            Should -Invoke Stop-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'FoxitUpdateTask' -and $TaskPath -eq '\Foxit Software\' }
        }
    }

    Context 'finding 15: results are verified, not assumed' {
        It 'exits 3 and names the key when a registry value does not read back' {
            Mock Set-FactsUpdaterValues { $Path -notlike '*WOW6432Node*' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match ([regex]::Escape('WOW6432Node\Foxit Software\Foxit PDF Editor\Updater'))
            $r.Text | Should -Match '\[WARN\].*exit 3'
        }
        It 'exits 3 when a Foxit update service cannot be disabled' {
            Mock Set-Service { throw 'access denied' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'FoxitReaderUpdateService'
        }
        It 'exits 3 when a Foxit update service is still not disabled afterwards' {
            Mock Get-Service { @([pscustomobject]@{ Name = 'FoxitReaderUpdateService'; StartType = 'Automatic' }) }
            $r = Get-Run
            $r.Code | Should -Be 3
        }
        It 'exits 3 when a Foxit update task is still enabled afterwards' {
            Mock Get-ScheduledTask {
                if ($TaskName -eq 'FACTS_Universal_Enforcement') { return (New-FakeTask 'FACTS_Universal_Enforcement' '\' 'Ready') }
                @((New-FakeTask 'FoxitUpdateTask' '\' 'Ready'))
            }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'FoxitUpdateTask'
        }
        It 'exits 3 when the registered maintenance task cannot be read back' {
            Mock Get-ScheduledTask {
                if ($TaskName -eq 'FACTS_Universal_Enforcement') { return $null }
                @((New-FakeTask 'FoxitUpdateTask' '\' 'Disabled'))
            }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'FACTS_Universal_Enforcement'
        }
        It 'exits 3 when the old maintenance task cannot be removed after Foxit is gone' {
            Mock Get-FactsFoxitApps { [pscustomobject]@{ Ok = $true; Names = @() } }
            Mock Unregister-ScheduledTask { throw 'access denied' }
            $r = Get-Run
            $r.Code | Should -Be 3
        }
    }

    Context 'finding 15: the installed maintenance task survives a failed update' {
        It 'does not delete the old task before the new one is registered' {
            Mock Register-ScheduledTask { throw 'task store unavailable' }
            $r = Get-Run
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\].*exit 1'
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
        }
        It 'replaces the task with one Register call using -Force and no prior Unregister' {
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Register-ScheduledTask -Times 1 -Exactly -ParameterFilter { $Force }
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
        }
    }
}
