#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# PEEK tests. PEEK is a read-only monitor, but every system call here is still mocked or replaced
# by a synthetic fixture: events, processes, Panther logs, compatibility reports and registry state.

Describe 'PEEK entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'PEEK.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $script:Source = [IO.File]::ReadAllText($path)
    }

    It 'defines Invoke-PEEKMain and a local Exit-KillerScript' {
        $names = @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-PEEKMain'
        $names | Should -Contain 'Exit-KillerScript'
    }

    It 'ends with the test-mode guard and has no bare exit statement outside Exit-KillerScript' {
        $Source | Should -Match '(?m)^if \(-not \$global:KillerScriptsTestMode\) \{ Invoke-PEEKMain \}\s*$'
        $exits = @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExitStatementAst] }, $true))
        $exits.Count | Should -Be 1
    }
}

Describe 'PEEK main' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'PEEK')

        function Get-PeekRun {
            # Runs Main and returns the exit code carried by Exit-KillerScript plus the output stream text.
            $code = -1
            $script:RunLines = New-Object System.Collections.ArrayList
            try { Invoke-PEEKMain *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($script:RunLines -join "`n") }
        }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $Once = [switch]$true
        $RefreshSeconds = 1
        Mock Test-PEEKElevated { $true }
        Mock Get-PEEKBitLockerState { 'Unavailable' }
        Mock Show-PEEKSnapshot { }
        Mock Write-Footer { }
    }

    It 'loads without running when dot-sourced in test mode' {
        Get-Command Invoke-PEEKMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'exits 2 with a FAIL line and shows nothing when not elevated' {
        Mock Test-PEEKElevated { $false }
        $r = Get-PeekRun
        $r.Code | Should -Be 2
        $r.Text | Should -Match '\[FAIL\].*elevation'
        Should -Invoke Show-PEEKSnapshot -Times 0 -Exactly
    }

    It 'exits 0 with an OK line after a one-shot snapshot' {
        $r = Get-PeekRun
        $r.Code | Should -Be 0
        $r.Text | Should -Match '\[OK\].*exit 0'
        Should -Invoke Show-PEEKSnapshot -Times 1 -Exactly
    }

    It 'exits 1 with a FAIL line when the snapshot throws, never claiming success' {
        Mock Show-PEEKSnapshot { throw 'boom' }
        $r = Get-PeekRun
        $r.Code | Should -Be 1
        $r.Text | Should -Match '\[FAIL\].*boom.*exit 1'
        $r.Text | Should -Not -Match '\[OK\]'
    }

    It 'does not throw when there is no console to clear' {
        Mock Clear-Host { throw 'The handle is invalid.' }
        { Clear-PEEKScreen } | Should -Not -Throw
    }
}

Describe 'PEEK finding 20: servicing state needs corroborating identity' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'PEEK')

        function New-Proc { param([string]$Name, [string]$Path = '') [pscustomobject]@{ Name = $Name; Id = 100; CPU = $null; RAMMB = 1; Start = $null; Priority = 2; Path = $Path } }
        function New-OS { param([string]$Version = '24H2') [pscustomobject]@{ DisplayVersion = $Version; Build = '26100'; UBR = '1'; Caption = 'Windows 11 Pro'; LastBoot = (Get-Date).AddDays(-3); Uptime = [timespan]::FromDays(3) } }
        function New-Setup { param([bool]$Active = $false) [pscustomobject]@{ ActFile = $null; ErrFile = $null; Active = $Active; ActAgeSeconds = $null; Mode = 'None'; CommandLine = ''; FeatureCode = ''; FeatureText = '' } }
        function New-Compat { [pscustomobject]@{ File = $null; LastWriteTime = [datetime]::MinValue; BlockCount = 0; Summary = 'None'; HasReport = $false; ParseFailed = $false } }
        function New-Reboot { [pscustomobject]@{ WindowsUpdate = $false; CBS = $false; MoSetup = $false; FileRename = $false; UpdateExe = $false; PendingXml = $false; Any = $false; Sources = @() } }
        function New-Ctx {
            # Replay of a feature-update context: event 43 for a 25H2 feature update with no terminal event.
            param([bool]$Active = $true)
            $start = [pscustomobject]@{ TimeCreated = (Get-Date).AddMinutes(-20); Id = 43; Status = 'TRY'; Label = 'Installing'; Title = 'Feature update to Windows 11, version 25H2'; ErrorCode = '' }
            [pscustomobject]@{ Current = $start; ActiveStart = $(if ($Active) { $start } else { $null }); IsFeature = $true; FeatureTarget = '25H2' }
        }
        function New-IdleCtx { [pscustomobject]@{ Current = $null; ActiveStart = $null; IsFeature = $false; FeatureTarget = '' } }
    }
    AfterAll { Disable-KillerTestMode }

    It 'does not call an unrelated setup.exe a servicing process on its name alone' {
        $r = Get-PEEKSystemState -OS (New-OS) -UpdateContext (New-IdleCtx) -Processes @(New-Proc 'setup' 'C:\Vendor\Acme\setup.exe') -Setup (New-Setup) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.InstallProcess | Should -BeFalse
        $r.State | Should -Be 'IDLE'
        $r.RebootAction | Should -Be 'NO REBOOT REQUIRED'
    }

    It 'does not call a setup process with an unreadable path servicing without other evidence' {
        $r = Get-PEEKSystemState -OS (New-OS) -UpdateContext (New-IdleCtx) -Processes @(New-Proc 'setup' '') -Setup (New-Setup) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.InstallProcess | Should -BeFalse
    }

    It 'counts setup.exe running from the $WINDOWS.~BT sources folder' {
        $r = Get-PEEKSystemState -OS (New-OS) -UpdateContext (New-IdleCtx) -Processes @(New-Proc 'setup' 'C:\$WINDOWS.~BT\Sources\setup.exe') -Setup (New-Setup) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.InstallProcess | Should -BeTrue
        $r.State | Should -Be 'WINDOWS SERVICING ACTIVE'
        $r.RebootAction | Should -Match '^DO NOT REBOOT'
    }

    It 'counts a setup process when the Panther setup log is being written' {
        $r = Get-PEEKSystemState -OS (New-OS) -UpdateContext (New-IdleCtx) -Processes @(New-Proc 'setup' 'E:\setup.exe') -Setup (New-Setup -Active $true) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.InstallProcess | Should -BeTrue
    }

    It 'lets live servicing win over a matching installed release, so state and reboot advice agree' {
        $procs = @(New-Proc 'setuphost' 'C:\$WINDOWS.~BT\Sources\SetupHost.exe')
        $r = Get-PEEKSystemState -OS (New-OS '25H2') -UpdateContext (New-Ctx) -Processes $procs -Setup (New-Setup -Active $true) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.State | Should -Be 'FEATURE UPGRADE ACTIVE'
        $r.RebootAction | Should -Match '^DO NOT REBOOT'
        $r.State | Should -Not -Match 'COMPLETE'
    }

    It 'reports the feature update complete, with no do-not-reboot advice, when the target is installed and nothing is live' {
        $r = Get-PEEKSystemState -OS (New-OS '25H2') -UpdateContext (New-Ctx) -Processes @() -Setup (New-Setup) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.State | Should -Be 'FEATURE UPDATE COMPLETE - 25H2 INSTALLED'
        $r.RebootAction | Should -Be 'NO REBOOT REQUIRED'
    }

    It 'keeps do-not-reboot for an install that started but has not reached its target release' {
        $r = Get-PEEKSystemState -OS (New-OS '24H2') -UpdateContext (New-Ctx) -Processes @() -Setup (New-Setup) -Compat (New-Compat) -Reboot (New-Reboot)
        $r.State | Should -Be 'INSTALL STARTED - NO LIVE ACTIVITY DETECTED'
        $r.RebootAction | Should -Match '^DO NOT REBOOT'
    }

    It 'records the executable path in the process snapshot' {
        Mock Get-Process { [pscustomobject]@{ ProcessName = 'setup'; Id = 4242; TotalProcessorTime = [timespan]::FromSeconds(1); WorkingSet64 = 10MB; StartTime = (Get-Date); Path = 'C:\$WINDOWS.~BT\Sources\setup.exe' } }
        $rows = @(Get-PEEKProcessSnapshot)
        $rows[0].Path | Should -Be 'C:\$WINDOWS.~BT\Sources\setup.exe'
    }
}
