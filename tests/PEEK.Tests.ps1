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

Describe 'PEEK finding 34: an unreadable compatibility report is not a compatibility block' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'PEEK')
        $script:FixtureDir = Join-Path $PSScriptRoot 'fixtures\peek'
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:CompatCacheKey = $null
        $script:CompatCacheValue = $null
        $script:CompatRetryAfter = [datetime]::MinValue
        $script:Report = Join-Path $TestDrive 'CompatData_1.xml'
        [IO.File]::Copy((Join-Path $FixtureDir 'CompatData-truncated.xml'), $Report, $true)
        (Get-Item -LiteralPath $Report).LastWriteTime = Get-Date
        Mock Get-ChildItem { Get-Item -LiteralPath $script:Report } -ParameterFilter { $Path -like '*CompatData*.xml' }
    }

    It 'reports a parse failure on its own, with no block count' {
        $c = Get-PEEKCompatState
        $c.HasReport | Should -BeTrue
        $c.ParseFailed | Should -BeTrue
        $c.BlockCount | Should -Be 0
        $c.Summary | Should -Match 'could not be parsed'
    }

    It 'does not let a fresh unreadable report produce COMPATIBILITY BLOCK DETECTED' {
        $c = Get-PEEKCompatState
        $ctx = [pscustomobject]@{ Current = $null; ActiveStart = $null; IsFeature = $false; FeatureTarget = '' }
        $os = [pscustomobject]@{ DisplayVersion = '24H2'; Build = '26100'; UBR = '1'; Caption = 'Windows 11 Pro'; LastBoot = (Get-Date).AddDays(-1); Uptime = [timespan]::FromDays(1) }
        $setup = [pscustomobject]@{ ActFile = $null; ErrFile = $null; Active = $false; ActAgeSeconds = $null; Mode = 'None'; CommandLine = ''; FeatureCode = ''; FeatureText = '' }
        $reboot = [pscustomobject]@{ WindowsUpdate = $false; CBS = $false; MoSetup = $false; FileRename = $false; UpdateExe = $false; PendingXml = $false; Any = $false; Sources = @() }
        $r = Get-PEEKSystemState -OS $os -UpdateContext $ctx -Processes @() -Setup $setup -Compat $c -Reboot $reboot
        $r.State | Should -Not -Match 'COMPATIBILITY BLOCK'
        $r.CompatibilityFresh | Should -BeFalse
    }

    It 'picks up the finished report once the producer has written it' {
        $null = Get-PEEKCompatState
        [IO.File]::Copy((Join-Path $FixtureDir 'CompatData-block.xml'), $Report, $true)
        (Get-Item -LiteralPath $Report).LastWriteTime = (Get-Date).AddSeconds(5)
        $script:CompatRetryAfter = [datetime]::MinValue
        $c = Get-PEEKCompatState
        $c.ParseFailed | Should -BeFalse
        $c.BlockCount | Should -Be 2
        $c.Summary | Should -Match 'HARD: Contoso Legacy VPN'
        $c.Summary | Should -Match 'DRIVER: oem42.inf'
    }

    It 'retries an unchanged file after the retry window, because the read itself may have failed' {
        $null = Get-PEEKCompatState
        [IO.File]::Copy((Join-Path $FixtureDir 'CompatData-block.xml'), $Report, $true)
        $stamp = (Get-Item -LiteralPath $Report).LastWriteTime
        # Same name, time and length as the cached failure, as when a read hit a sharing violation.
        $script:CompatCacheKey = "$Report|$($stamp.Ticks)|$((Get-Item -LiteralPath $Report).Length)"
        $script:CompatCacheValue = [pscustomobject]@{ File = (Get-Item -LiteralPath $Report); LastWriteTime = $stamp; BlockCount = 0; Summary = 'Compatibility report could not be parsed (retrying)'; HasReport = $true; ParseFailed = $true }
        $script:CompatRetryAfter = [datetime]::MinValue
        (Get-PEEKCompatState).BlockCount | Should -Be 2
    }

    It 'does not reread an unchanged unreadable report inside the retry window' {
        $null = Get-PEEKCompatState
        Mock Get-Content { throw 'should not be read again' }
        { Get-PEEKCompatState } | Should -Not -Throw
        (Get-PEEKCompatState).ParseFailed | Should -BeTrue
        Should -Invoke Get-Content -Times 0 -Exactly
    }

    It 'still reports a genuine hard block from a well-formed report' {
        [IO.File]::Copy((Join-Path $FixtureDir 'CompatData-block.xml'), $Report, $true)
        $c = Get-PEEKCompatState
        $c.ParseFailed | Should -BeFalse
        $c.BlockCount | Should -Be 2
    }

    It 'reports a well-formed report without blocks as clean' {
        [IO.File]::Copy((Join-Path $FixtureDir 'CompatData-clean.xml'), $Report, $true)
        $c = Get-PEEKCompatState
        $c.ParseFailed | Should -BeFalse
        $c.BlockCount | Should -Be 0
        $c.Summary | Should -Be 'None detected'
    }
}

Describe 'PEEK finding 73: every reboot source that counts is displayed' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'PEEK')

        function Get-FlagText {
            param($Reboot)
            (Write-RebootFlags -Reboot $Reboot 6>&1 | ForEach-Object { [string]$_ }) -join ''
        }
    }
    AfterAll { Disable-KillerTestMode }

    It 'detects UpdateExeVolatile as a reboot source' {
        Mock Test-Path { $false }
        Mock Get-ItemProperty { [pscustomobject]@{ UpdateExeVolatile = 1 } } -ParameterFilter { $Path -like '*\Microsoft\Updates' }
        $r = Get-PEEKRebootState
        $r.UpdateExe | Should -BeTrue
        $r.Any | Should -BeTrue
        $r.Sources | Should -Contain 'UpdateExe'
    }

    It 'shows UpdateExe=YES when it is the only reason a reboot is pending' {
        $reboot = [pscustomobject]@{ WindowsUpdate = $false; CBS = $false; MoSetup = $false; FileRename = $false; UpdateExe = $true; PendingXml = $false; Any = $true; Sources = @('UpdateExe') }
        $text = Get-FlagText $reboot
        $text | Should -Match 'UpdateExe=YES'
        $text | Should -Match 'WU=no'
    }

    It 'shows UpdateExe=no when it is clear' {
        $reboot = [pscustomobject]@{ WindowsUpdate = $false; CBS = $false; MoSetup = $false; FileRename = $false; UpdateExe = $false; PendingXml = $false; Any = $false; Sources = @() }
        Get-FlagText $reboot | Should -Match 'UpdateExe=no'
    }
}

Describe 'PEEK finding 44: the deep Panther scan is cached by source identity' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'PEEK')
        $script:FixtureDir = Join-Path $PSScriptRoot 'fixtures\peek'

    function Grow-Log {
        param([int]$SecondsAhead = 1)
        [IO.File]::AppendAllText($script:Act, "2026-11-18 09:01:00, Info                  SP     more synthetic output`r`n")
        [IO.File]::SetLastWriteTime($script:Act, (Get-Date).AddSeconds($SecondsAhead))
    }    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:SetupCommandLinePath = ""
        $script:SetupCommandLine     = ""
        $script:SetupDeepScanTime    = [datetime]::MinValue
        $script:SetupDeepScanPath    = $null
        $script:SetupDeepScanKey     = $null

        $script:Act = Join-Path $TestDrive 'setupact.log'
        [IO.File]::Copy((Join-Path $FixtureDir 'setupact-nocmdline.log'), $script:Act, $true)
        $script:ActLines = [IO.File]::ReadAllLines((Join-Path $FixtureDir 'setupact-nocmdline.log'))

        # Only the Windows\Panther setupact.log exists. Everything else on the machine is absent.
        Mock Test-Path { $true } -ParameterFilter { $LiteralPath -like '*Windows\Panther\setupact.log' -and $LiteralPath -notlike '*BT*' }
        Mock Test-Path { $false } -ParameterFilter { -not ($LiteralPath -like '*Windows\Panther\setupact.log' -and $LiteralPath -notlike '*BT*') }
        Mock Get-Item { New-Object System.IO.FileInfo $script:Act }
        Mock Get-Content { $script:ActLines }
    }

    function Grow-Log {
        param([int]$SecondsAhead = 1)
        [IO.File]::AppendAllText($script:Act, "2026-11-18 09:01:00, Info                  SP     more synthetic output`r`n")
        (Get-Item -LiteralPath $script:Act).LastWriteTime = (Get-Date).AddSeconds($SecondsAhead)
    }

    It 'reads the 50,000 line tail once, not every refresh, while a negative log is unchanged' {
        1..4 | ForEach-Object { $null = Get-PEEKSetupState }
        Should -Invoke Get-Content -Times 1 -Exactly -ParameterFilter { $Tail -eq 50000 }
    }

    It 'still reports no command line and no mode for the negative log' {
        $s = Get-PEEKSetupState
        $s.CommandLine | Should -Be ''
        $s.Mode | Should -Be 'None'
    }

    It 'does not rescan a growing negative log more often than every 15 seconds' {
        $null = Get-PEEKSetupState
        Grow-Log
        $null = Get-PEEKSetupState
        Grow-Log
        $null = Get-PEEKSetupState
        Should -Invoke Get-Content -Times 1 -Exactly -ParameterFilter { $Tail -eq 50000 }
    }

    It 'rescans a changed log once the 15 second interval has passed' {
        $null = Get-PEEKSetupState
        Grow-Log
        $script:SetupDeepScanTime = (Get-Date).AddSeconds(-20)
        $null = Get-PEEKSetupState
        Should -Invoke Get-Content -Times 2 -Exactly -ParameterFilter { $Tail -eq 50000 }
    }

    It 'rescans an unchanged log after the bounded refresh window' {
        $null = Get-PEEKSetupState
        $script:SetupDeepScanTime = (Get-Date).AddSeconds(-400)
        $null = Get-PEEKSetupState
        Should -Invoke Get-Content -Times 2 -Exactly -ParameterFilter { $Tail -eq 50000 }
    }

    It 'finds the command line once the log gains it, after the interval' {
        $null = Get-PEEKSetupState
        $script:ActLines = [IO.File]::ReadAllLines((Join-Path $FixtureDir 'setupact-cmdline.log'))
        Grow-Log
        $script:SetupDeepScanTime = (Get-Date).AddSeconds(-20)
        $s = Get-PEEKSetupState
        $s.CommandLine | Should -Match '/auto upgrade'
        $s.Mode | Should -Be 'In-Place Upgrade'
    }

    It 'does not reread a log whose command line is already known and unchanged' {
        $script:ActLines = [IO.File]::ReadAllLines((Join-Path $FixtureDir 'setupact-cmdline.log'))
        1..3 | ForEach-Object { $s = Get-PEEKSetupState }
        $s.Mode | Should -Be 'In-Place Upgrade'
        Should -Invoke Get-Content -Times 1 -Exactly -ParameterFilter { $Tail -eq 50000 }
    }
}
