#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# DEBLOAT tests. Every Appx, registry, service, scheduled task and file call is mocked or confined
# to TestDrive, so nothing here touches the machine. The fixture is a synthetic Users folder with
# two profiles, each holding a synthetic Edge profile.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath 'DEBLOAT')

    # PowerShell 7 has no native Appx or ScheduledTasks cmdlets without compatibility loading, so
    # give Pester something to mock when a command is missing.
    foreach ($n in 'Get-AppxPackage', 'Remove-AppxPackage', 'Get-AppxProvisionedPackage',
                   'Remove-AppxProvisionedPackage', 'Get-WindowsOptionalFeature',
                   'Get-ScheduledTask', 'Disable-ScheduledTask') {
        if (-not (Get-Command $n -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$n" -Value {
                [CmdletBinding()] param($Name, $PackageName, $FeatureName, $TaskPath, [switch]$AllUsers, [switch]$Online)
            }
        }
    }

    function Get-Exit {
        # Runs Main and returns the code carried by Exit-KillerScript.
        try { Invoke-DebloatMain *>&1 | Out-Null } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { return [int]$Matches[1] }
            throw
        }
        -1
    }

    function Get-Run {
        # Runs Main and returns the exit code plus everything written to the output stream.
        $code = -1
        $script:RunLines = New-Object System.Collections.ArrayList
        try { Invoke-DebloatMain *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        [pscustomobject]@{ Code = $code; Text = ($script:RunLines -join "`n") }
    }

    function New-FixtureProfile {
        # A synthetic profile: NTUSER.DAT plus an Edge Default profile with user data.
        param([string]$Root, [string]$Name)
        $home1 = Join-Path $Root $Name
        $edge = Join-Path $home1 'AppData\Local\Microsoft\Edge\User Data\Default'
        New-Item -ItemType Directory -Path $edge -Force | Out-Null
        Set-Content -Path (Join-Path $home1 'NTUSER.DAT') -Value 'hive'
        Set-Content -Path (Join-Path $edge 'Bookmarks') -Value "bookmarks of $Name"
        Set-Content -Path (Join-Path $edge 'Web Data') -Value "autofill of $Name"
        Set-Content -Path (Join-Path $edge 'Preferences') -Value "prefs of $Name"
        $home1
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'DEBLOAT' {
    BeforeEach {
        $script:Mfr = 'Contoso'
        $script:Users = Join-Path $TestDrive 'Users'
        New-Item -ItemType Directory -Path $script:Users -Force | Out-Null
        $script:Alice = New-FixtureProfile $script:Users 'alice'
        $script:Bob   = New-FixtureProfile $script:Users 'bob'
        New-Item -ItemType Directory -Path (Join-Path $script:Users 'Default') -Force | Out-Null
        Set-Content -Path (Join-Path $script:Users 'Default\NTUSER.DAT') -Value 'template'

        $script:SavedProgramData = $env:ProgramData
        $script:SavedProgramFiles = $env:ProgramFiles
        $env:ProgramData  = Join-Path $TestDrive 'ProgramData'
        $env:ProgramFiles = Join-Path $TestDrive 'ProgramFiles'
        New-Item -ItemType Directory -Path $env:ProgramData, $env:ProgramFiles -Force | Out-Null

        # Console and identity
        Mock Write-Host {}
        Mock Clear-Host {}
        Mock Test-IsAdmin { $true }
        # System and hardware
        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_ComputerSystem') { [pscustomobject]@{ Manufacturer = $script:Mfr } }
        }
        Mock Get-WindowsOptionalFeature { $null }
        Mock Start-Process {}
        Mock Wait-Process {}
        Mock Stop-Process {}
        Mock powercfg {}
        Mock Invoke-ScExe {}
        # Services and tasks
        Mock Get-Service { @() }
        Mock Stop-Service {}
        Mock Set-Service {}
        Mock Get-ScheduledTask { @() }
        Mock Disable-ScheduledTask {}
        # Appx
        Mock Get-AppxPackage { @() }
        Mock Remove-AppxPackage {}
        Mock Get-AppxProvisionedPackage { @() }
        Mock Remove-AppxProvisionedPackage {}
        # Registry and hives
        Mock Get-ItemProperty { $null }
        Mock Test-Path { $true } -ParameterFilter { "$Path" -like 'HKLM:*' -or "$Path" -like 'Registry::*' }
        Mock New-Item {} -ParameterFilter { "$Path" -like 'HKLM:*' -or "$Path" -like 'Registry::*' }
        Mock Set-ItemProperty {} -ParameterFilter { "$Path" -like 'HKLM:*' -or "$Path" -like 'Registry::*' }
        Mock Remove-Item {} -ParameterFilter { "$Path" -like 'HKLM:*' -or "$Path" -like 'Registry::*' }
        # No hive is mounted unless a test says so.
        Mock Test-Path { $false } -ParameterFilter { "$Path" -match '^Registry::HKEY_USERS\\(TempHive_\w+|DefaultUser)$' }
        Mock Start-Sleep {}
        Mock Invoke-RegExe { 0 }
        Mock Get-LoadedUserSidMap { @{} }
        Mock Get-UserFolder { Get-ChildItem -LiteralPath $script:Users -Directory }
        Mock Get-UserProfileFolder { Get-ChildItem -LiteralPath $script:Users -Directory | Where-Object { $_.Name -ne 'Default' } }
        Mock Get-DefaultUserHivePath { Join-Path $script:Users 'Default\NTUSER.DAT' }
        Mock Get-InstalledMsiPackage { @() }
    }

    AfterEach {
        $env:ProgramData = $script:SavedProgramData
        $env:ProgramFiles = $script:SavedProgramFiles
    }

    Context 'entry-point convention' {
        It 'defines Invoke-DebloatMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-DebloatMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Set-ItemProperty -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'exits 0 with an honest final line on a healthy run' {
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\] DEBLOAT complete \(exit 0\)'
        }
        It 'exits 3010 and says a reboot is required when DISM removes Recall and asks for one' {
            Mock Get-WindowsOptionalFeature { [pscustomobject]@{ State = 'Enabled' } }
            Mock Start-Process { [pscustomobject]@{ Id = 4242; HasExited = $true; ExitCode = 3010 } }
            $r = Get-Run
            $r.Code | Should -Be 3010
            $r.Text | Should -Match 'Reboot required.*\(exit 3010\)'
        }
        It 'exits 3 and says partial when DISM times out' {
            Mock Get-WindowsOptionalFeature { [pscustomobject]@{ State = 'Enabled' } }
            Mock Start-Process { [pscustomobject]@{ Id = 4242; HasExited = $false; ExitCode = $null } }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\] DEBLOAT partial.*\(exit 3\)'
        }
    }

    Context 'finding 08: HP cleanup leaves Edge user data alone' {
        BeforeEach { $script:Mfr = 'HP' }
        It 'keeps Bookmarks, Web Data and Preferences in every profile' {
            Get-Exit | Should -Be 0
            foreach ($u in $script:Alice, $script:Bob) {
                $edge = Join-Path $u 'AppData\Local\Microsoft\Edge\User Data\Default'
                Get-Content -LiteralPath (Join-Path $edge 'Bookmarks') | Should -Be "bookmarks of $(Split-Path -Leaf $u)"
                Get-Content -LiteralPath (Join-Path $edge 'Web Data') | Should -Be "autofill of $(Split-Path -Leaf $u)"
                Test-Path -LiteralPath (Join-Path $edge 'Preferences') | Should -BeTrue
            }
        }
        It 'still runs the HP per-user cleanup for both profiles' {
            Get-Exit | Should -Be 0
            Should -Invoke Invoke-RegExe -Times 2 -Exactly -ParameterFilter { $Arguments[0] -eq 'load' -and $Arguments[1] -eq 'HKU\TempHive_alice' }
            Should -Invoke Invoke-RegExe -Times 2 -Exactly -ParameterFilter { $Arguments[0] -eq 'load' -and $Arguments[1] -eq 'HKU\TempHive_bob' }
        }
    }

    Context 'finding 67: McAfee data and the Killer NIC driver folder' {
        BeforeEach {
            $script:Services = @()
            $script:McAfeeDir = Join-Path $env:ProgramData 'McAfee'
            $script:NortonDir = Join-Path $env:ProgramFiles 'Norton Security'
            New-Item -ItemType Directory -Path $script:McAfeeDir, $script:NortonDir -Force | Out-Null
            Set-Content -Path (Join-Path $script:McAfeeDir 'agent.db') -Value 'managed AV data'
            Mock Get-Service {
                if ($Name) { $script:Services | Where-Object { $_.Name -eq $Name } } else { $script:Services }
            }
            # The Rivet folder is a real system path: pretend it exists and record any delete attempt.
            Mock Test-Path { $true } -ParameterFilter { "$Path" -like '*RivetNetworks*' }
            Mock Remove-Item {} -ParameterFilter { "$Path" -like '*RivetNetworks*' }
        }

        It 'deletes ProgramData\McAfee when no McAfee or Trellix product or service is present' {
            Get-Exit | Should -Be 0
            Test-Path -LiteralPath $script:McAfeeDir | Should -BeFalse
        }
        It 'keeps ProgramData\McAfee when a McAfee MSI product is installed (any manufacturer)' {
            Mock Get-InstalledMsiPackage { [pscustomobject]@{ DisplayName = 'McAfee Agent'; UninstallString = 'msiexec /x' } } -ParameterFilter { $NamePattern -eq '*McAfee*' }
            Get-Exit | Should -Be 0
            Test-Path -LiteralPath (Join-Path $script:McAfeeDir 'agent.db') | Should -BeTrue
        }
        It 'keeps ProgramData\McAfee when a Trellix MSI product is installed' {
            Mock Get-InstalledMsiPackage { [pscustomobject]@{ DisplayName = 'Trellix Endpoint Security'; UninstallString = 'msiexec /x' } } -ParameterFilter { $NamePattern -eq '*Trellix*' }
            Get-Exit | Should -Be 0
            Test-Path -LiteralPath (Join-Path $script:McAfeeDir 'agent.db') | Should -BeTrue
        }
        It 'keeps ProgramData\McAfee when a McAfee service is present' {
            $script:Services = @([pscustomobject]@{ Name = 'masvc'; DisplayName = 'McAfee Agent Service' })
            Get-Exit | Should -Be 0
            Test-Path -LiteralPath (Join-Path $script:McAfeeDir 'agent.db') | Should -BeTrue
        }
        It 'keeps ProgramData\McAfee when a Trellix service is present' {
            $script:Services = @([pscustomobject]@{ Name = 'xagt'; DisplayName = 'Trellix Endpoint Agent' })
            Get-Exit | Should -Be 0
            Test-Path -LiteralPath (Join-Path $script:McAfeeDir 'agent.db') | Should -BeTrue
        }
        It 'still removes the Norton folder' {
            Get-Exit | Should -Be 0
            Test-Path -LiteralPath $script:NortonDir | Should -BeFalse
        }
        It 'never deletes System32\drivers\RivetNetworks on a Dell' {
            $script:Mfr = 'Dell Inc.'
            Get-Exit | Should -Be 0
            Should -Invoke Remove-Item -Times 0 -Exactly -ParameterFilter { "$Path" -like '*RivetNetworks*' }
        }
        It 'never deletes System32\drivers\RivetNetworks on any other manufacturer' {
            Get-Exit | Should -Be 0
            Should -Invoke Remove-Item -Times 0 -Exactly -ParameterFilter { "$Path" -like '*RivetNetworks*' }
        }
    }

    Context 'finding 09: temporary hives are always released, and only ours' {
        BeforeEach {
            $script:Sid = 'S-1-5-21-1111111111-2222222222-3333333333-1001'
        }

        It 'unloads every hive it loaded on a healthy run, exactly once each' {
            Get-Exit | Should -Be 0
            foreach ($h in 'DefaultUser', 'TempHive_alice', 'TempHive_bob') {
                Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'load' -and $Arguments[1] -eq "HKU\$h" }
                Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq "HKU\$h" }
            }
        }

        It 'unloads a profile hive when an operation throws, keeps going, and exits 3' {
            Mock Set-UserRegValue { throw 'simulated terminating error' } -ParameterFilter { "$Path" -like '*TempHive_alice*' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\] DEBLOAT partial.*\(exit 3\)'
            Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_alice' }
            Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'load' -and $Arguments[1] -eq 'HKU\TempHive_bob' }
            Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_bob' }
        }

        It 'unloads the Default User hive when an operation throws against it' {
            Mock Set-UserRegValue { throw 'simulated terminating error' } -ParameterFilter { "$Path" -like '*DefaultUser*' }
            Get-Exit | Should -Be 3
            Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\DefaultUser' }
            Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_bob' }
        }

        It 'never loads, writes to or unloads a hive somebody else already mounted' {
            Mock Test-Path { $true } -ParameterFilter { "$Path" -eq 'Registry::HKEY_USERS\TempHive_alice' }
            Get-Exit | Should -Be 3
            Should -Invoke Invoke-RegExe -Times 0 -Exactly -ParameterFilter { $Arguments -contains 'HKU\TempHive_alice' }
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter { "$Path" -like '*TempHive_alice*' }
            Should -Invoke Invoke-RegExe -Times 1 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_bob' }
        }

        It 'never unloads a hive that failed to load, and reports the skipped profile (exit 3)' {
            Mock Invoke-RegExe { 1 } -ParameterFilter { $Arguments[0] -eq 'load' -and $Arguments[1] -eq 'HKU\TempHive_alice' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'alice'
            Should -Invoke Invoke-RegExe -Times 0 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_alice' }
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter { "$Path" -like '*TempHive_alice*' }
        }

        It 'does not write into HKU\DefaultUser when the template hive failed to load' {
            Mock Invoke-RegExe { 1 } -ParameterFilter { $Arguments[0] -eq 'load' -and $Arguments[1] -eq 'HKU\DefaultUser' }
            Get-Exit | Should -Be 3
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter { "$Path" -like '*HKEY_USERS\DefaultUser*' }
            Should -Invoke New-Item -Times 0 -Exactly -ParameterFilter { "$Path" -like '*HKEY_USERS\DefaultUser*' }
            Should -Invoke Invoke-RegExe -Times 0 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\DefaultUser' }
        }

        It 'retries a refused unload, then reports the stuck hive instead of claiming success' {
            Mock Invoke-RegExe { 1 } -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_alice' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'TempHive_alice'
            Should -Invoke Invoke-RegExe -Times 3 -Exactly -ParameterFilter { $Arguments[0] -eq 'unload' -and $Arguments[1] -eq 'HKU\TempHive_alice' }
        }

        It 'writes through the SID for a signed-in user and never loads or unloads that hive' {
            Mock Get-LoadedUserSidMap { @{ alice = $script:Sid } }
            Get-Exit | Should -Be 0
            Should -Invoke Invoke-RegExe -Times 0 -Exactly -ParameterFilter { $Arguments -contains 'HKU\TempHive_alice' }
            Should -Invoke Set-ItemProperty -ParameterFilter { "$Path" -like "*HKEY_USERS\$($script:Sid)\*" }
        }
    }

    Context 'finding 69: not elevated' {
        It 'exits 2 with an honest status line and touches nothing' {
            Mock Test-IsAdmin { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\] DEBLOAT not elevated, nothing was changed \(exit 2\)'
            Should -Invoke Set-ItemProperty -Times 0 -Exactly
            Should -Invoke New-Item -Times 0 -Exactly
            Should -Invoke Invoke-RegExe -Times 0 -Exactly
            Should -Invoke Start-Process -Times 0 -Exactly
            Should -Invoke Remove-AppxProvisionedPackage -Times 0 -Exactly
            Should -Invoke powercfg -Times 0 -Exactly
        }
        It 'still runs when elevated' {
            Get-Exit | Should -Be 0
        }
    }

    Context 'issue 1: telemetry policy level' {
        It 'sets both DataCollection AllowTelemetry values to 0, never 1' {
            Get-Exit | Should -Be 0
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter {
                "$Path" -eq 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection' -and $Name -eq 'AllowTelemetry' -and $Value -eq 0
            }
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter {
                "$Path" -eq 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection' -and $Name -eq 'AllowTelemetry' -and $Value -eq 0
            }
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter {
                $Name -eq 'AllowTelemetry' -and $Value -ne 0
            }
        }
    }
}
