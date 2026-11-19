#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# DEFEND tests. Defender, firewall, registry, services and the scan process are all mocked,
# so nothing here touches the machine.

BeforeAll {
    # PowerShell 7 does not autoload the Defender, NetSecurity or DISM modules. Stub any command that
    # is missing so it can be mocked. Stubs are removed in AfterAll.
    $script:Stubbed = @()
    foreach ($n in 'Get-MpPreference', 'Set-MpPreference', 'Add-MpPreference', 'Remove-MpPreference', 'Get-MpComputerStatus',
                   'Get-MpThreat', 'Get-MpThreatDetection', 'Update-MpSignature', 'Start-MpScan', 'Get-WindowsOptionalFeature',
                   'Confirm-SecureBootUEFI', 'Get-Tpm', 'Get-NetFirewallProfile', 'Set-NetFirewallProfile',
                   'Get-NetConnectionProfile', 'Set-NetConnectionProfile') {
        if (-not (Get-Command $n -ErrorAction SilentlyContinue)) {
            Set-Item -Path "function:global:$n" -Value { }
            $script:Stubbed += $n
        }
    }
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath 'DEFEND')

    function Get-Run {
        # Runs Main and returns the exit code plus everything written to the output stream.
        $code = -1
        $script:RunLines = New-Object System.Collections.ArrayList
        try { Invoke-DefendMain *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        [pscustomobject]@{ Code = $code; Text = ($script:RunLines -join "`n"); Lines = @($script:RunLines) }
    }
    function Get-Exit { (Get-Run).Code }
}

AfterAll {
    Disable-KillerTestMode
    foreach ($n in $script:Stubbed) { Remove-Item -Path "function:global:$n" -ErrorAction SilentlyContinue }
}

Describe 'DEFEND' {
    BeforeEach {
        $script:TestPrefs = [pscustomobject]@{
            DisableRealtimeMonitoring    = $false
            PUAProtection                = 1
            MAPSReporting                = 2
            EnableNetworkProtection      = 1
            EnableControlledFolderAccess = 1
            CloudBlockLevel              = 2
            ExclusionPath                = @()
        }
        $script:TestStatus = [pscustomobject]@{ IsTamperProtected = $true; RealTimeProtectionEnabled = $true; AntivirusSignatureVersion = '1.2.3.4' }
        # Console and identity boundaries
        Mock Test-IsAdmin { $true }
        Mock Get-DefendPSMajor { 5 }
        Mock Test-DefendInteractive { $false }
        Mock Clear-DefendScreen {}
        Mock Import-Module {}
        Mock Start-Sleep {}
        # Hardware and registry boundaries
        Mock Get-CimInstance {
            switch ($ClassName) {
                'Win32_Processor'      { [pscustomobject]@{ VirtualizationFirmwareEnabled = $true } }
                'Win32_ComputerSystem' { [pscustomobject]@{ PartOfDomain = $false } }
                'MSFT_MpComputerStatus' { $script:TestStatus }
            }
        }
        Mock Get-WindowsOptionalFeature { [pscustomobject]@{ State = 'Enabled' } }
        Mock Confirm-SecureBootUEFI { $true }
        Mock Get-Tpm { [pscustomobject]@{ TpmPresent = $true; ManufacturerVersion = '2.0' } }
        Mock Get-DefendRegValue { 1 }
        Mock Test-DefendRegKey { $false }
        Mock Get-DefendHvciBlockers { @() }
        # Defender boundaries
        Mock Get-MpPreference { $script:TestPrefs }
        Mock Set-MpPreference {}
        Mock Add-MpPreference {}
        Mock Get-MpComputerStatus { $script:TestStatus }
        Mock Get-MpThreat { @() }
        Mock Get-MpThreatDetection { @() }
        Mock Update-MpSignature {}
        Mock Get-WinEvent { @() }
        Mock Invoke-DefendWinPS { [pscustomobject]@{ ExitCode = 0; Output = '' } }
        # Firewall boundaries
        Mock Set-NetFirewallProfile {}
        Mock Get-NetFirewallProfile { 'Domain', 'Private', 'Public' | ForEach-Object { [pscustomobject]@{ Name = $_; Enabled = 'True' } } }
        Mock Get-NetConnectionProfile { [pscustomobject]@{ Name = 'Net'; NetworkCategory = 'Public' } }
        Mock Set-NetConnectionProfile {}
    }

    Context 'entry-point convention' {
        It 'defines Invoke-DefendMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-DefendMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Set-MpPreference -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'exits 0 on a healthy run and the final line says so' {
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Lines[-1] | Should -Match '\[PASS\] DEFEND.*\(exit 0\)'
        }
    }

    Context 'finding 69: not elevated' {
        It 'exits 2 with a FAIL line and touches nothing when not elevated' {
            Mock Test-IsAdmin { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\] DEFEND.*\(exit 2\)'
            Should -Invoke Set-MpPreference -Times 0 -Exactly
            Should -Invoke Set-NetFirewallProfile -Times 0 -Exactly
            Should -Invoke Update-MpSignature -Times 0 -Exactly
        }
    }
}

Describe 'DEFEND console guard' {
    BeforeEach {
        Mock Get-DefendConsoleValue { throw 'The handle is invalid.' }
        Mock Set-DefendConsolePosition { throw 'The handle is invalid.' }
        Mock Clear-Host { throw 'The handle is invalid.' }
        Mock Import-Module {}
        Mock Test-IsAdmin { $true }
        Mock Get-DefendPSMajor { 5 }
        Mock Test-DefendInteractive { $false }
        Mock Get-CimInstance {
            switch ($ClassName) {
                'Win32_Processor'      { [pscustomobject]@{ VirtualizationFirmwareEnabled = $true } }
                'Win32_ComputerSystem' { [pscustomobject]@{ PartOfDomain = $false } }
                'MSFT_MpComputerStatus' { [pscustomobject]@{ IsTamperProtected = $true; RealTimeProtectionEnabled = $true } }
            }
        }
        Mock Get-WindowsOptionalFeature { [pscustomobject]@{ State = 'Enabled' } }
        Mock Confirm-SecureBootUEFI { $true }
        Mock Get-Tpm { [pscustomobject]@{ TpmPresent = $true; ManufacturerVersion = '2.0' } }
        Mock Get-DefendRegValue { 1 }
        Mock Test-DefendRegKey { $false }
        Mock Get-DefendHvciBlockers { @() }
        Mock Get-MpPreference { [pscustomobject]@{ DisableRealtimeMonitoring = $false; PUAProtection = 1; MAPSReporting = 2; EnableNetworkProtection = 1; EnableControlledFolderAccess = 1; CloudBlockLevel = 2; ExclusionPath = @() } }
        Mock Set-MpPreference {}
        Mock Get-MpComputerStatus { [pscustomobject]@{ AntivirusSignatureVersion = '1.2.3.4' } }
        Mock Get-MpThreat { @() }
        Mock Get-MpThreatDetection { @() }
        Mock Update-MpSignature {}
        Mock Get-WinEvent { @() }
        Mock Set-NetFirewallProfile {}
        Mock Get-NetFirewallProfile { 'Domain', 'Private', 'Public' | ForEach-Object { [pscustomobject]@{ Name = $_; Enabled = 'True' } } }
        Mock Get-NetConnectionProfile { @() }
    }

    It 'writes no errors, keeps exit 0 and the PASS line when every console call throws' {
        $code = -1; $errors = 0; $lines = New-Object System.Collections.ArrayList
        try {
            Invoke-DefendMain *>&1 | ForEach-Object {
                if ($_ -is [System.Management.Automation.ErrorRecord]) { $errors++ }
                [void]$lines.Add([string]$_)
            }
        } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        $errors | Should -Be 0
        $code | Should -Be 0
        ($lines -join "`n") | Should -Not -Match 'handle is invalid'
        $lines[-1] | Should -Match '\[PASS\] DEFEND.*\(exit 0\)'
    }
}


Describe 'DEFEND finding 55: active threats only' {
    BeforeEach {
        $script:TestPrefs = [pscustomobject]@{ DisableRealtimeMonitoring = $false; PUAProtection = 1; MAPSReporting = 2; EnableNetworkProtection = 1; EnableControlledFolderAccess = 1; CloudBlockLevel = 2; ExclusionPath = @() }
        $script:TestStatus = [pscustomobject]@{ IsTamperProtected = $true; RealTimeProtectionEnabled = $true; AntivirusSignatureVersion = '1.2.3.4' }
        Mock Test-IsAdmin { $true }
        Mock Get-DefendPSMajor { 5 }
        Mock Test-DefendInteractive { $false }
        Mock Clear-DefendScreen {}
        Mock Import-Module {}
        Mock Start-Sleep {}
        Mock Get-CimInstance {
            switch ($ClassName) {
                'Win32_Processor'      { [pscustomobject]@{ VirtualizationFirmwareEnabled = $true } }
                'Win32_ComputerSystem' { [pscustomobject]@{ PartOfDomain = $false } }
                'MSFT_MpComputerStatus' { $script:TestStatus }
            }
        }
        Mock Get-WindowsOptionalFeature { [pscustomobject]@{ State = 'Enabled' } }
        Mock Confirm-SecureBootUEFI { $true }
        Mock Get-Tpm { [pscustomobject]@{ TpmPresent = $true; ManufacturerVersion = '2.0' } }
        Mock Get-DefendRegValue { 1 }
        Mock Test-DefendRegKey { $false }
        Mock Get-DefendHvciBlockers { @() }
        Mock Get-MpPreference { $script:TestPrefs }
        Mock Set-MpPreference {}
        Mock Get-MpComputerStatus { $script:TestStatus }
        Mock Update-MpSignature {}
        Mock Get-WinEvent { @() }
        Mock Set-NetFirewallProfile {}
        Mock Get-NetFirewallProfile { 'Domain', 'Private', 'Public' | ForEach-Object { [pscustomobject]@{ Name = $_; Enabled = 'True' } } }
        Mock Get-NetConnectionProfile { @() }
        # Get-MpThreat is threat HISTORY: it still lists cleaned and quarantined items.
        Mock Get-MpThreat { @([pscustomobject]@{ ThreatID = 11; ThreatName = 'Trojan:Win32/Old'; IsActive = $false }, [pscustomobject]@{ ThreatID = 22; ThreatName = 'Virus:Win32/Live'; IsActive = $true }) }
        Mock Get-MpThreatDetection { @() }
    }

    It 'reports no active threats and exits 0 when history holds only cleaned, quarantined and removed items' {
        Mock Get-MpThreatDetection {
            @(2, 3, 4) | ForEach-Object { [pscustomobject]@{ ThreatID = 11; ThreatStatusID = $_; Resources = @('file:_C:\old.exe') } }
        }
        $r = Get-Run
        $r.Text | Should -Not -Match 'CRITICAL'
        $r.Text | Should -Match 'NO ACTIVE THREATS'
        $r.Code | Should -Be 0
    }

    It 'alerts only for the active detection, with its name and path, and exits 3' {
        Mock Get-MpThreatDetection {
            @(
                [pscustomobject]@{ ThreatID = 11; ThreatStatusID = 3; Resources = @('file:_C:\old.exe') },
                [pscustomobject]@{ ThreatID = 22; ThreatStatusID = 1; Resources = @('file:_C:\live.exe') }
            )
        }
        $r = Get-Run
        $r.Text | Should -Match 'CRITICAL: 1 Threat\(s\)'
        $r.Text | Should -Match 'Virus:Win32/Live'
        $r.Text | Should -Match 'live\.exe'
        $r.Text | Should -Not -Match 'Trojan:Win32/Old'
        $r.Code | Should -Be 3
        $r.Lines[-1] | Should -Match '\[WARN\] DEFEND.*active threat.*\(exit 3\)'
    }

    It 'treats a failed quarantine as still active' {
        Mock Get-MpThreatDetection { @([pscustomobject]@{ ThreatID = 22; ThreatStatusID = 102; Resources = @('file:_C:\live.exe') }) }
        (Get-Run).Text | Should -Match 'CRITICAL: 1 Threat\(s\)'
    }

    It 'reads active detections through the 5.1 child process under PowerShell 7' {
        Mock Get-DefendPSMajor { 7 }
        Mock Invoke-DefendWinPS {
            $o = switch -Wildcard ($Command) {
                '*Get-MpThreatDetection*' { '[{"ThreatID":22,"ThreatStatusID":1,"Resources":["file:_C:\\live.exe"]},{"ThreatID":11,"ThreatStatusID":2,"Resources":["file:_C:\\old.exe"]}]' }
                '*Get-MpThreat *'         { '[{"ThreatID":11,"ThreatName":"Trojan:Win32/Old"},{"ThreatID":22,"ThreatName":"Virus:Win32/Live"}]' }
                '*Get-MpPreference*'      { '{"DisableRealtimeMonitoring":false,"PUAProtection":1,"MAPSReporting":2,"EnableNetworkProtection":1,"EnableControlledFolderAccess":1,"CloudBlockLevel":2,"ExclusionPath":[]}' }
                '*Get-MpComputerStatus*'  { '{"AntivirusSignatureVersion":"1.2.3.4"}' }
                default { '' }
            }
            [pscustomobject]@{ ExitCode = 0; Output = $o }
        }
        $r = Get-Run
        $r.Text | Should -Match 'CRITICAL: 1 Threat\(s\)'
        $r.Text | Should -Match 'Virus:Win32/Live'
        $r.Text | Should -Not -Match 'Trojan:Win32/Old'
    }
}

Describe 'DEFEND findings 56, 57, 58, 10 and 11' {
    BeforeEach {
        $script:TestPrefs = [pscustomobject]@{ DisableRealtimeMonitoring = $false; PUAProtection = 1; MAPSReporting = 2; EnableNetworkProtection = 1; EnableControlledFolderAccess = 1; CloudBlockLevel = 2; ExclusionPath = @() }
        $script:TestStatus = [pscustomobject]@{ IsTamperProtected = $true; RealTimeProtectionEnabled = $true; AntivirusSignatureVersion = '1.2.3.4' }
        Mock Test-IsAdmin { $true }
        Mock Get-DefendPSMajor { 5 }
        Mock Test-DefendInteractive { $false }
        Mock Clear-DefendScreen {}
        Mock Import-Module {}
        Mock Start-Sleep {}
        Mock Get-CimInstance {
            switch ($ClassName) {
                'Win32_Processor'      { [pscustomobject]@{ VirtualizationFirmwareEnabled = $true } }
                'Win32_ComputerSystem' { [pscustomobject]@{ PartOfDomain = $false } }
                'MSFT_MpComputerStatus' { $script:TestStatus }
            }
        }
        Mock Get-WindowsOptionalFeature { [pscustomobject]@{ State = 'Enabled' } }
        Mock Confirm-SecureBootUEFI { $true }
        Mock Get-Tpm { [pscustomobject]@{ TpmPresent = $true; ManufacturerVersion = '2.0' } }
        Mock Get-DefendRegValue { 1 }
        Mock Test-DefendRegKey { $false }
        Mock Get-DefendHvciBlockers { @() }
        Mock Get-MpPreference { $script:TestPrefs }
        Mock Set-MpPreference {}
        Mock Get-MpComputerStatus { $script:TestStatus }
        Mock Get-MpThreat { @() }
        Mock Get-MpThreatDetection { @() }
        Mock Update-MpSignature {}
        Mock Get-WinEvent { @() }
        Mock Invoke-DefendWinPS { [pscustomobject]@{ ExitCode = 0; Output = '' } }
        Mock Set-NetFirewallProfile {}
        Mock Get-NetFirewallProfile { 'Domain', 'Private', 'Public' | ForEach-Object { [pscustomobject]@{ Name = $_; Enabled = 'True' } } }
        Mock Get-NetConnectionProfile { @() }
        Mock Set-NetConnectionProfile {}
    }

    Context 'finding 56: real-time monitoring is read back, not assumed' {
        It 'reports ENFORCED when the readback confirms it' {
            $r = Get-Run
            $r.Text | Should -Match 'RealTimeMonitoring\s+:\s+ENFORCED'
            $r.Code | Should -Be 0
        }
        It 'reports NOT ENFORCED and exits 3 when the setting stays off (Tamper Protection, policy or another AV)' {
            $script:TestPrefs.DisableRealtimeMonitoring = $true
            $r = Get-Run
            $r.Text | Should -Match 'RealTimeMonitoring\s+:\s+NOT ENFORCED'
            $r.Text | Should -Not -Match 'RealTimeMonitoring\s+:\s+ENFORCED'
            $r.Code | Should -Be 3
            $r.Lines[-1] | Should -Match '\[WARN\] DEFEND.*real-time.*\(exit 3\)'
        }
        It 'reports NOT ENFORCED when the engine says real-time protection is off even though the setting reads on' {
            $script:TestStatus.RealTimeProtectionEnabled = $false
            $r = Get-Run
            $r.Text | Should -Match 'RealTimeMonitoring\s+:\s+NOT ENFORCED'
            $r.Code | Should -Be 3
        }
        It 'surfaces the Set-MpPreference error instead of swallowing it' {
            Mock Set-MpPreference { throw 'Operation blocked by Tamper Protection' }
            $script:TestPrefs.DisableRealtimeMonitoring = $true
            $r = Get-Run
            $r.Text | Should -Match 'Tamper Protection'
            $r.Code | Should -Be 3
        }
        It 'surfaces a failed Windows PowerShell child under PowerShell 7' {
            Mock Get-DefendPSMajor { 7 }
            Mock Invoke-DefendWinPS {
                if ($Command -like '*Set-MpPreference*') { return [pscustomobject]@{ ExitCode = 1; Output = 'Set denied by policy' } }
                $o = switch -Wildcard ($Command) {
                    '*Get-MpPreference*'     { '{"DisableRealtimeMonitoring":false,"PUAProtection":1,"MAPSReporting":2,"EnableNetworkProtection":1,"EnableControlledFolderAccess":1,"CloudBlockLevel":2,"ExclusionPath":[]}' }
                    '*Get-MpComputerStatus*' { '{"AntivirusSignatureVersion":"1.2.3.4"}' }
                    default { '' }
                }
                [pscustomobject]@{ ExitCode = 0; Output = $o }
            }
            $r = Get-Run
            $r.Text | Should -Match 'Set denied by policy'
            $r.Code | Should -Be 3
        }
    }
}

