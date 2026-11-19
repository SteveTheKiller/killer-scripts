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

    Context 'finding 57: definition update result' {
        BeforeEach {
            $script:Ps7Json = {
                param($Command, $UpdateExit)
                if ($Command -like '*Update-MpSignature*') { return [pscustomobject]@{ ExitCode = $UpdateExit; Output = $(if ($UpdateExit -ne 0) { 'update server unreachable' } else { '' }) } }
                $o = switch -Wildcard ($Command) {
                    '*Get-MpPreference*'     { '{"DisableRealtimeMonitoring":false,"PUAProtection":1,"MAPSReporting":2,"EnableNetworkProtection":1,"EnableControlledFolderAccess":1,"CloudBlockLevel":2,"ExclusionPath":[]}' }
                    '*Get-MpComputerStatus*' { '{"AntivirusSignatureVersion":"1.2.3.4"}' }
                    default { '' }
                }
                [pscustomobject]@{ ExitCode = 0; Output = $o }
            }
        }
        It 'prints Definitions Updated and exits 0 when the update works' {
            $r = Get-Run
            $r.Text | Should -Match 'Definitions Updated'
            $r.Code | Should -Be 0
        }
        It 'reports Update Failed with the reason and exits 3 when Update-MpSignature throws' {
            Mock Update-MpSignature { throw 'update server unreachable' }
            $r = Get-Run
            $r.Text | Should -Match 'Update Failed'
            $r.Text | Should -Not -Match 'Definitions Updated'
            $r.Text | Should -Match 'update server unreachable'
            $r.Code | Should -Be 3
        }
        It 'checks the child exit code under PowerShell 7: nonzero is Update Failed, exit 3' {
            Mock Get-DefendPSMajor { 7 }
            Mock Invoke-DefendWinPS { & $script:Ps7Json $Command 1 }
            $r = Get-Run
            $r.Text | Should -Match 'Update Failed'
            $r.Text | Should -Not -Match 'Definitions Updated'
            $r.Code | Should -Be 3
        }
        It 'under PowerShell 7 a zero child exit is Definitions Updated and exit 0' {
            Mock Get-DefendPSMajor { 7 }
            Mock Invoke-DefendWinPS { & $script:Ps7Json $Command 0 }
            $r = Get-Run
            $r.Text | Should -Match 'Definitions Updated'
            $r.Code | Should -Be 0
        }
    }

    Context 'finding 58 (D7): exclusions are known-good only under the Huntress install roots' {
        It 'does not hide Scenarios, Priority or other paths that merely contain Rio or Huntress' {
            $script:TestPrefs.ExclusionPath = @('C:\Scenarios', 'C:\Users\bob\Priority\', 'C:\Temp\Huntress', 'C:\Windows\Temp\Rio\x', 'C:\Program Files\HuntressEvil\x')
            $r = Get-Run
            foreach ($p in 'C:\Scenarios', 'C:\Users\bob\Priority\', 'C:\Temp\Huntress', 'C:\Windows\Temp\Rio\x', 'C:\Program Files\HuntressEvil\x') {
                $r.Text | Should -Match ([regex]::Escape($p))
            }
            $r.Text | Should -Match '5 UNVERIFIED'
            $r.Text | Should -Not -Match 'MSP Stack'
        }
        It 'accepts the three Huntress roots and anything under them, in any case, as known-good' {
            $script:TestPrefs.ExclusionPath = @('C:\Program Files\Huntress', 'C:\Program Files\Huntress\Rio\Rio.exe', 'c:\program files (x86)\huntress\', 'C:\ProgramData\Huntress\Logs')
            $r = Get-Run
            $r.Text | Should -Match 'NONE \(Secure\)'
            $r.Text | Should -Match 'MSP Stack'
            $r.Text | Should -Not -Match 'UNVERIFIED'
        }
        It 'rejects traversal and wildcard tricks that escape the Huntress roots' {
            $script:TestPrefs.ExclusionPath = @('C:\Program Files\Huntress\..\Temp', 'C:\Program Files\Huntress*', 'C:\*\Huntress')
            (Get-Run).Text | Should -Match '3 UNVERIFIED'
        }
        It 'counts only the unverified paths when known-good and unknown paths are mixed' {
            $script:TestPrefs.ExclusionPath = @('C:\ProgramData\Huntress', 'C:\Users\bob\Downloads')
            $r = Get-Run
            $r.Text | Should -Match '1 UNVERIFIED'
            $r.Text | Should -Match 'C:\\Users\\bob\\Downloads'
            $r.Text | Should -Match 'MSP Stack'
        }
    }

    Context 'finding 10: firewall profiles are enforced and verified, networks are not reclassified' {
        It 'never changes a connection profile on a workgroup machine with a Public network' {
            Mock Get-NetConnectionProfile { [pscustomobject]@{ Name = 'Cafe'; NetworkCategory = 'Public' } }
            $r = Get-Run
            Should -Invoke Set-NetConnectionProfile -Times 0 -Exactly
            $r.Text | Should -Not -Match 'Private Security Enforced'
            $r.Text | Should -Match 'Cafe'
            $r.Text | Should -Match 'Public'
        }
        It 'never tries to set DomainAuthenticated on a domain member' {
            Mock Get-CimInstance {
                switch ($ClassName) {
                    'Win32_Processor'      { [pscustomobject]@{ VirtualizationFirmwareEnabled = $true } }
                    'Win32_ComputerSystem' { [pscustomobject]@{ PartOfDomain = $true } }
                    'MSFT_MpComputerStatus' { $script:TestStatus }
                }
            }
            Mock Get-NetConnectionProfile { [pscustomobject]@{ Name = 'corp.local'; NetworkCategory = 'DomainAuthenticated' } }
            $r = Get-Run
            Should -Invoke Set-NetConnectionProfile -Times 0 -Exactly
            $r.Text | Should -Match 'Domain Joined'
            $r.Text | Should -Not -Match 'Domain Security Enforced'
            $r.Code | Should -Be 0
        }
        It 'enables the Domain, Private and Public profiles and reports the verified result' {
            $r = Get-Run
            Should -Invoke Set-NetFirewallProfile -Times 1 -Exactly -ParameterFilter { $Enabled -eq 'True' -and (@($Profile) -join ',') -eq 'Domain,Private,Public' }
            $r.Text | Should -Match 'Firewall Profiles\s+:\s+Domain=On Private=On Public=On'
            $r.Code | Should -Be 0
        }
        It 'does not claim enforcement when the readback shows a profile still off, and exits 3' {
            Mock Get-NetFirewallProfile {
                [pscustomobject]@{ Name = 'Domain'; Enabled = 'True' }
                [pscustomobject]@{ Name = 'Private'; Enabled = 'True' }
                [pscustomobject]@{ Name = 'Public'; Enabled = 'False' }
            }
            $r = Get-Run
            $r.Text | Should -Match 'Public=OFF'
            $r.Text | Should -Not -Match 'Security Enforced'
            $r.Code | Should -Be 3
            $r.Lines[-1] | Should -Match '\[WARN\] DEFEND.*firewall.*\(exit 3\)'
        }
        It 'reports a refused firewall change and exits 3' {
            Mock Set-NetFirewallProfile { throw 'Access is denied (policy)' }
            $r = Get-Run
            $r.Text | Should -Match 'Access is denied'
            $r.Code | Should -Be 3
        }
        It 'does not claim enforcement when the profiles cannot be read back' {
            Mock Get-NetFirewallProfile { throw 'readback unavailable' }
            $r = Get-Run
            $r.Text | Should -Match 'UNVERIFIED'
            $r.Code | Should -Be 3
        }
    }

    Context 'finding 11: scan result, completion and deadlines (interactive session)' {
        BeforeAll {
            function New-FakeScan {
                param([int]$Polls = 2, [int]$ExitCode = 0, [int]$Id = 4242)
                $o = [pscustomobject]@{ Id = $Id; ExitCode = $ExitCode; Polls = 0; Limit = $Polls }
                Add-Member -InputObject $o -MemberType ScriptProperty -Name HasExited -Value { $this.Polls++; $this.Polls -gt $this.Limit }
                $o
            }
        }
        BeforeEach {
            Mock Test-DefendInteractive { $true }
            Mock Get-DefendConsoleValue { 5 }
            Mock Set-DefendConsolePosition {}
            Mock Test-DefendKeyAvailable { $false }
            Mock Stop-DefendScan {}
            Mock Test-DefendEngineBusy { $false }
            $script:Now = [datetime]'2030-01-01T00:00:00'
            $script:NowCalls = 0
            # One fake minute passes per call. A runaway wait loop throws instead of hanging the suite.
            Mock Get-DefendNow {
                $script:NowCalls++
                if ($script:NowCalls -gt 400) { throw 'runaway wait loop: no deadline' }
                $script:Now = $script:Now.AddMinutes(1); $script:Now
            }
            Mock Start-DefendScanProcess { New-FakeScan -Polls 2 -ExitCode 0 }
            $script:DoneEvent = [pscustomobject]@{ TimeCreated = [datetime]'2031-01-01'; Message = 'The antimalware scan has finished. Scan Type: Full Scan' }
            Mock Get-WinEvent { @($script:DoneEvent) }
        }

        It 'reports COMPLETE and exits 0 when the scan exits 0 and a completion event exists' {
            $r = Get-Run
            $r.Text | Should -Match 'Scan Status\s+:\s+COMPLETE'
            $r.Code | Should -Be 0
        }
        It 'reports FAILED with the exit code, not COMPLETE, when the scan process fails' {
            Mock Start-DefendScanProcess { New-FakeScan -Polls 2 -ExitCode 2 }
            $r = Get-Run
            $r.Text | Should -Match 'Scan Status\s+:\s+FAILED \(exit 2\)'
            $r.Text | Should -Not -Match 'Scan Status\s+:\s+COMPLETE'
            $r.Code | Should -Be 3
            $r.Lines[-1] | Should -Match '\[WARN\] DEFEND.*scan.*\(exit 3\)'
        }
        It 'does not report COMPLETE when exit 0 is not backed by a completion event' {
            Mock Get-WinEvent { @() }
            $r = Get-Run
            $r.Text | Should -Match 'Scan Status\s+:\s+NOT VERIFIED'
            $r.Text | Should -Not -Match 'Scan Status\s+:\s+COMPLETE'
            $r.Code | Should -Be 3
        }
        It 'ignores a completion event older than this scan' {
            $script:DoneEvent = [pscustomobject]@{ TimeCreated = [datetime]'2020-01-01'; Message = 'Full Scan finished' }
            (Get-Run).Text | Should -Match 'Scan Status\s+:\s+NOT VERIFIED'
        }
        It 'stops a quick scan that runs past its deadline and reports TIMED OUT' {
            $script:KeyCalls = 0
            Mock Test-DefendKeyAvailable { $script:KeyCalls++; $script:KeyCalls -eq 1 }
            Mock Read-DefendKey { [pscustomobject]@{ Key = 'Q'; KeyChar = 'q' } }
            $script:ProcN = 0
            Mock Start-DefendScanProcess {
                $script:ProcN++
                if ($script:ProcN -eq 1) { New-FakeScan -Polls 1000000 -ExitCode 0 -Id 1 } else { New-FakeScan -Polls 1000000 -ExitCode 0 -Id 2 }
            }
            $r = Get-Run
            $r.Text | Should -Match 'Scan Status\s+:\s+TIMED OUT'
            Should -Invoke Stop-DefendScan -Times 1 -Exactly -ParameterFilter { $Proc.Id -eq 2 }
            $r.Code | Should -Be 3
        }
        It 'gives up on an engine that stays busy after Q and does not start the quick scan' {
            $script:KeyCalls = 0
            Mock Test-DefendKeyAvailable { $script:KeyCalls++; $script:KeyCalls -eq 1 }
            Mock Read-DefendKey { [pscustomobject]@{ Key = 'Q'; KeyChar = 'q' } }
            $script:BusyCalls = 0
            Mock Test-DefendEngineBusy { $script:BusyCalls++; if ($script:BusyCalls -gt 300) { throw 'runaway busy wait: no deadline' }; $true }
            Mock Start-DefendScanProcess { New-FakeScan -Polls 1000000 -ExitCode 0 -Id 1 }
            $r = Get-Run
            Should -Invoke Start-DefendScanProcess -Times 1 -Exactly
            $r.Text | Should -Match 'Scan Status\s+:\s+FAILED'
            $r.Text | Should -Match 'busy'
            $r.Code | Should -Be 3
        }
        It 'treats Esc as a user skip: SKIPPED BY USER and exit 0' {
            Mock Test-DefendKeyAvailable { $true }
            Mock Read-DefendKey { [pscustomobject]@{ Key = 'Escape'; KeyChar = [char]27 } }
            Mock Start-DefendScanProcess { New-FakeScan -Polls 1000000 -ExitCode 0 }
            $r = Get-Run
            $r.Text | Should -Match 'SKIPPED BY USER'
            $r.Code | Should -Be 0
        }
        It 'skips the live scan in a non-interactive session and still exits 0' {
            Mock Test-DefendInteractive { $false }
            $r = Get-Run
            $r.Text | Should -Match 'SKIPPED \(Non-Interactive Session\)'
            Should -Invoke Start-DefendScanProcess -Times 0 -Exactly
            $r.Code | Should -Be 0
        }
    }
}

