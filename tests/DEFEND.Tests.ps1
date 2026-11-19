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

