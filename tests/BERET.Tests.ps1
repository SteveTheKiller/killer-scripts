#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# BERET tests. Every BitLocker cmdlet, manage-bde, scheduled task, clock and disk call is mocked.
# Nothing here touches a real volume, and BERET.ps1 is never started in a child process.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    # Commands BERET can call that the shared safety net does not list yet. Blocked unless a test mocks them.
    $block = { throw "UNMOCKED DESTRUCTIVE CALL: $($MyInvocation.MyCommand.Name) was not mocked by this test." }
    foreach ($c in 'Initialize-Tpm', 'Clear-BitLockerAutoUnlock', 'Lock-BitLocker', 'Unlock-BitLocker', 'Wait-Process') {
        if (Get-Command $c -ErrorAction SilentlyContinue) { Mock $c $block }
    }
    Enable-KillerTestMode
    . (Get-KillerScriptPath BERET)

    function New-BeretVolume {
        param([string]$Mount, [string]$Status, [string]$Method, [object[]]$Protectors = @(), [bool]$AutoUnlock = $false)
        [pscustomobject]@{
            MountPoint = $Mount; VolumeStatus = $Status; EncryptionMethod = $Method
            EncryptionPercentage = $(if ($Status -eq 'FullyDecrypted') { 0 } else { 100 })
            AutoUnlockEnabled = $AutoUnlock; KeyProtector = @($Protectors)
        }
    }

    function New-BeretProtector {
        param([string]$Type, [string]$Id = ('{' + [guid]::NewGuid().ToString().ToUpper() + '}'), [string]$Password = '')
        [pscustomobject]@{ KeyProtectorType = $Type; KeyProtectorId = $Id; RecoveryPassword = $Password }
    }

    function Invoke-BeretRun {
        # Runs Main, capturing every stream as text and the KILLERSCRIPT_EXIT code it throws.
        $script:Code = $null
        $script:Lines = New-Object System.Collections.ArrayList
        try {
            Invoke-BeretMain *>&1 | ForEach-Object { [void]$script:Lines.Add("$_") }
        } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $script:Code = [int]$Matches[1] } else { throw }
        }
        $script:Text = ($script:Lines -join "`n")
    }
}

AfterAll {
    Disable-KillerTestMode
}

Describe 'BERET' {
BeforeEach {
    # Fake machine state. Tests change these, the mocks below read them.
    $script:Vols = @{}
    $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyDecrypted' 'None'
    $script:Elevated = $true
    $script:Console = $true
    $script:Choice = 'N'
    $script:Caption = 'Microsoft Windows 11 Pro'
    $script:TpmReady = $true
    $script:TpmInit = [pscustomobject]@{ RestartRequired = $false }
    $script:TpmInitThrows = $false
    $script:Domain = $false
    $script:Dsreg = ''
    $script:Model = 'Latitude 7440'
    $global:BeretMode = $null

    Mock Clear-Host { }
    Mock Start-Sleep { }
    Mock Set-ItemProperty { } -ParameterFilter { $Path -like 'HKLM:*' }
    Mock Start-Process { [pscustomobject]@{ Id = 4242; HasExited = $true } }
    Mock Wait-Process { }
    Mock Stop-Process { }
    Mock Read-Host { throw 'Read-Host must not be reached in this test' }
    Mock Test-BeretElevated { $script:Elevated }
    Mock Test-BeretConsole { $script:Console }
    Mock Read-BeretFipsChoice { $script:Choice }
    Mock Get-BeretDsreg { $script:Dsreg }
    Mock Get-Tpm { [pscustomobject]@{ TpmReady = $script:TpmReady } }
    Mock Initialize-Tpm { if ($script:TpmInitThrows) { throw 'TPM is disabled in firmware' }; $script:TpmInit }
    Mock Get-CimInstance {
        if ($ClassName -eq 'Win32_OperatingSystem') { [pscustomobject]@{ Caption = $script:Caption } }
        else { [pscustomobject]@{ Name = 'TESTPC'; Model = $script:Model; Manufacturer = 'Contoso'; PartOfDomain = $script:Domain } }
    }
    Mock Get-BitLockerVolume {
        if ($MountPoint) { $script:Vols[$MountPoint] } else { @($script:Vols.Values) }
    }
}

Describe 'Entry-point convention' {
    It 'defines Invoke-BeretMain and Exit-KillerScript without running anything on dot-source' {
        Get-Command Invoke-BeretMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
        { Exit-KillerScript 3 } | Should -Throw 'KILLERSCRIPT_EXIT:3'
    }
}

Describe 'Preconditions exit with the contract codes and an honest final line' {
    It 'exits 2 when not elevated, before touching BitLocker' {
        $script:Elevated = $false
        Invoke-BeretRun
        $script:Code | Should -Be 2
        $script:Text | Should -Match 'BERET RESULT: exit code 2'
        Should -Invoke Get-BitLockerVolume -Times 0 -Exactly
    }

    It 'exits 2 on Windows Home' {
        $script:Caption = 'Microsoft Windows 11 Home'
        Invoke-BeretRun
        $script:Code | Should -Be 2
        $script:Text | Should -Match 'BERET RESULT: exit code 2'
        Should -Invoke Get-BitLockerVolume -Times 0 -Exactly
    }

    It 'exits 3010 when the TPM needs a restart to finish initializing' {
        $script:TpmReady = $false
        $script:TpmInit = [pscustomobject]@{ RestartRequired = $true }
        Invoke-BeretRun
        $script:Code | Should -Be 3010
        $script:Text | Should -Match 'BERET RESULT: exit code 3010'
        Should -Invoke Get-BitLockerVolume -Times 0 -Exactly
    }

    It 'exits 2 when the TPM cannot be initialized' {
        $script:TpmReady = $false
        $script:TpmInitThrows = $true
        Invoke-BeretRun
        $script:Code | Should -Be 2
        $script:Text | Should -Match 'BERET RESULT: exit code 2'
        Should -Invoke Get-BitLockerVolume -Times 0 -Exactly
    }

    It 'exits 2 without prompting when there is no console and no preset mode' {
        $script:Console = $false
        Mock Read-BeretFipsChoice { throw 'must not prompt' }
        Invoke-BeretRun
        $script:Code | Should -Be 2
        $script:Text | Should -Match 'BERET RESULT: exit code 2'
        Should -Invoke Read-Host -Times 0 -Exactly
        Should -Invoke Get-BitLockerVolume -Times 0 -Exactly
    }

    It 'uses a preset BeretMode without any prompt, even with no console' {
        $script:Console = $false
        $global:BeretMode = 'FIPS'
        Mock Read-BeretFipsChoice { throw 'must not prompt' }
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes256' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password '111111-222222-333333-444444-555555-666666-777777-888888'))
        Invoke-BeretRun
        $script:Text | Should -Match 'FIPS'
        $script:Text | Should -Not -Match 'exit code 2'
    }
}
}
