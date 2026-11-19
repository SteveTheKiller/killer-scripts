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
    $script:Now = [datetime]'2026-11-18T10:00:00'
    $script:Sleeps = 0
    $script:DecryptPollsNeeded = 2
    $script:DecryptPolls = 0
    $script:BdeCalls = New-Object System.Collections.ArrayList
    $script:BdeResults = New-Object System.Collections.Queue
    $script:BdeDefault = [pscustomobject]@{ ExitCode = 0; Output = 'BitLocker Drive Encryption: Configuration Tool' }
    $script:EncryptNeverStarts = $false

    # The fake clock moves only when the script sleeps, and a runaway loop fails the test fast.
    Mock Start-Sleep {
        $script:Sleeps++
        if ($script:Sleeps -gt 2000) { throw 'runaway loop: more than 2000 sleeps' }
        $script:Now = $script:Now.AddSeconds([int]$Seconds)
    }
    Mock Get-BeretNow { $script:Now }
    Mock Get-BitLockerVolume {
        $list = if ($MountPoint) { @($MountPoint | ForEach-Object { $script:Vols[[string]$_] }) } else { @($script:Vols.Values) }
        foreach ($v in $list) {
            if ($v.VolumeStatus -eq 'DecryptionInProgress') {
                $script:DecryptPolls++
                if ($script:DecryptPolls -ge $script:DecryptPollsNeeded) {
                    $v.VolumeStatus = 'FullyDecrypted'; $v.EncryptionMethod = 'None'; $v.EncryptionPercentage = 0; $v.KeyProtector = @()
                }
            }
            $v
        }
    }
    Mock Disable-BitLocker {
        $v = $script:Vols[[string]@($MountPoint)[0]]
        $v.VolumeStatus = 'DecryptionInProgress'; $v.EncryptionPercentage = 100
        $script:DecryptPolls = 0
    }
    Mock Add-BitLockerKeyProtector {
        $v = $script:Vols[[string]@($MountPoint)[0]]
        if ($TpmProtector) { $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'Tpm') }
        elseif ($RecoveryPasswordProtector) {
            $n = @($v.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' }).Count + 1
            $digits = (($n * 111111).ToString('D6') + '-') * 7 + ($n * 111111).ToString('D6')
            $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'RecoveryPassword' -Password $digits)
        }
        $v
    }
    Mock Invoke-BeretManageBde {
        [void]$script:BdeCalls.Add(($Arguments -join ' '))
        $r = if ($script:BdeResults.Count -gt 0) { $script:BdeResults.Dequeue() } else { $script:BdeDefault }
        if ($r.ExitCode -eq 0 -and $Arguments -contains '-on' -and -not $script:EncryptNeverStarts) {
            $names = @{ 'aes128' = 'Aes128'; 'aes256' = 'Aes256'; 'xts_aes128' = 'XtsAes128'; 'xts_aes256' = 'XtsAes256' }
            $i = [array]::IndexOf($Arguments, '-EncryptionMethod')
            $v = $script:Vols[$Arguments[1]]
            $v.EncryptionMethod = $names[[string]$Arguments[$i + 1]]; $v.VolumeStatus = 'EncryptionInProgress'; $v.EncryptionPercentage = 1
        }
        $r
    }
}

Describe 'Finding 06 and D4: the cipher must match exactly' {
    BeforeEach {
        $script:Key = '111111-222222-333333-444444-555555-666666-777777-888888'
    }

    It 'STANDARD leaves a healthy XtsAes128 volume with a recovery key alone' {
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes128' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $script:Key))
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 0 -Exactly
        $script:BdeCalls.Count | Should -Be 0
    }

    It 'STANDARD treats <Have> as non-compliant and runs the full cycle to XtsAes128' -ForEach @(
        @{ Have = 'Aes128' }, @{ Have = 'Aes256' }, @{ Have = 'XtsAes256' }
    ) {
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' $Have @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $script:Key))
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 1 -Exactly
        $script:Vols['C:'].EncryptionMethod | Should -Be 'XtsAes128'
    }

    It 'FIPS treats <Have> as non-compliant and runs the full cycle to XtsAes256' -ForEach @(
        @{ Have = 'Aes256' }, @{ Have = 'Aes128' }, @{ Have = 'XtsAes128' }
    ) {
        $global:BeretMode = 'FIPS'
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' $Have @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $script:Key))
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 1 -Exactly
        $script:Vols['C:'].EncryptionMethod | Should -Be 'XtsAes256'
    }

    It 'FIPS leaves a healthy XtsAes256 volume with a recovery key alone' {
        $global:BeretMode = 'FIPS'
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes256' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $script:Key))
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 0 -Exactly
    }
}

Describe 'Finding 07: manage-bde is called with the real XTS names and its exit code is checked' {
    It 'asks manage-bde for <Bde> in <Mode> mode' -ForEach @(
        @{ Mode = 'STANDARD'; Bde = 'xts_aes128' }, @{ Mode = 'FIPS'; Bde = 'xts_aes256' }
    ) {
        $global:BeretMode = $Mode
        Invoke-BeretRun
        $script:BdeCalls.Count | Should -Be 1
        $script:BdeCalls[0] | Should -Match "-EncryptionMethod $Bde( |$)"
    }

    It 'uses used-space-only on hardware and leaves the switch out on a VM' {
        Invoke-BeretRun
        $script:BdeCalls[0] | Should -Match '-UsedSpaceOnly'
        $script:BdeCalls[0] | Should -Not -Match 'UsedSpaceOnly:'
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyDecrypted' 'None'
        $script:Model = 'Virtual Machine'
        $script:BdeCalls.Clear()
        Invoke-BeretRun
        $script:BdeCalls[0] | Should -Not -Match 'UsedSpaceOnly'
    }

    It 'exits 1 and says so when manage-bde -on fails' {
        $script:BdeResults.Enqueue([pscustomobject]@{ ExitCode = 1; Output = 'ERROR: Volume cannot be encrypted.' })
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BERET RESULT: exit code 1'
    }

    It 'retries full-volume on 0x803100a5 and carries on when that works' {
        $script:BdeResults.Enqueue([pscustomobject]@{ ExitCode = 1; Output = 'Error 0x803100a5 used space only is not supported' })
        Invoke-BeretRun
        $script:BdeCalls.Count | Should -Be 2
        $script:BdeCalls[1] | Should -Not -Match 'UsedSpaceOnly'
        $script:Code | Should -Not -Be 1
        $script:Vols['C:'].EncryptionMethod | Should -Be 'XtsAes128'
    }

    It 'exits 1 when the full-volume retry also fails' {
        $script:BdeResults.Enqueue([pscustomobject]@{ ExitCode = 1; Output = 'Error 0x803100a5' })
        $script:BdeResults.Enqueue([pscustomobject]@{ ExitCode = 2147942402; Output = 'ERROR: not found' })
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BERET RESULT: exit code 1'
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
