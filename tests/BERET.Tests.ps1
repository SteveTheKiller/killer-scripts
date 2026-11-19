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
    $script:EncryptMethodOverride = $null
    $script:Events = New-Object System.Collections.ArrayList
    $script:AdVerify = $true
    $script:EntraVerify = $true
    $script:Net = $true
    $script:AdDirectFails = $false
    $script:EntraFails = $false
    $script:TaskResults = New-Object System.Collections.Queue
    $script:TaskUnregistered = 0
    $global:BeretKeyFile = $null

    Mock Test-BeretNetwork { $script:Net }
    Mock Test-BeretAdEscrow { $script:AdVerify }
    Mock Test-BeretEntraEscrow { $script:EntraVerify }
    Mock Backup-BitLockerKeyProtector {
        [void]$script:Events.Add('AD-BACKUP')
        if ($script:AdDirectFails) { throw 'Access is denied (0x80070005)' }
    }
    Mock BackupToAAD-BitLockerKeyProtector {
        [void]$script:Events.Add('AAD-BACKUP')
        if ($script:EntraFails) { throw 'The device is not registered' }
    }
    Mock New-ScheduledTaskAction { New-CimInstance -ClientOnly -ClassName 'MSFT_TaskAction' -Namespace 'Root/Microsoft/Windows/TaskScheduler' -Property @{ Execute = [string]$Execute } }
    Mock New-ScheduledTaskPrincipal { New-CimInstance -ClientOnly -ClassName 'MSFT_TaskPrincipal' -Namespace 'Root/Microsoft/Windows/TaskScheduler' -Property @{ UserId = [string]$UserId } }
    Mock Register-ScheduledTask { [void]$script:Events.Add('TASK-REGISTER') }
    Mock Start-ScheduledTask { [void]$script:Events.Add('TASK-START') }
    Mock Get-ScheduledTaskInfo {
        $r = if ($script:TaskResults.Count -gt 1) { $script:TaskResults.Dequeue() } elseif ($script:TaskResults.Count -eq 1) { $script:TaskResults.Peek() } else { 0 }
        [pscustomobject]@{ LastTaskResult = $r }
    }
    Mock Unregister-ScheduledTask { $script:TaskUnregistered++ }
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
        [void]$script:Events.Add('DECRYPT')
    }
    Mock Disable-BitLockerAutoUnlock {
        $m = [string]@($MountPoint)[0]
        [void]$script:Events.Add("AU-OFF $m")
        $script:Vols[$m].AutoUnlockEnabled = $false
    }
    Mock Enable-BitLockerAutoUnlock {
        $m = [string]@($MountPoint)[0]
        [void]$script:Events.Add("AU-ON $m")
        $script:Vols[$m].AutoUnlockEnabled = $true
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
        Write-Host "BDE-CALL: $($Arguments -join ' ')"
        [void]$script:Events.Add('BDE-ON')
        $r = if ($script:BdeResults.Count -gt 0) { $script:BdeResults.Dequeue() } else { $script:BdeDefault }
        if ($r.ExitCode -eq 0 -and $Arguments -contains '-on' -and -not $script:EncryptNeverStarts) {
            $names = @{ 'aes128' = 'Aes128'; 'aes256' = 'Aes256'; 'xts_aes128' = 'XtsAes128'; 'xts_aes256' = 'XtsAes256' }
            $i = [array]::IndexOf($Arguments, '-EncryptionMethod')
            $v = $script:Vols[$Arguments[1]]
            $v.EncryptionMethod = $(if ($script:EncryptMethodOverride) { $script:EncryptMethodOverride } else { $names[[string]$Arguments[$i + 1]] }); $v.VolumeStatus = 'EncryptionInProgress'; $v.EncryptionPercentage = 1
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

Describe 'Finding 07 and D3: the technician sees the recovery key before encryption and again at the end' {
    It 'prints the new key before manage-bde -on runs, and again in the final summary' {
        Invoke-BeretRun
        $key = '111111-111111-111111-111111-111111-111111-111111-111111'
        $first = $script:Text.IndexOf($key)
        $first | Should -BeGreaterThan -1
        $first | Should -BeLessThan $script:Text.IndexOf('BDE-CALL')
        [regex]::Matches($script:Text, [regex]::Escape($key)).Count | Should -BeGreaterOrEqual 2
    }

    It 'prints the key of an already compliant volume at the start and again at the end' {
        $key = '222222-333333-444444-555555-666666-777777-888888-999999'
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes128' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $key))
        Invoke-BeretRun
        [regex]::Matches($script:Text, [regex]::Escape($key)).Count | Should -BeGreaterOrEqual 2
    }
}

Describe 'Finding 06 and 07: waits are bounded and a timeout is a failure' {
    It 'prints DONE only after the volume has really finished decrypting' {
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'Aes128' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password '111111-222222-333333-444444-555555-666666-777777-888888'))
        Invoke-BeretRun
        $script:Text | Should -Match 'Decryption in progress\.\.\.\s+\[DONE\]'
    }

    It 'exits 1, never prints DONE and builds nothing when decryption does not finish in 30 minutes' {
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'Aes128' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password '111111-222222-333333-444444-555555-666666-777777-888888'))
        $script:DecryptPollsNeeded = [int]::MaxValue
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BERET RESULT: exit code 1'
        $script:Text | Should -Match 'still'
        $script:Text | Should -Not -Match '\[DONE\]'
        Should -Invoke Add-BitLockerKeyProtector -Times 0 -Exactly
        $script:BdeCalls.Count | Should -Be 0
    }

    It 'exits 1 when encryption never starts after manage-bde reports success' {
        $script:EncryptNeverStarts = $true
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BERET RESULT: exit code 1'
    }

    It 'exits 1 when the volume ends up with a different cipher than the mode asks for' {
        $script:EncryptMethodOverride = 'Aes128'
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'Aes128'
        $script:Text | Should -Match 'BERET RESULT: exit code 1'
    }

    It 'exits 1 when the volume state cannot be read' {
        Mock Get-BitLockerVolume { throw 'WMI provider failure' }
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BERET RESULT: exit code 1'
    }

    It 'exits 1 and starts no encryption when the recovery password protector cannot be added' {
        Mock Add-BitLockerKeyProtector { if ($RecoveryPasswordProtector) { throw 'Group Policy forbids recovery passwords' }; $v = $script:Vols['C:']; $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'Tpm'); $v }
        Invoke-BeretRun
        $script:Code | Should -Be 1
        $script:BdeCalls.Count | Should -Be 0
    }
}

Describe 'Finding 06 and D4: a missing recovery key alone is fixed in place' {
    BeforeEach {
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes128' @((New-BeretProtector 'Tpm'))
    }

    It 'adds the recovery protector, reads it back and shows it, without decrypting or calling manage-bde' {
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 0 -Exactly
        $script:BdeCalls.Count | Should -Be 0
        $rec = @($script:Vols['C:'].KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
        $rec.Count | Should -Be 1
        [regex]::Matches($script:Text, [regex]::Escape($rec[0].RecoveryPassword)).Count | Should -BeGreaterOrEqual 2
        $script:Text | Should -Match 'in place'
    }

    It 'falls back to the full decrypt and re-encrypt cycle when the read-back is not a real key' {
        $script:AddCalls = 0
        Mock Add-BitLockerKeyProtector {
            $script:AddCalls++
            $v = $script:Vols['C:']
            if ($TpmProtector) { $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'Tpm') }
            elseif ($script:AddCalls -eq 1) { $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'RecoveryPassword' -Password '') }
            else { $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'RecoveryPassword' -Password '123456-123456-123456-123456-123456-123456-123456-123456') }
            $v
        }
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 1 -Exactly
        $script:Vols['C:'].EncryptionMethod | Should -Be 'XtsAes128'
        @($script:Vols['C:'].KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' -and $_.RecoveryPassword }).Count | Should -Be 1
    }

    It 'falls back to the full cycle when adding the protector throws' {
        $script:AddCalls = 0
        Mock Add-BitLockerKeyProtector {
            $script:AddCalls++
            $v = $script:Vols['C:']
            if ($RecoveryPasswordProtector -and $script:AddCalls -eq 1) { throw 'The BitLocker service is busy' }
            if ($TpmProtector) { $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'Tpm') }
            else { $v.KeyProtector = @($v.KeyProtector) + (New-BeretProtector 'RecoveryPassword' -Password '123456-123456-123456-123456-123456-123456-123456-123456') }
            $v
        }
        Invoke-BeretRun
        Should -Invoke Disable-BitLocker -Times 1 -Exactly
    }
}

Describe 'Finding 06 and D4: data volume auto-unlock is restored and verified' {
    BeforeEach {
        $script:Key = '111111-222222-333333-444444-555555-666666-777777-888888'
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'Aes128' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $script:Key))
        $script:Vols['D:'] = New-BeretVolume 'D:' 'FullyEncrypted' 'XtsAes128' @((New-BeretProtector 'ExternalKey')) $true
    }

    It 'turns auto-unlock off before decrypting and back on after encryption has started' {
        Invoke-BeretRun
        $ev = $script:Events -join ','
        $ev | Should -Match 'AU-OFF D:,DECRYPT,BDE-ON,AU-ON D:'
        $script:Vols['D:'].AutoUnlockEnabled | Should -BeTrue
    }

    It 'exits 3 and names the volume when auto-unlock cannot be restored' {
        Mock Enable-BitLockerAutoUnlock { throw 'The volume is locked' }
        Invoke-BeretRun
        $script:Code | Should -Be 3
        $script:Text | Should -Match 'BERET RESULT: exit code 3'
        $script:Text | Should -Match 'D:'
        $script:Text | Should -Match 'Auto-Unlock'
    }

    It 'does not trust the cmdlet: exits 3 when the read-back still shows auto-unlock off' {
        Mock Enable-BitLockerAutoUnlock { }
        Invoke-BeretRun
        $script:Code | Should -Be 3
    }

    It 'restores auto-unlock when the decrypt times out, and still exits 1' {
        $script:DecryptPollsNeeded = [int]::MaxValue
        Invoke-BeretRun
        $script:Code | Should -Be 1
        ($script:Events -join ',') | Should -Match 'AU-ON D:'
    }

    It 'exits 1 without decrypting anything when auto-unlock cannot be turned off' {
        Mock Disable-BitLockerAutoUnlock { throw 'Access denied' }
        Invoke-BeretRun
        $script:Code | Should -Be 1
        Should -Invoke Disable-BitLocker -Times 0 -Exactly
    }

    It 'leaves auto-unlock alone when no decryption is needed' {
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes128' @((New-BeretProtector 'Tpm'), (New-BeretProtector 'RecoveryPassword' -Password $script:Key))
        Invoke-BeretRun
        Should -Invoke Disable-BitLockerAutoUnlock -Times 0 -Exactly
        Should -Invoke Enable-BitLockerAutoUnlock -Times 0 -Exactly
    }
}

Describe 'Finding 07: escrow is verified separately for every destination' {
    BeforeEach {
        $script:Key = '111111-111111-111111-111111-111111-111111-111111-111111'
    }

    Context 'no destination' {
        It 'exits 3 and says the key is only on screen and on the volume when there is nowhere to escrow it' {
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match 'BERET RESULT: exit code 3'
            $script:Text | Should -Match 'no escrow destination'
        }
    }

    Context 'Active Directory' {
        BeforeEach { $script:Domain = $true }

        It 'reports AD verified and exits 0 when the backup succeeds and the read-back finds the key' {
            Invoke-BeretRun
            $script:Code | Should -Be 0
            $script:Text | Should -Match '(?s)Active Directory.*\[VERIFIED\]'
            $script:Text | Should -Match 'BERET RESULT: exit code 0'
        }

        It 'exits 3 when the backup worked but the read-back cannot find the key in AD' {
            $script:AdVerify = $false
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match '(?s)Active Directory.*\[FAILED\]'
        }

        It 'reports success without a read-back when AD cannot be queried, and says so' {
            $script:AdVerify = $null
            Invoke-BeretRun
            $script:Code | Should -Be 0
            $script:Text | Should -Match 'not read back'
        }

        It 'retries through a SYSTEM task and accepts it only when the task result is 0' {
            $script:AdDirectFails = $true
            $script:TaskResults.Enqueue(0)
            Invoke-BeretRun
            $script:Code | Should -Be 0
            ($script:Events -join ',') | Should -Match 'TASK-REGISTER,TASK-START'
            $script:Text | Should -Match '(?s)Active Directory.*\[VERIFIED\]'
        }

        It 'exits 3 when the SYSTEM task ran and failed, instead of reporting success after a sleep' {
            $script:AdDirectFails = $true
            $script:TaskResults.Enqueue(1)
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match '(?s)Active Directory.*\[FAILED\]'
            $script:Text | Should -Match 'result 1'
        }

        It 'waits through the running state and exits 3 when the task never finishes' {
            $script:AdDirectFails = $true
            $script:TaskResults.Enqueue(267009)
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match 'did not finish'
        }

        It 'treats a task that never ran (267011) as failed' {
            $script:AdDirectFails = $true
            $script:TaskResults.Enqueue(267011)
            Invoke-BeretRun
            $script:Code | Should -Be 3
        }

        It 'waits for a running task to finish and then reads its result' {
            $script:AdDirectFails = $true
            $script:TaskResults.Enqueue(267009)
            $script:TaskResults.Enqueue(267009)
            $script:TaskResults.Enqueue(0)
            Invoke-BeretRun
            $script:Code | Should -Be 0
        }

        It 'always removes the scheduled task, also when it failed' {
            $script:AdDirectFails = $true
            $script:TaskResults.Enqueue(1)
            Invoke-BeretRun
            $script:TaskUnregistered | Should -Be 1
        }
    }

    Context 'Microsoft Entra ID' {
        BeforeEach { $script:Dsreg = "AzureAdJoined : YES`nEnterpriseJoined : NO" }

        It 'reports Entra verified when the backup event confirms it' {
            Invoke-BeretRun
            $script:Code | Should -Be 0
            $script:Text | Should -Match '(?s)Entra ID.*\[VERIFIED\]'
        }

        It 'exits 3 when the backup throws' {
            $script:EntraFails = $true
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match '(?s)Entra ID.*\[FAILED\]'
        }

        It 'exits 3 when the failure event is logged after a backup that did not throw' {
            $script:EntraVerify = $false
            Invoke-BeretRun
            $script:Code | Should -Be 3
        }

        It 'exits 3 without calling the backup when Entra cannot be reached' {
            $script:Net = $false
            Invoke-BeretRun
            $script:Code | Should -Be 3
            Should -Invoke BackupToAAD-BitLockerKeyProtector -Times 0 -Exactly
        }
    }

    Context 'several destinations' {
        It 'reports AD and Entra independently and exits 3 when only one of them failed' {
            $script:Domain = $true
            $script:Dsreg = 'AzureAdJoined : YES'
            $script:EntraFails = $true
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match '(?s)Active Directory.*\[VERIFIED\]'
            $script:Text | Should -Match '(?s)Entra ID.*\[FAILED\]'
            $script:Text | Should -Match 'BERET RESULT: exit code 3'
        }

        It 'escrows every recovery password on the volume' {
            $script:Domain = $true
            $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyEncrypted' 'XtsAes128' @((New-BeretProtector 'Tpm'),
                (New-BeretProtector 'RecoveryPassword' -Password '111111-111111-111111-111111-111111-111111-111111-111111'),
                (New-BeretProtector 'RecoveryPassword' -Password '222222-222222-222222-222222-222222-222222-222222-222222'))
            Invoke-BeretRun
            Should -Invoke Backup-BitLockerKeyProtector -Times 2 -Exactly
        }
    }

    Context 'local key file' {
        It 'writes the key to the file the operator asked for, reads it back and counts it as a destination' {
            $global:BeretKeyFile = Join-Path $TestDrive 'keys\C-recovery.txt'
            Invoke-BeretRun
            $script:Code | Should -Be 0
            (Get-Content -LiteralPath $global:BeretKeyFile -Raw) | Should -Match '111111-111111-111111-111111-111111-111111-111111-111111'
            $script:Text | Should -Match '(?s)Key file.*\[VERIFIED\]'
        }

        It 'exits 3 when the key file cannot be written, and still reports the other destinations' {
            $script:Domain = $true
            $blocker = Join-Path $TestDrive 'iamafile'
            [IO.File]::WriteAllText($blocker, 'x')
            $global:BeretKeyFile = Join-Path $blocker 'keys.txt'
            Invoke-BeretRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match '(?s)Key file.*\[FAILED\]'
            $script:Text | Should -Match '(?s)Active Directory.*\[VERIFIED\]'
        }
    }
}
}

Describe 'Finding 06: a 32-bit host reaches the real System32' {
    BeforeEach {
        $script:SavedArch = $env:PROCESSOR_ARCHITEW6432
        $script:SavedWin = $env:windir
    }
    AfterEach {
        $env:PROCESSOR_ARCHITEW6432 = $script:SavedArch
        $env:windir = $script:SavedWin
    }

    It 'uses System32 in a 64-bit host' {
        if ([IntPtr]::Size -ne 8) { Set-ItResult -Skipped -Because 'needs a 64-bit host'; return }
        $env:windir = 'C:\Windows'
        Get-BeretNativePath 'manage-bde.exe' | Should -Be 'C:\Windows\System32\manage-bde.exe'
    }

    It 'uses Sysnative in a 32-bit host on 64-bit Windows' {
        if ([IntPtr]::Size -ne 4) { Set-ItResult -Skipped -Because 'needs a 32-bit host'; return }
        $env:windir = 'C:\Windows'
        $env:PROCESSOR_ARCHITEW6432 = 'AMD64'
        Get-BeretNativePath 'manage-bde.exe' | Should -Be 'C:\Windows\Sysnative\manage-bde.exe'
    }
}
