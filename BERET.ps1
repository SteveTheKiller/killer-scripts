<#
.SYNOPSIS
    BitLocker Encryption, Recovery, & Escrow Tool (BERET) v6.0
    Developed by SteveTheKiller | Updated: 2026-03-12
.DESCRIPTION
    Interactive BitLocker lifecycle manager that prompts to choose
    between FIPS (AES-256) and Standard (AES-128) mode. Initializes
    TPM, enforces compliance-gated encryption, generates recovery keys,
    and escrows them to AD, Entra ID, or MSA depending on join state.
#>

#region 0 - UI Initialization
# ------------------------------------------------------------------------------
Clear-Host
$script:Width    = 85

$LineCol   = "DarkRed"
$MainCol   = "DarkYellow"
$BorderCol = "Red"
$ArtCol    = "White"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$OkCol        = "Green"
$InfoCol      = "Cyan"
$MainTextCol  = "Yellow"
function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "━" * $Width
    }
    $colors = @(
        [ConsoleColor]$BorderCol,
        [ConsoleColor]$ArtCol,
        [ConsoleColor]$AccentCol,
        [ConsoleColor]$DimCol
    )
    $useConsole = $true
    try { $saved = [Console]::ForegroundColor } catch { $useConsole = $false }
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') {
            $fg = [ConsoleColor]$DimCol
        } else {
            $fg = $colors[$i % $colors.Count]
            $i++
        }
        if ($useConsole) {
            [Console]::ForegroundColor = $fg
            [Console]::Write($char)
        } else {
            Write-Host $char -NoNewline -ForegroundColor $fg
        }
    }
    if ($useConsole) {
        [Console]::ForegroundColor = $saved
        [Console]::WriteLine()
    } else {
        Write-Host ""
    }
}

# Header
$_pfx  = "█  "
$_art1 = "╔╗  ╔═╗ ╦═╗ ╔═╗ ╔╦╗ "
$_art2 = "╠╩╗ ║╣  ╠╦╝ ║╣   ║  "
$_art3 = "╚═╝ ╚═╝ ╩╚═ ╚═╝  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "BITLOCKER ENCRYPTION, RECOVERY, & ESCROW TOOL"
$_tpad  = " " * ($_fillW - $_title.Length)
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $MainCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $MainCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $MainCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
function Backup-To-AD {
    param($KeyID, $Drive)
    $_msg = "Escrowing key to Active Directory..."
    Write-Host $_msg -ForegroundColor $InfoCol -NoNewline
    try {
        Backup-BitLockerKeyProtector -MountPoint $Drive -KeyProtectorId $KeyID -ErrorAction Stop | Out-Null
        $_pad = " " * [math]::Max(1, $script:Width - $_msg.Length - "[SUCCESS]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_msg -ForegroundColor $DimCol -NoNewline; Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
    } catch {
        $_pad = " " * [math]::Max(1, $script:Width - $_msg.Length - "[RETRYING]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_msg -ForegroundColor $DimCol -NoNewline; Write-Host "$_pad[RETRYING]" -ForegroundColor $AccentCol
        $_msg2 = "  Retrying via SYSTEM scheduled task..."
        Write-Host $_msg2 -ForegroundColor $InfoCol -NoNewline
        try {
            $TaskName = "BitLocker_AD_Escrow"
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-Command `"Backup-BitLockerKeyProtector -MountPoint '$Drive' -KeyProtectorId '$KeyID' -ErrorAction Stop`""
            $Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
            Register-ScheduledTask -TaskName $TaskName -Action $Action -Principal $Principal -Force | Out-Null
            Start-ScheduledTask -TaskName $TaskName
            Start-Sleep -Seconds 8
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
            $_pad2 = " " * [math]::Max(1, $script:Width - $_msg2.Length - "[SUCCESS]".Length)
            Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
            Write-Host $_msg2 -ForegroundColor $DimCol -NoNewline; Write-Host "$_pad2[SUCCESS]" -ForegroundColor $OkCol
        } catch {
            $_pad2 = " " * [math]::Max(1, $script:Width - $_msg2.Length - "[FAILED]".Length)
            Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
            Write-Host $_msg2 -ForegroundColor $DimCol -NoNewline; Write-Host "$_pad2[FAILED]" -ForegroundColor $BorderCol
            Write-Host "  $($_.Exception.Message)" -ForegroundColor $AccentCol
        }
    }
}
function Backup-To-Entra {
    param($KeyID, $Drive)
    $_msg = "Escrowing key to Microsoft Entra ID..."
    Write-Host $_msg -ForegroundColor $InfoCol -NoNewline
    try {
        BackupToAAD-BitLockerKeyProtector -MountPoint $Drive -KeyProtectorId $KeyID -ErrorAction Stop | Out-Null
        $_pad = " " * [math]::Max(1, $script:Width - $_msg.Length - "[SUCCESS]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_msg -ForegroundColor $DimCol -NoNewline; Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
    } catch {
        $_pad = " " * [math]::Max(1, $script:Width - $_msg.Length - "[FAILED]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_msg -ForegroundColor $DimCol -NoNewline; Write-Host "$_pad[FAILED]" -ForegroundColor $BorderCol
        Write-Host "  $($_.Exception.Message)" -ForegroundColor $AccentCol
    }
}
# --- Show-RecoveryInfo with Lint Suppression ---
function Show-RecoveryInfo {
    param(
        [string]$Drive = "C:",
        [string]$KeyID,
        [string]$RecoveryKey,
        [string]$RealName
    )
    # FIX: Using $script: scope ensures the function can see the UI lines
    Write-HLine -Style dashed
    $_label = "BITLOCKER RECOVERY INFO"; $_padL = " " * [math]::Floor(($script:Width - $_label.Length) / 2); $_padR = " " * ($script:Width - $_label.Length - $_padL.Length)
    Write-Host "$_padL$_label$_padR" -NoNewline -ForegroundColor $ArtCol -BackgroundColor DarkGreen
    Write-Host ""
    Write-Host " DEVICE NAME : " -NoNewline -ForegroundColor $DimCol; Write-Host $RealName -ForegroundColor $ArtCol
    Write-Host " KEY ID      : " -NoNewline -ForegroundColor $DimCol; Write-Host $KeyID -ForegroundColor $AccentCol
    Write-Host " PASSWORD    : " -NoNewline -ForegroundColor $DimCol; Write-Host $RecoveryKey -ForegroundColor $ArtCol
    Write-HLine -Style dashed
}
#endregion

#region 1 - Pre-Flight Checks
# =================================================================================
# --- FIPS Selection ---
Write-Host "Does this system require FIPS 140-2 compliance? " -ForegroundColor $BorderCol -NoNewline
Write-Host "[Y/N]  Esc = quit: " -ForegroundColor $DimCol -NoNewline
try {
    $FipsKey = [Console]::ReadKey($true)
    if ($FipsKey.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit 0 }
    Write-Host $FipsKey.KeyChar -NoNewline
    $FipsPrompt = $FipsKey.KeyChar.ToString()
} catch {
    $FipsPrompt = (Read-Host).Trim()
}
if ($FipsPrompt -eq "Y" -or $FipsPrompt -eq "y") {
    $EncryptionMethod = "Aes256"
    $ModeLabel = "FIPS"
    $EnforceFipsReg = $true
    Write-Host ("`r" + (" " * $script:Width) + "`r") -NoNewline
    Write-Host "MODE: " -ForegroundColor $InfoCol -NoNewline; Write-Host "FIPS COMPLIANCE (AES-256)" -ForegroundColor $MainTextCol
} else {
    $EncryptionMethod = "Aes128"
    $ModeLabel = "STANDARD"
    $EnforceFipsReg = $false
    Write-Host ("`r" + (" " * $script:Width) + "`r") -NoNewline
    Write-Host "MODE: " -ForegroundColor $InfoCol -NoNewline; Write-Host "STANDARD EDITION (AES-128)" -ForegroundColor $MainTextCol
}
# --- OS Edition Gatekeeper ---
$OSCaption = (Get-CimInstance Win32_OperatingSystem).Caption
if ($OSCaption -match "Home") {
    Write-Host "`n![ ERROR ]!: Windows Home Edition Detected ($OSCaption)." -ForegroundColor $BorderCol
    Write-Host "BitLocker management requires Pro, Business, or Enterprise. Upgrade required." -ForegroundColor $AccentCol
    Exit 1
}
# --- TPM Health & Auto-Initialization ---
$TPM = Get-Tpm
if (-not $TPM.TpmReady) {
    Write-Host "`n[!] TPM is NOT READY. Attempting self-healing..." -ForegroundColor $AccentCol
    $Vendor = (Get-CimInstance -Class Win32_ComputerSystem).Manufacturer
    if ($Vendor -match "Dell" -and (Get-Module -ListAvailable -Name DellBIOSProvider)) {
        Write-Host "Dell Hardware detected. Attempting BIOS-level TPM activation..." -ForegroundColor $InfoCol
        try {
            Set-Item -Path DellSmbios:\Security\TpmSecurity "Enabled" -ErrorAction SilentlyContinue
            Set-Item -Path DellSmbios:\Security\TpmActivation "Enabled" -ErrorAction SilentlyContinue
        } catch {
            Write-Host "[!] Dell BIOS Provider failed to set flags." -ForegroundColor $DimCol
        }
    }

    try {
        $InitResult = Initialize-Tpm -AllowClear -AllowPhysicalPresence -ErrorAction Stop
        if ($InitResult.RestartRequired) {
            Write-Host "![ RESTART REQUIRED ]!: TPM ownership initiated. Please reboot." -ForegroundColor $LineCol
            Exit 0
        }
    } catch {
        Write-Host "![ CRITICAL ]!: TPM could not be initialized: $($_.Exception.Message)" -ForegroundColor $BorderCol
    }
} else {
    Write-Host "TPM Health: " -ForegroundColor $InfoCol -NoNewline; Write-Host "Ready" -ForegroundColor $OkCol
}
# --- Policy Refresh ---
$_gpMsg = "Triggering Policy Update..."
Write-Host $_gpMsg -ForegroundColor $InfoCol -NoNewline
$GPProc = Start-Process gpupdate -ArgumentList "/target:computer /force" -PassThru -WindowStyle Hidden
$GPProc | Wait-Process -Timeout 20 -ErrorAction SilentlyContinue
if ($GPProc.HasExited -eq $false) {
    Stop-Process -Id $GPProc.Id -Force
    $_gpStatus = "[TIMEOUT]"; $_gpStatusCol = $AccentCol
} else {
    $_gpStatus = "[SUCCESS]"; $_gpStatusCol = $OkCol
}
$_gpPad = " " * [math]::Max(1, $script:Width - $_gpMsg.Length - $_gpStatus.Length)
Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
Write-Host $_gpMsg -ForegroundColor $DimCol -NoNewline
Write-Host "$_gpPad$_gpStatus" -ForegroundColor $_gpStatusCol
if ($_gpStatus -eq "[TIMEOUT]") {
    Write-Host "  Policy update timed out. Proceeding with local/cached settings." -ForegroundColor $AccentCol
}
Start-Sleep -Seconds 3
if ($EnforceFipsReg) {
    Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy" -Name "Enabled" -Value 1 -Type DWORD -Force -ErrorAction SilentlyContinue
}
$DriveLetter = "$env:SystemDrive"
$DeviceName = (Get-CimInstance Win32_ComputerSystem).Name
$_blMsg = "Checking BitLocker status on $DriveLetter..."
Write-Host $_blMsg -ForegroundColor $InfoCol -NoNewline
$BLV = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction Stop
$RecoveryProtector = $BLV.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'RecoveryPassword'}
$IsCompliant = ($BLV.EncryptionMethod -match $EncryptionMethod)
if ($BLV.VolumeStatus -ne 'FullyDecrypted') {
    $NeedsDecryption = $false
    if (-not $IsCompliant) {
        $_blPad = " " * [math]::Max(1, $script:Width - $_blMsg.Length - "[NON-COMPLIANT]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_blMsg -ForegroundColor $DimCol -NoNewline; Write-Host "$_blPad[NON-COMPLIANT]" -ForegroundColor $AccentCol
        Write-Host "  $ModeLabel requires $EncryptionMethod - found $($BLV.EncryptionMethod). Starting decryption..." -ForegroundColor $LineCol
        $NeedsDecryption = $true
    } elseif (-not $RecoveryProtector) {
        $_blPad = " " * [math]::Max(1, $script:Width - $_blMsg.Length - "[NO RECOVERY KEY]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_blMsg -ForegroundColor $DimCol -NoNewline; Write-Host "$_blPad[NO RECOVERY KEY]" -ForegroundColor $LineCol
        Write-Host "  Drive encrypted but recovery key missing. Starting decryption..." -ForegroundColor $LineCol
        $NeedsDecryption = $true
    } else {
        $ExplicitKeyID = $RecoveryProtector[0].KeyProtectorId
        $RecoveryKey = $RecoveryProtector[0].RecoveryPassword
        $_blPad = " " * [math]::Max(1, $script:Width - $_blMsg.Length - "[COMPLIANT]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_blMsg -ForegroundColor $DimCol -NoNewline; Write-Host "$_blPad[COMPLIANT]" -ForegroundColor $OkCol
        Show-RecoveryInfo -Drive $DriveLetter -KeyID $ExplicitKeyID -RecoveryKey $RecoveryKey -RealName $DeviceName
    }
}

if ($NeedsDecryption) {
    try {
        $AutoUnlockDrives = Get-BitLockerVolume | Where-Object { $_.AutoUnlockEnabled -eq $true }
        foreach ($Volume in $AutoUnlockDrives) {
            Write-Host "Disabling Auto-Unlock on $($Volume.MountPoint) to allow OS decryption..." -ForegroundColor $AccentCol
            Disable-BitLockerAutoUnlock -MountPoint $Volume.MountPoint
        }
        Disable-BitLocker -MountPoint $DriveLetter | Out-Null
        Write-Host "Decryption in progress..." -ForegroundColor $MainCol -NoNewline
        $Timeout = (Get-Date).AddMinutes(30)
        do {
            Start-Sleep -Seconds 10
            $CurrentBLV = Get-BitLockerVolume -MountPoint $DriveLetter
            $CurrentStatus = $CurrentBLV.VolumeStatus
            $Percent = $CurrentBLV.EncryptionPercentage
            $_decLine = "Decryption in progress...  $Percent%"
            Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
            Write-Host $_decLine -ForegroundColor $MainCol -NoNewline
            if ((Get-Date) -gt $Timeout) { Write-Host ""; Write-Warning "Decryption loop timed out."; break }
        } while ($CurrentStatus -ne 'FullyDecrypted')
        $_decDone = "Decryption in progress..."
        $_decPad = " " * [math]::Max(1, $script:Width - $_decDone.Length - "[DONE]".Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host $_decDone -ForegroundColor $DimCol -NoNewline; Write-Host "$_decPad[DONE]" -ForegroundColor $OkCol
        $ExplicitKeyID = $null
    }
    catch {
        Write-Error -Message "ERROR: Failed to decrypt volume: $($_.Exception.Message)"
        exit 1
    }
}
#endregion

#region 2: Key Generation & Encryption
# =================================================================================
if (-not $ExplicitKeyID) {
    Write-Host "Enabling BitLocker ($EncryptionMethod)..." -ForegroundColor $InfoCol -NoNewline
    try {
        $BLV_Check = Get-BitLockerVolume -MountPoint $DriveLetter
        if (-not ($BLV_Check.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'Tpm'})) {
            Add-BitLockerKeyProtector -MountPoint $DriveLetter -TpmProtector | Out-Null
        }
        Add-BitLockerKeyProtector -MountPoint $DriveLetter -RecoveryPasswordProtector -WarningAction SilentlyContinue | Out-Null
        $NewBLV = Get-BitLockerVolume -MountPoint $DriveLetter
        $TargetProtector = $NewBLV.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'RecoveryPassword'} | Select-Object -First 1
        $ExplicitKeyID = $TargetProtector.KeyProtectorId
        $RecoveryKey = $TargetProtector.RecoveryPassword
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Show-RecoveryInfo -Drive $DriveLetter -KeyID $ExplicitKeyID -RecoveryKey $RecoveryKey -RealName $DeviceName
        $SysInfo = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $IsVM = $SysInfo.Model -match "Virtual|QEMU|VMware"
        if ($IsVM) {
            manage-bde -on $DriveLetter -EncryptionMethod $EncryptionMethod -UsedSpaceOnly:0 -SkipHardwareTest | Out-Null
        } else {
            $Result = manage-bde -on $DriveLetter -EncryptionMethod $EncryptionMethod -UsedSpaceOnly -SkipHardwareTest 2>&1
            if ($Result -match "0x803100a5") {
                manage-bde -on $DriveLetter -EncryptionMethod $EncryptionMethod -UsedSpaceOnly:0 -SkipHardwareTest | Out-Null
            }
        }
    }
    catch {
        Write-Error "Failed to enable: $($_.Exception.Message)"
    }
}
#endregion

#region 3: Hybrid Backup Logic
# =================================================================================
if ($ExplicitKeyID) {
    $SysInfo = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
    $IsDomainJoined = $SysInfo.PartOfDomain
    $Dsreg = dsregcmd /status | Out-String
    $IsEntraJoined = $Dsreg -match "AzureAdJoined : YES"
    $IsMSALinked   = $Dsreg -match "MsaAccount : YES"
    $BackupPerformed = $false
    if ($IsDomainJoined) {
        Backup-To-AD -KeyID $ExplicitKeyID -Drive $DriveLetter
        $BackupPerformed = $true
    }
    if ($IsEntraJoined) {
        try {
            $null = [System.Net.Dns]::GetHostAddresses("enterpriseregistration.windows.net")
            Backup-To-Entra -KeyID $ExplicitKeyID -Drive $DriveLetter
            $BackupPerformed = $true
        } catch {
            Write-Host "[!] Entra Backup failed: Network Error." -ForegroundColor $BorderCol
        }
    }
    if ($IsMSALinked) {
        Write-Host "[INFO] Personal Microsoft Account detected." -ForegroundColor $InfoCol
        Backup-To-Entra -KeyID $ExplicitKeyID -Drive $DriveLetter
        $BackupPerformed = $true
    }
    if (-not $BackupPerformed) {
        Write-Host "[!] No Domain, Entra, or MSA link detected. Key exists only locally and on-screen." -ForegroundColor $BorderCol
    }
}
#endregion

#region 4: Final Status
# =================================================================================
Write-Host "Verifying encryption startup..." -ForegroundColor $DimCol -NoNewline
$RetryCount = 0
do {
    Start-Sleep -Seconds 3
    Write-Host "." -ForegroundColor $DimCol -NoNewline
    $FinalStatus = Get-BitLockerVolume -MountPoint $DriveLetter
    $RetryCount++
} while (($null -eq $FinalStatus.EncryptionMethod -or $FinalStatus.EncryptionMethod -eq "None") -and $RetryCount -lt 15)
Write-Host ("`r" + (" " * 60) + "`r") -NoNewline
Write-Host "Final Encryption Method: " -ForegroundColor $InfoCol -NoNewline
Write-Host $FinalStatus.EncryptionMethod -ForegroundColor $MainTextCol
# Footer
$_sfx   = "█"
$_ftr1 = " ╔╗  ╔═╗ ╦═╗ ╔═╗ ╔╦╗ "
$_ftr2 = " ╠╩╗ ║╣  ╠╦╝ ║╣   ║  "
$_ftr3 = " ╚═╝ ╚═╝ ╩╚═ ╚═╝  ╩  "
$_ftrW = [Math]::Max($_ftr1.Length, [Math]::Max($_ftr2.Length, $_ftr3.Length))
$_ftr1 = $_ftr1.PadRight($_ftrW); $_ftr2 = $_ftr2.PadRight($_ftrW); $_ftr3 = $_ftr3.PadRight($_ftrW)
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  BITLOCKER DEPLOYMENT SEQUENCE COMPLETE"
$_ver    = "| v6.0"
$_fpad   = " " * ($_ffillW - $_footer.Length - $_ver.Length)
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr1 -ForegroundColor $MainCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host $_ftr2 -ForegroundColor $MainCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr3 -ForegroundColor $MainCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Exit 0
#endregion
