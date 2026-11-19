<#
.SYNOPSIS
    BitLocker Encryption, Recovery, & Escrow Tool (BERET) v6.0
    Developed by SteveTheKiller | Updated: 2026-03-12
.DESCRIPTION
    Interactive BitLocker lifecycle manager that prompts to choose
    between FIPS (XTS-AES-256) and Standard (XTS-AES-128) mode. Initializes
    TPM, enforces compliance-gated encryption, generates recovery keys,
    and escrows them to AD, Entra ID, or MSA depending on join state.
#>

#region 0 - UI Initialization
# ------------------------------------------------------------------------------
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

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

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

function Show-BeretBanner {
    try { Clear-Host } catch { }
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
}

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
function Show-RecoveryInfo {
    param(
        [string]$Drive = "C:",
        [string]$KeyID,
        [string]$RecoveryKey,
        [string]$RealName
    )
    Write-HLine -Style dashed
    $_label = "BITLOCKER RECOVERY INFO"; $_padL = " " * [math]::Floor(($script:Width - $_label.Length) / 2); $_padR = " " * ($script:Width - $_label.Length - $_padL.Length)
    Write-Host "$_padL$_label$_padR" -NoNewline -ForegroundColor $ArtCol -BackgroundColor DarkGreen
    Write-Host ""
    Write-Host " DEVICE NAME : " -NoNewline -ForegroundColor $DimCol; Write-Host $RealName -ForegroundColor $ArtCol
    Write-Host " KEY ID      : " -NoNewline -ForegroundColor $DimCol; Write-Host $KeyID -ForegroundColor $AccentCol
    Write-Host " PASSWORD    : " -NoNewline -ForegroundColor $DimCol; Write-Host $RecoveryKey -ForegroundColor $ArtCol
    Write-HLine -Style dashed
}

function Exit-BeretRun {
    # One exit point. The final line goes to the success stream so an RMM captures it, and it
    # always agrees with the exit code.
    param([int]$Code, [string]$Text)
    Write-Output "BERET RESULT: exit code $Code - $Text"
    Exit-KillerScript $Code
}

function Get-BeretMode {
    # Returns 'FIPS', 'STANDARD', 'NOCONSOLE' or $null when the operator quit. A preset $BeretMode
    # wins and never prompts. With no console and no preset there is nobody to ask, so the caller
    # fails closed. Never exits from here: output inside an assigned call would be swallowed.
    $preset = [string]$global:BeretMode
    if ($preset -match '^(?i:fips)$') { return 'FIPS' }
    if ($preset -match '^(?i:standard)$') { return 'STANDARD' }
    if (-not (Test-BeretConsole)) { return 'NOCONSOLE' }
    $answer = Read-BeretFipsChoice
    if ($answer -eq 'ESC') { return $null }
    if ($answer -match '^(?i:y)') { return 'FIPS' }
    return 'STANDARD'
}

function Read-BeretFipsChoice {
    # Returns 'Y', 'N' (or any other typed text), or 'ESC'.
    Write-Host "Does this system require FIPS 140-2 compliance? " -ForegroundColor $BorderCol -NoNewline
    Write-Host "[Y/N]  Esc = quit: " -ForegroundColor $DimCol -NoNewline
    try {
        $FipsKey = [Console]::ReadKey($true)
        if ($FipsKey.Key -eq [ConsoleKey]::Escape) { return 'ESC' }
        Write-Host $FipsKey.KeyChar -NoNewline
        return $FipsKey.KeyChar.ToString()
    } catch {
        return (Read-Host).Trim()
    }
}

function Get-BeretNativePath {
    # A 32-bit host on 64-bit Windows sees System32 redirected to SysWOW64, which has no
    # manage-bde or dsregcmd. Sysnative reaches the real 64-bit folder.
    param([string]$Name)
    $dir = 'System32'
    if ([IntPtr]::Size -eq 4 -and $env:PROCESSOR_ARCHITEW6432) { $dir = 'Sysnative' }
    Join-Path (Join-Path $env:windir $dir) $Name
}

function Get-BeretDsreg {
    $ErrorActionPreference = 'Continue'
    & (Get-BeretNativePath 'dsregcmd.exe') /status 2>&1 | Out-String
}

function Invoke-BeretManageBde {
    # The only place manage-bde runs. Returns the exit code and the text, so callers can check both.
    param([string[]]$Arguments)
    $ErrorActionPreference = 'Continue'
    $text = & (Get-BeretNativePath 'manage-bde.exe') @Arguments 2>&1 | Out-String
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = $text }
}

function Get-BeretNow {
    Get-Date
}

function Test-BeretElevated {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        return (New-Object Security.Principal.WindowsPrincipal($id)).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    } catch { return $false }
}

function Test-BeretConsole {
    # True only when a person can answer a key prompt. A redirected or absent console cannot.
    try {
        return ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected)
    } catch { return $false }
}
#endregion

function Write-BeretStatusLine {
    param([string]$Message, [string]$Status, [string]$Color)
    $pad = " " * [math]::Max(1, $script:Width - $Message.Length - $Status.Length)
    Write-Host $Message -ForegroundColor $DimCol -NoNewline
    Write-Host "$pad$Status" -ForegroundColor $Color
}

function Show-BeretFooter {
    param([string]$Text)
    $_sfx   = "█"
    $_ftr1 = " ╔╗  ╔═╗ ╦═╗ ╔═╗ ╔╦╗ "
    $_ftr2 = " ╠╩╗ ║╣  ╠╦╝ ║╣   ║  "
    $_ftr3 = " ╚═╝ ╚═╝ ╩╚═ ╚═╝  ╩  "
    $_ftrW = [Math]::Max($_ftr1.Length, [Math]::Max($_ftr2.Length, $_ftr3.Length))
    $_ftr1 = $_ftr1.PadRight($_ftrW); $_ftr2 = $_ftr2.PadRight($_ftrW); $_ftr3 = $_ftr3.PadRight($_ftrW)
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_ver    = "| v6.0"
    $_fpad   = " " * [Math]::Max(0, $_ffillW - $Text.Length - $_ver.Length)
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr1 -ForegroundColor $MainCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host "$Text$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host $_ftr2 -ForegroundColor $MainCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr3 -ForegroundColor $MainCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
}

function Get-BeretRecoveryProtectors {
    param([string]$Drive)
    try {
        $v = Get-BitLockerVolume -MountPoint $Drive -ErrorAction Stop
        return @($v.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' })
    } catch { return @() }
}

function Show-BeretKeys {
    # Prints every recovery password protector on the volume. The technician needs to see them.
    param([string]$Drive, [string]$DeviceName)
    foreach ($k in @(Get-BeretRecoveryProtectors -Drive $Drive)) {
        Show-RecoveryInfo -Drive $Drive -KeyID $k.KeyProtectorId -RecoveryKey $k.RecoveryPassword -RealName $DeviceName
    }
}

function Invoke-BeretDecrypt {
    # Turns BitLocker off and waits, with a deadline, until the volume is fully decrypted.
    # Never exits. Returns Ok and a plain Error text.
    param([string]$Drive, [int]$TimeoutMinutes = 30)
    $err = $null
    try {
        $AutoUnlockDrives = @(Get-BitLockerVolume | Where-Object { $_.AutoUnlockEnabled -eq $true })
        foreach ($Volume in $AutoUnlockDrives) {
            Write-Host "Disabling Auto-Unlock on $($Volume.MountPoint) to allow OS decryption..." -ForegroundColor $AccentCol
            Disable-BitLockerAutoUnlock -MountPoint $Volume.MountPoint | Out-Null
        }
        Disable-BitLocker -MountPoint $Drive -ErrorAction Stop | Out-Null
        Write-Host "Decryption in progress..." -ForegroundColor $MainCol
        $deadline = (Get-BeretNow).AddMinutes($TimeoutMinutes)
        $done = $false
        while (-not $done) {
            Start-Sleep -Seconds 10
            $cur = Get-BitLockerVolume -MountPoint $Drive -ErrorAction Stop
            Write-Host "Decryption in progress...  $($cur.EncryptionPercentage)%" -ForegroundColor $MainCol
            if ($cur.VolumeStatus -eq 'FullyDecrypted') { $done = $true }
            elseif ((Get-BeretNow) -gt $deadline) { break }
        }
        if (-not $done) {
            $err = "decryption of $Drive did not finish within $TimeoutMinutes minutes and is still running. Let it finish, then run BERET again."
        }
    } catch {
        $err = "failed to decrypt ${Drive}: $($_.Exception.Message)"
    }
    if ($err) { return [pscustomobject]@{ Ok = $false; Error = $err } }
    Write-BeretStatusLine "Decryption in progress..." "[DONE]" $OkCol
    [pscustomobject]@{ Ok = $true; Error = $null }
}

function Invoke-BeretEnable {
    # Adds the TPM and recovery protectors, shows the key, starts encryption with manage-bde and
    # proves it started with the right cipher. Never exits. Returns Ok and a plain Error text.
    param([string]$Drive, [string]$DeviceName, [string]$Method, [string]$BdeMethod, [int]$StartTimeoutSeconds = 120)
    Write-Host "Enabling BitLocker ($Method)..." -ForegroundColor $InfoCol
    $err = $null
    try {
        $before = Get-BitLockerVolume -MountPoint $Drive -ErrorAction Stop
        if (-not ($before.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'Tpm'})) {
            Add-BitLockerKeyProtector -MountPoint $Drive -TpmProtector -ErrorAction Stop | Out-Null
        }
        Add-BitLockerKeyProtector -MountPoint $Drive -RecoveryPasswordProtector -WarningAction SilentlyContinue -ErrorAction Stop | Out-Null
        $after = Get-BitLockerVolume -MountPoint $Drive -ErrorAction Stop
        $rec = @($after.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'RecoveryPassword'})
        if ($rec.Count -eq 0 -or [string]$rec[0].RecoveryPassword -notmatch '^(\d{6}-){7}\d{6}$') {
            throw "the new recovery password could not be read back from $Drive"
        }
    } catch { $err = $_.Exception.Message }
    if ($err) { return [pscustomobject]@{ Ok = $false; Error = $err } }

    # The technician sees the key before encryption starts, so it is never lost if the run dies.
    Show-BeretKeys -Drive $Drive -DeviceName $DeviceName

    $SysInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
    $IsVM = $SysInfo.Model -match "Virtual|QEMU|VMware"
    $bdeBase = @('-on', $Drive, '-EncryptionMethod', $BdeMethod)
    $bdeSkip = @('-SkipHardwareTest')
    if ($IsVM) {
        # Full-volume encryption is manage-bde's default, so the used-space switch is simply left out.
        $Result = Invoke-BeretManageBde -Arguments ($bdeBase + $bdeSkip)
    } else {
        $Result = Invoke-BeretManageBde -Arguments ($bdeBase + @('-UsedSpaceOnly') + $bdeSkip)
        if ($Result.ExitCode -ne 0 -and $Result.Output -match "0x803100a5") {
            $Result = Invoke-BeretManageBde -Arguments ($bdeBase + $bdeSkip)
        }
    }
    if ($Result.ExitCode -ne 0) {
        $line = ($Result.Output -split "`n" | Where-Object { $_.Trim() } | Select-Object -First 1)
        return [pscustomobject]@{ Ok = $false; Error = "manage-bde -on exited $($Result.ExitCode): $line" }
    }

    # manage-bde returned 0, but only the volume itself can say encryption really started.
    $deadline = (Get-BeretNow).AddSeconds($StartTimeoutSeconds)
    $seen = $null
    $started = $false
    do {
        Start-Sleep -Seconds 3
        try { $seen = Get-BitLockerVolume -MountPoint $Drive -ErrorAction Stop } catch { $seen = $null }
        if ($seen -and $seen.VolumeStatus -ne 'FullyDecrypted' -and [string]$seen.EncryptionMethod -notin @('', 'None')) { $started = $true }
    } while (-not $started -and (Get-BeretNow) -lt $deadline)
    if (-not $started) {
        return [pscustomobject]@{ Ok = $false; Error = "encryption of $Drive did not start within $StartTimeoutSeconds seconds. The recovery key above is on the volume." }
    }
    if ([string]$seen.EncryptionMethod -ne $Method) {
        return [pscustomobject]@{ Ok = $false; Error = "encryption started with $($seen.EncryptionMethod) but this mode requires $Method. Check the BitLocker encryption method policy." }
    }
    [pscustomobject]@{ Ok = $true; Error = $null }
}
function Invoke-BeretMain {
    Show-BeretBanner
    if (-not (Test-BeretElevated)) {
        Exit-BeretRun 2 "not elevated. Run as Administrator or SYSTEM. Nothing was changed."
    }

    #region 1 - Pre-Flight Checks
    # =================================================================================
    $BeretModeChoice = Get-BeretMode
    if ($BeretModeChoice -eq 'NOCONSOLE') {
        Exit-BeretRun 2 "no console to ask the FIPS question. Set `$BeretMode = 'FIPS' or 'STANDARD' before running. Nothing was changed."
    }
    if ($null -eq $BeretModeChoice) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit-BeretRun 0 "cancelled by the operator. Nothing was changed." }
    if ($BeretModeChoice -eq 'FIPS') {
        $EncryptionMethod = "XtsAes256"
        $BdeMethod = "xts_aes256"
        $ModeLabel = "FIPS"
        $EnforceFipsReg = $true
        Write-Host ("`r" + (" " * $script:Width) + "`r") -NoNewline
        Write-Host "MODE: " -ForegroundColor $InfoCol -NoNewline; Write-Host "FIPS COMPLIANCE (XTS-AES-256)" -ForegroundColor $MainTextCol
    } else {
        $EncryptionMethod = "XtsAes128"
        $BdeMethod = "xts_aes128"
        $ModeLabel = "STANDARD"
        $EnforceFipsReg = $false
        Write-Host ("`r" + (" " * $script:Width) + "`r") -NoNewline
        Write-Host "MODE: " -ForegroundColor $InfoCol -NoNewline; Write-Host "STANDARD EDITION (XTS-AES-128)" -ForegroundColor $MainTextCol
    }
    # --- OS Edition Gatekeeper ---
    $OSCaption = (Get-CimInstance Win32_OperatingSystem).Caption
    if ($OSCaption -match "Home") {
        Write-Host "`n![ ERROR ]!: Windows Home Edition Detected ($OSCaption)." -ForegroundColor $BorderCol
        Write-Host "BitLocker management requires Pro, Business, or Enterprise. Upgrade required." -ForegroundColor $AccentCol
        Exit-BeretRun 2 "Windows Home does not support BitLocker management. Nothing was changed."
    }
    # --- TPM Health & Auto-Initialization ---
    $TPM = Get-Tpm
    $tpmRestart = $false
    $tpmFailed = $null
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
                $tpmRestart = $true
            }
        } catch {
            Write-Host "![ CRITICAL ]!: TPM could not be initialized: $($_.Exception.Message)" -ForegroundColor $BorderCol
            $tpmFailed = $_.Exception.Message
        }
        if ($tpmRestart) { Exit-BeretRun 3010 "TPM ownership initiated. Reboot, then run BERET again. Nothing was encrypted." }
        if ($tpmFailed) { Exit-BeretRun 2 "the TPM could not be initialized ($tpmFailed). Nothing was changed." }
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
    $DeviceName = (Get-CimInstance -ClassName Win32_ComputerSystem).Name
    $_blMsg = "Checking BitLocker status on $DriveLetter..."
    $readError = $null
    try { $BLV = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction Stop } catch { $readError = $_.Exception.Message }
    if ($readError -or -not $BLV) {
        Write-BeretStatusLine $_blMsg '[FAILED]' $BorderCol
        Exit-BeretRun 1 "could not read the BitLocker state of $DriveLetter ($readError). Nothing was changed."
    }
    $RecoveryProtector = @($BLV.KeyProtector | Where-Object {$_.KeyProtectorType -eq 'RecoveryPassword'})
    $IsCompliant = ([string]$BLV.EncryptionMethod -eq $EncryptionMethod)
    $NeedsDecryption = $false
    $HaveKey = $false
    if ($BLV.VolumeStatus -ne 'FullyDecrypted') {
        if (-not $IsCompliant) {
            Write-BeretStatusLine $_blMsg '[NON-COMPLIANT]' $AccentCol
            Write-Host "  $ModeLabel requires $EncryptionMethod - found $($BLV.EncryptionMethod). Starting decryption..." -ForegroundColor $LineCol
            $NeedsDecryption = $true
        } elseif ($RecoveryProtector.Count -eq 0) {
            Write-BeretStatusLine $_blMsg '[NO RECOVERY KEY]' $LineCol
            Write-Host "  Drive encrypted but recovery key missing. Starting decryption..." -ForegroundColor $LineCol
            $NeedsDecryption = $true
        } else {
            Write-BeretStatusLine $_blMsg '[COMPLIANT]' $OkCol
            Show-BeretKeys -Drive $DriveLetter -DeviceName $DeviceName
            $HaveKey = $true
        }
    } else {
        Write-BeretStatusLine $_blMsg '[NOT ENCRYPTED]' $AccentCol
    }

    if ($NeedsDecryption) {
        $dec = Invoke-BeretDecrypt -Drive $DriveLetter
        if (-not $dec.Ok) {
            Exit-BeretRun 1 $dec.Error
        }
    }
    #endregion

    #region 2: Key Generation & Encryption
    # =================================================================================
    if (-not $HaveKey) {
        $enable = Invoke-BeretEnable -Drive $DriveLetter -DeviceName $DeviceName -Method $EncryptionMethod -BdeMethod $BdeMethod
        if (-not $enable.Ok) {
            Write-Host "[FAILED] $($enable.Error)" -ForegroundColor $BorderCol
            Exit-BeretRun 1 "BitLocker was not enabled ($($enable.Error))."
        }
    }
    #endregion

    #region 3: Hybrid Backup Logic
    # =================================================================================
    $KeyRows = @(Get-BeretRecoveryProtectors -Drive $DriveLetter)
    $ExplicitKeyID = $null
    if ($KeyRows.Count -gt 0) { $ExplicitKeyID = $KeyRows[0].KeyProtectorId }
    if ($ExplicitKeyID) {
        $SysInfo = Get-CimInstance -ClassName Win32_ComputerSystem -ErrorAction SilentlyContinue
        $IsDomainJoined = $SysInfo.PartOfDomain
        $Dsreg = Get-BeretDsreg
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
    $finalError = $null
    try { $FinalStatus = Get-BitLockerVolume -MountPoint $DriveLetter -ErrorAction Stop } catch { $finalError = $_.Exception.Message }
    if ($finalError -or [string]$FinalStatus.EncryptionMethod -ne $EncryptionMethod) {
        Exit-BeretRun 1 "the final check found $($FinalStatus.EncryptionMethod) on $DriveLetter, expected $EncryptionMethod $finalError"
    }
    Write-Host "Final Encryption Method: " -ForegroundColor $InfoCol -NoNewline
    Write-Host $FinalStatus.EncryptionMethod -ForegroundColor $MainTextCol
    Write-Host "Recovery key(s) for this machine:" -ForegroundColor $InfoCol
    Show-BeretKeys -Drive $DriveLetter -DeviceName $DeviceName
    Show-BeretFooter "  BITLOCKER DEPLOYMENT SEQUENCE COMPLETE"
    Exit-BeretRun 0 "$ModeLabel BitLocker ($EncryptionMethod) is in place on $DriveLetter."
    #endregion
}

if (-not $global:KillerScriptsTestMode) { Invoke-BeretMain }
