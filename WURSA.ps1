<#
.SYNOPSIS
    Windows Update, Repair, & System Alignment (W.U.R.S.A.) v2.1
    Developed by Steve the Killer | Updated: 2026-06-15
.DESCRIPTION
    Enforces all essential and optional OS patches, OEM driver updates, and third-party
    app upgrades via Chocolatey. Skips apps that are currently in use to avoid
    disrupting the active user, and self-installs Chocolatey if not present. Performs
    unattended Windows feature upgrades via an ISO-based in-place upgrade (setup.exe)
    dispatched to a detached SYSTEM scheduled task, so the upgrade survives RMM /
    LiveConnect session disconnects. Includes BitLocker suspension, URL pre-flight,
    and partial download detection. Reboot is always deferred to the caller.
.NOTES
    Parameters:
      -InplaceUpgrade  Auto-confirms the feature upgrade prompt. Safe for unattended/RMM use.
                       Dispatches the ISO-based setup.exe upgrade to a detached SYSTEM
                       scheduled task; the WURSA run returns immediately and the upgrade
                       continues independently. Progress/result is written to
                       C:\Windows\Temp\25H2IPU\ipu_status.txt. BitLocker is suspended.
      -No3rdParty      Skips the Chocolatey third-party app update pass entirely.
      -NoUpgrade       Skips the feature version check and upgrade prompt entirely.

    Exit Codes:
      0    - Completed successfully, no reboot required
      3010 - Completed successfully, reboot required
      1    - Script terminated due to an unhandled error
      10   - IPU hard driver/compat block (0xC1900101)
      11   - IPU app/driver compat block (0xC1900208)
      12   - IPU machine does not meet minimum requirements (0xC1900200)

    Note: When the feature upgrade is dispatched it runs detached as SYSTEM, so its
    result (including compat blocks 10/11/12 and the 3010 reboot signal) is written
    to ipu_status.txt rather than returned as this script's exit code. Poll that file
    from the RMM to determine when to reboot.
#>
param(
    [switch]$InplaceUpgrade,   # Auto-confirm the feature upgrade prompt
    [switch]$No3rdParty,       # Skip Chocolatey / third-party app updates
    [switch]$NoUpgrade         # Skip the feature upgrade check entirely (region 5)
)

$_ver    = "| v2.1"

# Define the latest known Windows release
$LatestVersion = "25H2"

#region 0 - Pre-Flight & Helpers
$script:ExitCode = 0
trap { Write-Host "[!] Fatal: $($_.Exception.Message)" -ForegroundColor Red; exit 1 }
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit
}
# 1. Force session bypass to clear environment hurdles
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force
# 2. Re-register core paths to bypass the LiveConnect temp-path bug
$env:PSModulePath = [Environment]::GetEnvironmentVariable("PSModulePath","Machine") + ";" + [Environment]::GetEnvironmentVariable("PSModulePath","User")
# UI Helpers
$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($Success) {
        $savedTop = [Console]::CursorTop
        if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
            [Console]::SetCursorPosition(0, $script:StepRow)
            $combined = if ($CustomInfo) { "$($script:StepMsg) $CustomInfo" } else { $script:StepMsg }
            $_pad = " " * [math]::Max(1, $script:Width - $combined.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
            if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
            Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
            [Console]::SetCursorPosition(0, $savedTop)
        }
        $script:StepRow = -1
    } else {
        $script:StepRow = [Console]::CursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $InfoCol
    }
}

# Baseline Data for UI
$OS = Get-CimInstance Win32_OperatingSystem

Clear-Host
$script:Width  = 85

$LineCol   = "DarkCyan"
$MainCol   = "Cyan"
$WarnCol   = "DarkYellow"
$BorderCol = "Red"
$ArtCol    = "White"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"

function Write-SubResult {
    param([string]$Tag, [string]$Color)
    $_pad = " " * [math]::Max(1, $script:Width - [Console]::CursorLeft - $Tag.Length)
    Write-Host "$_pad$Tag" -ForegroundColor $Color
}

function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "-" * $Width
    }
    $colors = @(
        [ConsoleColor]$BorderCol,
        [ConsoleColor]$ArtCol,
        [ConsoleColor]$MainCol,
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
$_art1 = "╦ ╦ ╦ ╦ ╦═╗ ╔═╗ ╔═╗ "
$_art2 = "║║║ ║ ║ ╠╦╝ ╚═╗ ╠═╣ "
$_art3 = "╚╩╝ ╚═╝ ╩╚═ ╚═╝ ╩ ╩ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "WINDOWS UPDATE, REPAIR, & SYSTEM ALIGNMENT"
$_tpad  = " " * ($_fillW - $_title.Length)
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

Write-Host "[>] Device Name      : " -NoNewline -ForegroundColor $LineCol
Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $AccentCol
Write-Host "[>] Operating System : " -NoNewline -ForegroundColor $LineCol
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
Write-Host "$($OS.Caption) $WinVer (Build $($OS.BuildNumber))" -ForegroundColor $AccentCol
Write-HLine -Style dashed

#region 0.5 - System Restore Point
# ============================================================================
Write-StepUpdate "[1/5] Creating System Restore Point..."
try {
    # Bypass the 24-hour cooldown Windows enforces between restore points
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
    # Ensure System Protection is enabled on C: (often disabled on managed endpoints)
    Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue | Out-Null
    Checkpoint-Computer -Description "WURSA Pre-Update $($env:COMPUTERNAME) $(Get-Date -Format 'yyyy-MM-dd HH:mm')" -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop
    Write-StepUpdate -Success
} catch {
    Write-Host "      [!] Restore point failed: $($_.Exception.Message)" -ForegroundColor $WarnCol
}
#endregion

#region 1 - Service Registration
# ============================================================================
Write-StepUpdate "[2/5] Opting-in to Microsoft Product Updates..."
try {
    $ServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
    $ServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null
    Write-StepUpdate -Success
    Start-Sleep -Seconds 3
} catch {
    Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
}
#endregion

#region 2 - Discovery
# ============================================================================
Write-StepUpdate "[3/5] Scanning for Drivers & OS Patches..."
# Direct API search to bypass the 'remoteIpNoProxy' crash
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
try {
    $s1 = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0").Updates
    $s2 = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver' and IsHidden=0").Updates
    $UpdateList = New-Object -ComObject Microsoft.Update.UpdateColl
    foreach ($u in $s1) { $UpdateList.Add($u) | Out-Null }
    foreach ($u in $s2) { $UpdateList.Add($u) | Out-Null }
} catch {
    Write-StepUpdate -Success -CustomInfo "(Scan Error)"
    Write-Host "      [!] WU search failed: $($_.Exception.Message)" -ForegroundColor Red
    $UpdateList = $null
}
if (-not $UpdateList -or $UpdateList.Count -eq 0) {
    Write-StepUpdate -Success -CustomInfo "(System Up to Date)"
    Write-Host "[4/5] Windows Update Installation..." -NoNewline -ForegroundColor $DimCol
    Write-SubResult "[SKIPPED]" $WarnCol
} else {
    Write-StepUpdate -Success -CustomInfo "($($UpdateList.Count) Found)"
#endregion

#region 3 - Installation & Progress
# ========================================================================
    $ProgressPreference = 'SilentlyContinue'
    Write-StepUpdate "[4/5] Installing Windows Updates..."
    $Counter = 0
    foreach ($Update in $UpdateList) {
        $Counter++

        # Word-wrap the title at word boundaries so long names don't blow the
        # cursor-reposition math. All lines use the same indent; [SUCCESS]/[FAILED]
        # is appended to the last line only.
        $_uPrefix = "      > Deploying: "
        $_uIndent = " " * $_uPrefix.Length
        $_uAvail  = $script:Width - "[SUCCESS]".Length - $_uPrefix.Length
        $_uLines  = @()
        $_rem     = $Update.Title
        $_first   = $true
        while ($_rem.Length -gt 0) {
            $_pfx = if ($_first) { $_uPrefix } else { $_uIndent }
            if ($_rem.Length -le $_uAvail) {
                $_uLines += "$_pfx$_rem"
                break
            }
            $_chunk = $_rem.Substring(0, $_uAvail)
            $_split = $_chunk.LastIndexOf(' ')
            if ($_split -le 0) { $_split = $_uAvail }   # no space found — hard break
            $_uLines += "$_pfx$($_rem.Substring(0, $_split).TrimEnd())"
            $_rem    = $_rem.Substring($_split).TrimStart()
            $_first  = $false
        }

        $_uRow = [Console]::CursorTop
        foreach ($_line in $_uLines) { Write-Host $_line -ForegroundColor $InfoCol }

        try {
            # Bypassing Install-WindowsUpdate to clear path errors
            $UpdatesToInstall = New-Object -ComObject Microsoft.Update.UpdateColl
            $UpdatesToInstall.Add($Update) | Out-Null
            $Downloader = $UpdateSession.CreateUpdateDownloader()
            $Downloader.Updates = $UpdatesToInstall
            $null = $Downloader.Download()
            $Installer = $UpdateSession.CreateUpdateInstaller()
            $Installer.Updates = $UpdatesToInstall
            $Installer.AllowSourcePrompts = $false
            $InstallResult = $Installer.Install()
            $_savedTop = [Console]::CursorTop
            [Console]::SetCursorPosition(0, $_uRow)
            for ($_li = 0; $_li -lt $_uLines.Count; $_li++) {
                if ($_li -eq $_uLines.Count - 1) {
                    $_pad = " " * [math]::Max(1, $script:Width - $_uLines[$_li].Length - "[SUCCESS]".Length)
                    Write-Host "$($_uLines[$_li])$_pad" -NoNewline -ForegroundColor $DimCol
                    Write-Host "[SUCCESS]" -ForegroundColor $OkCol
                } else {
                    Write-Host $_uLines[$_li] -ForegroundColor $DimCol
                }
            }
            [Console]::SetCursorPosition(0, $_savedTop)
        } catch {
            $_savedTop = [Console]::CursorTop
            [Console]::SetCursorPosition(0, $_uRow)
            for ($_li = 0; $_li -lt $_uLines.Count; $_li++) {
                if ($_li -eq $_uLines.Count - 1) {
                    $_pad = " " * [math]::Max(1, $script:Width - $_uLines[$_li].Length - "[FAILED]".Length)
                    Write-Host "$($_uLines[$_li])$_pad" -NoNewline -ForegroundColor $DimCol
                    Write-Host "[FAILED]" -ForegroundColor Red
                } else {
                    Write-Host $_uLines[$_li] -ForegroundColor $DimCol
                }
            }
            [Console]::SetCursorPosition(0, $_savedTop)
        }
    }
    Write-Progress -Activity "W.U.R.S.A.: Deploying Updates" -Completed
    Write-StepUpdate "[4/5] Windows Updates" -Success
}
#endregion

#region 4 - Third-Party Software Updates
# ============================================================================
$ThirdParty = @(
    @{ Name = "Google Chrome";   ChocoID = "googlechrome --install-arguments='--system-level' --ignore-checksums"; Path = "C:\Program Files\Google\Chrome\Application\chrome.exe";                    Process = "chrome" },
    @{ Name = "Microsoft Edge";  ChocoID = "microsoft-edge";   Path = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";                     Process = "msedge" },
    @{ Name = "Mozilla Firefox"; ChocoID = "firefox";          Path = "C:\Program Files\Mozilla Firefox\firefox.exe";                                      Process = "firefox" },
    @{ Name = "Brave";           ChocoID = "brave";            Path = "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe";                Process = "brave" },
    @{ Name = "Adobe Acrobat";   ChocoID = "adobereader";      Path = @("C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe","C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe","C:\Program Files\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe","C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe"); Process = @("Acrobat","AcroRd32") },
    @{ Name = "Zoom";            ChocoID = "zoom --install-arguments='/quiet /norestart'"; Path = @("C:\Program Files\Zoom\bin\Zoom.exe","C:\Program Files (x86)\Zoom\bin\Zoom.exe","$env:APPDATA\Zoom\bin\Zoom.exe","$env:LOCALAPPDATA\Zoom\bin\Zoom.exe"); Process = "Zoom" },
    @{ Name = "Microsoft Teams"; ChocoID = "microsoft-teams";  Path = @("C:\Program Files\Microsoft\Teams\current\Teams.exe","C:\Program Files (x86)\Microsoft\Teams\current\Teams.exe","C:\Program Files\WindowsApps\MSTeams_*\ms-teams.exe"); Process = @("Teams","ms-teams") },
    @{ Name = "Webex";           ChocoID = "webex";            Path = "C:\Program Files\Webex\bin\CiscoCollabHost.exe";                                    Process = "CiscoCollabHost" },
    @{ Name = "Slack";           ChocoID = "slack";            Path = "C:\Program Files\Slack\slack.exe";                                                  Process = "slack" },
    @{ Name = "RingCentral";     ChocoID = "ringcentral";      Path = "C:\Program Files\RingCentral\RingCentral.exe";                                      Process = "RingCentral" },
    @{ Name = "Notepad++";       ChocoID = "notepadplusplus";  Path = "C:\Program Files\Notepad++\notepad++.exe";                                          Process = "notepad++" },
    @{ Name = "VLC";             ChocoID = "vlc";              Path = "C:\Program Files\VideoLAN\VLC\vlc.exe";                                             Process = "vlc" },
    @{ Name = "7-Zip";           ChocoID = "7zip";             Path = "C:\Program Files\7-Zip\7z.exe";                                                     Process = "7zFM" }
#    @{ Name = "KillerPDF";       ChocoID = "killerpdf";        Path = @("C:\Program Files\KillerPDF\KillerPDF.exe","$env:LOCALAPPDATA\Programs\KillerPDF\KillerPDF.exe"); Process = "KillerPDF",
#    @{ Name = "KillerScan";       ChocoID = "killerscan";        Path = @("C:\Program Files\KillerScan\KillerScan.exe","$env:LOCALAPPDATA\Programs\KillerScan\KillerScan.exe"); Process = "KillerPDF" }
)
$ChocoAvailable = Get-Command choco -ErrorAction SilentlyContinue
if (-not $ChocoAvailable) {
    Write-Host "`n[!] Chocolatey not found - Installing..." -ForegroundColor Yellow
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        $ChocoInstallDir = "$env:ProgramData\chocolatey"
        $ZipPath = "$env:TEMP\chocolatey.zip"
        $ExtractPath = "$env:TEMP\chocoInstall"
        # Download the zip directly
        (New-Object System.Net.WebClient).DownloadFile(
            "https://community.chocolatey.org/api/v2/package/chocolatey",
            $ZipPath
        )
        # Extract using pure .NET - bypasses Microsoft.PowerShell.Archive entirely
        # Clear any leftover extraction folder from a previous failed attempt
        if (Test-Path $ExtractPath) { Remove-Item $ExtractPath -Recurse -Force }
        Add-Type -AssemblyName System.IO.Compression.FileSystem
        [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $ExtractPath)
        # Run the embedded install script directly
        $ChocoInstallScript = Get-ChildItem "$ExtractPath" -Recurse -Filter "chocolateyInstall.ps1" | Select-Object -First 1
        if ($ChocoInstallScript) {
            $env:ChocolateyInstall = $ChocoInstallDir
            & $ChocoInstallScript.FullName *>&1 | Out-Null
        }
        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        $ChocoAvailable = Get-Command choco -ErrorAction SilentlyContinue
        if ($ChocoAvailable) {
            Write-Host "      Chocolatey installed successfully." -ForegroundColor Green
        } else {
            Write-Host "      [!] Chocolatey install failed - skipping third-party updates." -ForegroundColor Red
        }
    } catch {
        Write-Host "      [!] Chocolatey install failed: $($_.Exception.Message)" -ForegroundColor Red
    }
}
if ($ChocoAvailable -and -not $No3rdParty) {
    Write-StepUpdate "[5/5] Updating Installed Third-Party Software..."
    # Update Chocolatey itself first
    Write-Host "      > " -NoNewline -ForegroundColor $DimCol
    Write-Host "Chocolatey: " -NoNewline -ForegroundColor $DimCol
    $chocoSelfOut = choco upgrade chocolatey -y --no-progress 2>&1
    $chocoSelfMatch = $chocoSelfOut | Select-String -Pattern 'upgraded (\d+)/'
    if ($chocoSelfMatch -and [int]$chocoSelfMatch.Matches[0].Groups[1].Value -gt 0) {
        Write-SubResult "[UPDATED]" Green
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
    } else {
        Write-SubResult "[ALREADY UPDATED]" Cyan
    }
    foreach ($App in $ThirdParty) {
        # Resolve first valid path - supports arrays and wildcard paths (e.g. WindowsApps\MSTeams_*)
        $_appPath = $null
        foreach ($_candidate in @($App.Path)) {
            if ($_candidate -match '\*') {
                $_resolved = Get-Item $_candidate -ErrorAction SilentlyContinue | Select-Object -First 1
                if ($_resolved) { $_appPath = $_resolved.FullName; break }
            } elseif (Test-Path $_candidate) {
                $_appPath = $_candidate; break
            }
        }
        $_procs = @($App.Process)
        if (-not $_appPath) {
            Write-Host "      > " -NoNewline -ForegroundColor $DimCol
            Write-Host "$($App.Name):" -NoNewline -ForegroundColor $DimCol
            Write-SubResult "[NOT INSTALLED]" DarkGray
        } else {
            $IsRunning = $_procs | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
            if ($IsRunning) {
                Write-Host "      > " -NoNewline -ForegroundColor $DimCol
                Write-Host "$($App.Name):" -NoNewline -ForegroundColor $DimCol
                Write-SubResult "[IN USE - SKIPPED]" Yellow
            } else {
                Write-Host "      > " -NoNewline -ForegroundColor Gray
                Write-Host "$($App.Name): " -NoNewline -ForegroundColor White
                $chocoOut = choco upgrade $App.ChocoID -y --no-progress 2>&1
                if ($App.Name -eq "Google Chrome") {
                    $GUpdate = "C:\Program Files (x86)\Google\Update\GoogleUpdate.exe"
                    if (Test-Path $GUpdate) {
                        & $GUpdate /ua /installsource scheduler 2>&1 | Out-Null
                    }
                }
                $upgradeMatch = $chocoOut | Select-String -Pattern 'upgraded (\d+)/'
                $upgradeCount = if ($upgradeMatch) { [int]$upgradeMatch.Matches[0].Groups[1].Value } else { 0 }
                if ($upgradeCount -gt 0) {
                    Write-SubResult "[UPDATED]" Green
                } else {
                    Write-SubResult "[ALREADY UPDATED]" Cyan
                }
            }
        }
    }
    Write-StepUpdate "[5/5] Third-Party Updates" -Success
}
#endregion

#region 5 - Post-Update Version Check & Optional In-Place Upgrade
# ============================================================================
# ISO config - hosted on Cloudflare R2 behind a stable custom domain
# (iso.killertools.net) so the link does not expire. To target a different
# feature build, swap the object in the killer-isos bucket and update the URL.
$IPU_IsoUrl     = "https://iso.killertools.net/Win11_25H2_x64.iso"
$IPU_IsoSizeGB  = 7.9
$IPU_WorkDir    = "C:\Windows\Temp\25H2IPU"
$IPU_IsoPath    = Join-Path $IPU_WorkDir "Win11_25H2_x64.iso"
$IPU_LogDir     = Join-Path $IPU_WorkDir "SetupLogs"
$IPU_SetupLog   = Join-Path $IPU_WorkDir "setup_exit.log"
$IPU_StatusFile = Join-Path $IPU_WorkDir "ipu_status.txt"
$IPU_RunnerPath = Join-Path $IPU_WorkDir "Invoke-IPU.ps1"
$IPU_TaskName   = "WURSA-25H2-IPU"
$IPU_IsoSha256  = "66B7B4B71763ED6F9B2CE29326ED9284544DA6F5283D00329921540C01AAAEEA"

Write-HLine -Style dashed
Write-Host "[>] Checking Windows Feature Update Level..." -ForegroundColor $LineCol
$InstalledVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
Write-Host "      Installed Version : $InstalledVersion" -ForegroundColor Yellow
Write-Host "      Latest Version    : $LatestVersion"    -ForegroundColor Yellow
if ($NoUpgrade) {
    Write-Host "      [-NoUpgrade] Feature upgrade check skipped." -ForegroundColor $DimCol
} elseif ($InstalledVersion -ne $LatestVersion) {
    Write-Host "      > Feature update available." -ForegroundColor $WarnCol
    # Battery safety
    $OnBattery = $false
    try {
        $Battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue
        if ($Battery -and $Battery.BatteryStatus -ne 2) { $OnBattery = $true }
    } catch {}

    $proceed = $false
    if ($OnBattery) {
        Write-Host "      [!] Upgrade skipped: device is running on battery power." -ForegroundColor Yellow
    } elseif ($InplaceUpgrade) {
        Write-Host "      [-InplaceUpgrade] Auto-dispatching detached ISO-based upgrade." -ForegroundColor $DimCol
        $proceed = $true
    } else {
        # Flush buffered keystrokes; guarded so redirected (RMM) consoles do not throw
        try { while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null } } catch {}
        try {
            Write-Host "Would you like to perform an in-place upgrade to $($LatestVersion)? (Y/N): " -NoNewline -ForegroundColor $WarnCol
            $key    = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $choice = $key.Character
            if ($choice -ne "`n" -and $choice -ne "`r") { Write-Host $choice -NoNewline }
            Write-Host ""
            if ($choice -in @('Y','y')) { $proceed = $true }
        } catch {
            Write-Host "      [i] Non-interactive - defaulting to NO. Use -InplaceUpgrade for unattended runs." -ForegroundColor $DimCol
        }
    }

    if (-not $OnBattery -and $proceed) {
        Write-Host "[>] Dispatching detached In-Place Upgrade (survives session disconnect)..." -ForegroundColor $LineCol

        if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }

        # --- Build the standalone runner the scheduled task executes as SYSTEM ---
        # Config preamble (double-quoted here-string: values are injected now).
        $RunnerConfig = @"
`$IPU_IsoUrl     = '$IPU_IsoUrl'
`$IPU_IsoSizeGB  = $IPU_IsoSizeGB
`$IPU_WorkDir    = '$IPU_WorkDir'
`$IPU_IsoPath    = '$IPU_IsoPath'
`$IPU_LogDir     = '$IPU_LogDir'
`$IPU_SetupLog   = '$IPU_SetupLog'
`$IPU_StatusFile = '$IPU_StatusFile'
`$IPU_TaskName   = '$IPU_TaskName'
`$IPU_IsoSha256  = '$IPU_IsoSha256'
"@

        # Runner body (single-quoted here-string: written verbatim, runs later).
        $RunnerBody = @'
$ErrorActionPreference = "Continue"
function Set-Status { param([string]$Text) try { $Text | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force } catch {} }
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }
if (-not (Test-Path $IPU_LogDir))  { New-Item -ItemType Directory -Path $IPU_LogDir  -Force | Out-Null }
Set-Status "RUNNING $stamp"
Start-Transcript -Path (Join-Path $IPU_LogDir "ipu_runner.log") -Append -Force | Out-Null

# BitLocker suspension - 4 reboots covers SafeOS + first boot + second boot + buffer
try {
    $BLStatus = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
    if ($BLStatus -and $BLStatus.ProtectionStatus -eq 'On') {
        Suspend-BitLocker -MountPoint "C:" -RebootCount 4 -ErrorAction Stop
        Write-Output "BitLocker suspended."
    }
} catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }

$cancel = $false

# ISO download with partial-file detection
$NeedDownload = $true
if (Test-Path $IPU_IsoPath) {
    $ExistingGB = [math]::Round((Get-Item $IPU_IsoPath).Length / 1GB, 2)
    if ($ExistingGB -ge $IPU_IsoSizeGB) { $NeedDownload = $false }
    else { Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue }
}

if ($NeedDownload) {
    try {
        $HeadResponse = Invoke-WebRequest -Uri $IPU_IsoUrl -Method Head -UseBasicParsing -TimeoutSec 30
        Write-Output "URL OK (HTTP $($HeadResponse.StatusCode))."
    } catch {
        Write-Output "ISO URL not accessible: $_"
        Set-Status "FAILED URL_UNREACHABLE $stamp"
        $cancel = $true
    }
}

if (-not $cancel -and $NeedDownload) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $IPU_IsoUrl -OutFile $IPU_IsoPath -UseBasicParsing
    } catch {
        Write-Output "ISO download failed: $_"
        if (Test-Path $IPU_IsoPath) { Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue }
        Set-Status "FAILED DOWNLOAD $stamp"
        $cancel = $true
    }
    if (-not $cancel) {
        if (-not (Test-Path $IPU_IsoPath)) { Set-Status "FAILED NO_ISO $stamp"; $cancel = $true }
        else {
            $DownloadedGB = [math]::Round((Get-Item $IPU_IsoPath).Length / 1GB, 2)
            if ($DownloadedGB -lt $IPU_IsoSizeGB) {
                Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue
                Set-Status "FAILED PARTIAL_DOWNLOAD $stamp"
                $cancel = $true
            }
        }
    }
}

if (-not $cancel) {
    $actualHash = (Get-FileHash $IPU_IsoPath -Algorithm SHA256).Hash
    if ($actualHash -ne $IPU_IsoSha256) {
        Write-Output "ISO hash mismatch. Expected $IPU_IsoSha256 got $actualHash"
        Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue
        Set-Status "FAILED HASH_MISMATCH $stamp"
        $cancel = $true
    } else { Write-Output "ISO hash verified." }
}

if (-not $cancel) {
    $MountedDrive = $null
    try {
        $MountResult  = Mount-DiskImage -ImagePath $IPU_IsoPath -PassThru
        $MountedDrive = ($MountResult | Get-Volume).DriveLetter
    } catch { Write-Output "Could not mount ISO: $_"; Set-Status "FAILED MOUNT $stamp"; $cancel = $true }

    if (-not $cancel) {
        $SetupExe = "${MountedDrive}:\setup.exe"
        if (-not (Test-Path $SetupExe)) {
            Dismount-DiskImage -ImagePath $IPU_IsoPath -ErrorAction SilentlyContinue
            Set-Status "FAILED NO_SETUP $stamp"
            $cancel = $true
        }
    }

    if (-not $cancel) {
        $SetupArgs = "/auto upgrade /quiet /compat ignorewarning /DynamicUpdate disable /showoobe None /Telemetry Disable /EULA Accept /noreboot /Copylogs `"$IPU_LogDir`""
        $IPU_ExitCode = -1
        try {
            $proc = Start-Process -FilePath $SetupExe -ArgumentList $SetupArgs -Wait -PassThru
            $IPU_ExitCode = $proc.ExitCode
        } catch { Write-Output "setup.exe failed to launch: $_" }
        "setup.exe exit code: $IPU_ExitCode" | Out-File -FilePath $IPU_SetupLog -Encoding ASCII
        Dismount-DiskImage -ImagePath $IPU_IsoPath -ErrorAction SilentlyContinue

        $done = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        switch ($IPU_ExitCode) {
            0           { Set-Status "REBOOT_REQUIRED 3010 $done" }
            3010        { Set-Status "REBOOT_REQUIRED 3010 $done" }
            -1047526908 { Set-Status "REBOOT_REQUIRED 3010 $done" }
            -1047527167 { Set-Status "BLOCK_HARD_COMPAT 0xC1900101 $done" }
            -1047526904 { Set-Status "BLOCK_APP_DRIVER 0xC1900208 $done" }
            -1047526912 { Set-Status "BLOCK_MIN_REQ 0xC1900200 $done" }
            default     { Set-Status "UNEXPECTED $IPU_ExitCode $done" }
        }
    }
}

Stop-Transcript | Out-Null
# Remove the one-shot task so it does not linger
schtasks.exe /Delete /TN "$IPU_TaskName" /F | Out-Null
'@

        $RunnerConfig + "`r`n" + $RunnerBody | Out-File -FilePath $IPU_RunnerPath -Encoding ASCII -Force

        # --- Register + start the detached SYSTEM task ---
        $null = ('DISPATCHED ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
        $TR = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File $IPU_RunnerPath"
        $createOut = schtasks.exe /Create /TN $IPU_TaskName /TR $TR /SC ONCE /ST 00:00 /RU SYSTEM /RL HIGHEST /F 2>&1
        if ($LASTEXITCODE -eq 0) {
            $runOut = schtasks.exe /Run /TN $IPU_TaskName 2>&1
            if ($LASTEXITCODE -eq 0) {
                Write-Host "      > Upgrade dispatched as SYSTEM. It continues after you disconnect." -ForegroundColor Green
                Write-Host "      > Status file : $IPU_StatusFile" -ForegroundColor $DimCol
                Write-Host "      > Setup logs  : $IPU_LogDir" -ForegroundColor $DimCol
                Write-Host "      > A reboot will be required once the detached upgrade completes." -ForegroundColor $WarnCol
            } else {
                Write-Host "      [!] Task created but failed to start: $runOut" -ForegroundColor Red
            }
        } else {
            Write-Host "      [!] Failed to register scheduled task: $createOut" -ForegroundColor Red
        }
    }
} else {
    Write-Host "      System is already on the latest feature update." -ForegroundColor Green
}
#endregion


#region 6 - Finalization
# ============================================================================
# Checking reboot status via direct API to bypass Get-WURebootStatus crash
$RebootPending = if ($InstallResult) { $InstallResult.RebootRequired } else { $false }
if (-not $RebootPending) {
    $RebootPending = $null -ne (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue)
}

Write-HLine -Style dashed
if ($RebootPending) {
    Write-Host "[!] STATUS: REBOOT REQUIRED" -ForegroundColor Red
} else {
    Write-Host "STATUS: SYSTEM CURRENT" -ForegroundColor Green
}
$script:ExitCode = if ($RebootPending) { 3010 } else { 0 }
# Footer
$_sfx    = "█"
$_ftrW   = $_artW + 1
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  WINDOWS UPDATE SEQUENCE COMPLETE"
$_fpad   = " " * ($_ffillW - $_footer.Length - $_ver.Length)
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
#endregion

exit $script:ExitCode