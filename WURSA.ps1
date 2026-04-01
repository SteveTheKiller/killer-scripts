<#
.SYNOPSIS
    Windows Update, Repair, & System Alignment (W.U.R.S.A.) v1.7
    Developed by Steve the Killer | Updated: 2026-04-01
.DESCRIPTION
    Enforces all essential and optional OS patches, OEM driver updates, and third-party
    app upgrades via Chocolatey. Skips apps that are currently in use to avoid
    disrupting the active user, and self-installs Chocolatey if not present. Adds a
    reliable, unattended Windows feature upgrade using WindowsUpdateBox with
    ESC-to-cancel and a live heartbeat indicator across all contexts.
.NOTES
    Parameters:
      -InplaceUpgrade  Auto-confirms the feature upgrade prompt. Safe for unattended/RMM use.
      -No3rdParty      Skips the Chocolatey third-party app update pass entirely.
      -NoUpgrade       Skips the feature version check and upgrade prompt entirely.

    Exit Codes:
      0    - Completed successfully, no reboot required
      3010 - Completed successfully, reboot required
      1    - Script terminated due to an unhandled error
#>
param(
    [switch]$InplaceUpgrade,   # Auto-confirm the feature upgrade prompt
    [switch]$No3rdParty,       # Skip Chocolatey / third-party app updates
    [switch]$NoUpgrade         # Skip the feature upgrade check entirely (region 5)
)

$_ver    = "| v1.7"

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

powercfg /change standby-timeout-ac 120
powercfg /change monitor-timeout-ac 120

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
#endregion

#region 1 - Service Registration
# ============================================================================
Write-StepUpdate "[1/4] Opting-in to Microsoft Product Updates..."
try {
    $ServiceManager = New-Object -ComObject Microsoft.Update.ServiceManager
    $ServiceManager.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null
    Write-StepUpdate -Success
} catch {
    Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
}
#endregion

#region 2 - Discovery
# ============================================================================
Write-StepUpdate "[2/4] Scanning for Drivers & OS Patches..."
# Direct API search to bypass the 'remoteIpNoProxy' crash
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
$SearchResult = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0")
$UpdateList = $SearchResult.Updates
if (-not $UpdateList -or $UpdateList.Count -eq 0) {
    Write-StepUpdate -Success -CustomInfo "(System Up to Date)"
    Write-Host "[3/4] Windows Update Installation..." -NoNewline -ForegroundColor $DimCol
    Write-SubResult "[SKIPPED]" $WarnCol
} else {
    Write-StepUpdate -Success -CustomInfo "($($UpdateList.Count) Found)"
#endregion

#region 3 - Installation & Progress
# ========================================================================
    $ProgressPreference = 'SilentlyContinue'
    Write-StepUpdate "[3/4] Installing Windows Updates..."
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
    Write-StepUpdate "[3/4] Windows Updates" -Success
}
#endregion

#region 4 - Third-Party Software Updates
# ============================================================================
$ThirdParty = @(
    @{ Name = "Google Chrome";   ChocoID = "googlechrome --install-arguments='--system-level' --ignore-checksums"; Path = "C:\Program Files\Google\Chrome\Application\chrome.exe";                    Process = "chrome" },
    @{ Name = "Microsoft Edge";  ChocoID = "microsoft-edge";   Path = "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe";                     Process = "msedge" },
    @{ Name = "Mozilla Firefox"; ChocoID = "firefox";          Path = "C:\Program Files\Mozilla Firefox\firefox.exe";                                      Process = "firefox" },
    @{ Name = "Brave";           ChocoID = "brave";            Path = "C:\Program Files\BraveSoftware\Brave-Browser\Application\brave.exe";                Process = "brave" },
    @{ Name = "Adobe Acrobat";   ChocoID = "adobereader";      Path = "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe";               Process = "AcroRd32" },
    @{ Name = "Zoom";            ChocoID = "zoom --install-arguments='/quiet /norestart'"; Path = "C:\Program Files\Zoom\bin\Zoom.exe";                     Process = "Zoom" },
    @{ Name = "Microsoft Teams"; ChocoID = "microsoft-teams";  Path = "C:\Program Files\Microsoft\Teams\current\Teams.exe";                               Process = "Teams" },
    @{ Name = "Webex";           ChocoID = "webex";            Path = "C:\Program Files\Webex\bin\CiscoCollabHost.exe";                                    Process = "CiscoCollabHost" },
    @{ Name = "Slack";           ChocoID = "slack";            Path = "C:\Program Files\Slack\slack.exe";                                                  Process = "slack" },
    @{ Name = "RingCentral";     ChocoID = "ringcentral";      Path = "C:\Program Files\RingCentral\RingCentral.exe";                                      Process = "RingCentral" },
    @{ Name = "Notepad++";       ChocoID = "notepadplusplus";  Path = "C:\Program Files\Notepad++\notepad++.exe";                                          Process = "notepad++" },
    @{ Name = "VLC";             ChocoID = "vlc";              Path = "C:\Program Files\VideoLAN\VLC\vlc.exe";                                             Process = "vlc" },
    @{ Name = "7-Zip";           ChocoID = "7zip";             Path = "C:\Program Files\7-Zip\7z.exe";                                                     Process = "7zFM" }
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
    Write-StepUpdate "[4/4] Updating Installed Third-Party Software..."
    foreach ($App in $ThirdParty) {
        if (-not (Test-Path $App.Path)) {
            Write-Host "      > " -NoNewline -ForegroundColor $DimCol
            Write-Host "$($App.Name):" -NoNewline -ForegroundColor $DimCol
            Write-SubResult "[NOT INSTALLED]" DarkGray
        } else {
            $IsRunning = Get-Process -Name $App.Process -ErrorAction SilentlyContinue
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
    Write-StepUpdate "[4/4] Third-Party Updates" -Success
}
#endregion

#region 5 - Post-Update Version Check & Optional In-Place Upgrade
# ============================================================================
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
    # Heartbeat indicator - defined here so it's available for both attach and new-launch paths
    function Show-UpgradeHeartbeat {
        param([string[]]$PathsToWatch, [int]$LauncherPid)
        $e = [char]27
        $upgradeProcs   = @('Windows11InstallationAssistant','Windows10UpgraderApp','WindowsUpdateAssistant','SetupHost','SetupPrep')
        $friendlyNames  = @{ 'Windows11InstallationAssistant' = 'Installation Assistant'; 'Windows10UpgraderApp' = 'Installation Assistant'; 'SetupHost' = 'Windows Setup'; 'SetupPrep' = 'Windows Setup (Prep)' }
        Write-Host -NoNewline "${e}[2K`r      > Upgrade in progress... (Press ESC to detach)" -ForegroundColor Yellow
        $spinner = @('|','/','-','\')
        $tick = 0
        function Get-UpgradeRunning {
            if ($LauncherPid -and (Get-Process -Id $LauncherPid -ErrorAction SilentlyContinue)) { return $true }
            foreach ($n in $upgradeProcs) {
                if (Get-Process -Name $n -ErrorAction SilentlyContinue) { return $true }
            }
            return $false
        }
        $shared = [hashtable]::Synchronized(@{ Escape = $false; SizeMB = 0.0; Label = 'Installation Assistant' })

        # ESC watcher runspace
        $rsEsc = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rsEsc.Open()
        $psEsc = [System.Management.Automation.PowerShell]::Create()
        $psEsc.Runspace = $rsEsc
        $psEsc.AddScript({
            param($shared)
            try {
                while (-not $shared.Escape) {
                    if ([Console]::KeyAvailable) {
                        $k = [Console]::ReadKey($true)
                        if ($k.Key -eq [System.ConsoleKey]::Escape) { $shared.Escape = $true; return }
                    }
                    [System.Threading.Thread]::Sleep(50)
                }
            } catch {}
        }).AddArgument($shared) | Out-Null
        $psEsc.BeginInvoke() | Out-Null

        # Folder size + label watcher runspace (runs independently so it never blocks the spinner)
        $rsSize = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace()
        $rsSize.Open()
        $psSize = [System.Management.Automation.PowerShell]::Create()
        $psSize.Runspace = $rsSize
        $psSize.AddScript({
            param($shared, $PathsToWatch, $upgradeProcs, $friendlyNames)
            while (-not $shared.Escape) {
                $total = 0
                foreach ($p in $PathsToWatch) {
                    if (Test-Path $p) {
                        $s = (Get-ChildItem $p -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
                        if ($s) { $total += $s }
                    }
                }
                $shared.SizeMB = [math]::Round($total / 1MB, 1)
                $active = $upgradeProcs | Where-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
                if ($active) { $shared.Label = if ($friendlyNames[$active]) { $friendlyNames[$active] } else { $active } }
                Start-Sleep -Seconds 4
            }
        }).AddArgument($shared).AddArgument($PathsToWatch).AddArgument($upgradeProcs).AddArgument($friendlyNames) | Out-Null
        $psSize.BeginInvoke() | Out-Null

        # Lightweight spinner loop - just reads shared state, never blocks
        while (Get-UpgradeRunning -and -not $shared.Escape) {
            $sizeStr = if ($shared.SizeMB -gt 0) { "  |  $($shared.SizeMB) MB staged" } else { "" }
            Write-Host -NoNewline ("${e}[2K`r        $($spinner[$tick % $spinner.Length])  $($shared.Label)$sizeStr") -ForegroundColor DarkGray
            $tick++
            Start-Sleep -Milliseconds 200
        }
        $didEscape = $shared.Escape
        $shared.Escape = $true
        $psEsc.Stop(); $psEsc.Dispose(); $rsEsc.Close(); $rsEsc.Dispose()
        $psSize.Stop(); $psSize.Dispose(); $rsSize.Close(); $rsSize.Dispose()
        if ($didEscape) {
            Write-Host "${e}[2K`r      > Detached. Upgrade continues in background." -ForegroundColor Yellow
            return
        }
        $rebootPending = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue
        if ($rebootPending) {
            Write-Host "${e}[2K`r      > Staging complete — reboot required to apply upgrade." -ForegroundColor Yellow
        } else {
            Write-Host "${e}[2K`r      > Upgrade process has exited." -ForegroundColor Green
        }
    }
    # Check if an upgrade process is already running from a previous session
    $upgradeProcsCheck = @('Windows11InstallationAssistant','Windows10UpgraderApp','WindowsUpdateAssistant','SetupHost','SetupPrep')
    $alreadyRunning = $upgradeProcsCheck | Where-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
    if ($alreadyRunning) {
        $alreadyPid = (Get-Process -Name $alreadyRunning -ErrorAction SilentlyContinue | Select-Object -First 1).Id
        Write-Host "      [i] Upgrade already in progress ($alreadyRunning, PID $alreadyPid) - attaching..." -ForegroundColor $DimCol
        Show-UpgradeHeartbeat -PathsToWatch @("C:\`$WINDOWS.~BT", "C:\ESD") -LauncherPid $alreadyPid
        $proceed = $false
    } elseif ($OnBattery) {
        Write-Host "      [!] Upgrade skipped: device is running on battery power." -ForegroundColor Yellow
        $proceed = $false
    } elseif ($InplaceUpgrade) {
        Write-Host "      [-InplaceUpgrade] Auto-proceeding with upgrade." -ForegroundColor $DimCol
        $proceed = $true
    } else {
        $proceed = $false
        # Flush any buffered keystrokes from earlier steps
        while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null }
        try {
            Write-Host "Would you like to perform an in-place upgrade to $($LatestVersion)? (Y/N): " -NoNewline -ForegroundColor $WarnCol
            $key    = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $choice = $key.Character
            if ($choice -ne "`n" -and $choice -ne "`r") { Write-Host $choice -NoNewline }
            Write-Host -NoNewline "`r"
            if ($choice -in @('Y','y')) { $proceed = $true }
            $esc = [char]27
            Write-Host -NoNewline "${esc}[2K`r"
        } catch {
            Write-Host "      [i] Non-interactive - defaulting to NO." -ForegroundColor $DimCol
        }
    }
    if ($proceed) {
        # Preparing Windows Upgrade
        Write-Host "[>] Preparing Windows Upgrade..." -ForegroundColor $LineCol
        while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null }
        Write-Host "     > Upgrade will begin in 10 seconds. " -NoNewline -ForegroundColor Yellow
        Write-Host "(Press ESC to cancel...)" -ForegroundColor DarkGray
        $cancel = $false
        [Console]::TreatControlCAsInput = $true
        for ($i = 10; $i -gt 0; $i--) {
            for ($t = 0; $t -lt 10; $t++) {
                Start-Sleep -Milliseconds 100
                if ([Console]::KeyAvailable) {
                    $key = [Console]::ReadKey($true)
                    if ($key.Key -eq "Escape") {
                        $cancel = $true
                        break
                    }
                }
            }
            if ($cancel) { break }
            Write-Host -NoNewline "`r     $i... " -ForegroundColor DarkGray
        }
        $esc = [char]27
        Write-Host -NoNewline "${esc}[2K`r${esc}[1A${esc}[2K`r"
        [Console]::TreatControlCAsInput = $false
        while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null }
        if ($cancel) {
            Write-Host "     [!] Upgrade canceled by user." -ForegroundColor Yellow
        } else {
        function Invoke-FeatureUpdate {
            $e             = [char]27
            $AssistantUrl  = "https://go.microsoft.com/fwlink/?linkid=2171764"
            $AssistantPath = "$env:TEMP\Windows11InstallationAssistant.exe"
            try {
                Write-Host -NoNewline "${e}[2K`r      > Downloading Windows 11 Installation Assistant..." -ForegroundColor Gray
                Invoke-WebRequest -Uri $AssistantUrl -OutFile $AssistantPath -UseBasicParsing

                if (-not (Test-Path $AssistantPath)) {
                    Write-Host "${e}[2K`r      [!] Download failed - file not found at $AssistantPath" -ForegroundColor Red
                    return
                }
                $sizeMB = [math]::Round((Get-Item $AssistantPath).Length / 1MB, 1)
                if ($sizeMB -lt 1) {
                    Write-Host "${e}[2K`r      [!] Download suspect - only $sizeMB MB at $AssistantPath" -ForegroundColor Red
                    return
                }
                Write-Host -NoNewline "${e}[2K`r      > Launching Installation Assistant ($sizeMB MB)..." -ForegroundColor Gray
                $proc = Start-Process -FilePath $AssistantPath -ArgumentList "/QuietInstall /SkipEULA /Auto Upgrade" -PassThru
                Start-Sleep -Milliseconds 3000
                $childProcs = @('Windows10UpgraderApp','WindowsUpdateAssistant','SetupHost','SetupPrep')
                if (-not (Get-Process -Id $proc.Id -ErrorAction SilentlyContinue)) {
                    # Launcher exited - expected if it hands off to a child process
                    if ($proc.ExitCode -ne 0) {
                        Write-Host "${e}[2K`r      [!] Installation Assistant failed (code: $($proc.ExitCode))" -ForegroundColor Red
                        return
                    }
                    # Wait up to 20 seconds for the child upgrade process to appear
                    Write-Host -NoNewline "${e}[2K`r      > Waiting for upgrade process..." -ForegroundColor Gray
                    $found = $null
                    for ($w = 0; $w -lt 20 -and -not $found; $w++) {
                        $found = $childProcs | Where-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
                        if (-not $found) { Start-Sleep -Seconds 1 }
                    }
                    if (-not $found) {
                        Write-Host "${e}[2K`r      [!] Launcher exited cleanly but no upgrade process appeared" -ForegroundColor Yellow
                        return
                    }
                    $foundPid = (Get-Process -Name $found -ErrorAction SilentlyContinue | Select-Object -First 1).Id
                    Write-Host "${e}[2K`r      > Handed off to $found (PID $foundPid)" -ForegroundColor DarkGray
                    Show-UpgradeHeartbeat -PathsToWatch @("C:\`$WINDOWS.~BT", "C:\ESD") -LauncherPid $foundPid
                } else {
                    Write-Host "${e}[2K`r      > Installation Assistant running (PID $($proc.Id))" -ForegroundColor DarkGray
                    Show-UpgradeHeartbeat -PathsToWatch @("C:\`$WINDOWS.~BT", "C:\ESD") -LauncherPid $proc.Id
                }
                Write-StepUpdate "[X] In-Place Upgrade" -Success
            }
            catch {
                Write-Host "${e}[2K`r      [!] Failed to start upgrade: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
        Invoke-FeatureUpdate
        } # end else (not canceled)
    } else {
        Write-Host "      Upgrade skipped." -ForegroundColor Yellow
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