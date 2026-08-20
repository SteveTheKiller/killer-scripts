<#
.SYNOPSIS
    Windows Update, Repair, & System Alignment (W.U.R.S.A.) v2.7
    Developed by Steve the Killer | Updated: 2026-08-20
.DESCRIPTION
    Enforces OS patches, OEM driver updates, and Chocolatey third-party app upgrades
    (skips in-use apps). Performs unattended feature upgrades to 25H2, dispatched to a
    detached SYSTEM task so it survives RMM / LiveConnect disconnects: 24H2 takes the
    KB5054156 eKB, Windows 10 and older Windows 11 take the ISO. Reboot is deferred to
    the caller.
.NOTES
    Params: -InplaceUpgrade (auto-confirm, unattended/RMM), -No3rdParty, -NoUpgrade.
    Feature-upgrade result is written to C:\Windows\Temp\25H2IPU\ipu_status.txt; poll it.
    Status: REBOOT_REQUIRED (staged, reboot to finish), DISPATCHED/RUNNING/DOWNLOADING/
    REPAIRING/VERIFY_PENDING (upgrade active), VERIFIED (committed), BLOCK_*/FAILED_*/
    UNEXPECTED (not complete).
    Only REBOOT_REQUIRED should trigger a reboot, always left to the caller.
#>
param(
    [switch]$InplaceUpgrade,   # Auto-confirm the feature upgrade prompt
    [switch]$No3rdParty,       # Skip Chocolatey / third-party app updates
    [switch]$NoUpgrade         # Skip the feature upgrade check entirely (region 5)
)

$_ver    = "| v2.7"

# Define the latest known Windows release
$LatestVersion = "25H2"

# Feature-upgrade state is tracked separately from ordinary Windows Update.
# A surfaced/downloaded feature update is NOT considered staged. Region 5 owns
# the deterministic 25H2 upgrade path unless Windows reports a real reboot-ready
# feature update.
$script:FeatureUpdateStaged = $false
$script:FeatureUpgradeState = "UNKNOWN"

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
$script:Width  = 95

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
# Bypass the 24-hour cooldown Windows enforces between restore points
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
# The checkpoint calls into VSS/SrClient, which can stall or hard-fault the host
# process under RMM and take the whole run down before try/catch can fire. Run it
# in a timeboxed child job so a VSS hang can't kill WURSA; treat it as non-fatal.
$_rpTimeout = 90
$_rpDesc    = "WURSA Pre-Update $($env:COMPUTERNAME) $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
$_rpJob = Start-Job -ScriptBlock {
    param($Desc)
    try {
        $ProgressPreference = 'SilentlyContinue'
        # Ensure System Protection is enabled on C: (often disabled on managed endpoints)
        Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue | Out-Null
        Checkpoint-Computer -Description $Desc -RestorePointType "MODIFY_SETTINGS" -ErrorAction Stop | Out-Null
        "OK"
    } catch {
        "ERR: $($_.Exception.Message)"
    }
} -ArgumentList $_rpDesc
if (Wait-Job -Job $_rpJob -Timeout $_rpTimeout) {
    $_rpResult = Receive-Job -Job $_rpJob
    if ($_rpResult -match '^OK') {
        Write-StepUpdate -Success
    } else {
        Write-Host "      [!] Restore point failed: $($_rpResult -replace '^ERR: ','')" -ForegroundColor $WarnCol
    }
} else {
    Write-Host "      [!] Restore point timed out after ${_rpTimeout}s (VSS stalled); continuing." -ForegroundColor $WarnCol
    Stop-Job -Job $_rpJob -ErrorAction SilentlyContinue
}
Remove-Job -Job $_rpJob -Force -ErrorAction SilentlyContinue
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

#region 1.4 - Open Windows Update policy for the update pass
# ============================================================================
# Some managed endpoints carry NoAutoUpdate=1, which can prevent the direct Windows
# Update API from servicing ordinary updates. Open the policy here, but DO NOT kick
# a USO feature-update scan. Region 5 owns the 25H2 upgrade path. Starting USO here
# can pre-download 25H2 in parallel and race the deterministic ISO/eKB path. Enterprise
# and ARM64 fallback paths explicitly start USO later only when Windows Update is the
# selected upgrade mechanism. Original policy values are recorded for restoration.
$IPU_WorkDir          = "C:\Windows\Temp\25H2IPU"
$IPU_AURestoreFile    = Join-Path $IPU_WorkDir "au_restore.txt"
$IPU_ActiveHoursStart = 7     # WU will not auto-reboot the feature update between these hours (0-23, max 18h span)
$IPU_ActiveHoursEnd   = 19
if (-not $NoUpgrade -and $WinVer -ne $LatestVersion) {
    Write-Host "[>] Opening Windows Update policy for the update pass..." -ForegroundColor $LineCol
    $_polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $_auk  = "$_polk\AU"
    if (-not (Test-Path $_polk)) { New-Item -Path $_polk -Force | Out-Null }
    if (-not (Test-Path $_auk))  { New-Item -Path $_auk  -Force | Out-Null }
    if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }
    # Record originals once, in the same format Restore-WUUpgradeOverrides expects.
    if (-not (Test-Path $IPU_AURestoreFile)) {
        $_noAuto = (Get-ItemProperty -Path $_auk  -Name NoAutoUpdate     -ErrorAction SilentlyContinue).NoAutoUpdate
        $_auOpt  = (Get-ItemProperty -Path $_auk  -Name AUOptions        -ErrorAction SilentlyContinue).AUOptions
        $_setAH  = (Get-ItemProperty -Path $_polk -Name SetActiveHours   -ErrorAction SilentlyContinue).SetActiveHours
        $_ahS    = (Get-ItemProperty -Path $_polk -Name ActiveHoursStart -ErrorAction SilentlyContinue).ActiveHoursStart
        $_ahE    = (Get-ItemProperty -Path $_polk -Name ActiveHoursEnd   -ErrorAction SilentlyContinue).ActiveHoursEnd
        @(
            "NoAutoUpdate=$(if ($null -eq $_noAuto) {'ABSENT'} else {$_noAuto})"
            "AUOptions=$(if ($null -eq $_auOpt) {'ABSENT'} else {$_auOpt})"
            "SetActiveHours=$(if ($null -eq $_setAH) {'ABSENT'} else {$_setAH})"
            "ActiveHoursStart=$(if ($null -eq $_ahS) {'ABSENT'} else {$_ahS})"
            "ActiveHoursEnd=$(if ($null -eq $_ahE) {'ABSENT'} else {$_ahE})"
        ) | Set-Content -Path $IPU_AURestoreFile -Encoding ASCII -Force
    }
    # NoAutoUpdate=0 alone does not release a pending feature update on the re-keyed
    # Enterprise-composition boxes; AUOptions=4 (auto download + schedule install) is
    # what actually moves it. Pin active hours so any WU-driven reboot lands out of hours.
    Set-ItemProperty -Path $_auk  -Name NoAutoUpdate     -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $_auk  -Name AUOptions        -Value 4 -Type DWord -Force
    Set-ItemProperty -Path $_polk -Name SetActiveHours   -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $_polk -Name ActiveHoursStart -Value $IPU_ActiveHoursStart -Type DWord -Force
    Set-ItemProperty -Path $_polk -Name ActiveHoursEnd   -Value $IPU_ActiveHoursEnd   -Type DWord -Force
    Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Write-Host "      WU opened (NoAutoUpdate=0, AUOptions=4); direct API scan follows. 25H2 is reserved for Region 5." -ForegroundColor $DimCol
}
#endregion

#region 2 - Discovery
# ============================================================================
Write-StepUpdate "[3/5] Scanning for Drivers & OS Patches..."
# Resume Windows Update if it is paused. While paused, WU returns nothing, so an
# update run would silently find zero updates and the feature upgrade is never
# offered. Clears the local "Pause updates" values (UX\Settings); leaves any
# policy-driven pause alone.
$_uxk = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
if (Test-Path $_uxk) {
    $_wasPaused = $false
    foreach ($_pn in 'PauseUpdatesExpiryTime','PauseUpdatesStartTime','PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime','PauseQualityUpdatesStartTime','PauseQualityUpdatesEndTime') {
        if ($null -ne (Get-ItemProperty -Path $_uxk -Name $_pn -ErrorAction SilentlyContinue)) {
            Remove-ItemProperty -Path $_uxk -Name $_pn -ErrorAction SilentlyContinue
            $_wasPaused = $true
        }
    }
    if ($_wasPaused) { Write-Host "      [i] Windows Update was paused - resumed." -ForegroundColor $WarnCol }
}
# Direct API search to bypass the 'remoteIpNoProxy' crash
$UpdateSession = New-Object -ComObject Microsoft.Update.Session
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
try {
    $s1 = $UpdateSearcher.Search("IsInstalled=0 and Type='Software' and IsHidden=0").Updates
    $s2 = $UpdateSearcher.Search("IsInstalled=0 and Type='Driver' and IsHidden=0").Updates
    $UpdateList = New-Object -ComObject Microsoft.Update.UpdateColl
    $script:FeatureUpdateOffered = $false
    foreach ($u in $s1) {
        # Do not install a feature update through the generic WU loop. A download or
        # pre-download is not a completed upgrade and used to cause false WU_HANDOFF
        # success. Region 5 handles 25H2 with a deterministic upgrade mechanism.
        if ($u.Title -match [regex]::Escape($LatestVersion) -or $u.Title -match 'Feature update to Windows') {
            $script:FeatureUpdateOffered = $true
            continue
        }
        $UpdateList.Add($u) | Out-Null
    }
    foreach ($u in $s2) { $UpdateList.Add($u) | Out-Null }
    if ($script:FeatureUpdateOffered) {
        Write-Host "      [i] $LatestVersion feature update detected; reserved for the dedicated feature-upgrade path." -ForegroundColor $DimCol
    }
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
            if ($_split -le 0) { $_split = $_uAvail }   # no space found - hard break
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
    @{ Name = "7-Zip";           ChocoID = "7zip";             Path = "C:\Program Files\7-Zip\7z.exe";                                                     Process = "7zFM" },
    @{ Name = "KillerPDF";       ChocoID = "killerpdf";        Path = @("C:\Program Files\KillerPDF\KillerPDF.exe","$env:LOCALAPPDATA\Programs\KillerPDF\KillerPDF.exe"); Process = "KillerPDF" }
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
            Write-Host "      > " -NoNewline -ForegroundColor Gray
            Write-Host "$($App.Name): " -NoNewline -ForegroundColor White
            Write-SubResult "[NOT INSTALLED]" DarkGray
        } else {
            $IsRunning = $_procs | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
            if ($IsRunning) {
                Write-Host "      > " -NoNewline -ForegroundColor Gray
                Write-Host "$($App.Name): " -NoNewline -ForegroundColor White
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
#
# Language-aware: the in-place upgrade requires the ISO language to match the
# running OS UI language, or setup.exe returns 0xC1900204 and the keep files
# path is silently blocked. Read the base install language from the NLS registry
# (reliable under SYSTEM) and pick the matching image: en-GB (0809) gets the
# International ISO, everything else defaults to en-US for the US fleet.
$IPU_InstallLang = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language' -ErrorAction SilentlyContinue).InstallLanguage
switch ($IPU_InstallLang) {
    '0809'  { $IPU_IsoName = "Win11_25H2_x64_en-GB.iso"; $IPU_IsoSha256 = "66B7B4B71763ED6F9B2CE29326ED9284544DA6F5283D00329921540C01AAAEEA" }
    default { $IPU_IsoName = "Win11_25H2_x64_en-US.iso"; $IPU_IsoSha256 = "768984706B909479417B2368438909440F2967FF05C6A9195ED2667254E465E3" }
}
$IPU_IsoUrl     = "https://iso.killertools.net/$IPU_IsoName"
$IPU_IsoSizeGB  = 6.0
# $IPU_WorkDir, $IPU_AURestoreFile, and $IPU_ActiveHoursStart/End are defined earlier
# (region 1.4 WU ungag, which runs before the scan); reused here.
$IPU_IsoPath    = Join-Path $IPU_WorkDir $IPU_IsoName
$IPU_LogDir     = Join-Path $IPU_WorkDir "SetupLogs"
$IPU_SetupLog   = Join-Path $IPU_WorkDir "setup_exit.log"
$IPU_StatusFile = Join-Path $IPU_WorkDir "ipu_status.txt"
$IPU_RunnerPath = Join-Path $IPU_WorkDir "Invoke-IPU.ps1"
$IPU_TaskName   = "WURSA-25H2-IPU"
$IPU_VerifierPath  = Join-Path $IPU_WorkDir "Verify-25H2.ps1"
$IPU_VerifyTaskName = "WURSA-25H2-Verify"

# 24H2 (26100) -> 25H2 enablement package (KB5054156). x64 only; host MSU in killer-isos bucket.
$IPU_EkbName    = "Win11_25H2_eKB_KB5054156_x64.msu"
$IPU_EkbUrl     = "https://iso.killertools.net/$IPU_EkbName"
$IPU_EkbSha256  = "92EDDA7EEAA19B60D15CCDF777556BF0662EE9FEA1DCC9AEC281FCF12068044C"
$IPU_EkbPath    = Join-Path $IPU_WorkDir $IPU_EkbName

# Installation Assistant fallback (online, edition-correct). The local ISO is a
# consumer image with no Enterprise SKU, so an Enterprise-composition box fails the
# ISO match with 0xC1900204. The Assistant pulls the matching edition straight from
# Microsoft. Host the exe alongside the ISO in the killer-isos bucket; the hash pin
# is optional (leave empty to skip).
$IPU_IAUrl      = "https://iso.killertools.net/Windows11InstallationAssistant.exe"
$IPU_IASha256   = ""

function Restore-WUUpgradeOverrides {
    # Runs in the MAIN script (not the detached runner), when the box is found already
    # on the latest build. Undoes Set-WUUpgradeOverrides: each value is put back exactly,
    # or removed if it did not exist before. The runner only ever sets the overrides;
    # the main script is what restores them.
    if (-not (Test-Path $IPU_AURestoreFile)) { return }
    $polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auk  = "$polk\AU"
    $map  = @{}
    foreach ($line in (Get-Content $IPU_AURestoreFile -ErrorAction SilentlyContinue)) {
        if ($line -match '^([^=]+)=(.+)$') { $map[$matches[1]] = $matches[2] }
    }
    $targets = @{ NoAutoUpdate = $auk; AUOptions = $auk; SetActiveHours = $polk; ActiveHoursStart = $polk; ActiveHoursEnd = $polk }
    foreach ($name in @($targets.Keys)) {
        $key = $targets[$name]; $orig = $map[$name]
        if ($null -eq $orig -or $orig -eq 'ABSENT') { Remove-ItemProperty -Path $key -Name $name -ErrorAction SilentlyContinue }
        else { Set-ItemProperty -Path $key -Name $name -Value ([int]$orig) -Type DWord -Force }
    }
    Remove-Item $IPU_AURestoreFile -Force -ErrorAction SilentlyContinue
    Write-Host "      Restored original Windows Update policy values (auto-update and active hours)." -ForegroundColor $DimCol
}

function Test-WUFeatureUpdateStaged {
    # Only a feature update that Windows reports as reboot-required is considered staged.
    # Offered, downloading, and pre-downloaded updates do not suppress the deterministic
    # Region 5 upgrade path. This prevents WURSA from declaring a handoff while the box
    # remains on the old feature release.
    try {
        $_res = $UpdateSearcher.Search("RebootRequired=1").Updates
        foreach ($_u in $_res) {
            if ($_u.Title -match [regex]::Escape($LatestVersion) -or $_u.Title -match 'Feature update to Windows') { return $true }
        }
    } catch {}
    return $false
}

Write-HLine -Style dashed
Write-Host "[>] Checking Windows Feature Update Level..." -ForegroundColor $LineCol
$InstalledVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
Write-Host "      Installed Version : $InstalledVersion" -ForegroundColor Yellow
Write-Host "      Latest Version    : $LatestVersion"    -ForegroundColor Yellow
if ($NoUpgrade) {
    $script:FeatureUpgradeState = "SKIPPED"
    Write-Host "      [-NoUpgrade] Feature upgrade check skipped." -ForegroundColor $DimCol
} elseif ($InstalledVersion -ne $LatestVersion -and (Test-WUFeatureUpdateStaged)) {
    # This is the only Windows Update state that suppresses the dedicated upgrade path.
    # Windows has actually staged the feature update and is asking for a reboot.
    $script:FeatureUpdateStaged = $true
    $script:FeatureUpgradeState = "REBOOT_REQUIRED"
    Write-Host "      > $LatestVersion is staged by Windows Update and requires a reboot." -ForegroundColor $DimCol
    if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }
    ('REBOOT_REQUIRED WU ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
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
        $script:FeatureUpgradeState = "BLOCKED_BATTERY"
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

    if (-not $OnBattery -and -not $proceed) {
        $script:FeatureUpgradeState = "NOT_STARTED"
    }

    if (-not $OnBattery -and $proceed) {
        Write-Host "[>] Dispatching detached In-Place Upgrade (survives session disconnect)..." -ForegroundColor $LineCol

        if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }

        # Build a one-shot post-reboot verifier. It runs as SYSTEM at startup, waits
        # for normal boot to settle, verifies DisplayVersion, and restores the temporary
        # Windows Update policy overrides only after 25H2 is actually committed.
        $VerifierBody = @"
`$ErrorActionPreference = 'SilentlyContinue'
Start-Sleep -Seconds 180
`$cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
`$ver = `$cv.DisplayVersion
`$build = `$cv.CurrentBuild
`$ubr = `$cv.UBR
`$stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
if (`$ver -eq '$LatestVersion') {
    "VERIFIED $LatestVersion `$build.`$ubr `$stamp" | Out-File -FilePath '$IPU_StatusFile' -Encoding ASCII -Force
    if (Test-Path '$IPU_AURestoreFile') {
        `$polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
        `$auk  = "`$polk\AU"
        `$map  = @{}
        foreach (`$line in (Get-Content '$IPU_AURestoreFile' -ErrorAction SilentlyContinue)) {
            if (`$line -match '^([^=]+)=(.+)$') { `$map[`$matches[1]] = `$matches[2] }
        }
        `$targets = @{ NoAutoUpdate = `$auk; AUOptions = `$auk; SetActiveHours = `$polk; ActiveHoursStart = `$polk; ActiveHoursEnd = `$polk }
        foreach (`$name in @(`$targets.Keys)) {
            `$key = `$targets[`$name]; `$orig = `$map[`$name]
            if (`$null -eq `$orig -or `$orig -eq 'ABSENT') { Remove-ItemProperty -Path `$key -Name `$name -ErrorAction SilentlyContinue }
            else { Set-ItemProperty -Path `$key -Name `$name -Value ([int]`$orig) -Type DWord -Force }
        }
        Remove-Item '$IPU_AURestoreFile' -Force -ErrorAction SilentlyContinue
    }
    schtasks.exe /Delete /TN '$IPU_VerifyTaskName' /F | Out-Null
    exit 0
}
`$setupActive = Test-Path 'HKLM:\SYSTEM\Setup\MoSetup\Volatile'
if (`$setupActive) {
    "VERIFY_PENDING VERSION_`$ver BUILD_`$build.`$ubr `$stamp" | Out-File -FilePath '$IPU_StatusFile' -Encoding ASCII -Force
    exit 0
}
"FAILED_POSTREBOOT VERSION_`$ver BUILD_`$build.`$ubr `$stamp" | Out-File -FilePath '$IPU_StatusFile' -Encoding ASCII -Force
schtasks.exe /Delete /TN '$IPU_VerifyTaskName' /F | Out-Null
exit 1
"@
        $VerifierBody | Out-File -FilePath $IPU_VerifierPath -Encoding ASCII -Force

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
`$IPU_VerifierPath  = '$IPU_VerifierPath'
`$IPU_VerifyTaskName = '$IPU_VerifyTaskName'
`$IPU_IsoSha256  = '$IPU_IsoSha256'
`$IPU_IAUrl     = '$IPU_IAUrl'
`$IPU_IASha256  = '$IPU_IASha256'
`$IPU_AURestoreFile    = '$IPU_AURestoreFile'
`$IPU_ActiveHoursStart = $IPU_ActiveHoursStart
`$IPU_ActiveHoursEnd   = $IPU_ActiveHoursEnd
`$IPU_EkbUrl     = '$IPU_EkbUrl'
`$IPU_EkbSha256  = '$IPU_EkbSha256'
`$IPU_EkbPath    = '$IPU_EkbPath'
"@

        # Runner body (single-quoted here-string: written verbatim, runs later).
        $RunnerBody = @'
$ErrorActionPreference = "Continue"
function Set-Status { param([string]$Text) try { $Text | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force } catch {} }
$IPU_IAExe  = Join-Path $IPU_WorkDir "Windows11InstallationAssistant.exe"
$IPU_IAArgs = "/QuietInstall /SkipEULA /auto upgrade /NoReboot"
function Register-IPUPostRebootVerifier {
    $verifyCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$IPU_VerifierPath`""
    $verifyOut = schtasks.exe /Create /TN $IPU_VerifyTaskName /TR $verifyCmd /SC ONSTART /RU SYSTEM /RL HIGHEST /F 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Output "Post-reboot verifier registered as $IPU_VerifyTaskName."
        return $true
    }
    Write-Output "WARN: post-reboot verifier could not be registered: $verifyOut"
    return $false
}
function Invoke-IPUInstallationAssistant {
    param([string]$Reason)
    Write-Output "Routing to Windows 11 Installation Assistant ($Reason)."
    $iaStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Suspend BitLocker for the reboots the Assistant will drive
    try {
        $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($bl -and $bl.ProtectionStatus -eq 'On') { Suspend-BitLocker -MountPoint "C:" -RebootCount 4 -ErrorAction Stop; Write-Output "BitLocker suspended." }
    } catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }
    if (Get-Process -Name 'Windows11InstallationAssistant' -ErrorAction SilentlyContinue) {
        Write-Output "Installation Assistant already running; leaving it to finish."
        Set-Status "RUNNING IA_INPROGRESS $iaStamp"
        return
    }
    if (-not (Test-Path $IPU_IAExe)) {
        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $IPU_IAUrl -OutFile $IPU_IAExe -UseBasicParsing
        } catch {
            Write-Output "Installation Assistant download failed: $_"
            Invoke-IPUWindowsUpdate -Reason "Assistant download failed"
            return
        }
    }
    if ($IPU_IASha256) {
        $iaHash = (Get-FileHash $IPU_IAExe -Algorithm SHA256).Hash
        if ($iaHash -ne $IPU_IASha256) {
            Write-Output "Installation Assistant hash mismatch. Expected $IPU_IASha256 got $iaHash"
            Remove-Item $IPU_IAExe -Force -ErrorAction SilentlyContinue
            Set-Status "FAILED IA_HASH $iaStamp"
            return
        }
        Write-Output "Installation Assistant hash verified."
    }
    # Run to completion and wait. /NoReboot stages the upgrade without restarting,
    # so the reboot is deferred to the caller exactly like the ISO path. We never
    # auto-reboot here. ( /NoRestartUI is deliberately NOT used: as SYSTEM it forces
    # an immediate silent restart, which is the opposite of what we want. )
    try {
        $iaProc = Start-Process -FilePath $IPU_IAExe -ArgumentList $IPU_IAArgs -Wait -PassThru
        $iaCode = $iaProc.ExitCode
        Write-Output "Installation Assistant exited with code $iaCode."
        if ($iaCode -eq 0 -or $iaCode -eq 3010) {
            $null = Register-IPUPostRebootVerifier
            Set-Status "REBOOT_REQUIRED IA $iaStamp"
        } else {
            Write-Output "Installation Assistant could not complete (exit $iaCode); falling back to Windows Update."
            Invoke-IPUWindowsUpdate -Reason "Assistant exit $iaCode"
        }
    } catch {
        Write-Output "Installation Assistant launch failed: $_"
        Invoke-IPUWindowsUpdate -Reason "Assistant launch failed"
    }
}
function Set-WUUpgradeOverrides {
    # Bypass with reboot protection. Some boxes can only upgrade through Windows Update
    # (Enterprise composition, ARM64), but a NoAutoUpdate=1 policy stops WU from ever
    # delivering the feature update. Clear it and pin active hours so the WU-driven
    # reboot lands out of hours. Originals are recorded and restored once the box is on
    # the latest build; an enforced policy also self-reverts, and each run re-opens the
    # window until the upgrade lands.
    $polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auk  = "$polk\AU"
    if (-not (Test-Path $polk)) { New-Item -Path $polk -Force | Out-Null }
    if (-not (Test-Path $auk))  { New-Item -Path $auk  -Force | Out-Null }
    if (-not (Test-Path $IPU_AURestoreFile)) {
        $noAuto = (Get-ItemProperty -Path $auk  -Name NoAutoUpdate     -ErrorAction SilentlyContinue).NoAutoUpdate
        $auOpt  = (Get-ItemProperty -Path $auk  -Name AUOptions        -ErrorAction SilentlyContinue).AUOptions
        $setAH  = (Get-ItemProperty -Path $polk -Name SetActiveHours   -ErrorAction SilentlyContinue).SetActiveHours
        $ahS    = (Get-ItemProperty -Path $polk -Name ActiveHoursStart -ErrorAction SilentlyContinue).ActiveHoursStart
        $ahE    = (Get-ItemProperty -Path $polk -Name ActiveHoursEnd   -ErrorAction SilentlyContinue).ActiveHoursEnd
        @(
            "NoAutoUpdate=$(if ($null -eq $noAuto) {'ABSENT'} else {$noAuto})"
            "AUOptions=$(if ($null -eq $auOpt) {'ABSENT'} else {$auOpt})"
            "SetActiveHours=$(if ($null -eq $setAH) {'ABSENT'} else {$setAH})"
            "ActiveHoursStart=$(if ($null -eq $ahS) {'ABSENT'} else {$ahS})"
            "ActiveHoursEnd=$(if ($null -eq $ahE) {'ABSENT'} else {$ahE})"
        ) | Set-Content -Path $IPU_AURestoreFile -Encoding ASCII -Force
    }
    Set-ItemProperty -Path $auk  -Name NoAutoUpdate     -Value 0 -Type DWord -Force
    # NoAutoUpdate=0 alone does not release an already-pending feature update on the
    # re-keyed Enterprise-composition boxes; AUOptions=4 (auto download + schedule
    # install) is what actually moves it. Set both.
    Set-ItemProperty -Path $auk  -Name AUOptions        -Value 4 -Type DWord -Force
    Set-ItemProperty -Path $polk -Name SetActiveHours   -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $polk -Name ActiveHoursStart -Value $IPU_ActiveHoursStart -Type DWord -Force
    Set-ItemProperty -Path $polk -Name ActiveHoursEnd   -Value $IPU_ActiveHoursEnd -Type DWord -Force
    Write-Output "Cleared NoAutoUpdate, set AUOptions=4, and pinned active hours $IPU_ActiveHoursStart-$IPU_ActiveHoursEnd so WU can deliver and reboot out of hours (originals recorded)."
}

function Invoke-IPUStaleCleanup {
    # Clear anything a prior aborted attempt left behind: orphaned setup processes,
    # mounted media (which pile up FsDepends/WIMMount filter instances), and setup
    # scratch folders. Any of these can lock files during the next downlevel finalize
    # and produce a 0xC1900101 sharing-violation failure, or leave a stale ~BT that
    # misleads Windows Update. Scoped to our work dir so an unrelated admin-mounted
    # image is never touched. Runs before routing, so every path (WU, Assistant, ISO)
    # starts from a clean state.
    Write-Output "Clearing stale upgrade state..."
    foreach ($pname in 'setup','setuphost','setupprep','setupplatform') {
        Get-Process -Name $pname -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Output "  Stopping leftover process: $($_.Name) (PID $($_.Id))"
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    }
    Get-CimInstance -Namespace 'root\Microsoft\Windows\Storage' -ClassName 'MSFT_DiskImage' -ErrorAction SilentlyContinue |
        Where-Object { $_.Attached -and $_.ImagePath -like "$IPU_WorkDir\*" } |
        ForEach-Object {
            try {
                Dismount-DiskImage -ImagePath $_.ImagePath -ErrorAction Stop | Out-Null
                Write-Output "  Dismounted stale image: $($_.ImagePath)"
            } catch {
                Write-Output "  WARN: could not dismount $($_.ImagePath): $_"
            }
        }
    foreach ($d in "$env:SystemDrive\`$WINDOWS.~BT","$env:SystemDrive\`$WINDOWS.~WS","$env:SystemDrive\`$WINDOWS.~Q") {
        if (Test-Path -LiteralPath $d) {
            Write-Output "  Removing setup scratch: $d"
            Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $d) { cmd /c "rd /s /q `"$d`"" 2>$null }
        }
    }
    Write-Output "Stale state cleanup complete."
}

function Invoke-IPUWindowsUpdate {
    # Universal last-resort path. Windows Update is edition- and language-aware and is
    # the only mechanism that upgrades Enterprise composition and ARM64. It does not
    # stage locally and WU controls the timing, so clear any pause, nudge a scan, and
    # flag it. No reboot flag is set because nothing is staged.
    param([string]$Reason)
    $wuStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "Handing off to Windows Update ($Reason)."
    $uxk = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
    if (Test-Path $uxk) {
        foreach ($pn in 'PauseUpdatesExpiryTime','PauseUpdatesStartTime','PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime','PauseQualityUpdatesStartTime','PauseQualityUpdatesEndTime') {
            Remove-ItemProperty -Path $uxk -Name $pn -ErrorAction SilentlyContinue
        }
    }
    Set-WUUpgradeOverrides
    Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Start-Process -FilePath "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartInteractiveScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
    Set-Status "FAILED_WU_HANDOFF $wuStamp"
}
function Invoke-IPUComponentRepair {
    # Repair component store (StartComponentCleanup, RestoreHealth, SFC). Fixes 0xC1900204 / eKB apply failures. Needs internet.
    Set-Status "REPAIRING $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Repairing component store (DISM StartComponentCleanup, RestoreHealth, SFC)..."
    try { & dism.exe /Online /Cleanup-Image /StartComponentCleanup | Out-Null } catch {}
    try { & dism.exe /Online /Cleanup-Image /RestoreHealth        | Out-Null } catch {}
    try { & sfc.exe /scannow                                       | Out-Null } catch {}
    Write-Output "Component store repair complete."
}
function Invoke-IPUEnablementPackage {
    # 24H2 (26100) -> 25H2 via KB5054156 eKB. Verify servicing baseline, apply, repair+retry once, else WU fallback. x64 only.
    $ekCV = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $ekBuild = [int]$ekCV.CurrentBuild
    $ekUBR = [int]$ekCV.UBR
    if ($ekBuild -eq 26100 -and $ekUBR -lt 5074) {
        Write-Output "24H2 servicing baseline $ekBuild.$ekUBR is below the 25H2 eKB prerequisite 26100.5074. Handing off to Windows Update."
        Invoke-IPUWindowsUpdate -Reason "24H2 servicing baseline $ekBuild.$ekUBR below 26100.5074"
        return
    }

    Write-Output "24H2 detected (build $ekBuild.$ekUBR). Applying the 25H2 enablement package (KB5054156)."
    try {
        $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($bl -and $bl.ProtectionStatus -eq 'On') { Suspend-BitLocker -MountPoint "C:" -RebootCount 2 -ErrorAction Stop; Write-Output "BitLocker suspension for the eKB reboot." }
    } catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }

    Set-Status "DOWNLOADING $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    if (Test-Path $IPU_EkbPath) { Remove-Item $IPU_EkbPath -Force -ErrorAction SilentlyContinue }
    $ekOK = $true
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $IPU_EkbUrl -OutFile $IPU_EkbPath -UseBasicParsing -TimeoutSec 180
    } catch { Write-Output "eKB download failed from $IPU_EkbUrl : $_"; $ekOK = $false }
    if ($ekOK -and $IPU_EkbSha256) {
        $h = (Get-FileHash $IPU_EkbPath -Algorithm SHA256).Hash
        if ($h -ne $IPU_EkbSha256) { Write-Output "eKB hash mismatch. Expected $IPU_EkbSha256 got $h."; Remove-Item $IPU_EkbPath -Force -ErrorAction SilentlyContinue; $ekOK = $false }
        else { Write-Output "eKB hash verified." }
    }
    if ($ekOK) {
        $codes = @()
        for ($t = 1; $t -le 2; $t++) {
            Write-Output "Applying eKB via wusa (attempt $t of 2)..."
            $wp = Start-Process -FilePath "$env:SystemRoot\System32\wusa.exe" -ArgumentList "`"$IPU_EkbPath`" /quiet /norestart" -Wait -PassThru
            $ec = $wp.ExitCode
            $codes += ("0x{0:X}" -f $ec)
            "wusa exit code: $ec (0x$('{0:X}' -f $ec))" | Out-File -FilePath $IPU_SetupLog -Encoding ASCII
            # 0 = applied; 3010 = applied, reboot required; 0x240006 (2359302) = already installed
            if ($ec -eq 0 -or $ec -eq 3010) {
                $null = Register-IPUPostRebootVerifier
                Set-Status "REBOOT_REQUIRED 3010 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Write-Output "eKB applied (exit $($codes[-1])). Reboot to finish 25H2."
                return
            }
            if ($ec -eq 2359302) {
                $ekPending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or (Test-Path "$env:SystemRoot\WinSxS\pending.xml")
                if ($ekPending) {
                    $null = Register-IPUPostRebootVerifier
                    Set-Status "REBOOT_REQUIRED 3010 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    Write-Output "eKB is already installed and servicing reports a pending reboot. Reboot to finish 25H2."
                    return
                }
                Write-Output "eKB is already installed, but no pending reboot is present and the OS is still 24H2. Handing off to Windows Update."
                Invoke-IPUWindowsUpdate -Reason "24H2 eKB already installed without pending reboot"
                return
            }
            if ($t -eq 1) { Write-Output "eKB apply returned $($codes[-1]). Repairing component store and retrying once."; Invoke-IPUComponentRepair }
        }
        Write-Output "eKB did not apply after repair and retry (codes: $($codes -join ', ')). Handing off to Windows Update."
    } else {
        Write-Output "eKB not available locally. Handing off to Windows Update (it delivers prerequisite CUs and the 25H2 feature update)."
    }
    Invoke-IPUWindowsUpdate -Reason "24H2 eKB unavailable or failed"
}
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }
if (-not (Test-Path $IPU_LogDir))  { New-Item -ItemType Directory -Path $IPU_LogDir  -Force | Out-Null }
$IPU_LegacyIso = Join-Path $IPU_WorkDir "Win11_25H2_x64.iso"
if (Test-Path $IPU_LegacyIso) { Remove-Item $IPU_LegacyIso -Force -ErrorAction SilentlyContinue }
Set-Status "RUNNING $stamp"
Start-Transcript -Path (Join-Path $IPU_LogDir "ipu_runner.log") -Append -Force | Out-Null

# Hygiene first: clear stale setup scratch / processes / mounts a prior aborted run
# left behind, so it cannot lock files at finalize or leave a stale ~BT that misleads
# Windows Update. Runs ahead of routing, so it benefits the WU, Assistant and ISO paths.
Invoke-IPUStaleCleanup

$cancel = $false

# Universal routing: pick a mechanism that can actually serve THIS box. WURSA runs
# worldwide, so the device may be any architecture, edition, UI language, or build.
# The local ISO only covers x64 + retail editions + the languages we host (en-US/en-GB).
#   LTSC / LTSB            -> skip and report (not feature-upgraded this way)
#   ARM64                  -> Windows Update (x64 ISO and x64-only Assistant cannot run)
#   24H2 x64 (build 26100) -> KB5054156 enablement package (light, one reboot)
#   Enterprise             -> Windows Update (no retail media carries Enterprise)
#   non-hosted UI language -> Installation Assistant (MS serves matching language+edition)
#   hosted language, retail-> the local ISO below (covers Win10 and older Win11 x64)
# The Assistant itself falls back to Windows Update if it cannot serve the box.
$IPU_CV      = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
$IPU_CompEd  = $IPU_CV.CompositionEditionID
$IPU_EdId    = $IPU_CV.EditionID
$IPU_Prod    = $IPU_CV.ProductName
$IPU_Build   = [int]$IPU_CV.CurrentBuild
$IPU_RunLang = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language' -ErrorAction SilentlyContinue).InstallLanguage
$IPU_IsArm   = ($env:PROCESSOR_ARCHITECTURE -match 'ARM') -or ($env:PROCESSOR_ARCHITEW6432 -match 'ARM')
$IPU_IsLtsc  = ($IPU_EdId -match 'EnterpriseS|IoTEnterpriseS') -or ($IPU_Prod -match 'LTSC|LTSB')
$IPU_Hosted  = ($IPU_RunLang -eq '0409' -or $IPU_RunLang -eq '0809')

if ($IPU_IsLtsc) {
    Write-Output "Edition '$IPU_EdId' ($IPU_Prod) is LTSC/LTSB; not feature-upgraded this way. Skipping for manual handling."
    Set-Status "BLOCK_LTSC $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $cancel = $true
}
elseif ($IPU_IsArm) {
    Write-Output "ARM64 device; the x64 ISO and Installation Assistant cannot run here."
    Invoke-IPUWindowsUpdate -Reason "ARM64"
    $cancel = $true
}
elseif ($IPU_Build -eq 26100) {
    # 24H2 (x64, since ARM was handled above): take the enablement package, not the ISO.
    Invoke-IPUEnablementPackage
    $cancel = $true
}
elseif ($IPU_EdId -match 'Enterprise') {
    Write-Output "Edition '$IPU_EdId'; no retail media (ISO or Assistant) carries Enterprise."
    Invoke-IPUWindowsUpdate -Reason "edition $IPU_EdId"
    $cancel = $true
}
elseif (-not $IPU_Hosted) {
    Write-Output "UI language '$IPU_RunLang' is not one we host an ISO for; using the Installation Assistant."
    Invoke-IPUInstallationAssistant -Reason "non-hosted language $IPU_RunLang"
    $cancel = $true
}
# else: hosted language + retail edition -> fall through to the ISO path below

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
        # BitLocker suspension - 4 reboots covers SafeOS + first/second boot + buffer.
        # Done here, after all gates pass, so an aborted run never leaves protection off.
        try {
            $BLStatus = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
            if ($BLStatus -and $BLStatus.ProtectionStatus -eq 'On') {
                Suspend-BitLocker -MountPoint "C:" -RebootCount 4 -ErrorAction Stop
                Write-Output "BitLocker suspended."
            }
        } catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }

        $SetupArgs = "/auto upgrade /quiet /compat ignorewarning /DynamicUpdate disable /showoobe None /Telemetry Disable /EULA Accept /noreboot /Copylogs `"$IPU_LogDir`""
        $IPU_ExitCode = -1
        $storeRepaired = $false
        for ($setupTry = 1; $setupTry -le 2; $setupTry++) {
            Write-Output "Launching Windows Setup (downlevel phase, attempt $setupTry of 2)..."
            $IPU_ExitCode = -1
            try {
                $proc = Start-Process -FilePath $SetupExe -ArgumentList $SetupArgs -Wait -PassThru
                $IPU_ExitCode = $proc.ExitCode
            } catch { Write-Output "setup.exe failed to launch: $_" }
            "setup.exe exit code: $IPU_ExitCode" | Out-File -FilePath $IPU_SetupLog -Encoding ASCII
            if ($IPU_ExitCode -eq 0 -or $IPU_ExitCode -eq 3010) { break }
            # 0xC1900204 with matched edition/language = corrupt store, not a real block. Repair once and retry.
            if ($IPU_ExitCode -eq -1047526908 -and -not $storeRepaired -and $setupTry -lt 2) {
                Write-Output "0xC1900204 with matched edition/language: repairing component store and retrying."
                Invoke-IPUComponentRepair
                $storeRepaired = $true
                continue
            }
            break
        }
        Dismount-DiskImage -ImagePath $IPU_IsoPath -ErrorAction SilentlyContinue

        $done = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        switch ($IPU_ExitCode) {
            0           { $null = Register-IPUPostRebootVerifier; Set-Status "REBOOT_REQUIRED 3010 $done" }
            3010        { $null = Register-IPUPostRebootVerifier; Set-Status "REBOOT_REQUIRED 3010 $done" }
            -1047526908 { Invoke-IPUInstallationAssistant -Reason "0xC1900204 persists after store repair" }
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
                $script:FeatureUpgradeState = "DISPATCHED"
                Write-Host "      > Upgrade dispatched as SYSTEM. It continues after you disconnect." -ForegroundColor Green
                Write-Host "      > Status file : $IPU_StatusFile" -ForegroundColor $DimCol
                Write-Host "      > Setup logs  : $IPU_LogDir" -ForegroundColor $DimCol
                Write-Host "      > A reboot will be required once the detached upgrade completes." -ForegroundColor $WarnCol
            } else {
                $script:FeatureUpgradeState = "FAILED"
                ('FAILED TASK_START ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
                Write-Host "      [!] Task created but failed to start: $runOut" -ForegroundColor Red
            }
        } else {
            $script:FeatureUpgradeState = "FAILED"
            ('FAILED TASK_CREATE ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
            Write-Host "      [!] Failed to register scheduled task: $createOut" -ForegroundColor Red
        }
    }
} else {
    $script:FeatureUpgradeState = "CURRENT"
    Write-Host "      System is already on the latest feature update." -ForegroundColor Green
    Restore-WUUpgradeOverrides
}
#endregion


#region 6 - Finalization
# ============================================================================
# Checking reboot status via direct API to bypass Get-WURebootStatus crash.
$RebootPending = if ($InstallResult) { $InstallResult.RebootRequired } else { $false }
if (-not $RebootPending) {
    $RebootPending = $null -ne (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue)
}

# Never report SYSTEM CURRENT solely because there is no reboot flag. A feature
# upgrade is complete only when the installed DisplayVersion actually equals 25H2.
$_finalCV      = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
$_finalVersion = $_finalCV.DisplayVersion
$_ipuState     = ""
if (Test-Path $IPU_StatusFile) {
    try { $_ipuState = (Get-Content $IPU_StatusFile -ErrorAction Stop | Select-Object -First 1).Trim() } catch {}
}

Write-HLine -Style dashed
if (-not $NoUpgrade -and $_finalVersion -eq $LatestVersion) {
    Write-Host "STATUS: SYSTEM CURRENT ($LatestVersion)" -ForegroundColor Green
    $script:ExitCode = 0
} elseif ($NoUpgrade) {
    if ($RebootPending) {
        Write-Host "[!] STATUS: REBOOT REQUIRED" -ForegroundColor Red
        $script:ExitCode = 3010
    } else {
        Write-Host "STATUS: UPDATE PASS COMPLETE" -ForegroundColor Green
        $script:ExitCode = 0
    }
} elseif ($script:FeatureUpgradeState -eq "REBOOT_REQUIRED") {
    Write-Host "[!] STATUS: REBOOT REQUIRED TO FINISH $LatestVersion" -ForegroundColor Red
    $script:ExitCode = 3010
} elseif ($script:FeatureUpgradeState -match '^(FAILED|BLOCKED|NOT_STARTED)') {
    Write-Host "[!] STATUS: $LatestVersion FEATURE UPGRADE NOT COMPLETE" -ForegroundColor Red
    if ($script:FeatureUpgradeState) { Write-Host "      Upgrade state: $($script:FeatureUpgradeState)" -ForegroundColor $WarnCol }
    $script:ExitCode = 1
} elseif ($script:FeatureUpgradeState -eq "DISPATCHED") {
    if ($_ipuState -match '^REBOOT_REQUIRED') {
        Write-Host "[!] STATUS: REBOOT REQUIRED TO FINISH $LatestVersion" -ForegroundColor Red
        Write-Host "      Upgrade state: $_ipuState" -ForegroundColor $DimCol
        $script:ExitCode = 3010
    } elseif ($_ipuState -match '^(FAILED|BLOCK_|UNEXPECTED)') {
        Write-Host "[!] STATUS: $LatestVersion FEATURE UPGRADE NOT COMPLETE" -ForegroundColor Red
        Write-Host "      Upgrade state: $_ipuState" -ForegroundColor $WarnCol
        $script:ExitCode = 1
    } else {
        Write-Host "[>] STATUS: $LatestVersion FEATURE UPGRADE IN PROGRESS" -ForegroundColor $WarnCol
        if ($_ipuState) { Write-Host "      Upgrade state: $_ipuState" -ForegroundColor $DimCol }
        $script:ExitCode = 0
    }
} elseif ($_ipuState -match '^REBOOT_REQUIRED') {
    Write-Host "[!] STATUS: REBOOT REQUIRED TO FINISH $LatestVersion" -ForegroundColor Red
    Write-Host "      Upgrade state: $_ipuState" -ForegroundColor $DimCol
    $script:ExitCode = 3010
} elseif ($_ipuState -match '^(DISPATCHED|RUNNING|DOWNLOADING|REPAIRING|VERIFY_PENDING)') {
    Write-Host "[>] STATUS: $LatestVersion FEATURE UPGRADE IN PROGRESS" -ForegroundColor $WarnCol
    Write-Host "      Upgrade state: $_ipuState" -ForegroundColor $DimCol
    $script:ExitCode = 0
} elseif ($_ipuState -match '^(FAILED|BLOCK_|UNEXPECTED)') {
    Write-Host "[!] STATUS: $LatestVersion FEATURE UPGRADE NOT COMPLETE" -ForegroundColor Red
    Write-Host "      Upgrade state: $_ipuState" -ForegroundColor $WarnCol
    $script:ExitCode = 1
} else {
    Write-Host "[!] STATUS: FEATURE UPGRADE NOT COMPLETE (still $_finalVersion)" -ForegroundColor Red
    $script:ExitCode = 1
}
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