<#
.NAME
    Deployment Environment Bloatware Liquidator & Optimized Automated Toolkit (D.E.B.L.O.A.T.) v2.2
    Developed by Steve the Killer | Updated: 2026-05-12
.DESCRIPTION
    Standardizes Windows 11 by removing OEM bloat (HP, Dell,
    ASUS/Acer), AI/Recall features, and sponsored consumer content.
    Hardens privacy via telemetry caps, Edge policy enforcement, and
    taskbar/Start menu lockdown applied across all user profiles
    including the Default User template. All per-user settings are
    routed through hive loading so they apply correctly when run as
    SYSTEM via LiveConnect/RMM.
#>

#region 0: INITIALIZATION AND HELPER FUNCTIONS
# ============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "DEBLOAT requires Administrative privileges. Please relaunch as Admin."
    exit
}

$ProgressPreference = 'SilentlyContinue'

# Probe for real console support. LiveConnect / RMM / redirected hosts
# throw on cursor API; fall back to plain step output when unavailable.
$script:UseCursor = $false
try {
    $_t = [Console]::CursorTop
    [Console]::SetCursorPosition(0, $_t)
    $script:UseCursor = $true
} catch {
    $script:UseCursor = $false
}

if ($script:UseCursor) { Clear-Host }
$script:Width = 85

$LineCol   = "White"
$MainCol   = "Yellow"
$WarnCol   = "DarkYellow"
$ArtCol    = "DarkRed"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$InfoCol   = "Cyan"
$OkCol     = "Green"

# Register HKU: PSDrive so HKU:\<Hive>\... paths resolve in Test-Path / Set-ItemProperty.
if (-not (Get-PSDrive -Name HKU -PSProvider Registry -ErrorAction SilentlyContinue)) {
    New-PSDrive -Name HKU -PSProvider Registry -Root HKEY_USERS -Scope Script | Out-Null
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
        [ConsoleColor]$LineCol,
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

# Header Art & Logic
$_pfx  = "█  "
$_art1 = "╔═╗ ╔═╗ ╔╗  ╦   ╔═╗ ╔═╗ ╔╦╗ "
$_art2 = "║ ║ ║╣  ╠╩╗ ║   ║ ║ ╠═╣  ║  "
$_art3 = "╩═╝ ╚═╝ ╩═╝ ╩═╝ ╚═╝ ╩ ╩  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "DEPLOYMENT ENV BLOAT LIQUIDATOR & OPTIMIZED TOOLKIT"
$_ver   = "| v2.2"

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
$Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer

# --- Step Indicator Functions (LiveConnect-safe) ---
$script:StepRow = -1
$script:StepMsg = ""

function Write-Step {
    param([string]$Msg)
    $script:StepMsg = $Msg
    if ($script:UseCursor) {
        try { $script:StepRow = [Console]::CursorTop } catch { $script:StepRow = -1 }
    } else {
        $script:StepRow = -1
    }
    Write-Host $Msg -ForegroundColor $InfoCol
}

function Complete-Step {
    $marker = "[SUCCESS]"
    if ($script:UseCursor -and $script:StepRow -ge 0) {
        try {
            $savedTop = [Console]::CursorTop
            if ($script:StepRow -lt $savedTop) {
                [Console]::SetCursorPosition(0, $script:StepRow)
                $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - $marker.Length)
                Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
                Write-Host "$_pad$marker" -ForegroundColor $OkCol
                [Console]::SetCursorPosition(0, $savedTop)
            }
        } catch {
            Write-Host "  $marker $($script:StepMsg)" -ForegroundColor $OkCol
        }
    } else {
        $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - $marker.Length)
        Write-Host "$($script:StepMsg)$_pad$marker" -ForegroundColor $OkCol
    }
    $script:StepRow = -1
}

# Universal registry value setter. Creates the key if missing.
function Set-UserRegValue {
    param(
        [Parameter(Mandatory)][string]$Path,
        [Parameter(Mandatory)][string]$Name,
        [Parameter(Mandatory)]$Value,
        [string]$Type = "DWord"
    )
    try {
        if (-not (Test-Path $Path)) {
            New-Item -Path $Path -Force -ErrorAction Stop | Out-Null
        }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction Stop
    } catch {
        # Path may be inside an unloaded hive or otherwise unreachable; suppress.
    }
}

# Per-user hive cleanup. Loads each profile (or targets already-loaded SIDs
# for logged-in users), runs each provided scriptblock against the hive,
# then unloads. Safe to call multiple times; loaded hives are detected.
function Invoke-ComprehensiveUserCleanup {
    param([scriptblock[]]$RegistryOperations)

    # 1. Target the Default User (Template for all FUTURE users)
    Write-Host "[*]   Updating Default User Template..." -ForegroundColor $WarnCol
    reg load HKU\DefaultUser "C:\Users\Default\NTUSER.DAT" *>&1 | Out-Null
    foreach ($Op in $RegistryOperations) {
        Invoke-Command -ScriptBlock $Op -ArgumentList "DefaultUser"
    }
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    reg unload HKU\DefaultUser *>&1 | Out-Null

    # 2. Map loaded hives by username so we can target logged-in users by SID
    #    instead of trying to reg-load their in-use NTUSER.DAT.
    $LoadedHives = @{}
    Get-ChildItem "Registry::HKEY_USERS" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-' -and $_.PSChildName -notmatch '_Classes$' } |
        ForEach-Object {
            $sid = $_.PSChildName
            try {
                $acct = ([System.Security.Principal.SecurityIdentifier]$sid).Translate([System.Security.Principal.NTAccount]).Value
                $uname = $acct.Split('\')[-1]
                $LoadedHives[$uname] = $sid
            } catch { }
        }

    # 3. Iterate every profile under C:\Users
    $UserFolders = Get-ChildItem "C:\Users" -Directory |
        Where-Object { $_.Name -notmatch "Public|Default|All Users|DefaultAppPool|WDAGUtilityAccount" }

    foreach ($Folder in $UserFolders) {
        if ($LoadedHives.ContainsKey($Folder.Name)) {
            # Hive is already loaded (user logged in). Write through SID; no load/unload.
            $HiveName = $LoadedHives[$Folder.Name]
            Write-Host "[*]   Cleaning Profile (live): $($Folder.Name)..." -ForegroundColor $WarnCol
            foreach ($Op in $RegistryOperations) {
                Invoke-Command -ScriptBlock $Op -ArgumentList $HiveName
            }
        } else {
            $NTUserPath = "$($Folder.FullName)\NTUSER.DAT"
            if (Test-Path $NTUserPath) {
                $HiveName = "TempHive_$($Folder.Name)"
                Write-Host "[*]   Cleaning Profile: $($Folder.Name)..." -ForegroundColor $WarnCol
                reg load "HKU\$HiveName" $NTUserPath *>&1 | Out-Null
                if ($LASTEXITCODE -eq 0) {
                    foreach ($Op in $RegistryOperations) {
                        Invoke-Command -ScriptBlock $Op -ArgumentList $HiveName
                    }
                    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
                    reg unload "HKU\$HiveName" *>&1 | Out-Null
                } else {
                    Write-Host "[!]   Skipped $($Folder.Name): hive in use or inaccessible." -ForegroundColor $DimCol
                }
            }
        }
    }
}

# Walk Uninstall registry directly; Get-Package is unreliable under SYSCTX.
function Get-InstalledMsiPackage {
    param([Parameter(Mandatory)][string]$NamePattern)
    $UninstallPaths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $matched = @()
    foreach ($p in $UninstallPaths) {
        $matched += Get-ItemProperty $p -ErrorAction SilentlyContinue |
            Where-Object { $_.DisplayName -like $NamePattern -and $_.UninstallString }
    }
    return $matched
}

function Uninstall-MsiPackage {
    param([Parameter(Mandatory)]$Package)
    $name = $Package.DisplayName
    if ($Package.PSChildName -match '^\{[0-9A-Fa-f-]+\}$') {
        Write-Host "[*]   Uninstalling MSI: $name" -ForegroundColor $WarnCol
        Start-Process "msiexec.exe" -ArgumentList "/x $($Package.PSChildName) /qn /norestart" -Wait -NoNewWindow -ErrorAction SilentlyContinue
    } elseif ($Package.QuietUninstallString) {
        Write-Host "[*]   Uninstalling (quiet): $name" -ForegroundColor $WarnCol
        Start-Process "cmd.exe" -ArgumentList "/c $($Package.QuietUninstallString)" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
    } elseif ($Package.UninstallString) {
        $us = $Package.UninstallString
        if ($us -match 'msiexec') {
            $us = $us -replace '(?i)/i\{', '/x{' -replace '(?i)/i ', '/x '
            if ($us -notmatch '(?i)/qn|/quiet') { $us += ' /qn /norestart' }
            Write-Host "[*]   Uninstalling: $name" -ForegroundColor $WarnCol
            Start-Process "cmd.exe" -ArgumentList "/c $us" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
        } else {
            Write-Host "[!]   No silent uninstall available: $name" -ForegroundColor $DimCol
        }
    }
}
#endregion

#region 1A: PER-USER OPERATION DEFINITIONS
# ============================================================================
# All HKCU-equivalent settings are defined here as scriptblocks so they can
# be applied against every loaded hive (Default User + every real user)
# instead of writing to whichever HKCU happens to be in scope at runtime
# (which under LiveConnect is SYSTEM's hive).

$UserHKCUOps = {
    param($Hive)
    $base = "HKU:\$Hive"

    # 1.3: Privacy / Tailored Experiences / Language Opt-Out
    Set-UserRegValue "$base\Software\Microsoft\Windows\CurrentVersion\Privacy" "TailoredExperiencesWithDiagnosticDataEnabled" 0
    Set-UserRegValue "$base\Control Panel\International\User Profile" "HttpAcceptLanguageOptOut" 1

    # 1.5: Bing Search / Search Highlights
    Set-UserRegValue "$base\Software\Policies\Microsoft\Windows\Explorer" "DisableSearchBoxSuggestions" 1
    Set-UserRegValue "$base\Software\Microsoft\Windows\CurrentVersion\SearchSettings" "IsDynamicSearchBoxPresent" 0
    Set-UserRegValue "$base\Software\Microsoft\Windows\CurrentVersion\Search" "BingSearchEnabled" 0

    # 1.7: ContentDeliveryManager (lock screen ads, sponsored apps, spotlight)
    $cdm = "$base\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
    Set-UserRegValue $cdm "SubscribedContent-338387Enabled" 0
    Set-UserRegValue $cdm "SubscribedContent-338388Enabled" 0
    Set-UserRegValue $cdm "SubscribedContent-338389Enabled" 0
    Set-UserRegValue $cdm "SubscribedContent-353694Enabled" 0
    Set-UserRegValue $cdm "SubscribedContent-353696Enabled" 0
    Set-UserRegValue $cdm "SubscribedContent-338393Enabled" 0
    Set-UserRegValue $cdm "SubscribedContent-310093Enabled" 0
    Set-UserRegValue $cdm "SystemPaneSuggestionsEnabled"    0
    Set-UserRegValue $cdm "SoftLandingEnabled"              0
    Set-UserRegValue $cdm "RotatingLockScreenEnabled"       0
    Set-UserRegValue $cdm "RotatingLockScreenOverlayEnabled" 0
    Set-UserRegValue $cdm "PreInstalledAppsEnabled"         0
    Set-UserRegValue $cdm "ContentDeliveryAllowed"          0
    Set-UserRegValue $cdm "SilentInstalledAppsEnabled"      0

    # 1.7: Taskbar Widgets / Chat / Start tracking / Explorer LaunchTo
    $adv = "$base\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
    Set-UserRegValue $adv "TaskbarDa"         0   # Widgets
    Set-UserRegValue $adv "TaskbarMn"         0   # Chat/Teams Consumer
    Set-UserRegValue $adv "ShowCopilotButton" 0   # Copilot taskbar button
    Set-UserRegValue $adv "Start_TrackProgs"  0   # Most used apps
    Set-UserRegValue $adv "LaunchTo"          1   # Open Explorer to "This PC"

    # 1.7: Suggested notifications
    Set-UserRegValue "$base\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested" "Enabled" 0
}

$HPUserOps = {
    param($Hive)
    $base = "HKU:\$Hive"

    # Wipe HP-injected Start pinning + cloud cache. CDM deny flags are owned by
    # $UserHKCUOps so we don't wipe the whole CDM key here.
    foreach ($p in @(
        "$base\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage2",
        "$base\Software\Microsoft\Windows\CurrentVersion\CloudStore"
    )) {
        if (Test-Path $p) {
            Write-Host "Removing $p..." -ForegroundColor $DimCol
            Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # Wipe HP-injected silent-install subscriptions specifically.
    $Subs = "$base\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager\Subscriptions"
    if (Test-Path $Subs) {
        Remove-Item -Path $Subs -Recurse -Force -ErrorAction SilentlyContinue
    }
}
#endregion

#region 1: WINDOWS CORE (RECALL, AI, PRIVACY & PERFORMANCE)
# ============================================================================
Write-Step "[1]   DEBLOAT: Purging Windows AI, Hardening Privacy & Optimizing UI..."

# 1.1: Recall & AI Removal (Added Error Handling)
$RecallFeature = Get-WindowsOptionalFeature -Online -FeatureName "Recall" -ErrorAction SilentlyContinue
if ($RecallFeature -and $RecallFeature.State -eq "Enabled") {
    Write-Host "[*]   Removing Windows Recall via DISM (may take 30-60s)..." -ForegroundColor $WarnCol
    $DismProc = Start-Process "dism.exe" -ArgumentList "/online /Disable-Feature /FeatureName:Recall /Remove /NoRestart /Quiet /English" -PassThru -WindowStyle Hidden
    $DismProc | Wait-Process -Timeout 60 -ErrorAction SilentlyContinue
    if (-not $DismProc.HasExited) {
        Stop-Process -Id $DismProc.Id -Force
        Write-Host "[!]   DISM timed out." -ForegroundColor $WarnCol
    }
} else {
    Write-Host "[*]   Windows Recall not present on this build. Skipping." -ForegroundColor $WarnCol
}
$aiReg = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"
if (!(Test-Path $aiReg)) { New-Item -Path $aiReg -Force | Out-Null }
Set-ItemProperty -Path $aiReg -Name "DisableAIDataAnalysis" -Value 1

# 1.2: Kill Telemetry & Diagnostic Data (Using sc.exe to bypass RPC 1726)
$TargetSvcs = @("DiagTrack", "dmwappushservice")
foreach ($Svc in $TargetSvcs) {
    if (Get-Service -Name $Svc -ErrorAction SilentlyContinue) {
        # sc.exe is a native tool; it won't trigger the RPC red text like Stop-Service
        & sc.exe stop $Svc >$null 2>&1
        & sc.exe config $Svc start= disabled >$null 2>&1
        # Keep your registry line as a backup
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$Svc" -Name "Start" -Value 4 -ErrorAction SilentlyContinue 2>$null
    }
}
# Telemetry level caps
$dcPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
if (!(Test-Path $dcPath)) { New-Item -Path $dcPath -Force | Out-Null }
Set-ItemProperty -Path $dcPath -Name "AllowTelemetry"                 -Value 1
Set-ItemProperty -Path $dcPath -Name "DoNotShowFeedbackNotifications" -Value 1
$dcPath2 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
if (!(Test-Path $dcPath2)) { New-Item -Path $dcPath2 -Force | Out-Null }
Set-ItemProperty -Path $dcPath2 -Name "AllowTelemetry"      -Value 0
Set-ItemProperty -Path $dcPath2 -Name "MaxTelemetryAllowed" -Value 0
# Windows Error Reporting
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting" -Name "Disabled" -Value 1 -ErrorAction SilentlyContinue
$werPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting"
if (!(Test-Path $werPath)) { New-Item -Path $werPath -Force | Out-Null }
Set-ItemProperty -Path $werPath -Name "Disabled" -Value 1
# Customer Experience Improvement Program
$sqmPath = "HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows"
if (!(Test-Path $sqmPath)) { New-Item -Path $sqmPath -Force | Out-Null }
Set-ItemProperty -Path $sqmPath -Name "CEIPEnable" -Value 0
# Advertising ID - machine-level policy
$adPolPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"
if (!(Test-Path $adPolPath)) { New-Item -Path $adPolPath -Force | Out-Null }
Set-ItemProperty -Path $adPolPath -Name "DisabledByGroupPolicy" -Value 1

# 1.4: Disable Activity History & Clipboard Sync (HKLM policy)
$sysPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
if (!(Test-Path $sysPath)) { New-Item -Path $sysPath -Force | Out-Null }
Set-ItemProperty -Path $sysPath -Name "PublishUserActivities" -Value 0
Set-ItemProperty -Path $sysPath -Name "EnableActivityFeed"    -Value 0
Set-ItemProperty -Path $sysPath -Name "UploadUserActivities"  -Value 0

# 1.6: Global OEM Re-injection & Content Prevention
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata" -Name "PreventDeviceMetadataFromNetwork" -Value 1
$cdmPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
if (!(Test-Path $cdmPath)) { New-Item -Path $cdmPath -Force | Out-Null }
Set-ItemProperty -Path $cdmPath -Name "DisableWindowsConsumerFeatures" -Value 1
# Windows Spotlight & soft landing
Set-ItemProperty -Path $cdmPath -Name "DisableWindowsSpotlightFeatures" -Value 1
Set-ItemProperty -Path $cdmPath -Name "DisableSoftLanding"              -Value 1

# 1.7 (HKLM portion): Chat/Teams icon policy & News/Feeds
$chatPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
if (!(Test-Path $chatPolicyPath)) { New-Item -Path $chatPolicyPath -Force | Out-Null }
Set-ItemProperty -Path $chatPolicyPath -Name "ChatIcon" -Value 3
$feedsPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
if (!(Test-Path $feedsPath)) { New-Item -Path $feedsPath -Force | Out-Null }
Set-ItemProperty -Path $feedsPath -Name "EnableFeeds" -Value 0
# Start menu recommendations (HKLM policy)
$startPolPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
if (!(Test-Path $startPolPath)) { New-Item -Path $startPolPath -Force | Out-Null }
Set-ItemProperty -Path $startPolPath -Name "HideRecentlyAddedApps"  -Value 1
$explorerPolPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
if (!(Test-Path $explorerPolPath)) { New-Item -Path $explorerPolPath -Force | Out-Null }

# 1.8: Microsoft Edge Hardening
$EdgePolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
if (!(Test-Path $EdgePolicy)) { New-Item -Path $EdgePolicy -Force | Out-Null }
Set-ItemProperty -Path $EdgePolicy -Name "HubsSidebarEnabled"             -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "CopilotPageContext"              -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "EdgeEntraCopilotPageContext"     -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "EdgeSidebarEnabled"              -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "BackgroundModeEnabled"           -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "StartupBoostEnabled"             -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "ShowMicrosoftRewards"            -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "EdgeShoppingAssistantEnabled"    -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "PersonalizationReportingEnabled" -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "EdgeFollowEnabled"               -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "ShowRecommendationsEnabled"      -Value 0
Set-ItemProperty -Path $EdgePolicy -Name "DiscoverPageContextEnabled"      -Value 0

# 1.9b: Windows Copilot (machine-wide policy; per-user ShowCopilotButton handled in UserHKCUOps)
$copilotPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"
if (!(Test-Path $copilotPath)) { New-Item -Path $copilotPath -Force | Out-Null }
Set-ItemProperty -Path $copilotPath -Name "TurnOffWindowsCopilot" -Value 1

# 1.10: General Performance & Power
# Set Power Plan to High Performance ONLY when plugged in (AC)
$highPerf = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan | Where-Object {$_.ElementName -eq "High performance"}
if ($highPerf) {
    $planGuid = ($highPerf.InstanceId -split '\{' | Select-Object -Last 1).TrimEnd('}')
    powercfg /setactive $planGuid
}

# Set Sleep and Monitor Timers
powercfg /change standby-timeout-ac 0    # Never sleep on AC
powercfg /change standby-timeout-dc 60   # 60 mins on Battery
powercfg /change monitor-timeout-ac 60   # 60 mins on AC
powercfg /change monitor-timeout-dc 15   # 15 mins on Battery

# Disable Hibernation
powercfg /hibernate off

# 1.9: Apply all per-user HKCU settings across Default User + every real profile
Write-Host "[*]   Applying per-user settings across all profiles..." -ForegroundColor $WarnCol
Invoke-ComprehensiveUserCleanup -RegistryOperations $UserHKCUOps

# 1.11: Office Language Cleanup (Purge non-English C2R cultures)
Write-Host "[*]   Checking for non-English Office Language Packs..." -ForegroundColor $WarnCol
$C2RPath   = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
$C2RConfig = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if ((Test-Path $C2RPath) -and (Test-Path $C2RConfig)) {
    $InstalledCultures = (Get-ItemProperty $C2RConfig -Name "ClientCulture" -ErrorAction SilentlyContinue).ClientCulture
    if ($InstalledCultures) {
        $Cultures = $InstalledCultures -split "," |
            ForEach-Object { $_.Trim().ToLower() } |
            Where-Object { $_ -and $_ -ne "en-us" }
        if ($Cultures.Count -gt 0) {
            foreach ($LangToRemove in $Cultures) {
                Write-Host "[*]   Removing Office language: $LangToRemove" -ForegroundColor $DimCol
                Start-Process $C2RPath -ArgumentList @(
                    "scenario=install",
                    "scenariosubtype=ARP",
                    "sourcetype=None",
                    "productstoremove=LanguagePack.$LangToRemove",
                    "culture=$LangToRemove",
                    "DisplayLevel=False",
                    "forceappshutdown=True"
                ) -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
            }
        } else {
            Write-Host "[*]   Only en-us installed. Nothing to remove." -ForegroundColor $DimCol
        }
    } else {
        Write-Host "[*]   ClientCulture key not present. Skipping." -ForegroundColor $DimCol
    }
} else {
    Write-Host "[*]   Office C2R not installed. Skipping." -ForegroundColor $DimCol
}
#endregion

#region 2: HP
# ============================================================================
Complete-Step
Write-Step "[2]   DEBLOAT: OEM Hardware Cleanup..."
if ($Manufacturer -match "HP" -or $Manufacturer -match "Hewlett-Packard") {
    Write-Host "[2.1] HP Hardware Detected. Commencing Deep Cleanup..." -ForegroundColor DarkYellow

    # 2.1: Services & Tasks
    $hpSvc = @("HPTouchpointAnalyticsService", "HPAppHelperCap", "HPDiagsCap", "HPSysInfoCap", "HPNetworkCap", "HPSupportAssistant")
    foreach ($s in $hpSvc) {
        if (Get-Service -Name $s -ErrorAction SilentlyContinue) {
            Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
            Set-Service -Name $s -StartupType Disabled
        }
    }
    Get-ScheduledTask -TaskPath "\HP*" -ErrorAction SilentlyContinue | Disable-ScheduledTask

    # 2.2: Per-user HP pinning/cache wipe
    Write-Host "[*]   Cleaning User Hives for HP..." -ForegroundColor $WarnCol
    Invoke-ComprehensiveUserCleanup -RegistryOperations $HPUserOps

    # Remove HP Edge Bookmarks from all profile directories
    Get-ChildItem "C:\Users" -Directory | ForEach-Object {
        $EdgePath = "$($_.FullName)\AppData\Local\Microsoft\Edge\User Data\Default"
        if (Test-Path $EdgePath) {
            Remove-Item -Path "$EdgePath\Bookmarks" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$EdgePath\Web Data" -Force -ErrorAction SilentlyContinue
        }
    }

    # 2.3: HP Wolf Security & Bloatware Purge
    Write-Host "[*]   Purging HP Wolf Security via Uninstall registry..." -ForegroundColor $WarnCol
    $WolfNames = @("HP Wolf Security", "HP Wolf Security - Console", "HP Security Update Service")
    foreach ($Name in $WolfNames) {
        $Pkgs = Get-InstalledMsiPackage -NamePattern $Name
        foreach ($p in $Pkgs) { Uninstall-MsiPackage -Package $p }
    }

    $hpAppx = @("*HPJumpStarts*", "*HPPrivacySettings*", "*HPSupportAssistant*", "*HPQuickDrop*", "*myHP*", "*HPEasyClean*", "*HPSmart*")
    foreach ($app in $hpAppx) {
        Get-AppxPackage -AllUsers -Name $app -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
    }
} else {
    Write-Host "[2.1] HP: Not applicable. Skipping." -ForegroundColor $DimCol
}
#endregion

#region 3: DELL
# ============================================================================
if ($Manufacturer -match "Dell") {
    Write-Host "[2.2] Dell Hardware Detected. Commencing Full Purge..." -ForegroundColor DarkYellow

    # 3.1: Services & Tasks
    $dellSvc = @(
        "SupportAssistAgent", "DellHardwareSupport", "DellDigitalDeliveryService",
        "DellOptimizer", "DellClientManagementService", "DellUpdate",
        "KNDBWM", "Killer Network Service", "Killer Selection Service"
    )
    foreach ($s in $dellSvc) {
        if (Get-Service -Name $s -ErrorAction SilentlyContinue) {
            Write-Host "Disabling Service: $s" -ForegroundColor Yellow
            Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
            Set-Service -Name $s -StartupType Disabled
        }
    }

    $dellTasks = @("\Dell", "\Dell\SupportAssist")
    foreach ($tPath in $dellTasks) {
        Get-ScheduledTask -TaskPath "$tPath*" -ErrorAction SilentlyContinue | Disable-ScheduledTask
    }

    # 3.2: Win32 Bloatware Purge
    Write-Host "[*]   Purging Dell Win32 Bloatware via Uninstall registry..." -ForegroundColor $WarnCol
    $DellWin32 = @("Dell SupportAssist*", "Dell Optimizer*", "Dell Digital Delivery*", "Dell Update*", "Dell Customer Connect*", "Dell Help and Support*")
    foreach ($name in $DellWin32) {
        $Pkgs = Get-InstalledMsiPackage -NamePattern $name
        foreach ($p in $Pkgs) { Uninstall-MsiPackage -Package $p }
    }

    # 3.3: Appx Cleanup
    $dellAppx = @(
        "*DellInc.DellDigitalDelivery*", "*DellInc.DellSupportAssist*",
        "*DellInc.DellOptimizer*", "*DellInc.DellPowerManager*",
        "*DellInc.MyDell*", "*DellInc.DellCommandUpdate*",
        "*DellInc.DellRegistration*", "*WavesAudio.WavesMaxxAudio*"
    )
    foreach ($app in $dellAppx) {
        Get-AppxPackage -AllUsers -Name $app -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
    }

    # 3.4: Kill Killer Networking
    $killerDir = "C:\Windows\System32\drivers\RivetNetworks"
    if (Test-Path $killerDir) {
        Remove-Item -Path $killerDir -Recurse -Force -ErrorAction SilentlyContinue
    }
} else {
    Write-Host "[2.2] Dell: Not applicable. Skipping." -ForegroundColor $DimCol
}
#endregion

#region 4: ASUS & ACER
# ============================================================================
if ($Manufacturer -match "ASUS" -or $Manufacturer -match "Acer") {
    Write-Host "[2.3] ASUS/Acer Hardware Detected. Commencing Binary Purge..." -ForegroundColor DarkYellow

    # 4.1: Services & Tasks
    $vendorSvc = @(
        "AsusAppService", "ASUSSystemAnalysis", "ASUSSystemDiagnosis",
        "ArmouryCrateService", "AsusROGLSLService", "ASUSLinkRemote",
        "AcerAgentService", "AcerConfigurationManager", "AcerSvc",
        "AOP_UtilityService", "LiveUpdateSvc"
    )
    foreach ($s in $vendorSvc) {
        if (Get-Service -Name $s -ErrorAction SilentlyContinue) {
            Write-Host "Disabling Service: $s" -ForegroundColor Yellow
            Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
            Set-Service -Name $s -StartupType Disabled
        }
    }

    $vendorTasks = @("\ASUS", "\Acer", "\ASUS\Link to MyASUS", "\Acer\Acer Care Center")
    foreach ($tPath in $vendorTasks) {
        Get-ScheduledTask -TaskPath "$tPath*" -ErrorAction SilentlyContinue | Disable-ScheduledTask
    }

    # 4.2: Win32 Bloatware Purge
    Write-Host "[*]   Purging ASUS/Acer Win32 via Uninstall registry..." -ForegroundColor $WarnCol
    $VendorWin32 = @(
        "Armoury Crate*", "MyASUS*", "ASUS System Control Interface*",
        "Acer Care Center*", "Acer Configuration Manager*", "Acer Portal*",
        "Quick Access*", "AbFiles*", "AOP Framework*"
    )
    foreach ($name in $VendorWin32) {
        $Pkgs = Get-InstalledMsiPackage -NamePattern $name
        foreach ($p in $Pkgs) { Uninstall-MsiPackage -Package $p }
    }

    # 4.3: Appx Cleanup
    $vendorAppx = @(
        "*AsusSystemAnalysis*", "*MyASUS*", "*ArmouryCrate*", "*ASUSGlideX*",
        "*AcerCareCenter*", "*AcerConfigurationManager*", "*AcerQuickAccess*",
        "*AcerUserExperience*", "*AcerProductRegistration*"
    )
    foreach ($app in $vendorAppx) {
        Get-AppxPackage -AllUsers -Name $app -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
    }

    # 4.4: Block BIOS Injection (ASUS Armoury Crate)
    if ($Manufacturer -match "ASUS") {
        $asusGridReg = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Control Panel\Cursors\AsusGrid"
        if (Test-Path $asusGridReg) {
            Write-Host "Blocking ASUS BIOS Grid Auto-Injection..." -ForegroundColor Red
            Remove-Item -Path $asusGridReg -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
} else {
    Write-Host "[2.3] ASUS/Acer: Not applicable. Skipping." -ForegroundColor $DimCol
}
#endregion

#region 4B: LENOVO
# ============================================================================
if ($Manufacturer -match "LENOVO") {
    Write-Host "[2.4] Lenovo Hardware Detected. Commencing Full Purge..." -ForegroundColor DarkYellow

    # Services
    $lenovoSvc = @(
        "ImControllerService", "LenovoVantageService", "LenovoUtilityService",
        "LenovoFnAndFunctionKeys", "LenovoSystemUpdateAddin", "UDClientService"
    )
    foreach ($s in $lenovoSvc) {
        if (Get-Service -Name $s -ErrorAction SilentlyContinue) {
            Stop-Service -Name $s -Force -ErrorAction SilentlyContinue
            Set-Service -Name $s -StartupType Disabled
        }
    }
    Get-ScheduledTask -TaskPath "\Lenovo*" -ErrorAction SilentlyContinue | Disable-ScheduledTask

    # Win32 Bloatware
    Write-Host "[*]   Purging Lenovo Win32 Bloatware via Uninstall registry..." -ForegroundColor $WarnCol
    $LenovoWin32 = @(
        "Lenovo Vantage*", "Lenovo Service Bridge*", "Lenovo System Update*",
        "Lenovo Utility*", "Lenovo Hotkeys*", "Lenovo Now*"
    )
    foreach ($name in $LenovoWin32) {
        $Pkgs = Get-InstalledMsiPackage -NamePattern $name
        foreach ($p in $Pkgs) { Uninstall-MsiPackage -Package $p }
    }

    # Appx Cleanup
    $lenovoAppx = @(
        "*E046963F.LenovoCompanion*", "*LenovoVantage*",
        "*LenovoCompanion*", "*LenovoUtility*", "*LenovoSettings*"
    )
    foreach ($app in $lenovoAppx) {
        Get-AppxPackage -AllUsers -Name $app -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Out-Null
    }
} else {
    Write-Host "[2.4] Lenovo: Not applicable. Skipping." -ForegroundColor $DimCol
}
#endregion

#region 5: THIRD-PARTY CRAPWARE
# ============================================================================
Complete-Step
Write-Step "[3]   DEBLOAT: Scrubbing Common 3rd-Party & Consumer Bloat..."

$crapware = @(
    # Trial Antivirus & Security
    "*McAfee*", "*Norton*", "*Avast*", "*AVG*", "*ExpressVPN*",
    # Social & Streaming
    "*TikTok*", "*Instagram*", "*Facebook*", "*LinkedIn*", "*Netflix*", "*PrimeVideo*", "*Disney*",
    # Games & Consumer Apps
    "*CandyCrush*", "*Roblox*", "*Spotify*", "*SolitaireCollection*", "*WildTangent*", "*ByteDance*",
    # Partner Stubs
    "*Amazon*", "*eBay*", "*Pinterest*", "*Todoist*", "*Clipchamp*", "*MicrosoftNews*"
)

foreach ($item in $crapware) {
    # 1. Remove from all existing user profiles
    Get-AppxPackage -AllUsers -Name $item -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue 2>$null

    # 2. Target the Provisioned (System-wide) package
    $ProvisionedApps = Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $item}
    foreach ($App in $ProvisionedApps) {
        $n = if ($App.DisplayName.Length -gt 50) { $App.DisplayName.Substring(0,47) + "..." } else { $App.DisplayName }
        Write-Host "[*]   Purging: $n..." -ForegroundColor $WarnCol -NoNewline
        try {
            Remove-AppxProvisionedPackage -Online -PackageName $App.PackageName -ErrorAction Stop 2>$null
            Write-Host ""
        }
        catch {
            Write-Host "`r$(" " * $script:Width)`r[!]   DISM: files missing for $n." -ForegroundColor $WarnCol
        }
    }
}

# 5.1: Clean up remaining folders
$crapPaths = @(
    "$env:ProgramData\McAfee",
    "$env:ProgramFiles\Norton Security",
    "C:\Windows\System32\drivers\RivetNetworks"
)
foreach ($path in $crapPaths) {
    if (Test-Path $path) { Remove-Item -Path $path -Recurse -Force -ErrorAction SilentlyContinue }
}
#endregion

Complete-Step

# Footer (reuses header art - art on right, fill on left)
$_ffillW = $script:Width - $_artW - 2  # 1 leading space + 1 sfx char
$_footer = "  WINDOWS LIQUIDATION & HARDENING COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Exit 0
