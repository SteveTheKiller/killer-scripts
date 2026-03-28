<#
.NAME
    Deployment Environment Bloatware Liquidator & Optimized Automated Toolkit (D.E.B.L.O.A.T.) v2.0
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Standardizes Windows 11 by removing OEM bloat (HP, Dell,
    ASUS/Acer), AI/Recall features, and sponsored consumer content.
    Hardens privacy via telemetry caps, Edge policy enforcement, and
    taskbar/Start menu lockdown applied across all user profiles
    including the Default User template.
#>

#region 0: INITIALIZATION AND HELPER FUNCTIONS
# ============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "DEBLOAT requires Administrative privileges. Please relaunch as Admin."
    exit
}

Clear-Host
$script:Width = 85

$LineCol   = "White"
$MainCol   = "Yellow"
$WarnCol   = "DarkYellow"
$ArtCol    = "DarkRed"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$InfoCol   = "Cyan"
$OkCol     = "Green"
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

# Header Art & Logic
$_pfx  = "█  "
$_art1 = "╔═╗ ╔═╗ ╔╗  ╦   ╔═╗ ╔═╗ ╔╦╗ "
$_art2 = "║ ║ ║╣  ╠╩╗ ║   ║ ║ ╠═╣  ║  "
$_art3 = "╩═╝ ╚═╝ ╩═╝ ╩═╝ ╚═╝ ╩ ╩  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "DEPLOYMENT ENV BLOAT LIQUIDATOR & OPTIMIZED TOOLKIT"
$_ver   = "| v2.0"

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
$Manufacturer = (Get-CimInstance -ClassName Win32_ComputerSystem).Manufacturer

# --- Step Indicator Functions ---
$script:StepRow = -1
$script:StepMsg = ""

function Write-Step {
    param([string]$Msg)
    $script:StepRow = [Console]::CursorTop
    $script:StepMsg = $Msg
    Write-Host $Msg -ForegroundColor $InfoCol
}

function Complete-Step {
    $savedTop = [Console]::CursorTop
    if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
        [Console]::SetCursorPosition(0, $script:StepRow)
        $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - "[SUCCESS]".Length)
        Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
        Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
        [Console]::SetCursorPosition(0, $savedTop)
    }
    $script:StepRow = -1
}

# --- Helper Functions ---
function Invoke-ComprehensiveUserCleanup {
    param([scriptblock]$RegistryOperations)
    
    # 1. Target the Default User (Template for all FUTURE users)
    Write-Host "[*]   Updating Default User Template..." -ForegroundColor $WarnCol
    reg load HKU\DefaultUser "C:\Users\Default\NTUSER.DAT" | Out-Null
    Invoke-Command -ScriptBlock $RegistryOperations -ArgumentList "DefaultUser"
    [gc]::Collect(); [gc]::WaitForPendingFinalizers()
    reg unload HKU\DefaultUser | Out-Null

    # 2. Target all EXISTING Users
    $UserFolders = Get-ChildItem "C:\Users" -Directory | Where-Object { $_.Name -notmatch "Public|Default|All Users" }
    foreach ($Folder in $UserFolders) {
        $NTUserPath = "$($Folder.FullName)\NTUSER.DAT"
        if (Test-Path $NTUserPath) {
            $HiveName = "TempHive_$($Folder.Name)"
            Write-Host "[*]   Cleaning Profile: $($Folder.Name)..." -ForegroundColor $WarnCol
            
            # Load, Execute, and Unload
            reg load "HKU\$HiveName" $NTUserPath | Out-Null
            Invoke-Command -ScriptBlock $RegistryOperations -ArgumentList $HiveName
            [gc]::Collect(); [gc]::WaitForPendingFinalizers()
            reg unload "HKU\$HiveName" | Out-Null
        }
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

# 1.3: Disable Advertising ID & Tailored Experiences
$privacyPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy"
Set-ItemProperty -Path $privacyPath -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0
$genPath = "HKCU:\Control Panel\International\User Profile"
Set-ItemProperty -Path $genPath -Name "HttpAcceptLanguageOptOut" -Value 1

# 1.4: Disable Activity History & Clipboard Sync
$sysPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"
if (!(Test-Path $sysPath)) { New-Item -Path $sysPath -Force | Out-Null }
Set-ItemProperty -Path $sysPath -Name "PublishUserActivities" -Value 0
Set-ItemProperty -Path $sysPath -Name "EnableActivityFeed"    -Value 0
Set-ItemProperty -Path $sysPath -Name "UploadUserActivities"  -Value 0

# 1.5: Disable Bing Search (Web Results) & Search Highlights
# ------------------------------------------------------------------------------
# 1.5.1. Turn off Search Suggestions (Bing Web Results)
$searchPath = "HKCU:\Software\Policies\Microsoft\Windows\Explorer"
if (!(Test-Path $searchPath)) { New-Item -Path $searchPath -Force | Out-Null }
Set-ItemProperty -Path $searchPath -Name "DisableSearchBoxSuggestions" -Value 1 -ErrorAction SilentlyContinue
# 1.5.2. Turn off Search Highlights (The random icons/pictures in the search bar)
$shPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\SearchSettings"
if (!(Test-Path $shPath)) { New-Item -Path $shPath -Force | Out-Null }
# This keeps the box but removes the daily "Doodle" icons
Set-ItemProperty -Path $shPath -Name "IsDynamicSearchBoxPresent" -Value 0 -ErrorAction SilentlyContinue
# 1.5.3. Explicitly disable web search (Keep local search only)
$webSearchPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"
if (!(Test-Path $webSearchPath)) { New-Item -Path $webSearchPath -Force | Out-Null }
Set-ItemProperty -Path $webSearchPath -Name "BingSearchEnabled" -Value 0 -ErrorAction SilentlyContinue

# 1.6: Global OEM Re-injection & Content Prevention
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata" -Name "PreventDeviceMetadataFromNetwork" -Value 1
$cdmPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
if (!(Test-Path $cdmPath)) { New-Item -Path $cdmPath -Force | Out-Null }
Set-ItemProperty -Path $cdmPath -Name "DisableWindowsConsumerFeatures" -Value 1

# 1.7: Disable Lock Screen "Fun Facts" and Taskbar Widgets
$cdmUserPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
if (!(Test-Path $cdmUserPath)) { New-Item -Path $cdmUserPath -Force | Out-Null }
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338387Enabled" -Value 0 
Set-ItemProperty -Path $cdmUserPath -Name "RotatingLockScreenOverlayEnabled" -Value 0
# Disable Widgets (News) and Chat (Teams Consumer) icons
$webPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
# TaskbarDa = Widgets | TaskbarMn = Chat
Set-ItemProperty -Path $webPath -Name "TaskbarDa"      -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $webPath -Name "TaskbarMn"      -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $webPath -Name "Start_TrackProgs" -Value 0 -ErrorAction SilentlyContinue
# Chat/Teams icon policy & News/Feeds
$chatPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
if (!(Test-Path $chatPolicyPath)) { New-Item -Path $chatPolicyPath -Force | Out-Null }
Set-ItemProperty -Path $chatPolicyPath -Name "ChatIcon" -Value 3
$feedsPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
if (!(Test-Path $feedsPath)) { New-Item -Path $feedsPath -Force | Out-Null }
Set-ItemProperty -Path $feedsPath -Name "EnableFeeds" -Value 0
# Additional ContentDelivery Manager keys
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338387Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338388Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338389Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-353694Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-353696Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338393Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-310093Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SystemPaneSuggestionsEnabled"    -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SoftLandingEnabled"              -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "RotatingLockScreenEnabled"       -Value 0 -ErrorAction SilentlyContinue
# Windows Spotlight & soft landing
Set-ItemProperty -Path $cdmPath -Name "DisableWindowsSpotlightFeatures" -Value 1
Set-ItemProperty -Path $cdmPath -Name "DisableSoftLanding"              -Value 1
# Start menu recommendations
$startPolPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
if (!(Test-Path $startPolPath)) { New-Item -Path $startPolPath -Force | Out-Null }
Set-ItemProperty -Path $startPolPath -Name "HideRecentlyAddedApps"  -Value 1
$explorerPolPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"
if (!(Test-Path $explorerPolPath)) { New-Item -Path $explorerPolPath -Force | Out-Null }
# Notification suggestions
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested" -Name "Enabled" -Value 0 -ErrorAction SilentlyContinue

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

# 1.10: General Performance & Explorer Tweaks
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

# Set Explorer to open to "This PC" & Disable Hibernation
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -Value 1 -ErrorAction SilentlyContinue
powercfg /hibernate off

# 1.11: Office Language Cleanup (Purge non-English stubs)
Write-Host "[*]   Checking for non-English Office Language Packs..." -ForegroundColor $WarnCol
$C2RPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path $C2RPath) {
    $RegKeys = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" | Where-Object { $_.DisplayName -like "Microsoft 365*" -and $_.DisplayName -notmatch "en-us" }
    foreach ($Key in $RegKeys) {
        if ($Key.DisplayName -match " - ([a-z]{2}-[a-z]{2})") {
            $LangToRemove = $Matches[1]
            Write-Host "Removing: $LangToRemove" -ForegroundColor $DimCol
            Start-Process $C2RPath -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365ProPlusRetail.16_$($LangToRemove)_x-none culture=$LangToRemove version.16=16.0 DisplayLevel=False" -Wait
        }
    }
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

    # 2.2: The "Pre-emptive Strike" (Targeting New and Existing Users)
    Write-Host "[*]   Cleaning User Hives for HP..." -ForegroundColor $WarnCol
    Invoke-ComprehensiveUserCleanup -RegistryOperations {
        param($Hive)
        # Remove HP-specific pinning and content delivery keys
        $PathsToScrub = @(
            "HKU\$Hive\Software\Microsoft\Windows\CurrentVersion\Explorer\StartPage2",
            "HKU\$Hive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager",
            "HKU\$Hive\Software\Microsoft\Windows\CurrentVersion\CloudStore"
        )

        foreach ($p in $PathsToScrub) {
            if (Test-Path $p) {
                Write-Host "Removing $p..." -ForegroundColor $DimCol
                Remove-Item -Path $p -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        # Prevent HP from pushing consumer apps
        $CDManager = "HKU\$Hive\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
        if (Test-Path $CDManager) {
            Set-ItemProperty -Path $CDManager -Name "PreInstalledAppsEnabled" -Value 0 -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $CDManager -Name "ContentDeliveryAllowed" -Value 0 -ErrorAction SilentlyContinue
            Set-ItemProperty -Path $CDManager -Name "SilentInstalledAppsEnabled" -Value 0 -ErrorAction SilentlyContinue
        }
    }

    # Remove HP Edge Bookmarks from all profile directories
    Get-ChildItem "C:\Users" -Directory | ForEach-Object {
        $EdgePath = "$($_.FullName)\AppData\Local\Microsoft\Edge\User Data\Default"
        if (Test-Path $EdgePath) {
            Remove-Item -Path "$EdgePath\Bookmarks" -Force -ErrorAction SilentlyContinue
            Remove-Item -Path "$EdgePath\Web Data" -Force -ErrorAction SilentlyContinue
        }
    }

    # 2.3: HP Wolf Security & Bloatware Purge
    Write-Host "Purging HP Wolf Security via MSI GUID..." -ForegroundColor Yellow
    $WolfNames = @("HP Wolf Security", "HP Wolf Security - Console", "HP Security Update Service")
    foreach ($Name in $WolfNames) {
        $Pkg = Get-Package | Where-Object { $_.Name -eq $Name }
        if ($Pkg -and $Pkg.FastPackageId) {
            msiexec.exe /x $($Pkg.FastPackageId.Split('|')[0]) /qn /norestart
        }
    }

    $hpAppx = @("*HPJumpStarts*", "*HPPrivacySettings*", "*HPSupportAssistant*", "*HPQuickDrop*", "*myHP*", "*HPEasyClean*", "*HPSmart*")
    foreach ($app in $hpAppx) {
        Get-AppxPackage -AllUsers -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online
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
    Write-Host "Purging Dell Win32 Bloatware via MSI..." -ForegroundColor Yellow
    $DellWin32 = @("Dell SupportAssist*", "Dell Optimizer*", "Dell Digital Delivery*", "Dell Update*", "Dell Customer Connect*", "Dell Help and Support*")
    foreach ($name in $DellWin32) {
        $Pkgs = Get-Package -Name $name -ErrorAction SilentlyContinue
        if ($Pkgs) {
            foreach ($p in $Pkgs) {
                if ($p.FastPackageId) {
                    Write-Host "Uninstalling $($p.Name)..."
                    msiexec.exe /x $($p.FastPackageId.Split('|')[0]) /qn /norestart
                }
            }
        }
    }

    # 3.3: Appx Cleanup
    $dellAppx = @(
        "*DellInc.DellDigitalDelivery*", "*DellInc.DellSupportAssist*", 
        "*DellInc.DellOptimizer*", "*DellInc.DellPowerManager*", 
        "*DellInc.MyDell*", "*DellInc.DellCommandUpdate*",
        "*DellInc.DellRegistration*", "*WavesAudio.WavesMaxxAudio*"
    )
    foreach ($app in $dellAppx) {
        Get-AppxPackage -AllUsers -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online
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
    Write-Host "Purging ASUS/Acer Win32 Bloatware..." -ForegroundColor Yellow
    $VendorWin32 = @(
        "Armoury Crate*", "MyASUS*", "ASUS System Control Interface*", 
        "Acer Care Center*", "Acer Configuration Manager*", "Acer Portal*",
        "Quick Access*", "AbFiles*", "AOP Framework*"
    )
    foreach ($name in $VendorWin32) {
        $Pkgs = Get-Package -Name $name -ErrorAction SilentlyContinue
        if ($Pkgs) {
            foreach ($p in $Pkgs) {
                if ($p.FastPackageId) {
                    Write-Host "Uninstalling $($p.Name)..."
                    msiexec.exe /x $($p.FastPackageId.Split('|')[0]) /qn /norestart
                }
            }
        }
    }

    # 4.3: Appx Cleanup
    $vendorAppx = @(
        "*AsusSystemAnalysis*", "*MyASUS*", "*ArmouryCrate*", "*ASUSGlideX*",
        "*AcerCareCenter*", "*AcerConfigurationManager*", "*AcerQuickAccess*",
        "*AcerUserExperience*", "*AcerProductRegistration*"
    )
    foreach ($app in $vendorAppx) {
        Get-AppxPackage -AllUsers -Name $app | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online | Where-Object {$_.DisplayName -like $app} | Remove-AppxProvisionedPackage -Online
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