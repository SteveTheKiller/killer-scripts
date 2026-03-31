<#
.SYNOPSIS
    System Hardening Against Data Exposure (SHADE) v1.0
    Developed by Steve the Killer | Updated: 2026-03-30
.DESCRIPTION
    Comprehensive Windows 10/11 privacy hardening for security-conscious
    environments. Disables location tracking, telemetry services, advertising
    profiling, camera and microphone access, activity history, clipboard
    logging, feedback collection, Delivery Optimization, and network-level
    phone-home behavior. Safe for MSP deployment on Pro, Business, and
    Enterprise SKUs. A reboot is recommended after running.
#>
$_fver = "| v1.0"
#region Pre-Flight Checks
# ============================================================================
# Force UTF-8 output so box-drawing characters render correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding            = [System.Text.Encoding]::UTF8

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit
}

# Helper to handle WMI/CIM switching
function Get-SystemData {
    param([string]$Class)
    return Get-CimInstance -ClassName $Class -ErrorAction SilentlyContinue
}

# Standardized Console Output
$script:StepRow = 0
$script:LastStepMessage = ""

function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    $isDone = $Success -or ($CustomInfo -eq "[SKIPPED]")

    if ($Message -match '^\[[\d.]+/') { $script:LastStepMessage = $Message }
    $printMsg = if ($Message) { $Message } else { $script:LastStepMessage }

    $writeMsg = {
        param([string]$msg, [bool]$done)
        if ($done -and $msg -match '^(\[[\d./]+\])(\s+.+)$') {
            Write-Host $Matches[1] -NoNewline -ForegroundColor DarkGray
            Write-Host $Matches[2] -NoNewline -ForegroundColor White
        } else {
            Write-Host $msg -NoNewline -ForegroundColor Cyan
        }
    }

    if ($Message -and -not $isDone) {
        $script:StepRow = [Console]::CursorTop
        & $writeMsg $printMsg $false
        Write-Host ""
    }
    elseif ($isDone) {
        $currentPos = [Console]::CursorTop
        [Console]::SetCursorPosition(0, $script:StepRow)
        Write-Host (" " * ([Console]::WindowWidth - 1)) -NoNewline
        [Console]::SetCursorPosition(0, $script:StepRow)
        & $writeMsg $printMsg $true

        if ($CustomInfo) {
            if ($CustomInfo -eq "[SKIPPED]") {
                $tag = "[SKIPPED]"
                $currentCol = [Console]::CursorLeft
                $targetCol  = $script:Width - $tag.Length
                if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
                Write-Host $tag -ForegroundColor Yellow
            } else {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Gray
            }
        }

        if ($Success) {
            $tag = "[SUCCESS]"
            $currentCol = [Console]::CursorLeft
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor Green
        }

        if ($currentPos -gt $script:StepRow) {
            [Console]::SetCursorPosition(0, $currentPos)
        }
    }
}

# Environment Setup
$ProgressPreference = 'SilentlyContinue'
$Sys       = Get-SystemData Win32_ComputerSystem
$Baseboard = Get-SystemData Win32_BaseBoard
if ($Sys.Manufacturer -eq $Sys.Model) {
    $ArchitectureDisplay = "$($Baseboard.Manufacturer) $($Baseboard.Product)"
} else {
    $ArchitectureDisplay = "$($Sys.Manufacturer) $($Sys.Model)"
}
$OS     = Get-SystemData Win32_OperatingSystem
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue).DisplayVersion

Clear-Host
$script:Width = 85
$LineCol   = "DarkMagenta"
$MainCol   = "DarkCyan"
$BorderCol = "Magenta"
$ArtCol    = "White"
$AccentCol = "Cyan"
$DimCol    = "DarkGray"
$InfoCol   = "Cyan"

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
$_art1 = "╔═╗ ╦ ╦ ╔═╗ ╦═╗ ╔══ "
$_art2 = "╚═╗ ╠═╣ ╠═╣ ║ ║ ╠═  "
$_art3 = "╚═╝ ╩ ╩ ╩ ╩ ╩═╝ ╚══ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "SYSTEM HARDENING AGAINST DATA EXPOSURE"

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

# System Info Banner
Write-Host "Device Name         : " -ForegroundColor $InfoCol -NoNewline; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor Yellow
Write-Host "System Architecture : " -ForegroundColor $InfoCol -NoNewline; Write-Host "$ArchitectureDisplay" -ForegroundColor Yellow
Write-Host "Operating System    : " -ForegroundColor $InfoCol -NoNewline; Write-Host "$($OS.Caption) ($WinVer)" -ForegroundColor Yellow
Write-HLine -Style dashed
#endregion

# ── Helpers ──────────────────────────────────────────────────────────────────
function Set-RegKeys {
    param([array]$Keys)
    foreach ($K in $Keys) {
        try {
            if (-not (Test-Path $K.Path)) { New-Item -Path $K.Path -Force | Out-Null }
            Set-ItemProperty -Path $K.Path -Name $K.Name -Value $K.Value -Type DWord -Force -ErrorAction Stop | Out-Null
        } catch { }
    }
}

function Disable-Svc {
    param([string]$Name)
    $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if ($svc) {
        Stop-Service  -Name $Name -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Set-Service   -Name $Name -StartupType Disabled -ErrorAction SilentlyContinue
        & sc.exe config $Name start= disabled 2>$null | Out-Null
    }
}

#region 1. Location & Sensors
# ============================================================================
Write-StepUpdate "[01/08] Disabling Location Tracking & Sensors..."

Set-RegKeys @(
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"; Name="Value"; Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\location"; Name="Value"; Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors"; Name="DisableLocation";         Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors"; Name="DisableLocationScripting";Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\LocationAndSensors"; Name="DisableSensors";          Value=1},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search";       Name="SearchPlatformContext";    Value=0}
)
Write-Host "        [!] Stopping Geolocation Service (lfsvc)..." -ForegroundColor Cyan
Disable-Svc "lfsvc"
Write-StepUpdate -Success
#endregion

#region 2. Telemetry & Diagnostic Services
# ============================================================================
Write-StepUpdate "[02/08] Minimizing Telemetry & Diagnostic Services..."

Set-RegKeys @(
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection";                Name="AllowTelemetry";                  Value=0},
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"; Name="AllowTelemetry";                  Value=0},
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"; Name="MaxTelemetryAllowed";             Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection";                Name="DoNotShowFeedbackNotifications";   Value=1},
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting";                Name="Disabled";                        Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting";       Name="Disabled";                        Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows";                     Name="CEIPEnable";                      Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppCompat";                     Name="DisableInventory";                Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppCompat";                     Name="DisableUAR";                      Value=1}
)
foreach ($svc in @("DiagTrack", "dmwappushservice", "WerSvc", "PcaSvc")) {
    Write-Host "        [!] Stopping $svc..." -ForegroundColor Cyan
    Disable-Svc $svc
}
Write-StepUpdate -Success
#endregion

#region 3. Advertising & User Profiling
# ============================================================================
Write-StepUpdate "[03/08] Disabling Advertising & User Profiling..."

Set-RegKeys @(
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo";            Name="Enabled";                                       Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo";                  Name="DisabledByGroupPolicy";                         Value=1},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy";                    Name="TailoredExperiencesWithDiagnosticDataEnabled";   Value=0},
    @{Path="HKCU:\Control Panel\International\User Profile";                             Name="HttpAcceptLanguageOptOut";                       Value=1},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-338387Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-338388Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-338389Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-353694Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-353696Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-338393Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SubscribedContent-310093Enabled";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SystemPaneSuggestionsEnabled";                   Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SoftLandingEnabled";                             Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="OemPreInstalledAppsEnabled";                     Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="PreInstalledAppsEnabled";                        Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager";     Name="SilentInstalledAppsEnabled";                     Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search";                     Name="BingSearchEnabled";                              Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search";                     Name="CortanaConsent";                                 Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search";                     Name="CanCortanaSeeBrowserHistory";                    Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search";                   Name="DisableWebSearch";                               Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search";                   Name="ConnectedSearchUseWeb";                          Value=0},
    @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata";            Name="PreventDeviceMetadataFromNetwork";               Value=1}
)
Write-StepUpdate -Success
#endregion

#region 4. Camera & Microphone
# ============================================================================
Write-StepUpdate "[04/08] Denying Camera & Microphone Access..."

# ConsentStore uses string "Deny" not DWORD
foreach ($p in @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone",
    "HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"
)) {
    if (-not (Test-Path $p)) { New-Item -Path $p -Force | Out-Null }
    Set-ItemProperty -Path $p -Name "Value" -Value "Deny" -Force -ErrorAction SilentlyContinue
}

Set-RegKeys @(
    # Value=1: User decides — allows Teams/Zoom to prompt for permission normally
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy"; Name="LetAppsAccessCamera";     Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy"; Name="LetAppsAccessMicrophone"; Value=1}
)
Write-StepUpdate -Success
#endregion

#region 5. Activity History, Clipboard & App Tracking
# ============================================================================
Write-StepUpdate "[05/08] Disabling Activity History, Clipboard & App Tracking..."

Set-RegKeys @(
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System";                          Name="EnableActivityFeed";                  Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System";                          Name="PublishUserActivities";               Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System";                          Name="UploadUserActivities";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Clipboard";                                        Name="EnableClipboardHistory";              Value=0},
    @{Path="HKCU:\Software\Microsoft\Clipboard";                                        Name="EnableClipboardRoaming";              Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System";                          Name="AllowClipboardHistory";               Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\System";                          Name="AllowCrossDeviceClipboard";           Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced";         Name="Start_TrackProgs";                    Value=0},
    @{Path="HKCU:\Software\Microsoft\Input\TIPC";                                       Name="Enabled";                             Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization";                    Name="AllowInputPersonalization";           Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization";                    Name="RestrictImplicitInkCollection";       Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\InputPersonalization";                    Name="RestrictImplicitTextCollection";      Value=1}
)
Write-StepUpdate -Success
#endregion

#region 6. Feedback, Spotlight & Diagnostics UI
# ============================================================================
Write-StepUpdate "[06/08] Suppressing Feedback, Spotlight & Diagnostics UI..."

Set-RegKeys @(
    @{Path="HKCU:\Software\Microsoft\Siuf\Rules";                                       Name="NumberOfSIUFInPeriod";                Value=0},
    @{Path="HKCU:\Software\Microsoft\Siuf\Rules";                                       Name="PeriodInNanoSeconds";                 Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection";                  Name="DoNotShowFeedbackNotifications";      Value=1},
    @{Path="HKCU:\Software\Microsoft\Speech_OneCore\Settings\OnlineSpeechPrivacy";      Name="HasAccepted";                         Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\FindMyDevice";                            Name="AllowFindMyDevice";                   Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent";                    Name="DisableWindowsSpotlightFeatures";     Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent";                    Name="DisableWindowsConsumerFeatures";      Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent";                    Name="DisableSoftLanding";                  Value=1}
)
Write-StepUpdate -Success
#endregion

#region 7. Delivery Optimization
# ============================================================================
Write-StepUpdate "[07/08] Disabling Delivery Optimization (P2P Upload)..."

Set-RegKeys @(
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DeliveryOptimization";                        Name="DODownloadMode"; Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\DeliveryOptimization\Settings";         Name="DownloadMode";   Value=0}
)
Write-Host "        [!] Setting Delivery Optimization service (DoSvc) to Manual..." -ForegroundColor Cyan
$dosvc = Get-Service -Name "DoSvc" -ErrorAction SilentlyContinue
if ($dosvc) {
    Stop-Service  -Name "DoSvc" -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Set-Service   -Name "DoSvc" -StartupType Manual -ErrorAction SilentlyContinue
}

$DOCache = "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
if (Test-Path $DOCache) {
    Remove-Item "$DOCache\*" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "        [!] Delivery Optimization cache purged." -ForegroundColor Cyan
}
Write-StepUpdate -Success
#endregion

#region 8. Network Phone-Home & LLMNR
# ============================================================================
Write-StepUpdate "[08/08] Blocking Network Phone-Home & LLMNR..."

Set-RegKeys @(
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\NetworkConnectivityStatusIndicator"; Name="NoActiveProbe";          Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient";                       Name="EnableMulticast";         Value=0},
    @{Path="HKLM:\SYSTEM\CurrentControlSet\Services\NetBT\Parameters";                     Name="NodeType";                Value=2},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WCN\UI";                             Name="DisableWcnUi";            Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Peernet";                                    Name="Disabled";                Value=1}
)

$wlan = Get-Service -Name "WlanSvc" -ErrorAction SilentlyContinue
if ($wlan) {
    Set-RegKeys @(
        @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Wireless\GPTWirelessPolicy"; Name="EnableAutoConfig"; Value=0}
    )
    Write-Host "        [!] WLAN AutoConfig reporting suppressed." -ForegroundColor Cyan
}
Write-StepUpdate -Success
#endregion

#region Final Summary
# ============================================================================
Write-HLine -Style dashed
Write-Host "Reboot Recommended  : " -ForegroundColor $InfoCol -NoNewline
Write-Host "Camera/mic deny and LLMNR changes require a restart." -ForegroundColor DarkYellow

# Footer
$_sfx  = "█"
$_ftr1 = $_art1; $_ftr2 = $_art2; $_ftr3 = $_art3
$_ftrW = $_artW
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  HARDENING COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr1 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host $_ftr2 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr3 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
#endregion