<#
.SYNOPSIS
    System Hardening Against Data Exposure (SHADE) v2.1
    Developed by Steve the Killer | Updated: 2026-11-18
.DESCRIPTION
    Comprehensive Windows 10/11 privacy hardening for security-conscious
    environments. Disables location tracking, telemetry services, advertising
    profiling, camera and microphone access, activity history, clipboard
    logging, feedback collection, Delivery Optimization, and network-level
    phone-home behavior. Safe for MSP deployment on Pro, Business, and
    Enterprise SKUs. A reboot is recommended after running.
#>
$_fver = "| v2.1"

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

function Test-IsSystem {
    [Security.Principal.WindowsIdentity]::GetCurrent().IsSystem
}

function Get-LoadedUserSids {
    # SIDs of real user accounts whose registry hive is loaded, meaning they are signed in.
    Get-ChildItem -Path "Registry::HKEY_USERS" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match '^S-1-5-21-\d+-\d+-\d+-\d+$' } |
        ForEach-Object { $_.PSChildName }
}

function Get-UserHiveRoots {
    # Where per-user (HKCU) settings must be written. As SYSTEM, HKCU is the SYSTEM profile, which
    # nobody uses, so each signed-in user's loaded hive is targeted instead.
    if (Test-IsSystem) {
        return @(Get-LoadedUserSids | ForEach-Object { "Registry::HKEY_USERS\$_" })
    }
    return @("HKCU:")
}

# Helper to handle WMI/CIM switching
function Get-SystemData {
    param([string]$Class)
    return Get-CimInstance -ClassName $Class -ErrorAction SilentlyContinue
}

# Standardized Console Output
$script:StepRow = 0
$script:LastStepMessage = ""

# Console access is guarded. With output redirected (RMM, Kaseya LiveConnect) the console handle
# is invalid, so every call falls back to plain line output instead of printing errors.
function Get-ShadeConsoleValue {
    param([string]$Name)
    [Console]::$Name
}
function Set-ShadeConsolePosition {
    param([int]$Left, [int]$Top)
    [Console]::SetCursorPosition($Left, $Top)
}
function Set-ShadeConsoleUtf8 {
    [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
}
function Get-ShadeCursorTop   { try { [int](Get-ShadeConsoleValue 'CursorTop') }   catch { -1 } }
function Get-ShadeCursorLeft  { try { [int](Get-ShadeConsoleValue 'CursorLeft') }  catch { -1 } }
function Get-ShadeWindowWidth { try { [int](Get-ShadeConsoleValue 'WindowWidth') } catch { $script:Width + 1 } }
function Set-ShadeCursor {
    param([int]$Left, [int]$Top)
    try { Set-ShadeConsolePosition $Left $Top; $true } catch { $false }
}
function Set-ShadeUtf8Output { try { Set-ShadeConsoleUtf8 } catch { } }
function Clear-ShadeScreen   { try { Clear-Host -ErrorAction Stop } catch { } }

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
        $script:StepRow = Get-ShadeCursorTop
        & $writeMsg $printMsg $false
        Write-Host ""
    }
    elseif ($isDone) {
        $currentPos = Get-ShadeCursorTop
        # Rewrite the pending line in place when the console allows it. Otherwise print the
        # completed line on a new line, so every step still shows in a redirected session.
        $inPlace = ($currentPos -ge 0) -and ($script:StepRow -ge 0)
        if ($inPlace) { $inPlace = Set-ShadeCursor 0 $script:StepRow }
        if ($inPlace) {
            Write-Host (" " * [Math]::Max(0, (Get-ShadeWindowWidth) - 1)) -NoNewline
            [void](Set-ShadeCursor 0 $script:StepRow)
        }
        & $writeMsg $printMsg $true
        $col = $printMsg.Length

        if ($CustomInfo) {
            if ($CustomInfo -eq "[SKIPPED]") {
                $tag = "[SKIPPED]"
                $currentCol = if ($inPlace) { Get-ShadeCursorLeft } else { -1 }
                if ($currentCol -lt 0) { $currentCol = $col }
                $targetCol  = $script:Width - $tag.Length
                if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
                Write-Host $tag -ForegroundColor Yellow
            } else {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Gray
                $col += 1 + $CustomInfo.Length
            }
        }

        if ($Success) {
            $tag = "[SUCCESS]"
            $currentCol = if ($inPlace) { Get-ShadeCursorLeft } else { -1 }
            if ($currentCol -lt 0) { $currentCol = $col }
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor Green
        }

        if ($inPlace -and $currentPos -gt $script:StepRow) {
            [void](Set-ShadeCursor 0 $currentPos)
        }
    }
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

# ── Helpers ──────────────────────────────────────────────────────────────────
function Add-ShadeFailure {
    param([string]$What)
    [void]$script:ShadeFailures.Add($What)
}

function Set-RegKeys {
    # Writes each value. HKCU paths follow Get-UserHiveRoots, so SYSTEM or RMM runs reach the
    # signed-in users instead of the SYSTEM profile. Failures and skips are counted, never hidden.
    param([array]$Keys)
    foreach ($K in $Keys) {
        $paths = @()
        if ($K.Path -like 'HKCU:*') {
            $roots = @(Get-UserHiveRoots)
            if ($roots.Count -eq 0) { $script:ShadeSkipped++; continue }
            foreach ($r in $roots) { $paths += ($r + $K.Path.Substring(5)) }
        } else {
            $paths = @($K.Path)
        }
        $type = if ($K.Type) { $K.Type } else { 'DWord' }
        foreach ($p in $paths) {
            try {
                if (-not (Test-Path $p)) { New-Item -Path $p -Force -ErrorAction Stop | Out-Null }
                Set-ItemProperty -Path $p -Name $K.Name -Value $K.Value -Type $type -Force -ErrorAction Stop | Out-Null
                $script:ShadeApplied++
            } catch {
                Add-ShadeFailure ("{0} : {1} ({2})" -f $p, $K.Name, $_.Exception.Message)
            }
        }
    }
}

function Invoke-ScDisable {
    param([string]$Name)
    & sc.exe config $Name start= disabled 2>$null | Out-Null
    return $LASTEXITCODE
}

function Disable-Svc {
    param([string]$Name)
    $svc = Get-Service -Name $Name -ErrorAction SilentlyContinue
    if ($svc) {
        Stop-Service  -Name $Name -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        $ok = $true
        $err = ''
        try { Set-Service -Name $Name -StartupType Disabled -ErrorAction Stop } catch { $ok = $false; $err = $_.Exception.Message }
        $rc = [int](Invoke-ScDisable $Name)
        if (-not $ok -and $rc -ne 0) {
            Add-ShadeFailure ("service {0} could not be disabled ({1}; sc.exe exit {2})" -f $Name, $err, $rc)
        } else {
            $script:ShadeApplied++
        }
    }
}

function Invoke-ShadeMain {
    $script:ShadeFailures = New-Object System.Collections.ArrayList
    $script:ShadeApplied  = 0
    $script:ShadeSkipped  = 0
#region Pre-Flight Checks
# ============================================================================
    # Force UTF-8 output so box-drawing characters render correctly
    Set-ShadeUtf8Output
    $OutputEncoding            = [System.Text.Encoding]::UTF8


    if (-not (Test-IsAdmin)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Exit-KillerScript 1
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

    Clear-ShadeScreen
    $script:Width = 85
    $LineCol   = "DarkMagenta"
    $MainCol   = "DarkCyan"
    $BorderCol = "Magenta"
    $ArtCol    = "White"
    $AccentCol = "Cyan"
    $DimCol    = "DarkGray"
    $InfoCol   = "Cyan"


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
    # dmwappushservice is left alone on purpose: MDM push and device sync depend on it.
    foreach ($svc in @("DiagTrack", "WerSvc", "PcaSvc")) {
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
    Set-RegKeys @(
        @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam";     Name="Value"; Value="Deny"; Type="String"},
        @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\webcam";     Name="Value"; Value="Deny"; Type="String"},
        @{Path="HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"; Name="Value"; Value="Deny"; Type="String"},
        @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\CapabilityAccessManager\ConsentStore\microphone"; Name="Value"; Value="Deny"; Type="String"}
    )

    Set-RegKeys @(
        # Policy values: 0 = user in control (apps may prompt, so Teams and Zoom work), 1 = force allow, 2 = force deny.
        # 0 leaves the choice to the user. The ConsentStore "Deny" above is the only deny that SHADE sets.
        @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy"; Name="LetAppsAccessCamera";     Value=0},
        @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy"; Name="LetAppsAccessMicrophone"; Value=0}
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

    # Effective outcome. The final line always agrees with the exit code.
    if (Test-IsSystem) {
        $sids = @(Get-LoadedUserSids)
        Write-Output ("Per-user settings   : running as SYSTEM, applied to {0} signed-in user(s)" -f $sids.Count)
        if ($sids.Count -eq 0) {
            Write-Output "                      No signed-in user found. Per-user values were skipped and do not apply to anyone."
        } else {
            Write-Output "                      Users who sign in later do not get the per-user values until SHADE is run again."
        }
    } else {
        Write-Output ("Per-user settings   : applied to the executing user ({0})" -f [Security.Principal.WindowsIdentity]::GetCurrent().Name)
    }
    Write-Output ("Values applied      : {0}" -f $script:ShadeApplied)
    if ($script:ShadeSkipped -gt 0) { Write-Output ("Values skipped      : {0} per-user value(s), no signed-in user to apply them to" -f $script:ShadeSkipped) }
    foreach ($f in $script:ShadeFailures) { Write-Output ("[FAILED] {0}" -f $f) }

    Exit-KillerScript 0
}

if (-not $global:KillerScriptsTestMode) { Invoke-ShadeMain }
