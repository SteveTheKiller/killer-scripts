<#
.SYNOPSIS
    Advanced Maintenance, Optimization, and Repair Tool (AMORT) v15.0
    Developed by SteveTheKiller | Updated: 2026-03-20
.DESCRIPTION
    Automated Windows 10/11 tune-up for MSP field and remote deployment.
    Hardens AI, privacy, and browser settings; strips OEM and consumer
    bloat; purges browser and system caches; resets the Windows Update
    database; runs DISM and SFC repair; and performs SSD TRIM while
    reporting disk space recovered at each stage.
#>
$_fver   = "| v15.0"
#region Pre-Flight Checks
# ============================================================================
# Force UTF-8 output so box-drawing characters render correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding            = [System.Text.Encoding]::UTF8

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit
}
# Kill any stuck processes from a previous run
$StaleProcs = @("DISM", "sfc", "cleanmgr", "TiWorker")
foreach ($ProcName in $StaleProcs) {
    $Found = Get-Process -Name $ProcName -ErrorAction SilentlyContinue
    if ($Found) {
        Write-Host "[Pre-Flight] Stopping stuck process: $ProcName" -ForegroundColor Yellow
        $Found | Stop-Process -Force -ErrorAction SilentlyContinue
    }
} 

# Helper to handle WMI/CIM switching
function Get-SystemData {
    param([string]$Class)
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        # PowerShell 6/7 MUST use CIM
        return Get-CimInstance -ClassName $Class -ErrorAction SilentlyContinue
    } else {
        # PowerShell 5.1 can use either; CIM is preferred
        return Get-CimInstance -ClassName $Class -ErrorAction SilentlyContinue
    }
}
# Standardized Console Output
$script:StepRow = 0
$script:LastStepMessage = ""

function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [switch]$Reprint, [string]$CustomInfo)
    $isDone = $Success -or ($CustomInfo -eq "[SKIPPED]")
    
    # Store the header message
    if ($Message -match '^\[[\d.]+/') { $script:LastStepMessage = $Message }
    $printMsg = if ($Message) { $Message } else { $script:LastStepMessage }

    # Coloring Logic
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
        # STARTING A STEP: Save the current cursor row so we can return to it later
        $script:StepRow = [Console]::CursorTop
        & $writeMsg $printMsg $false
        Write-Host "" # Move cursor to next line so WARNINGS have a place to go
    } 
    elseif ($isDone) {
        # COMPLETING A STEP: Jump back to the saved row to overwrite the Cyan text
        $currentPos = [Console]::CursorTop
        [Console]::SetCursorPosition(0, $script:StepRow)
        
        # Clear the original Cyan line
        Write-Host (" " * ([Console]::WindowWidth - 1)) -NoNewline
        [Console]::SetCursorPosition(0, $script:StepRow)
        
        # Reprint the line in the "Done" (Gray/White) style
        & $writeMsg $printMsg $true

        # Add Custom Info (Saved MB/GB)
        if ($CustomInfo) {
            if ($CustomInfo -eq "[SKIPPED]") {
                $tag = "[SKIPPED]"
                $currentCol = [Console]::CursorLeft
                $targetCol  = $script:Width - $tag.Length
                if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
                Write-Host $tag -ForegroundColor Yellow
            } elseif ($CustomInfo.StartsWith("(Saved:")) {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Red
            } else {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Gray
            }
        }

        # Final Success Tag (right-aligned to console width)
        if ($Success) {
            $tag = "[SUCCESS]"
            $currentCol = [Console]::CursorLeft
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor Green
        }

        # Return the cursor to where it was (below any warnings that appeared)
        if ($currentPos -gt $script:StepRow) {
            [Console]::SetCursorPosition(0, $currentPos)
        #} else {
        #    Write-Host ""
        }
    }
}
# Service Management Helper
function Start-ServiceSilent {
    param([string]$ServiceName)
    Start-Service $ServiceName -WarningAction SilentlyContinue -ErrorAction SilentlyContinue
    $Timer = 0
    while ((Get-Service $ServiceName).Status -ne 'Running' -and $Timer -lt 15) {
        Start-Sleep -Seconds 1
        $Timer++
    }
}
# Environment Setup
$IsCore = ($PSVersionTable.PSEdition -eq "Core")
$Drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$StartSpace = $Drive.FreeSpace
$TotalSize = $Drive.Size
# Initialize cumulative yield (bytes)
if (-not $TotalYieldBytes) { $TotalYieldBytes = 0 }
# Ensure TotalSize is valid
$TotalSize = [double]$TotalSize
$script:RegionHistory = @()
if ($TotalSize -le 0) { throw "TotalSize is zero or undefined. Aborting." }

$StartUsagePct = [Math]::Round(((($TotalSize - $StartSpace) / $TotalSize) * 100), 2)
$LastRegionSpace = $Drive.FreeSpace # Rolling baseline for step-by-step reporting
$CS = Get-CimInstance Win32_ComputerSystem
$Vendor = $CS.Manufacturer
$IsVM = ($Vendor -match "QEMU|VMware|Virtual|Hyper-V")
# --- UPDATED LOGIC FOR CUSTOM BUILDS ---
$Sys = Get-SystemData Win32_ComputerSystem
$Baseboard = Get-SystemData Win32_BaseBoard
# Rule: If Manufacturer and Model are the same (typical of "To Be Filled By O.E.M."), 
# fallback to Motherboard Manufacturer and Product.
if ($Sys.Manufacturer -eq $Sys.Model) {
    $ArchitectureDisplay = "$($Baseboard.Manufacturer) $($Baseboard.Product)"
} else {
    $ArchitectureDisplay = "$($Sys.Manufacturer) $($Sys.Model)"
}
$OS = Get-SystemData Win32_OperatingSystem
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue).DisplayVersion
$CS = Get-SystemData Win32_LogicalDisk | Where-Object { $_.DeviceID -eq 'C:' }
# Suppress standard progress bars for speed in RMM/VSA
$ProgressPreference = 'SilentlyContinue'

Clear-Host
$script:Width    = 85
$LineCol   = "DarkCyan"
$MainCol   = "DarkYellow"
$BorderCol = "Cyan"
$ArtCol    = "White"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$InfoCol      = "Cyan"
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
$_art1 = "╔═╗ ╔╦╗ ╔═╗ ╦═╗ ╔╦╗ "
$_art2 = "╠═╣ ║║║ ║ ║ ╠╦╝  ║  "
$_art3 = "╩ ╩ ╩ ╩ ╚═╝ ╩╚═  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "ADVANCED MAINTENANCE, OPTIMIZATION, & RESTORATION TOOL"

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
# System Info Banner
# System Info with split coloring (Cyan Labels, Yellow Data)
Write-Host "Device Name         : " -ForegroundColor $InfoCol -NoNewline; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor Yellow
Write-Host "System Architecture : " -ForegroundColor $InfoCol -NoNewline; Write-Host "$ArchitectureDisplay" -ForegroundColor Yellow
Write-Host "Operating System    : " -ForegroundColor $InfoCol -NoNewline; Write-Host "$($OS.Caption) ($WinVer)" -ForegroundColor Yellow
$StartUsedGB = [Math]::Round(($TotalSize - $StartSpace) / 1GB, 2)
$StartTotalGB = [Math]::Round($TotalSize / 1GB, 0)
$DiskColor = if ($StartUsagePct -ge 90) { "Red" } elseif ($StartUsagePct -ge 80) { "DarkYellow" } else { "Green" }
Write-Host "Disk Usage          : " -ForegroundColor $InfoCol -NoNewline; Write-Host "${StartUsedGB}GB Used of ${StartTotalGB}GB ($StartUsagePct%)" -ForegroundColor $DiskColor
Write-HLine -Style dashed
if ($IsVM) {
    Write-Host "      Mode: Virtual Machine" -ForegroundColor Yellow 
}
#endregion

#region 1. AI, Recall, Privacy & Widget Block
# ============================================================================
Write-StepUpdate "[01/10] Killing AI, Widgets, Recall & Bing..."
$Policies = @(
    # AI & Copilot
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"; Name="DisableAIDataAnalysis"; Value=1},
    @{Path="HKCU:\Software\Policies\Microsoft\Windows\WindowsAI"; Name="DisableAIDataAnalysis"; Value=1},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Dsh"; Name="AllowNewsAndInterests"; Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\Widgets"; Name="AllowWidgets"; Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsCopilot"; Name="TurnOffWindowsCopilot"; Value=1},
    
    # Telemetry & Recall
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"; Name="AllowTelemetry"; Value=0},
    @{Path="HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsExperience"; Name="AllowRecall"; Value=0},
    
    # Search Privacy
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"; Name="BingSearchEnabled"; Value=0},
    @{Path="HKCU:\Software\Microsoft\Windows\CurrentVersion\Search"; Name="CanCortanaSeeBrowserHistory"; Value=0}
)

foreach ($Pol in $Policies) {
    if (-not (Test-Path $Pol.Path)) { New-Item $Pol.Path -Force | Out-Null }
    Set-ItemProperty -Path $Pol.Path -Name $Pol.Name -Value $Pol.Value -Force -ErrorAction SilentlyContinue | Out-Null
}
# Telemetry services - stop and permanently disable
foreach ($Svc in @("DiagTrack", "dmwappushservice")) {
    if (Get-Service -Name $Svc -ErrorAction SilentlyContinue) {
        & sc.exe stop $Svc 2>$null | Out-Null
        & sc.exe config $Svc start= disabled 2>$null | Out-Null
        Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$Svc" -Name "Start" -Value 4 -ErrorAction SilentlyContinue
    }
}
# Additional telemetry caps, WER, and CEIP
$dcPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"
if (!(Test-Path $dcPath)) { New-Item -Path $dcPath -Force | Out-Null }
Set-ItemProperty -Path $dcPath -Name "DoNotShowFeedbackNotifications" -Value 1 -ErrorAction SilentlyContinue
$dcPath2 = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection"
if (!(Test-Path $dcPath2)) { New-Item -Path $dcPath2 -Force | Out-Null }
Set-ItemProperty -Path $dcPath2 -Name "AllowTelemetry"      -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $dcPath2 -Name "MaxTelemetryAllowed" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting" -Name "Disabled" -Value 1 -ErrorAction SilentlyContinue
$werPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting"
if (!(Test-Path $werPath)) { New-Item -Path $werPath -Force | Out-Null }
Set-ItemProperty -Path $werPath -Name "Disabled" -Value 1 -ErrorAction SilentlyContinue
$sqmPath = "HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows"
if (!(Test-Path $sqmPath)) { New-Item -Path $sqmPath -Force | Out-Null }
Set-ItemProperty -Path $sqmPath -Name "CEIPEnable" -Value 0 -ErrorAction SilentlyContinue
# Advertising ID & tailored experiences
$adPolPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"
if (!(Test-Path $adPolPath)) { New-Item -Path $adPolPath -Force | Out-Null }
Set-ItemProperty -Path $adPolPath -Name "DisabledByGroupPolicy" -Value 1 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Privacy" -Name "TailoredExperiencesWithDiagnosticDataEnabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Control Panel\International\User Profile" -Name "HttpAcceptLanguageOptOut" -Value 1 -ErrorAction SilentlyContinue
# OEM content re-injection prevention
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata" -Name "PreventDeviceMetadataFromNetwork" -Value 1 -ErrorAction SilentlyContinue
$cdmPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"
if (!(Test-Path $cdmPath)) { New-Item -Path $cdmPath -Force | Out-Null }
Set-ItemProperty -Path $cdmPath -Name "DisableWindowsConsumerFeatures" -Value 1 -ErrorAction SilentlyContinue
# Lock screen, taskbar, Start menu & notification clutter
$cdmUserPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\ContentDeliveryManager"
if (!(Test-Path $cdmUserPath)) { New-Item -Path $cdmUserPath -Force | Out-Null }
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338387Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338388Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338389Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-353694Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-353696Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-338393Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SubscribedContent-310093Enabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SystemPaneSuggestionsEnabled"    -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "SoftLandingEnabled"              -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "RotatingLockScreenOverlayEnabled"-Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmUserPath -Name "RotatingLockScreenEnabled"       -Value 0 -ErrorAction SilentlyContinue
$chatPolicyPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"
if (!(Test-Path $chatPolicyPath)) { New-Item -Path $chatPolicyPath -Force | Out-Null }
Set-ItemProperty -Path $chatPolicyPath -Name "ChatIcon" -Value 3 -ErrorAction SilentlyContinue
$feedsPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"
if (!(Test-Path $feedsPath)) { New-Item -Path $feedsPath -Force | Out-Null }
Set-ItemProperty -Path $feedsPath -Name "EnableFeeds" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmPath -Name "DisableWindowsSpotlightFeatures" -Value 1 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $cdmPath -Name "DisableSoftLanding"              -Value 1 -ErrorAction SilentlyContinue
$startPolPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"
if (!(Test-Path $startPolPath)) { New-Item -Path $startPolPath -Force | Out-Null }
Set-ItemProperty -Path $startPolPath -Name "HideRecentlyAddedApps" -Value 1 -ErrorAction SilentlyContinue
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested" -Name "Enabled" -Value 0 -ErrorAction SilentlyContinue
# Explorer: open to This PC instead of Quick Access
Set-ItemProperty -Path "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced" -Name "LaunchTo" -Value 1 -ErrorAction SilentlyContinue
# Recall feature removal via DISM (full uninstall, not just policy block)
$RecallCheck = Get-WindowsOptionalFeature -Online -FeatureName "Recall" -ErrorAction SilentlyContinue
if ($RecallCheck -and $RecallCheck.State -ne "Disabled") {
    $DismArgs = "/online /Disable-Feature /FeatureName:Recall /Remove /NoRestart /Quiet /English"
    $RecallJob = Start-Process "dism.exe" -ArgumentList $DismArgs -PassThru -WindowStyle Hidden
    
    # 20-second "Fail-Fast" Timer
    $Timer = 0
    while (-not $RecallJob.HasExited -and $Timer -lt 20) {
        Start-Sleep -Seconds 1
        $Timer++
    }

    if (-not $RecallJob.HasExited) {
        Stop-Process -Id $RecallJob.Id -Force -ErrorAction SilentlyContinue
        Write-Host "        [!] Recall removal timed out (System Locked). Continuing..." -ForegroundColor DarkYellow
    }
}
# Edge policy hardening
$EdgePolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
if (!(Test-Path $EdgePolicy)) { New-Item -Path $EdgePolicy -Force | Out-Null }
Set-ItemProperty -Path $EdgePolicy -Name "HubsSidebarEnabled"             -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "CopilotPageContext"              -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "EdgeEntraCopilotPageContext"     -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "EdgeSidebarEnabled"              -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "BackgroundModeEnabled"           -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "StartupBoostEnabled"             -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "ShowMicrosoftRewards"            -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "EdgeShoppingAssistantEnabled"    -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "PersonalizationReportingEnabled" -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "EdgeFollowEnabled"               -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "ShowRecommendationsEnabled"      -Value 0 -ErrorAction SilentlyContinue
Set-ItemProperty -Path $EdgePolicy -Name "DiscoverPageContextEnabled"      -Value 0 -ErrorAction SilentlyContinue
# Office language pack cleanup (remove non-English stubs)
$C2RPath = "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\OfficeClickToRun.exe"
if (Test-Path $C2RPath) {
    $RegKeys = Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "Microsoft 365*" -and $_.DisplayName -notmatch "en-us" }
    foreach ($Key in $RegKeys) {
        if ($Key.DisplayName -match " - ([a-z]{2}-[a-z]{2})") {
            Start-Process $C2RPath -ArgumentList "scenario=install scenariosubtype=ARP sourcetype=None productstoremove=O365ProPlusRetail.16_$($Matches[1])_x-none culture=$($Matches[1]) version.16=16.0 DisplayLevel=False" -Wait -ErrorAction SilentlyContinue
        }
    }
}
# --- Calculate Regional Savings ---
$CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
if ($RegionSaved -gt 0) {
    $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
    Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
} else {
    Write-StepUpdate -Success
}
# Add the raw bytes to the running total
if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
$LastRegionSpace = $CurrentDrive.FreeSpace
#endregion

#region 2. Security: Browser Hardening
# ============================================================================
$HardeningStep = "[02/10] Security: Hardening Browsers (uBlock Origin)..."
Write-StepUpdate $HardeningStep
# Edge Hardening
$EdgeRegPath = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist"
$uBlockID = "odfafepnkmbhccpbejgmiehpchacaeak;https://edge.microsoft.com/extensionwebstorebase/v1/crx"
if (-not (Test-Path $EdgeRegPath)) { New-Item -Path $EdgeRegPath -Force | Out-Null }
$ExistingEdge = Get-ItemProperty -Path $EdgeRegPath -ErrorAction SilentlyContinue 2>$null
$NextEdgeIdx = ($ExistingEdge.PSObject.Properties.Name | Where-Object { $_ -match '^\d+$' } | Measure-Object -Maximum).Maximum + 1
if ($null -eq $NextEdgeIdx) { $NextEdgeIdx = 1 }
Set-ItemProperty -Path $EdgeRegPath -Name "$NextEdgeIdx" -Value $uBlockID -Force | Out-Null
# Chrome Hardening (Conditional)
if (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") {
    $ChromeRegPath = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
    $uBlockChromeID = "ddkjiahejlhfcafbddmgiahcphecmpfh;https://clients2.google.com/service/update2/crx"
    if (-not (Test-Path $ChromeRegPath)) { New-Item -Path $ChromeRegPath -Force | Out-Null }
    $ExistingChrome = Get-ItemProperty -Path $ChromeRegPath -ErrorAction SilentlyContinue 2>$null
    $NextChromeIdx = ($ExistingChrome.PSObject.Properties.Name | Where-Object { $_ -match '^\d+$' } | Measure-Object -Maximum).Maximum + 1
    if ($null -eq $NextChromeIdx) { $NextChromeIdx = 1 }
    Set-ItemProperty -Path $ChromeRegPath -Name "$NextChromeIdx" -Value $uBlockChromeID -Force | Out-Null
}
# --- Calculate Regional Savings ---
$CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
if ($RegionSaved -gt 0) {
    $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
    Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
} else {
    Write-StepUpdate -Success
}
# Add the raw bytes to the running total
if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
# Update the marker for the next region
$LastRegionSpace = $CurrentDrive.FreeSpace
#endregion

#region 3. Debloat: System & Software Purge
# ============================================================================
$DebloatStep = "[03/10] Executing System Debloat & Software Purge..."
Write-StepUpdate $DebloatStep
# --- Phase 1: Universal Appx Strip ---
$Bloat = @("*BingNews*", "*BingWeather*", "*ZuneVideo*", "*ZuneMusic*", "*Office.OneNote*", "*SkypeApp*", "*YourPhone*", "*WindowsCommunicationsApps*", "*PowerAutomate*", "*Todos*", "*BingSearch*")
if ($IsCore) { try { Remove-Module Appx -ErrorAction SilentlyContinue 2>$null } catch {} }
foreach ($App in $Bloat) { 
    try {
        $Pkg = Get-AppxPackage -AllUsers -Name $App -ErrorAction Stop 
        if ($Pkg) { 
            $Pkg | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue | Out-Null 
        }
    } catch {
        # If one fails, the Appx service might be locked; skip the rest to prevent the log crash
        Write-Host "        [!] Appx Service busy. Skipping further purges." -ForegroundColor Yellow
        break
    }
}
if ($IsCore) { 0..5 | ForEach-Object { Write-Progress -Id $_ -Activity "Done" -Completed } } else { Write-Progress -Activity "Stripping Bloatware" -Completed }
# --- Phase 2: Aggressive Dell Purge ---
if ($Vendor -like "*Dell*" -and -not $IsVM) {
    # Terminate and Disable Persistent Dell Services
    $DellServices = @("Dell*", "SupportAssist*", "DDV*", "DCF*", "DellHardwareSupport*")
    foreach ($SvcName in $DellServices) {
        $TargetSvcs = Get-Service -Name $SvcName -ErrorAction SilentlyContinue
        foreach ($S in $TargetSvcs) {
            Stop-Service $S.Name -Force -ErrorAction SilentlyContinue
            Set-Service $S.Name -StartupType Disabled -ErrorAction SilentlyContinue
            & sc.exe delete $S.Name | Out-Null
        }
    }
    # Targeted Software Removal using Package & WMI
    $DellPatterns = @("*SupportAssist*", "*DellUpdate*", "*DellCommand*", "*PremierColor*", "*DigitalDelivery*", "*Dell Optimizer*", "*DellCoreServices*", "*Alienware*")
    foreach ($DP in $DellPatterns) {
        $Found = @()
        $Found += Get-Package -Name $DP -ErrorAction SilentlyContinue
        $Found += Get-CimInstance -ClassName Win32_Product -ErrorAction SilentlyContinue | Where-Object { $_.Name -like $DP }
        
        foreach ($Pkg in $Found | Select-Object -Unique) {
            if ($Pkg.PSObject.Properties.Name -contains 'Uninstall') { $Pkg.Uninstall() | Out-Null } 
            else { Uninstall-Package -Name $Pkg.Name -Force -ErrorAction SilentlyContinue }
        }
    }

# --- Phase 3: Office Language Pack Purge ---
    # Target MSI-based Language Packs (Office 2016/2019/Current)
$MSILangPacks = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*Microsoft Office*Language Pack*" -and $_.DisplayName -notlike "*English*" }
if ($MSILangPacks) {
    Write-Host "        Debloat: Scrubbing foreign Office language packs..." -ForegroundColor Yellow
}
foreach ($Pack in $MSILangPacks) {
    if ($Pack.UninstallString -match "msiexec") {
        $Guid = ([regex]::Match($Pack.UninstallString, '{[A-Z0-9-]+}').Value)
        if ($Guid) { 
            Start-Process "msiexec.exe" -ArgumentList "/x $Guid /qn /norestart" -Wait -WindowStyle Hidden
        }
    }
}
    # Target Modern Office (Appx) Language Components
    # We keep 1033 (en-US) and remove everything else
Get-AppxPackage -AllUsers -Name "Microsoft.Office.Desktop.*" | Where-Object { $_.Name -match "Language" -and $_.Name -notmatch "1033" } | 
    Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    # Filesystem Scrub
    $DellFolders = @("C:\ProgramData\Dell", "C:\Program Files\Dell", "C:\Program Files (x86)\Dell", "C:\Windows\System32\Drivers\Dell")
    foreach ($DF in $DellFolders) {
        if (Test-Path $DF) { Remove-Item $DF -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
# --- Phase 3: General Win32 Software Purge ---
$KeepList = @("*Peripheral Manager*", "*Lenovo Vantage*", "*Quick Access*", "*Power Manager*", "*HP Smart*", "*HP Wolf Security*", "*Brother*", "*Canon*", "*Spotify*", "*Outlook*")
$PurgeList = @("*Teams Machine-Wide*", "*Microsoft Teams Classic*", "*Adobe Flash*", "*McAfee*", "*Norton*", "*Avast*", "*WildTangent*", "*WebAdvisor*")
# Standard Vendor Purge (Non-Dell handled here)
if (-not $IsVM -and $Vendor -notlike "*Dell*") {
    $PurgeList += switch -Wildcard ($Vendor) {
        "*HP*"     { @("*HP Support Assistant*", "*HP Support Solutions*", "*HP Wolf Security*") }
        "*Lenovo*" { @("*LenovoWelcome*", "*LenovoExperience*", "*LenovoSystemUpdate*") }
        "*Acer*"   { @("*Acer Care Center*", "*Acer Configuration Manager*", "*Acer Quick Access*") }
        "*ASUS*"   { @("*Armoury Crate*", "*MyASUS*", "*ASUS System Control Interface*") }
    }
}
$UnKeys = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
Get-ItemProperty $UnKeys -ErrorAction SilentlyContinue | Where-Object {
    $Name = $_.DisplayName; $Match = $false
    if ($Name) {
        if ($Name -like "*HP Wolf Security*") { return $false } # Essential MSP Safety Check
        foreach ($P in $PurgeList) { if ($Name -like $P) { $Match = $true } }
        foreach ($K in $KeepList) { if ($Name -like $K) { $Match = $false } }
    }
    $Match -and ($_.UninstallString)
} | ForEach-Object {
    if ($_.UninstallString -match "msiexec") {
        $Guid = ([regex]::Match($_.UninstallString, '{[A-Z0-9-]+}').Value)
        if ($Guid) { 
            $Process = Start-Process "msiexec.exe" -ArgumentList "/x $Guid /qn /norestart" -PassThru -WindowStyle Hidden
            $Process.PriorityClass = 'BelowNormal'
            $Process | Wait-Process -Timeout 30 -ErrorAction SilentlyContinue
        }
    }
}
# --- Phase 4: Registry Scrub ---
$RegPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall",
    "HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall"
)
foreach ($Reg in $RegPaths) {
    if (Test-Path $Reg) {
        Get-ChildItem $Reg -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -match "Dell|McAfee|WildTangent|Candy|Dropbox|Xbox|Spotify|TikTok" } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }
}
# --- Calculate Regional Savings ---
$CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
if ($RegionSaved -gt 0) {
    $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
    Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
} else {
    Write-StepUpdate -Success
}
# Add the raw bytes to the running total
if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
# Update the marker for the next region
$CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$LastRegionSpace = $CurrentDrive.FreeSpace
#endregion

#region 4. Snapshot & Storage Purge
# ============================================================================
Write-StepUpdate "[04/10] Purging VSS, Installer Cache & Search Index..."
# VSS & Dell Remediation
if ($Vendor -like "*Dell*" -and -not $IsVM) {
    $DellPath = "C:\ProgramData\Dell\SARemediation"
    if (Test-Path $DellPath) { Remove-Item $DellPath -Recurse -Force -ErrorAction SilentlyContinue | Out-Null }
}
& vssadmin.exe delete shadows /all /quiet 2>&1 | Out-Null
# Delivery Optimization
if (Test-Path "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache") {
    Remove-Item "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
}

# --- IMPROVED SEARCH INDEX RESET ---
$SvcName = "WSearch"
$Svc = Get-Service $SvcName -ErrorAction SilentlyContinue
if ($Svc -and $Svc.Status -ne 'Stopped') {
    Stop-Service $SvcName -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    $RetryCount = 0
    while ((Get-Service $SvcName).Status -ne 'Stopped' -and $RetryCount -lt 10) {
        $dots = '.' * (($RetryCount % 3) + 1)
        $savedRow = [Console]::CursorTop
        [Console]::SetCursorPosition(0, $script:StepRow)
        Write-Host ("$($script:LastStepMessage) [Stopping WSearch$dots]").PadRight([Console]::WindowWidth - 1) -NoNewline -ForegroundColor Cyan
        [Console]::SetCursorPosition(0, $savedRow)
        Start-Sleep -Seconds 2
        $RetryCount++
    }
    # Restore clean step line (strip the "[Stopping WSearch...]" suffix)
    [Console]::SetCursorPosition(0, $script:StepRow)
    Write-Host $script:LastStepMessage.PadRight([Console]::WindowWidth - 1) -NoNewline -ForegroundColor Cyan
    [Console]::SetCursorPosition(0, $script:StepRow + 1)
    if ((Get-Service $SvcName).Status -ne 'Stopped') {
        Get-Process "SearchIndexer" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

$SearchPath = "C:\ProgramData\Microsoft\Search\Data\Applications\Windows"
if (Test-Path $SearchPath) { 
    # FIX: Using cmd /c del bypasses the PowerShell ArgumentException 
    # if files disappear during the recursive delete.
    Start-Process "cmd.exe" -ArgumentList "/c del /s /f /q `"$SearchPath\*`"" -WindowStyle Hidden -Wait
}
Start-Service $SvcName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

# Windows Installer Cache
$InstallerPath = "C:\Windows\Installer"
if (Test-Path $InstallerPath) {
    $InstalledProducts = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object { $_.LocalPackage } | Select-Object -ExpandProperty LocalPackage
    Get-ChildItem $InstallerPath -Filter "*.ms[ip]" -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notin $InstalledProducts -and $_.LastWriteTime -lt (Get-Date).AddDays(-90) } | Remove-Item -Force -ErrorAction SilentlyContinue
}

# --- Windows.old cleanup ---
if (Test-Path "C:\Windows.old") {
    & DISM.exe /Online /Remove-OSUninstall /NoRestart | Out-Null
    $CleanupKey = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Previous Installations"
    if (Test-Path $CleanupKey) {
        Set-ItemProperty -Path $CleanupKey -Name "StateFlags1337" -Value 2 -Type DWord -ErrorAction SilentlyContinue
        $CleanupJob = Start-Process "cleanmgr.exe" -ArgumentList "/sagerun:1337" -WindowStyle Hidden -PassThru
        $CleanupJob | Wait-Process -Timeout 600 -ErrorAction SilentlyContinue
    }
    if (Test-Path "C:\Windows.old") {
        Write-StepUpdate "`n      [!] Cleanmgr timed out. Forcing Windows.old removal..."
        & takeown /F "C:\Windows.old" /R /A /D Y 2>$null | Out-Null
        & icacls "C:\Windows.old" /grant Administrators:F /T /C /Q 2>$null | Out-Null
        $WinOldItems = @(Get-ChildItem "C:\Windows.old" -Recurse -Force -ErrorAction SilentlyContinue)
        $WinOldTotal = $WinOldItems.Count
        $WinOldDone  = 0
        foreach ($WinOldItem in ($WinOldItems | Sort-Object FullName -Descending)) {
            Remove-Item $WinOldItem.FullName -Force -Recurse -ErrorAction SilentlyContinue
            $WinOldDone++
            if ($WinOldDone % 200 -eq 0 -or $WinOldDone -eq $WinOldTotal) {
                $Pct = if ($WinOldTotal -gt 0) { [int]($WinOldDone / $WinOldTotal * 100) } else { 100 }
                Write-Progress -Activity "Removing Windows.old" -Status "$WinOldDone / $WinOldTotal items removed" -PercentComplete $Pct
            }
        }
        Write-Progress -Activity "Removing Windows.old" -Completed
        Remove-Item "C:\Windows.old" -Recurse -Force -Confirm:$false -ErrorAction SilentlyContinue 2>$null
    }
}

# --- Region 4: finalize and accumulate (robust) ---
# Get current free space once and compute bytes delta
$CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$CurrentSpace = [int64]$CurrentDrive.FreeSpace

# Defensive: ensure $LastRegionSpace exists and is numeric
if ($null -eq $LastRegionSpace) { $LastRegionSpace = $CurrentSpace }

$RegionSavedBytes = [int64]($CurrentSpace - $LastRegionSpace)

# Format for display (no rounding for accumulator)
if ($RegionSavedBytes -gt 0) {
    $SavedStr = if ($RegionSavedBytes -ge 1GB) { "{0:N2} GB" -f ($RegionSavedBytes / 1GB) } else { "{0:N2} MB" -f ($RegionSavedBytes / 1MB) }
    Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
} else {
    Write-StepUpdate -Success -CustomInfo "Saved: 0 MB"
    $RegionSavedBytes = 0
}

# Ensure accumulator exists in script scope and add only positive gains
if (-not $script:TotalYieldBytes) { $script:TotalYieldBytes = 0 }
$script:TotalYieldBytes = [int64]$script:TotalYieldBytes
if ($RegionSavedBytes -gt 0) { $script:TotalYieldBytes += $RegionSavedBytes }

# Keep a small history for debugging
if (-not $script:RegionHistory) { $script:RegionHistory = @() }
$script:RegionHistory += [pscustomobject]@{ Region = 'Region4'; Bytes = $RegionSavedBytes; Time = (Get-Date) }

# Update rolling baseline after accumulation
$LastRegionSpace = $CurrentSpace
#endregion

#region 5. Deep Cache Purge
# ============================================================================
Write-StepUpdate "[05/10] Purging Browser, Office, and GPU Caches..."
$GlobalCaches = @("C:\Windows\Temp\*", "C:\Windows\Prefetch\*", "C:\Windows\SystemTemp\*")
foreach ($P in $GlobalCaches) { if (Test-Path $P) { Remove-Item $P -Recurse -Force -ErrorAction SilentlyContinue | Out-Null } }
Get-ChildItem "C:\Users" -Directory | ForEach-Object {
    $UP = $_.FullName
    $ShaderPaths = @("$UP\AppData\Local\D3DSCache", "$UP\AppData\Local\AMD\DxCache", "$UP\AppData\Local\NVIDIA\GLCache")
    $OffPaths = @("$UP\AppData\Local\Microsoft\Office\16.0\OfficeFileCache", "$UP\AppData\Local\Microsoft\Office\OTele")
    $TargetDirs = @(
        "$UP\AppData\Local\Google\Chrome\User Data\*\Cache\*",
        "$UP\AppData\Local\Microsoft\Edge\User Data\*\Cache\*",
        "$UP\AppData\Local\Mozilla\Firefox\Profiles\*\cache2\*",
        "$UP\AppData\Local\BraveSoftware\Brave-Browser\User Data\*\Cache\*",
        "$UP\AppData\Local\Opera Software\Opera Stable\Cache\*",
        "$UP\AppData\Local\Temp\*"
    )
    # Office & Outlook (NST) Cleanup
    foreach ($OP in $OffPaths) { if (Test-Path $OP) { Remove-Item $OP -Recurse -Force -ErrorAction SilentlyContinue } }
    # Outlook NST (Search Index) files
    $OutlookPath = "$UP\AppData\Local\Microsoft\Outlook"
    if (Test-Path $OutlookPath) {
        Get-ChildItem $OutlookPath -Filter "*.nst" -Force | Remove-Item -Force -ErrorAction SilentlyContinue
    }
    # GPU Shader Caches
    foreach ($SP in $ShaderPaths) { if (Test-Path $SP) { Remove-Item "$SP\*" -Recurse -Force -ErrorAction SilentlyContinue } }
    foreach ($T in $TargetDirs) { 
        if (Test-Path $T) { 
            try {
                Remove-Item $T -Recurse -Force -ErrorAction SilentlyContinue -ErrorVariable DeleteError
                # VSA STABILITY: Brief pause to allow RMM heartbeat and disk breathing
                Start-Sleep -Milliseconds 50
            } catch {
                continue
            }
        } 
    }
}
# --- Calculate Regional Savings ---
$CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
if ($RegionSaved -gt 0) {
    $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
    Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
} else {
    Write-StepUpdate -Success
}
# Add the raw bytes to the running total
if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
# Update the marker for the next region
$LastRegionSpace = $CurrentDrive.FreeSpace
#endregion

#region 6. Windows Update Database Reset
# ============================================================================
Write-StepUpdate "[06/10] Resetting Windows Update Database..."
# CHECK: If a reboot is pending, SoftwareDistribution is likely locked. 
# Skip to prevent the script from hanging.
$PendingReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"

if ($PendingReboot) {
    Write-StepUpdate -CustomInfo "[SKIPPED]"
    Write-Host "        (Reboot pending - SoftwareDistribution locked)" -ForegroundColor DarkYellow
} else {
    # REMOVED "Bits" from this list to prevent VSA disconnects
    $Svcs = @("Wuauserv", "CryptSvc", "Msiserver")
    foreach ($S in $Svcs) { Stop-Service $S -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue | Out-Null }
    
    if (Test-Path "C:\Windows\SoftwareDistribution") { 
        Remove-Item "C:\Windows\SoftwareDistribution" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null 
    }
    
    foreach ($S in $Svcs) { 
        Set-Service $S -StartupType Automatic -ErrorAction SilentlyContinue | Out-Null
        Start-ServiceSilent $S
    }   

    # --- Calculate Regional Savings ---
    $CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
    if ($RegionSaved -gt 0) {
        $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
        Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
    } else {
        Write-StepUpdate -Success
    }
    # Add the raw bytes to the running total
    if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
    # Update the marker for the next region
    $LastRegionSpace = $CurrentDrive.FreeSpace
}
#endregion

#region 7. Repair & Integrity
# ============================================================================
if ($IsVM) { [System.GC]::Collect() }
Write-Progress -Activity "Cleaning up" -Completed

# Pre-repair: stability check - warn only, never skip (PendingFileRenameOperations
# is routinely re-created by Windows/installers and does not block DISM or SFC)
$SkipRepair = $false
$PendingRename = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
$HasPendingRename = $null -ne $PendingRename
if ($PendingReboot) {
    Write-Host "        [!] Reboot pending - DISM and SFC repair steps will be skipped." -ForegroundColor DarkYellow
}
elseif ($HasPendingRename) {
    Write-Host "        [!] PendingFileRenameOperations found - skipping repair steps." -ForegroundColor DarkYellow
}
# Ensure TrustedInstaller is available (prevents DISM Error 87 / SFC failures)
if (-not $SkipRepair) {
    $TI = Get-Service -Name "TrustedInstaller" -ErrorAction SilentlyContinue
    if ($TI.StartType -eq 'Disabled') { Set-Service -Name "TrustedInstaller" -StartupType Manual }
    if ($TI.Status -ne 'Running') { Start-Service -Name "TrustedInstaller" -ErrorAction SilentlyContinue }
}
if (-not $SkipRepair) {
# Helper: clear console lines from $startRow to current row, then reprint a step result
    function Clear-AndReprintStep {
        param([int]$StartRow, [string]$Message, [switch]$Success, [string]$CustomInfo)
        try {
            $endRow = [Console]::CursorTop
            $width  = [Console]::WindowWidth - 1
            for ($r = $StartRow; $r -le $endRow; $r++) {
                [Console]::SetCursorPosition(0, $r)
                [Console]::Write(' ' * $width)
            }
            [Console]::SetCursorPosition(0, $StartRow)
        } catch {}
        if ($Success) { Write-StepUpdate $Message -Success }
        elseif ($CustomInfo -eq "[SKIPPED]") { Write-StepUpdate $Message -CustomInfo "[SKIPPED]" }
        elseif ($CustomInfo -match '^\[FAILED') {
            # Print step label in Gray, description in White, error in Red
            if ($Message -match '^(\[[\d./]+\])(\s+.+)$') {
                Write-Host $Matches[1] -NoNewline -ForegroundColor DarkGray
                Write-Host $Matches[2] -NoNewline -ForegroundColor White
            } else { Write-Host $Message -NoNewline -ForegroundColor White }
            $tag = $CustomInfo
            $currentCol = [Console]::CursorLeft
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor Red
        }
        elseif ($CustomInfo -eq "[WARNING]") {
            # Print step label in Gray, description in White, warning in Yellow
            if ($Message -match '^(\[[\d./]+\])(\s+.+)$') {
                Write-Host $Matches[1] -NoNewline -ForegroundColor DarkGray
                Write-Host $Matches[2] -NoNewline -ForegroundColor White
            } else { Write-Host $Message -NoNewline -ForegroundColor White }
            $tag = $CustomInfo
            $currentCol = [Console]::CursorLeft
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor Yellow
        }
        elseif ($CustomInfo) { Write-StepUpdate $Message -CustomInfo $CustomInfo }
    }
# Helper: flush buffered console keypresses using raw .NET Console API (bypasses PSReadLine)
    function Clear-InputBuffer { try { while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null } } catch {} }
    # Kills cmd.exe and any DISM/TiWorker children it left behind
    function Stop-DismTree {
        Get-Process -Name "DISM","TiWorker" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
    # After DISM exits, give TiWorker 5s to exit naturally, then kill it to release the console stdin handle
    function Stop-TiWorker {
        $tw = Get-Process -Name "TiWorker" -ErrorAction SilentlyContinue
        if ($tw) {
            $tw | Wait-Process -Timeout 5 -ErrorAction SilentlyContinue
            Get-Process -Name "TiWorker" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
            Start-Sleep -Milliseconds 300
        }
    }
# --- PHASE 7: RestoreHealth ---
            Clear-InputBuffer
            $S72 = "[07/10] DISM RestoreHealth..."
            Write-StepUpdate $S72 -CustomInfo "[Press ESC to Skip]"
            $Row72 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($PendingReboot -or $HasPendingRename) {
                Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[SKIPPED]"
            }
            else {
                $DismSpin = [char[]]@('|','/','-','\')
                $DismTmp1 = [System.IO.Path]::GetTempFileName()
                $Proc1 = Start-Process -FilePath "Dism.exe" -ArgumentList "/Online /Cleanup-Image /RestoreHealth /NoRestart" -NoNewWindow -PassThru -RedirectStandardOutput $DismTmp1
                $Skipped1 = $false
                $DismSpinIdx1 = 0
                $DismTimer1 = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc1.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                $Proc1 | Stop-Process -Force
                                Stop-DismTree
                                $Skipped1 = $true
                                break
                            }
                        }
                    } catch {}
                    try {
                        [Console]::SetCursorPosition(0, $Row72)
                        $SpinLine = "$S72 $($DismSpin[$DismSpinIdx1 % 4]) $($DismTimer1.Elapsed.ToString('mm\:ss'))"
                        Write-Host $SpinLine.PadRight([Console]::WindowWidth - 1) -NoNewline -ForegroundColor Cyan
                    } catch {}
                    $DismSpinIdx1++
                    Start-Sleep -Milliseconds 250
                }
                $DismTimer1.Stop()

                if (-not $Skipped1) {
                    $Proc1.WaitForExit()  # Flush exit code before reading
                    Stop-TiWorker
                }
                Remove-Item $DismTmp1 -Force -ErrorAction SilentlyContinue

                if ($Skipped1) {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[SKIPPED]"
                }
                elseif ($Proc1.ExitCode -in @(0, 3010)) {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -Success
                }
                else {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[FAILED:0x$($Proc1.ExitCode.ToString('X'))]"
                }
            }

# --- STEP 8: DISM ComponentCleanup ---
            $S73 = "[08/10] DISM ComponentCleanup..."
            Write-StepUpdate $S73 -CustomInfo "[Press ESC to Skip]"
            $Row73 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($PendingReboot -or $HasPendingRename) {
                Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[SKIPPED]"
            }
            else {
                Stop-Service wuauserv -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Stop-Service bits -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Stop-Service TrustedInstaller -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Service TrustedInstaller -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Sleep -Seconds 3

                Clear-InputBuffer

                $DismTmp2 = [System.IO.Path]::GetTempFileName()
                $Proc2 = Start-Process -FilePath "Dism.exe" -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup /NoRestart" -NoNewWindow -PassThru -RedirectStandardOutput $DismTmp2
                $Skipped2 = $false
                $DismSpinIdx2 = 0
                $DismTimer2 = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc2.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                $Proc2 | Stop-Process -Force
                                Stop-DismTree
                                $Skipped2 = $true
                                break
                            }
                        }
                    } catch {}
                    try {
                        [Console]::SetCursorPosition(0, $Row73)
                        $SpinLine = "$S73 $($DismSpin[$DismSpinIdx2 % 4]) $($DismTimer2.Elapsed.ToString('mm\:ss'))"
                        Write-Host $SpinLine.PadRight([Console]::WindowWidth - 1) -NoNewline -ForegroundColor Cyan
                    } catch {}
                    $DismSpinIdx2++
                    Start-Sleep -Milliseconds 250
                }
                $DismTimer2.Stop()

                if (-not $Skipped2) {
                    $Proc2.WaitForExit()  # Flush exit code before reading
                    Stop-TiWorker
                }
                Remove-Item $DismTmp2 -Force -ErrorAction SilentlyContinue

                Start-Service bits -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Service wuauserv -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

                if ($Skipped2) {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[SKIPPED]"
                }
                elseif ($Proc2.ExitCode -in @(0, 3010)) {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -Success
                }
                else {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[FAILED:0x$($Proc2.ExitCode.ToString('X'))]"
                }
            }

# --- STEP 9: SFC /scannow ---
            $S74 = "[09/10] SFC /scannow..."
            Write-StepUpdate $S74 -CustomInfo "[Press ESC to Skip]"
            $Row74 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($PendingReboot -or $HasPendingRename) {
                Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[SKIPPED]"
            }
            else {
                Clear-InputBuffer
                $SfcTmp = [System.IO.Path]::GetTempFileName()
                $Proc3 = Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -PassThru -RedirectStandardOutput $SfcTmp
                $Skipped3 = $false
                $SfcSpinIdx = 0
                $SfcTimer = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc3.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                $Proc3 | Stop-Process -Force
                                Get-Process -Name "sfc" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                                $Skipped3 = $true
                                break
                            }
                        }
                    } catch {}
                    try {
                        [Console]::SetCursorPosition(0, $Row74)
                        $SpinLine = "$S74 $($DismSpin[$SfcSpinIdx % 4]) $($SfcTimer.Elapsed.ToString('mm\:ss'))"
                        Write-Host $SpinLine.PadRight([Console]::WindowWidth - 1) -NoNewline -ForegroundColor Cyan
                    } catch {}
                    $SfcSpinIdx++
                    Start-Sleep -Milliseconds 250
                }
                $SfcTimer.Stop()

                if (-not $Skipped3) {
                    $Proc3.WaitForExit()  # Flush exit code before reading
                }
                Remove-Item $SfcTmp -Force -ErrorAction SilentlyContinue

                if ($Skipped3) {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[SKIPPED]"
                }
                elseif ($Proc3.ExitCode -in @(0, 1)) {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -Success
                }
                elseif ($Proc3.ExitCode -eq 2) {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[WARNING]"
                }
                else {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[FAILED:0x$($Proc3.ExitCode.ToString('X'))]"
                }
            }
}
$CurrentSpace = (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace
$RegionSaved = $CurrentSpace - $LastRegionSpace
#if ($RegionSaved -gt 0) {
#    $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
#    Write-Host "     Repair Yield: $SavedStr" -ForegroundColor Gray
#}
# Add the raw bytes to the running total
if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
# Update marker for the final Region 8
$LastRegionSpace = $CurrentSpace
#endregion

#region 8. Final Optimization
# ============================================================================
Write-StepUpdate "[10/10] Finalizing Network & SSD TRIM..."
& ipconfig.exe /flushdns | Out-Null
& powercfg.exe /h off | Out-Null
# High Performance power plan (AC only)
$highPerf = Get-CimInstance -Namespace root\cimv2\power -ClassName Win32_PowerPlan -ErrorAction SilentlyContinue | Where-Object { $_.ElementName -eq "High performance" }
if ($highPerf) { $planGuid = ($highPerf.InstanceId -split '\{' | Select-Object -Last 1).TrimEnd('}'); powercfg /setactive $planGuid }
& powercfg.exe /change standby-timeout-ac 0  | Out-Null
& powercfg.exe /change standby-timeout-dc 60 | Out-Null
& powercfg.exe /change monitor-timeout-ac 60 | Out-Null
& powercfg.exe /change monitor-timeout-dc 15 | Out-Null
try { Optimize-Volume -DriveLetter C -ReTrim -ErrorAction SilentlyContinue | Out-Null } catch { }
Write-StepUpdate -Success

# --- FINAL SUMMARY ---
# Ensure disk info and total size are valid
$FinalDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$TotalSize = [double]$TotalSize
if ($TotalSize -le 0) { throw "TotalSize is zero or undefined. Aborting final summary." }

# Use script-scoped accumulator (robust across functions/jobs)
if (-not $script:TotalYieldBytes) { $script:TotalYieldBytes = 0 }
$script:TotalYieldBytes = [int64]$script:TotalYieldBytes

# Compute final usage percent
$FinalFree = [int64]$FinalDrive.FreeSpace
$FinalUsedPct = [Math]::Round(((($TotalSize - $FinalFree) / $TotalSize) * 100), 2)

# Format recovered total for display
if ($script:TotalYieldBytes -ge 1GB) {
    $TotalStr = "{0:N2} GB" -f ($script:TotalYieldBytes / 1GB)
} else {
    $TotalStr = "{0:N2} MB" -f ($script:TotalYieldBytes / 1MB)
}

# Print summary with simple color logic
$FinalUsedGB = [Math]::Round(($TotalSize - $FinalFree) / 1GB, 2)
$FinalTotalGB = [Math]::Round($TotalSize / 1GB, 0)
$FinalColor = if ($FinalUsedPct -ge 90) { "Red" } elseif ($FinalUsedPct -ge 80) { "DarkYellow" } else { "Green" }
Write-HLine -Style dashed
Write-Host "Final Disk Usage    : " -NoNewline -ForegroundColor $InfoCol
Write-Host "$FinalUsedGB GB used of $FinalTotalGB GB ($FinalUsedPct%)" -ForegroundColor $FinalColor
Write-Host "Space Recovered     : " -NoNewline -ForegroundColor $InfoCol
Write-Host "$TotalStr" -ForegroundColor Yellow
# Footer
$_sfx   = "█"
$_ftr1 = " ╔═╗ ╔╦╗ ╔═╗ ╦═╗ ╔╦╗ "
$_ftr2 = " ╠═╣ ║║║ ║ ║ ╠╦╝  ║  "
$_ftr3 = " ╩ ╩ ╩ ╩ ╚═╝ ╩╚═  ╩  "
$_ftrW = [Math]::Max($_ftr1.Length, [Math]::Max($_ftr2.Length, $_ftr3.Length))
$_ftr1 = $_ftr1.PadRight($_ftrW); $_ftr2 = $_ftr2.PadRight($_ftrW); $_ftr3 = $_ftr3.PadRight($_ftrW)
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  MAINTENANCE COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr1 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host $_ftr2 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr3 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
#endregion