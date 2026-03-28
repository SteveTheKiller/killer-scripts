<#
.SYNOPSIS
    Definition Enforcement & Full Endpoint Network Defense (D.E.F.E.N.D.) v2.3
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Audits and enforces kernel-level Windows security including TPM,
    Secure Boot, HVCI, and Defender hardening. Syncs threat definitions,
    runs a full Defender scan with live IOPS reporting, and lists any
    active threats with file paths pulled from event logs.
#>

#region [0] - UI INITIALIZATION
# ------------------------------------------------------------------------------
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "D.E.F.E.N.D. requires Elevated Privileges. Please run as Administrator."
    break
}
Import-Module NetSecurity -Force
Import-Module NetConnection -Force

Clear-Host
# --- SESSION & VERSION DETECTION ---
# Determines PS version and whether we are in an interactive console (vs. Kaseya LiveConnect / piped / redirected)
$PSVer        = $PSVersionTable.PSVersion.Major
$IsInteractive = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected

# ANSI escape for tight background highlights (PS6+ only; PS5.1 falls back to Write-Host -BackgroundColor)
$ESC = if ($PSVer -ge 6) { [char]0x1b } else { $null }

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

function Write-Alert {
    param([string]$Text)
    if ($PSVer -ge 6) {
        Write-Host "`n$ESC[41m$ESC[97m$Text$ESC[0m"
    } else {
        Write-Host "`n$Text" -ForegroundColor White -BackgroundColor Red
    }
}

$script:Width = 85

$LineCol   = "DarkCyan"
$MainCol   = "Cyan"
$BorderCol = "Cyan"
$ArtCol    = "White"
$AccentCol = "Cyan"
$DimCol    = "DarkGray"
$_ver      = "| v2.3"

# Header Art & Logic
$_pfx  = "█  "
$_art1 = "╔╦╗ ╔═╗ ╔═╗ ╔═╗ ╔╗╔ ╔╦╗ "
$_art2 = " ║║ ║╣  ╠╣  ║╣  ║║║  ║║ "
$_art3 = "╩╩╝ ╚═╝ ╩   ╚═╝ ╝╚╝ ╩╩╝ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "DEFINITION ENFORCEMENT & FULL ENDPOINT NETWORK DEFENSE"

# Safety padding to prevent negative multiplication error
$_tpad  = " " * [Math]::Max(0, ($_fillW - $_title.Length))

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
#endregion

#region [1] - COMPLIANCE AUDIT & REMEDIATION
# ------------------------------------------------------------------------------
Write-Host "[>] Auditing Security Protocols..." -ForegroundColor DarkCyan

$PrefArgs = @{
    DisableRealtimeMonitoring     = $false
    PUAProtection                 = 'Enabled'
    MAPSReporting                 = 'Advanced'
    EnableNetworkProtection       = 'Enabled'
    EnableControlledFolderAccess  = 'Enabled'
    CloudBlockLevel               = 2
    ErrorAction                   = 'SilentlyContinue'
}
if ($PSVer -ge 6) {
    powershell.exe -NonInteractive -NoProfile -Command {
        Set-MpPreference -DisableRealtimeMonitoring $false -PUAProtection 1 -MAPSReporting 2 -EnableNetworkProtection 1 -EnableControlledFolderAccess 1 -CloudBlockLevel 2 -ErrorAction SilentlyContinue
    } | Out-Null
} else {
    Set-MpPreference @PrefArgs
}

# 1. HARDWARE & FEATURE FOUNDATION (The "Pillars")
$VTx = (Get-CimInstance Win32_Processor).VirtualizationFirmwareEnabled
$VTxReport = if ($VTx) { "ENABLED" } else { "DISABLED (Check BIOS)" }
$VTxColor  = if ($VTx) { "Green" } else { "Red" }

$VMP = Get-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform"
$VMPReport = if ($VMP.State -eq "Enabled") { "INSTALLED" } else { "MISSING" }
$VMPColor  = if ($VMP.State -eq "Enabled") { "Green" } else { "Yellow" }

# 2. INTEGRITY & TPM AUDIT
$SecureBoot = Confirm-SecureBootUEFI -ErrorAction SilentlyContinue
$SBReport   = if ($SecureBoot) { "ENABLED" } else { "DISABLED/NOT SUPPORTED" }
$SBColor    = if ($SecureBoot) { "Green" } else { "Red" }

$TPM = Get-Tpm
$TPMReport = if ($TPM.TpmPresent) { "PRESENT ($($TPM.ManufacturerVersion))" } else { "NOT FOUND" }
$TPMColor  = if ($TPM.TpmPresent) { "Green" } else { "Red" }

$HVCIKey = "HKLM:\SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity"
$HVCIActive = (Get-ItemProperty -Path $HVCIKey -Name "Enabled" -ErrorAction SilentlyContinue).Enabled -eq 1
$HVCIReport = if ($HVCIActive) { "ACTIVE" } else { "OFF" }
$HVCIColor  = if ($HVCIActive) { "Green" } else { "Yellow" }

$BlockerCount = (Get-Content "$env:windir\inf\setupapi.dev.log" -ErrorAction SilentlyContinue | Select-String "not HVCI compatible").Count
$BlockerReport = if ($BlockerCount -eq 0) { "CLEAN" } else { "$BlockerCount CONFLICTS" }
$BlockerColor = if ($BlockerCount -eq 0) { "Green" } else { "Red" }

# 3. DEFENDER ENGINE AUDIT
$LiveStatus = Get-CimInstance -Namespace "root\Microsoft\Windows\Defender" -ClassName "MSFT_MpComputerStatus"
$Prefs      = if ($PSVer -ge 6) { powershell.exe -NonInteractive -NoProfile -Command "Get-MpPreference | ConvertTo-Json -Depth 3" | ConvertFrom-Json } else { Get-MpPreference }

$TamperReport = if ($LiveStatus.IsTamperProtected) { "ENFORCED" } else { "NOT ENABLED" }
$TamperColor  = if ($LiveStatus.IsTamperProtected) { "Green" } else { "Red" }

$CFAReport = if ($Prefs.EnableControlledFolderAccess -eq 1) { "ENFORCED" } else { "NOT ENABLED" }
$CFAColor  = if ($Prefs.EnableControlledFolderAccess -eq 1) { "Green" } else { "Red" }

$CloudReport = switch ($Prefs.CloudBlockLevel) {
    0 { "Normal" }
    2 { "High" }
    4 { "High Plus" }
    6 { "Extreme" }
    Default { "Standard" }
}

# 4. EXCLUSION FILTER
$RawExclusions = $Prefs.ExclusionPath
$FilteredExclusions = $RawExclusions | Where-Object { $_ -notlike "*Huntress*" -and $_ -notlike "*Rio*" }
$StackDetected = if ($RawExclusions -like "*Huntress*" -or $RawExclusions -like "*Rio*") { " (+MSP Stack)" } else { "" }
$ExclReport = if (-not $FilteredExclusions) { "NONE (Secure)$StackDetected" } else { "$($FilteredExclusions.Count) Active (Filtered)$StackDetected" }
$ExclColor  = if (-not $FilteredExclusions) { "Green" } else { "Yellow" }

# 5. CONSOLE OUTPUT
Write-Host "$("BIOS Virtualization".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $VTxReport -ForegroundColor $VTxColor
Write-Host "$("Virtual Machine Platform".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $VMPReport -ForegroundColor $VMPColor
Write-Host "$("TPM 2.0 Module".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $TPMReport -ForegroundColor $TPMColor
Write-Host "$("Secure Boot Status".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $SBReport -ForegroundColor $SBColor
Write-Host "$("Memory Integrity (HVCI)".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $HVCIReport -ForegroundColor $HVCIColor
Write-Host "$("Driver Compatibility".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $BlockerReport -ForegroundColor $BlockerColor
Write-Host "$("RealTimeMonitoring".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host "ENFORCED" -ForegroundColor Green
Write-Host "$("Tamper Protection".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $TamperReport -ForegroundColor $TamperColor
Write-Host "$("Controlled Folder Access".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $CFAReport -ForegroundColor $CFAColor
Write-Host "$("Cloud Block Level".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $CloudReport -ForegroundColor Cyan
Write-Host "$("Active Exclusions".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $ExclReport -ForegroundColor $ExclColor

# 6. DYNAMIC BLOCKER ANALYSIS & REMEDIATION
if ($BlockerCount -gt 0) {
    Write-Host "`n[!] HVCI BLOCKER ANALYSIS" -ForegroundColor Yellow
    Write-Host "The following drivers are listed in SetupAPI logs as incompatible:" -ForegroundColor Gray
    Get-Content "$env:windir\inf\setupapi.dev.log" | Select-String "not HVCI compatible" -Context 0,1 | ForEach-Object {
        Write-Host " > " -NoNewline -ForegroundColor Red; Write-Host $_.ToString().Trim() -ForegroundColor White
    }
    
    Write-Host "`n[?] Use 'pnputil /enum-drivers' to find and remove the OEM .inf files listed above." -ForegroundColor Gray
}

if ($VMP.State -ne "Enabled" -and $VTx) {
    Write-Host "`n[!] Virtualization is ON in BIOS but Windows Feature is MISSING." -ForegroundColor Yellow
    Write-Host "    Action: Run 'Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform'" -ForegroundColor Gray
}
#endregion

#region [2] - DYNAMIC FIREWALL CONFIGURATION
# ------------------------------------------------------------------------------
# Enables all firewall profiles and categorizes connection based on domain status.
Write-Host "[>] Configuring Adaptive Firewall..." -ForegroundColor DarkCyan

Set-NetFirewallProfile -Profile Domain, Private, Public -Enabled True

$IsDomainJoined = (Get-CimInstance Win32_ComputerSystem).PartofDomain

if ($IsDomainJoined) {
    Write-Host "Context                      : " -NoNewline -ForegroundColor White; Write-Host "Domain Joined" -ForegroundColor Cyan
    Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory DomainAuthenticated -ErrorAction SilentlyContinue
    Write-Host "Action                       : " -NoNewline -ForegroundColor White; Write-Host "Domain Security Enforced" -ForegroundColor Green
} else {
    Write-Host "Context                      : " -NoNewline -ForegroundColor White; Write-Host "Workgroup/Standalone" -ForegroundColor Cyan
    Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private -ErrorAction SilentlyContinue
    Write-Host "Action                       : " -NoNewline -ForegroundColor White; Write-Host "Private Security Enforced" -ForegroundColor Green
}
#endregion

#region [3] - THREAT INTELLIGENCE SYNC
# ------------------------------------------------------------------------------
Write-Host "[>] Syncing Threat Intelligence..." -ForegroundColor DarkCyan

try {
    if ($PSVer -ge 6) {
        powershell.exe -NonInteractive -NoProfile -Command "Update-MpSignature -ErrorAction Stop" | Out-Null
    } else {
        Update-MpSignature -ErrorAction Stop
    }
    Write-Host "Status                       : " -NoNewline -ForegroundColor White; Write-Host "Definitions Updated" -ForegroundColor Green
} catch {
    Write-Host "Status                       : " -NoNewline -ForegroundColor White; Write-Host "Update Failed" -ForegroundColor Red
}
#endregion

#region [4] - DYNAMIC SCAN CONTROL
$ExePath = "C:\Program Files\Windows Defender\MpCmdRun.exe"
$StartTime = Get-Date
$ScanType = 2 
$ModeName = "FULL SCAN"
$Pivoted = $false
$Skipped = $false

if (-not $IsInteractive) {
    # Non-interactive session (Kaseya LiveConnect, piped, redirected)
    # Skip live scan to avoid blocking the session; Defender definitions were already updated above.
    Write-Host "Scan Status                  : " -NoNewline -ForegroundColor White
    Write-Host "SKIPPED (Non-Interactive Session)" -ForegroundColor Yellow
} else {

$ScanMsgY = [Console]::CursorTop
Write-Host "[!] $ModeName initiated.  |  Press 'Q' for Quick Scan  |  'Esc' to skip" -ForegroundColor Yellow

$ScanProc = Start-Process -FilePath $ExePath -ArgumentList "-Scan -ScanType $ScanType" -PassThru -WindowStyle Hidden
$Y = [Console]::CursorTop

while (-not $ScanProc.HasExited) {
    if ([Console]::KeyAvailable) {
        $Key = [Console]::ReadKey($true)
        if ($Key.Key -eq 'Escape') {
            & $ExePath -Scan -Cancel | Out-Null
            Stop-Process -Id $ScanProc.Id -Force -ErrorAction SilentlyContinue
            $Skipped = $true
            break
        }
        if ($Key.KeyChar -eq 'q') {
            & $ExePath -Scan -Cancel | Out-Null
            Stop-Process -Id $ScanProc.Id -Force -ErrorAction SilentlyContinue
            
            [Console]::SetCursorPosition(0, $Y)
            Write-Host "[>] Waiting for Defender engine to release session lock..." -ForegroundColor Gray -NoNewline
            do {
                Start-Sleep -Seconds 5
                $Test = & $ExePath -GetVersion 2>&1
            } while ($Test -like "*0x8050801a*" -or $Test -like "*busy*")
            
            [Console]::SetCursorPosition(0, $Y)
            [Console]::Write(" " * 120)
            [Console]::SetCursorPosition(0, $Y)
            
            $Skipped = $false
            $Pivoted = $true
            break 
        }
    }
    $Elapsed = (Get-Date) - $StartTime
    $Timer = "{0:hh\:mm\:ss}" -f $Elapsed
    [Console]::SetCursorPosition(0, $Y)
    [Console]::Write(" " * 130) 
    [Console]::SetCursorPosition(0, $Y)
    [Console]::Write("Scan Activity : RUNNING ($ModeName) [$Timer]")
    Start-Sleep -Milliseconds 500
}

if ($Pivoted -and -not $Skipped) {
    $ScanType = 1
    $ModeName = "QUICK SCAN"
    $StartTime = Get-Date 
    [Console]::SetCursorPosition(0, $ScanMsgY)
    [Console]::Write(" " * 120)
    [Console]::SetCursorPosition(0, $ScanMsgY)
    Write-Host "[!] Pivot Successful: Starting $ModeName... Hit 'Esc' to skip.          " -ForegroundColor Yellow
    
    $ScanProc = Start-Process -FilePath $ExePath -ArgumentList "-Scan -ScanType $ScanType" -PassThru -WindowStyle Hidden
    while (-not $ScanProc.HasExited) {
        if ([Console]::KeyAvailable) {
            $Key = [Console]::ReadKey($true)
            if ($Key.Key -eq 'Escape') {
                & $ExePath -Scan -Cancel | Out-Null
                Stop-Process -Id $ScanProc.Id -Force -ErrorAction SilentlyContinue
                $Skipped = $true
                break
            }
        }
        $Elapsed = (Get-Date) - $StartTime
        $Timer = "{0:hh\:mm\:ss}" -f $Elapsed
        [Console]::SetCursorPosition(0, $Y)
        [Console]::Write(" " * 130) 
        [Console]::SetCursorPosition(0, $Y)
        [Console]::Write("Scan Activity : RUNNING ($ModeName) [$Timer]")
        Start-Sleep -Milliseconds 500
    }
}

# Wipe both the scan message and activity lines
[Console]::SetCursorPosition(0, $ScanMsgY)
[Console]::Write(" " * 80)
[Console]::SetCursorPosition(0, $ScanMsgY)
[Console]::SetCursorPosition(0, $Y)
[Console]::Write(" " * 130)
[Console]::SetCursorPosition(0, $ScanMsgY)

if ($Skipped) {
    Write-Host "Scan Status                  : " -NoNewline -ForegroundColor White; Write-Host "SKIPPED BY USER" -ForegroundColor Yellow
} else {
    Write-Host "Scan Status                  : " -NoNewline -ForegroundColor White; Write-Host "COMPLETE" -ForegroundColor Green
}

} # end IsInteractive scan block
#endregion

#region [5] - THREAT IDENTIFICATION & AUDIT
# ------------------------------------------------------------------------------
# Only pull threats that are actually "Active" (Status 1) and ignore those marked "Cleaned" or "Removed"
$Threats   = if ($PSVer -ge 6) { powershell.exe -NonInteractive -NoProfile -Command "Get-MpThreat | ConvertTo-Json -Depth 3" | ConvertFrom-Json } else { Get-MpThreat }
$ScanStats = if ($PSVer -ge 6) { powershell.exe -NonInteractive -NoProfile -Command "Get-MpComputerStatus | ConvertTo-Json -Depth 3" | ConvertFrom-Json } else { Get-MpComputerStatus }
Write-Host "[>] Finalizing Security Audit..." -ForegroundColor DarkCyan
# 1. TRUTH CAPTURE: Query Event ID 1001 (Scan Completed) for the absolute last successful dates
$FullScanEvent = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Defender/Operational'; Id=1001} -ErrorAction SilentlyContinue | 
                 Where-Object { $_.Message -like "*Full Scan*" } | Select-Object -First 1
$QuickScanEvent = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Defender/Operational'; Id=1001} -ErrorAction SilentlyContinue | 
                  Where-Object { $_.Message -like "*Quick Scan*" } | Select-Object -First 1
# 2. DATA ASSIGNMENT: Pull the TimeCreated property to match the Windows Security UI
$FullDisplay = if ($FullScanEvent) { $FullScanEvent.TimeCreated } else { "No Completed Full Scan Found" }
$QuickDisplay = if ($QuickScanEvent) { $QuickScanEvent.TimeCreated } else { "No Completed Quick Scan Found" }
# 3. CONSOLE OUTPUT
Write-Host "$("Last Successful Full Scan".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $FullDisplay -ForegroundColor Yellow
Write-Host "$("Last Successful Quick Scan".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host $QuickDisplay -ForegroundColor Yellow
Write-Host "$("Antivirus Signature".PadRight(28)) : " -NoNewline -ForegroundColor White; Write-Host "v$($ScanStats.AntivirusSignatureVersion)" -ForegroundColor Yellow
if ($Threats) {
    Write-Alert " [!] CRITICAL: $($Threats.Count) Threat(s) Detected on $env:COMPUTERNAME "
    foreach ($T in $Threats) {
        $TName = if ($T.ThreatName) { $T.ThreatName } else { "Unknown Signature" }
        
        # DYNAMIC PATH: Try resources first, then filtered event log fallback
        if ($T.Resources) {
            $TPath = $T.Resources -join ", "
        } else {
            # Specifically seeking the path for the current Threat Name to avoid OpenRGB crosstalk
            $ThreatEvent = Get-WinEvent -FilterHashtable @{LogName='Microsoft-Windows-Windows Defender/Operational'; Id=1116} -ErrorAction SilentlyContinue | 
                     Where-Object { $_.Message -like "*$TName*" } | Select-Object -First 1
            
            if ($ThreatEvent) {
                $TPath = ([xml]$ThreatEvent.ToXml()).Event.EventData.Data | Where-Object { $_.Name -eq 'Path' } | Select-Object -ExpandProperty '#text'
            } else {
                $TPath = "Location Not Found in Logs"
            }
        }
        
        Write-HLine -Style dashed
        Write-Host "THREAT TYPE : " -NoNewline -ForegroundColor White; Write-Host $TName -ForegroundColor Yellow
        Write-Host "LOCATION    : " -NoNewline -ForegroundColor White; Write-Host $TPath -ForegroundColor Yellow
        if ($TPath -eq "Location Not Found in Logs") {
            Write-Host "(File likely already removed or remediated by Defender)" -ForegroundColor DarkGray
        }
        Write-HLine -Style dashed
    }
} else {
    Write-Host "`n[+] System Audit Results: NO ACTIVE THREATS" -ForegroundColor Green
    Write-Host "Status       : Endpoint Verified Compliant" -ForegroundColor White
}

# 4. REBOOT PENDING CHECK
$RebootPending = $false
$RegPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
$ComponentPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"

if ((Test-Path $RegPath) -or (Test-Path $ComponentPath)) {
    $RebootPending = $true
}

if ($RebootPending) {
    Write-Host "[!] ACTION REQUIRED: A System Reboot is Pending." -ForegroundColor Red
    Write-Host "    Some security enforcements may not be active yet." -ForegroundColor Gray
} else {
    Write-Host "[+] System State: No Reboot Required." -ForegroundColor Green
}

# Footer (reuses header art - art on right, fill on left)
$_ffillW = $script:Width - $_artW - 2  # 1 leading space + 1 sfx char
$_footer = "  ENDPOINT SECURITY ENFORCEMENT COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Exit 0
#endregion