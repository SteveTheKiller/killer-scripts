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
$_ver = "| v2.3"

#region [0] - HELPERS AND SYSTEM BOUNDARIES
# ------------------------------------------------------------------------------
function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

function Test-IsAdmin {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

function Get-DefendPSMajor { $PSVersionTable.PSVersion.Major }

function Get-DefendNow { Get-Date }

# Console access is guarded. With output redirected (RMM, Kaseya LiveConnect) the console handle
# is invalid, so every call falls back to plain line output instead of throwing.
function Test-DefendInteractive {
    try { [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected } catch { $false }
}
function Get-DefendConsoleValue {
    param([string]$Name)
    [Console]::$Name
}
function Set-DefendConsolePosition {
    param([int]$Left, [int]$Top)
    [Console]::SetCursorPosition($Left, $Top)
}
function Get-DefendCursorTop { try { [int](Get-DefendConsoleValue 'CursorTop') } catch { -1 } }
function Set-DefendCursor {
    param([int]$Left, [int]$Top)
    try { Set-DefendConsolePosition $Left $Top; $true } catch { $false }
}
function Test-DefendKeyAvailable { try { [bool][Console]::KeyAvailable } catch { $false } }
function Read-DefendKey { try { [Console]::ReadKey($true) } catch { $null } }
function Clear-DefendScreen { try { Clear-Host -ErrorAction Stop } catch { } }

# Registry reads go to the 64-bit view, so a 32-bit host (some RMM agents) sees the real keys.
function Get-DefendRegValue {
    param([string]$SubKey, [string]$Name)
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        try {
            $k = $base.OpenSubKey($SubKey)
            if ($k) { try { $k.GetValue($Name) } finally { $k.Close() } }
        } finally { $base.Close() }
    } catch { $null }
}
function Test-DefendRegKey {
    param([string]$SubKey)
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        try {
            $k = $base.OpenSubKey($SubKey)
            if ($k) { $k.Close(); $true } else { $false }
        } finally { $base.Close() }
    } catch { $false }
}

function Get-DefendHvciBlockers {
    # Lines of the SetupAPI log that name drivers as not HVCI compatible.
    $log = Join-Path $env:windir 'inf\setupapi.dev.log'
    @(Get-Content -LiteralPath $log -ErrorAction SilentlyContinue | Where-Object { $_ -match 'not HVCI compatible' } | ForEach-Object { $_.Trim() })
}

# Runs a command in Windows PowerShell 5.1. Used when this script runs under PowerShell 7,
# because the Defender cmdlets are native to 5.1.
function Invoke-DefendWinPS {
    param([string]$Command)
    $out = & powershell.exe -NonInteractive -NoProfile -Command $Command 2>&1
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = (($out | ForEach-Object { [string]$_ }) -join "`n") }
}

function Start-DefendScanProcess {
    param([int]$ScanType)
    $p = Start-Process -FilePath $script:DefendExe -ArgumentList "-Scan -ScanType $ScanType" -PassThru -WindowStyle Hidden
    # Touching the handle keeps ExitCode readable after the process exits.
    $null = $p.Handle
    $p
}
function Stop-DefendScan {
    param($Proc)
    & $script:DefendExe -Scan -Cancel | Out-Null
    Stop-Process -Id $Proc.Id -Force -ErrorAction SilentlyContinue
}
function Test-DefendEngineBusy {
    $t = & $script:DefendExe -GetVersion 2>&1
    ("$t" -like '*0x8050801a*') -or ("$t" -like '*busy*')
}

$script:Width = 85
$LineCol   = "DarkCyan"
$MainCol   = "Cyan"
$BorderCol = "Cyan"
$ArtCol    = "White"
$AccentCol = "Cyan"
$DimCol    = "DarkGray"

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
    $colors = @([ConsoleColor]$BorderCol, [ConsoleColor]$ArtCol, [ConsoleColor]$AccentCol, [ConsoleColor]$DimCol)
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') {
            Write-Host $char -NoNewline -ForegroundColor ([ConsoleColor]$DimCol)
        } else {
            Write-Host $char -NoNewline -ForegroundColor $colors[$i % $colors.Count]
            $i++
        }
    }
    Write-Host ""
}

function Write-Alert {
    param([string]$Text)
    Write-Host "`n$Text" -ForegroundColor White -BackgroundColor Red
}

function Write-Field {
    param([string]$Label, [string]$Value, [string]$Color = "White")
    Write-Host "$($Label.PadRight(28)) : " -NoNewline -ForegroundColor White
    Write-Host $Value -ForegroundColor $Color
}

# Defender preference helpers. Under PowerShell 7 they run in a Windows PowerShell child.
# Set-DefendPreferences returns $null on success, or the error text when the change was refused.
function Set-DefendPreferences {
    if ((Get-DefendPSMajor) -ge 6) {
        $cmd = 'try { Set-MpPreference -DisableRealtimeMonitoring $false -PUAProtection 1 -MAPSReporting 2 -EnableNetworkProtection 1 -EnableControlledFolderAccess 1 -CloudBlockLevel 2 -ErrorAction Stop; exit 0 } catch { Write-Output $_.Exception.Message; exit 1 }'
        $r = Invoke-DefendWinPS $cmd
        if ($r.ExitCode -ne 0) { return ("exit {0}: {1}" -f $r.ExitCode, $r.Output) }
        return $null
    }
    try {
        Set-MpPreference -DisableRealtimeMonitoring $false -PUAProtection Enabled -MAPSReporting Advanced -EnableNetworkProtection Enabled -EnableControlledFolderAccess Enabled -CloudBlockLevel 2 -ErrorAction Stop
        $null
    } catch { $_.Exception.Message }
}# Update-DefendSignatures returns $null on success, or the error text.
function Update-DefendSignatures {
    if ((Get-DefendPSMajor) -ge 6) {
        $r = Invoke-DefendWinPS 'try { Update-MpSignature -ErrorAction Stop; exit 0 } catch { Write-Output $_.Exception.Message; exit 1 }'
        if ($r.ExitCode -ne 0) { return ("exit {0}: {1}" -f $r.ExitCode, $r.Output) }
        return $null
    }
    try { Update-MpSignature -ErrorAction Stop; $null } catch { $_.Exception.Message }
}
function Get-DefendPreferences {
    if ((Get-DefendPSMajor) -ge 6) {
        $o = (Invoke-DefendWinPS 'Get-MpPreference | ConvertTo-Json -Depth 3').Output
        if (-not [string]::IsNullOrWhiteSpace($o)) { $o | ConvertFrom-Json }
    } else {
        Get-MpPreference
    }
}

function ConvertFrom-DefendJson {
    # JSON text from a child process to an array of objects. Empty output is an empty array.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @() }
    $x = $Text | ConvertFrom-Json
    @($x | ForEach-Object { $_ })
}

# Get-MpThreat is threat HISTORY and still lists cleaned and quarantined items. The detections carry
# the status. ThreatStatusID: 1 Detected, 2 Cleaned, 3 Quarantined, 4 Removed, 5 Allowed, 6 Blocked,
# 102 QuarantineFailed, 103 RemoveFailed, 104 AllowFailed, 105 Abandoned, 107 BlockedFailed.
# A threat is active when it was detected and not remediated, or when remediation failed.
$script:DefendActiveStatus = @(1, 102, 103, 104, 105, 107)

function Get-DefendActiveThreats {
    if ((Get-DefendPSMajor) -ge 6) {
        $det  = ConvertFrom-DefendJson (Invoke-DefendWinPS 'Get-MpThreatDetection | ConvertTo-Json -Depth 3').Output
        $hist = ConvertFrom-DefendJson (Invoke-DefendWinPS 'Get-MpThreat | ConvertTo-Json -Depth 3').Output
    } else {
        $det  = @(Get-MpThreatDetection -ErrorAction SilentlyContinue)
        $hist = @(Get-MpThreat -ErrorAction SilentlyContinue)
    }
    foreach ($d in @($det)) {
        if ($null -eq $d -or $script:DefendActiveStatus -notcontains [int]$d.ThreatStatusID) { continue }
        $h = @($hist | Where-Object { $_.ThreatID -eq $d.ThreatID }) | Select-Object -First 1
        [pscustomobject]@{
            ThreatName = $(if ($h -and $h.ThreatName) { $h.ThreatName } else { $null })
            Resources  = @($d.Resources)
        }
    }
}#endregion

function Invoke-DefendMain {
    $script:DefendFailures = New-Object System.Collections.ArrayList
    $script:DefendExe      = Join-Path $(if ($env:ProgramW6432) { $env:ProgramW6432 } else { $env:ProgramFiles }) 'Windows Defender\MpCmdRun.exe'

    #region [0] - UI INITIALIZATION
    # --------------------------------------------------------------------------
    if (-not (Test-IsAdmin)) {
        Write-Warning "D.E.F.E.N.D. requires Elevated Privileges. Please run as Administrator."
        Write-Output "[FAIL] DEFEND not elevated, nothing was changed (exit 2)"
        Exit-KillerScript 2
    }
    Import-Module NetSecurity -ErrorAction SilentlyContinue
    Import-Module NetConnection -ErrorAction SilentlyContinue

    Clear-DefendScreen
    $PSVer         = Get-DefendPSMajor
    $IsInteractive = Test-DefendInteractive

    $_pfx  = "█  "
    $_art1 = "╔╦╗ ╔═╗ ╔═╗ ╔═╗ ╔╗╔ ╔╦╗ "
    $_art2 = " ║║ ║╣  ╠╣  ║╣  ║║║  ║║ "
    $_art3 = "╩╩╝ ╚═╝ ╩   ╚═╝ ╝╚╝ ╩╩╝ "
    $_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
    $_fillW = $script:Width - $_pfx.Length - $_artW
    $_title = "DEFINITION ENFORCEMENT & FULL ENDPOINT NETWORK DEFENSE"
    $_tpad  = " " * [Math]::Max(0, ($_fillW - $_title.Length))

    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    #endregion

    #region [1] - COMPLIANCE AUDIT & REMEDIATION
    # --------------------------------------------------------------------------
    Write-Host "[>] Auditing Security Protocols..." -ForegroundColor DarkCyan

    $SetError = Set-DefendPreferences
    if ($SetError) { [void]$script:DefendFailures.Add("Set-MpPreference was refused: $SetError") }

    # 1. HARDWARE & FEATURE FOUNDATION
    $VTx = (Get-CimInstance Win32_Processor | Select-Object -First 1).VirtualizationFirmwareEnabled
    $VTxReport = if ($VTx) { "ENABLED" } else { "DISABLED (Check BIOS)" }
    $VTxColor  = if ($VTx) { "Green" } else { "Red" }

    $VMP = $null
    try { $VMP = Get-WindowsOptionalFeature -Online -FeatureName "VirtualMachinePlatform" -ErrorAction Stop } catch { }
    $VMPReport = if (-not $VMP) { "UNKNOWN" } elseif ($VMP.State -eq "Enabled") { "INSTALLED" } else { "MISSING" }
    $VMPColor  = if (-not $VMP) { "Yellow" } elseif ($VMP.State -eq "Enabled") { "Green" } else { "Yellow" }

    # 2. INTEGRITY & TPM AUDIT
    $SecureBoot = $null
    try { $SecureBoot = Confirm-SecureBootUEFI -ErrorAction Stop } catch { }
    $SBReport   = if ($SecureBoot) { "ENABLED" } else { "DISABLED/NOT SUPPORTED" }
    $SBColor    = if ($SecureBoot) { "Green" } else { "Red" }

    $TPM = $null
    try { $TPM = Get-Tpm -ErrorAction Stop } catch { }
    $TPMReport = if ($TPM.TpmPresent) { "PRESENT ($($TPM.ManufacturerVersion))" } else { "NOT FOUND" }
    $TPMColor  = if ($TPM.TpmPresent) { "Green" } else { "Red" }

    $HVCIActive = (Get-DefendRegValue 'SYSTEM\CurrentControlSet\Control\DeviceGuard\Scenarios\HypervisorEnforcedCodeIntegrity' 'Enabled') -eq 1
    $HVCIReport = if ($HVCIActive) { "ACTIVE" } else { "OFF" }
    $HVCIColor  = if ($HVCIActive) { "Green" } else { "Yellow" }

    $Blockers     = @(Get-DefendHvciBlockers)
    $BlockerCount = $Blockers.Count
    $BlockerReport = if ($BlockerCount -eq 0) { "CLEAN" } else { "$BlockerCount CONFLICTS" }
    $BlockerColor = if ($BlockerCount -eq 0) { "Green" } else { "Red" }

    # 3. DEFENDER ENGINE AUDIT
    $LiveStatus = Get-CimInstance -Namespace "root\Microsoft\Windows\Defender" -ClassName "MSFT_MpComputerStatus"
    $Prefs      = Get-DefendPreferences

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

    # Real-time monitoring is judged by what the engine reports after the change, not by what was asked.
    $RtLive     = $LiveStatus.RealTimeProtectionEnabled
    $RtEnforced = ($Prefs.DisableRealtimeMonitoring -eq $false) -and (($null -eq $RtLive) -or ($RtLive -eq $true))
    $RtReport   = if ($RtEnforced) { "ENFORCED" } else { "NOT ENFORCED (blocked by Tamper Protection, policy or another AV?)" }
    $RtColor    = if ($RtEnforced) { "Green" } else { "Red" }
    if (-not $RtEnforced) { [void]$script:DefendFailures.Add("real-time monitoring is not enforced") }

    # 4. EXCLUSION FILTER
    $RawExclusions = $Prefs.ExclusionPath
    $FilteredExclusions = $RawExclusions | Where-Object { $_ -notlike "*Huntress*" -and $_ -notlike "*Rio*" }
    $StackDetected = if ($RawExclusions -like "*Huntress*" -or $RawExclusions -like "*Rio*") { " (+MSP Stack)" } else { "" }
    $ExclReport = if (-not $FilteredExclusions) { "NONE (Secure)$StackDetected" } else { "$($FilteredExclusions.Count) Active (Filtered)$StackDetected" }
    $ExclColor  = if (-not $FilteredExclusions) { "Green" } else { "Yellow" }

    # 5. CONSOLE OUTPUT
    Write-Field "BIOS Virtualization" $VTxReport $VTxColor
    Write-Field "Virtual Machine Platform" $VMPReport $VMPColor
    Write-Field "TPM 2.0 Module" $TPMReport $TPMColor
    Write-Field "Secure Boot Status" $SBReport $SBColor
    Write-Field "Memory Integrity (HVCI)" $HVCIReport $HVCIColor
    Write-Field "Driver Compatibility" $BlockerReport $BlockerColor
    Write-Field "RealTimeMonitoring" $RtReport $RtColor
    Write-Field "Tamper Protection" $TamperReport $TamperColor
    Write-Field "Controlled Folder Access" $CFAReport $CFAColor
    Write-Field "Cloud Block Level" $CloudReport "Cyan"
    Write-Field "Active Exclusions" $ExclReport $ExclColor

    # 6. DYNAMIC BLOCKER ANALYSIS & REMEDIATION
    if ($BlockerCount -gt 0) {
        Write-Host "`n[!] HVCI BLOCKER ANALYSIS" -ForegroundColor Yellow
        Write-Host "The following drivers are listed in SetupAPI logs as incompatible:" -ForegroundColor Gray
        foreach ($b in $Blockers) {
            Write-Host " > " -NoNewline -ForegroundColor Red; Write-Host $b -ForegroundColor White
        }
        Write-Host "`n[?] Use 'pnputil /enum-drivers' to find and remove the OEM .inf files listed above." -ForegroundColor Gray
    }

    if ($VMP -and $VMP.State -ne "Enabled" -and $VTx) {
        Write-Host "`n[!] Virtualization is ON in BIOS but Windows Feature is MISSING." -ForegroundColor Yellow
        Write-Host "    Action: Run 'Enable-WindowsOptionalFeature -Online -FeatureName VirtualMachinePlatform'" -ForegroundColor Gray
    }
    #endregion

    #region [2] - DYNAMIC FIREWALL CONFIGURATION
    # --------------------------------------------------------------------------
    Write-Host "[>] Configuring Adaptive Firewall..." -ForegroundColor DarkCyan

    Set-NetFirewallProfile -Profile Domain, Private, Public -Enabled True

    $IsDomainJoined = (Get-CimInstance Win32_ComputerSystem).PartofDomain

    if ($IsDomainJoined) {
        Write-Field "Context" "Domain Joined" "Cyan"
        Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory DomainAuthenticated -ErrorAction SilentlyContinue
        Write-Field "Action" "Domain Security Enforced" "Green"
    } else {
        Write-Field "Context" "Workgroup/Standalone" "Cyan"
        Get-NetConnectionProfile | Set-NetConnectionProfile -NetworkCategory Private -ErrorAction SilentlyContinue
        Write-Field "Action" "Private Security Enforced" "Green"
    }
    #endregion

    #region [3] - THREAT INTELLIGENCE SYNC
    # --------------------------------------------------------------------------
    Write-Host "[>] Syncing Threat Intelligence..." -ForegroundColor DarkCyan

    $UpdateError = Update-DefendSignatures
    if ($UpdateError) {
        Write-Field "Status" "Update Failed" "Red"
        Write-Host "                               $UpdateError" -ForegroundColor Gray
        [void]$script:DefendFailures.Add("definition update failed: $UpdateError")
    } else {
        Write-Field "Status" "Definitions Updated" "Green"
    }
    #endregion

    #region [4] - DYNAMIC SCAN CONTROL
    # --------------------------------------------------------------------------
    $ScanType = 2
    $ModeName = "FULL SCAN"
    $Pivoted = $false
    $Skipped = $false

    if (-not $IsInteractive) {
        # Non-interactive session (Kaseya LiveConnect, piped, redirected): skip the live scan
        # so the session is not blocked.
        Write-Field "Scan Status" "SKIPPED (Non-Interactive Session)" "Yellow"
    } else {
        Write-Host "[!] $ModeName initiated.  |  Press 'Q' for Quick Scan  |  'Esc' to skip" -ForegroundColor Yellow
        $StartTime = Get-DefendNow
        $ScanProc = Start-DefendScanProcess -ScanType $ScanType
        $Y = Get-DefendCursorTop

        $res = Wait-DefendScan -Proc $ScanProc -Mode $ModeName -Y $Y -AllowQuick $true
        if ($res -eq 'Skipped') { $Skipped = $true }
        if ($res -eq 'Pivot') {
            Write-Host "[>] Waiting for Defender engine to release session lock..." -ForegroundColor Gray
            do {
                Start-Sleep -Seconds 5
            } while (Test-DefendEngineBusy)
            $Pivoted = $true
        }

        if ($Pivoted -and -not $Skipped) {
            $ScanType = 1
            $ModeName = "QUICK SCAN"
            $StartTime = Get-DefendNow
            Write-Host "[!] Pivot Successful: Starting $ModeName... Hit 'Esc' to skip." -ForegroundColor Yellow
            $ScanProc = Start-DefendScanProcess -ScanType $ScanType
            $res = Wait-DefendScan -Proc $ScanProc -Mode $ModeName -Y $Y -AllowQuick $false
            if ($res -eq 'Skipped') { $Skipped = $true }
        }

        if ($Skipped) {
            Write-Field "Scan Status" "SKIPPED BY USER" "Yellow"
        } else {
            Write-Field "Scan Status" "COMPLETE" "Green"
        }
    }
    #endregion

    #region [5] - THREAT IDENTIFICATION & AUDIT
    # --------------------------------------------------------------------------
    $Threats   = @(Get-DefendActiveThreats)
    $ScanStats = if ($PSVer -ge 6) { (Invoke-DefendWinPS 'Get-MpComputerStatus | ConvertTo-Json -Depth 3').Output | ConvertFrom-Json } else { Get-MpComputerStatus }
    Write-Host "[>] Finalizing Security Audit..." -ForegroundColor DarkCyan
    $Log = 'Microsoft-Windows-Windows Defender/Operational'
    $FullScanEvent = Get-WinEvent -FilterHashtable @{LogName=$Log; Id=1001} -ErrorAction SilentlyContinue |
                     Where-Object { $_.Message -like "*Full Scan*" } | Select-Object -First 1
    $QuickScanEvent = Get-WinEvent -FilterHashtable @{LogName=$Log; Id=1001} -ErrorAction SilentlyContinue |
                      Where-Object { $_.Message -like "*Quick Scan*" } | Select-Object -First 1
    $FullDisplay = if ($FullScanEvent) { [string]$FullScanEvent.TimeCreated } else { "No Completed Full Scan Found" }
    $QuickDisplay = if ($QuickScanEvent) { [string]$QuickScanEvent.TimeCreated } else { "No Completed Quick Scan Found" }
    Write-Field "Last Successful Full Scan" $FullDisplay "Yellow"
    Write-Field "Last Successful Quick Scan" $QuickDisplay "Yellow"
    Write-Field "Antivirus Signature" "v$($ScanStats.AntivirusSignatureVersion)" "Yellow"
    if ($Threats.Count -gt 0) {
        [void]$script:DefendFailures.Add("$($Threats.Count) active threat(s) detected")
        Write-Alert " [!] CRITICAL: $($Threats.Count) Threat(s) Detected on $env:COMPUTERNAME "
        foreach ($T in $Threats) {
            $TName = if ($T.ThreatName) { $T.ThreatName } else { "Unknown Signature" }
            if ($T.Resources) {
                $TPath = $T.Resources -join ", "
            } else {
                $ThreatEvent = Get-WinEvent -FilterHashtable @{LogName=$Log; Id=1116} -ErrorAction SilentlyContinue |
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

    # REBOOT PENDING CHECK
    $RebootPending = (Test-DefendRegKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or
                     (Test-DefendRegKey 'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending')
    if ($RebootPending) {
        Write-Host "[!] ACTION REQUIRED: A System Reboot is Pending." -ForegroundColor Red
        Write-Host "    Some security enforcements may not be active yet." -ForegroundColor Gray
    } else {
        Write-Host "[+] System State: No Reboot Required." -ForegroundColor Green
    }

    # Footer (reuses header art - art on right, fill on left)
    $_ffillW = $script:Width - $_artW - 2
    $_footer = "  ENDPOINT SECURITY ENFORCEMENT COMPLETE"
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    #endregion

    if ($script:DefendFailures.Count -gt 0) {
        foreach ($m in $script:DefendFailures) { Write-Output "[!] $m" }
        Write-Output ("[WARN] DEFEND partial: {0} (exit 3)" -f ($script:DefendFailures -join "; "))
        Exit-KillerScript 3
    }
    if ($RebootPending) {
        Write-Output "[PASS] DEFEND complete, reboot required (exit 3010)"
        Exit-KillerScript 3010
    }
    Write-Output "[PASS] DEFEND complete (exit 0)"
    Exit-KillerScript 0
}

function Wait-DefendScan {
    # Waits for a scan process. Returns Exited, Skipped (Esc) or Pivot (Q).
    param($Proc, [string]$Mode, [int]$Y, [bool]$AllowQuick)
    $start = Get-DefendNow
    while (-not $Proc.HasExited) {
        if (Test-DefendKeyAvailable) {
            $key = Read-DefendKey
            if ($key -and $key.Key -eq 'Escape') { Stop-DefendScan $Proc; return 'Skipped' }
            if ($key -and $AllowQuick -and $key.KeyChar -eq 'q') { Stop-DefendScan $Proc; return 'Pivot' }
        }
        $timer = "{0:hh\:mm\:ss}" -f ((Get-DefendNow) - $start)
        if ($Y -ge 0 -and (Set-DefendCursor 0 $Y)) {
            Write-Host ("Scan Activity : RUNNING ($Mode) [$timer]".PadRight(100)) -NoNewline
        }
        Start-Sleep -Milliseconds 500
    }
    if ($Y -ge 0 -and (Set-DefendCursor 0 $Y)) { Write-Host (" " * 100) -NoNewline; [void](Set-DefendCursor 0 $Y) }
    'Exited'
}

if (-not $global:KillerScriptsTestMode) { Invoke-DefendMain }
