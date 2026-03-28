<#
.SYNOPSIS
    Foxit Audit and Control Task Script (FACTS) v2.3
    Developed by SteveTheKiller | Updated: 2026-03-10
.DESCRIPTION
    Audits for Foxit PDF installations and permanently blocks
    auto-updates via registry hardening, service suppression, and
    scheduled task enforcement. Creates a self-healing hourly
    maintenance task when Foxit is detected, and returns exit code 0
    for RMM compatibility.
#>

#region Pre-Flight Checks
# ============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit 1  # Exit with error code if not admin
}

$script:StepRow = -1
$script:StepMsg = ""

function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)

    if ($Success) {
        # Complete: go back to saved row, rewrite in DimCol + WarnCol CustomInfo + right-aligned [SUCCESS]
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
    } elseif ($Message -and -not $CustomInfo) {
        # Announcement: save row, print in InfoCol
        $script:StepRow = [Console]::CursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $InfoCol
    } else {
        # Single-call skip line (Message + CustomInfo, no Success)
        Write-Host $Message -ForegroundColor $DimCol -NoNewline
        if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
        Write-Host ""
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
Clear-Host
$script:Width    = 85

$LineCol   = "DarkMagenta"
$WarnCol   = "DarkYellow"
$BorderCol = "Magenta"
$ArtCol    = "Yellow"
$AccentCol = "Green"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"
$_ver      = "| v2.5"

# Header Art & Logic
$_pfx  = "█  "
$_art1 = "╔═╗ ╔═╗ ╔═╗ ╔╦╗ ╔═╗ "
$_art2 = "╠╣  ╠═╣ ║    ║  ╚═╗ "
$_art3 = "╩   ╩ ╩ ╚═╝  ╩  ╚═╝ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "FOXIT AUDIT AND CONTROL TASK SCRIPT"

# Safety padding to prevent negative multiplication error
$_tpad  = " " * [Math]::Max(0, ($_fillW - $_title.Length))

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $BorderCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
#endregion

#region 1. Audit & Discovery
# ============================================================================
Write-StepUpdate "[1/4] Auditing Foxit Installations..."
$FoxitRegistryPaths = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
)

$FoundApps = Get-ItemProperty $FoxitRegistryPaths | Where-Object { $_.DisplayName -like "*Foxit*" }
if ($FoundApps) {
    Write-StepUpdate -Success -CustomInfo "($($FoundApps.Count) Found)"
} else {
    Write-StepUpdate -Success -CustomInfo "(None Detected)"
}
#endregion

#region 2. Universal Registry Hardening
# ============================================================================
if ($FoundApps) {
    Write-StepUpdate "[2/4] Applying Recursive Registry Blocks..."
    $MachineHives = @("HKLM:\SOFTWARE\Foxit Software", "HKLM:\SOFTWARE\WOW6432Node\Foxit Software")
    foreach ($Hive in $MachineHives) {
        if (Test-Path $Hive) {
            $UpdaterKeys = Get-ChildItem -Path $Hive -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSChildName -eq "Updater" }
            foreach ($Key in $UpdaterKeys) {
                New-ItemProperty -Path $Key.PSPath -Name "UpdateMode" -Value "0" -PropertyType String -Force | Out-Null
                New-ItemProperty -Path $Key.PSPath -Name "bNoUpdate" -Value 1 -PropertyType DWord -Force | Out-Null
            }
        }
    }
    Write-StepUpdate -Success
} else {
    Write-StepUpdate "[2/4] Skipping Registry Hardening..." -CustomInfo "(Not Required)"
}
#endregion

#region 3. Service & Task Enforcement
# ============================================================================
if ($FoundApps) {
    Write-StepUpdate "[3/4] Disabling Native Services and Tasks..."
    Get-Service | Where-Object { $_.Name -like "*Foxit*Update*" } | ForEach-Object {
        Set-Service -Name $_.Name -StartupType Disabled
        Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
    }
    Get-ScheduledTask -TaskName "*Foxit*" | Where-Object { $_.TaskName -like "*Update*" } | ForEach-Object {
        Stop-ScheduledTask -TaskName $_.TaskName -ErrorAction SilentlyContinue
        Disable-ScheduledTask -TaskName $_.TaskName -ErrorAction SilentlyContinue
    }
    Write-StepUpdate -Success
} else {
    Write-StepUpdate "[3/4] Skipping Service/Task Block..." -CustomInfo "(Not Required)"
}
#endregion

#region 4. Hourly Self-Healing Task
# ============================================================================
$TaskName = "FACTS_Universal_Enforcement"

if ($FoundApps) {
    Write-StepUpdate "[4/4] Configuring Persistence (Hourly Task)..."
    $MaintenanceScript = {
        $AuditPaths = @("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*")
        if (Get-ItemProperty $AuditPaths | Where-Object { $_.DisplayName -like "*Foxit*" }) {
            Get-Service | Where-Object {$_.Name -like "*Foxit*Update*"} | ForEach-Object {Set-Service -Name $_.Name -StartupType Disabled; Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue}
            Get-ScheduledTask -TaskName "*Foxit*" | Where-Object {$_.TaskName -like "*Update*"} | ForEach-Object {Stop-ScheduledTask -TaskName $_.TaskName -ErrorAction SilentlyContinue; Disable-ScheduledTask -TaskName $_.TaskName -ErrorAction SilentlyContinue}
        }
    }.ToString()
    $EncodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($MaintenanceScript))
    $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -WindowStyle Hidden -EncodedCommand $EncodedCommand"
    $Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)
    try {
        Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
        Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -User "SYSTEM" -RunLevel Highest -Force | Out-Null
        Write-StepUpdate -Success
    }
    catch {
        Write-Host " [FAILED]" -ForegroundColor Red
        exit 1 # Exit with error if task registration fails
    }
} else {
    Write-StepUpdate "[4/4] Removing existing Maintenance Tasks..."
    Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction SilentlyContinue
    Write-StepUpdate -Success -CustomInfo "(No Foxit Detected)"
}
# Footer (reuses header art - art on right, fill on left)
$_ffillW = $script:Width - $_artW - 2  # 1 leading space + 1 sfx char
$_footer = "  FOXIT MAINTENANCE ENFORCEMENT COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $BorderCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Exit 0
