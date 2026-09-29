<#
.SYNOPSIS
    Advanced Maintenance, Optimization, and Repair Tool (AMORT) v17.0
    Developed by Steve the Killer | Updated: 2026-09-29
.DESCRIPTION
    One-step Windows 10/11 disk-space reclamation and integrity repair for MSP
    field and remote use. Built to be safe on personal and business machines
    while a user is signed in and working: no user documents, mail stores,
    sync caches, or search indexes are touched, and anything that could still
    be needed is age-gated or retained.

    Removes Dell SupportAssist snapshots, feature-upgrade leftovers, OEM driver
    extracts, Windows.old, Delivery Optimization and Windows Update download
    caches, browser/Teams/GPU caches, and aged temp files, crash dumps, error
    reports, and Recycle Bin items. Trims restore points to the newest few,
    compresses system logs in place, runs DISM and SFC repair, disables
    hibernation, and performs SSD TRIM. Each step reports the measured space
    it recovered, and the summary is sized to screenshot into a ticket.
.PARAMETER DryRun
    Read-only estimate mode. Makes no changes: every destructive step is skipped
    and each target is sized instead, reporting estimated reclaim per category
    plus a projected free-space total.
.PARAMETER RetainDays
    Crash dumps, error reports, and Recycle Bin items newer than this are kept. Default 14.
.PARAMETER KeepRestorePoints
    Number of newest restore points (shadow copies of C:) to keep. Default 2.
#>
param(
    [switch]$DryRun,
    [ValidateRange(1, 365)][int]$RetainDays = 14,
    [ValidateRange(1, 64)][int]$KeepRestorePoints = 2
)
$TempAgeDays = 2
$_fver   = "| v17.0"
#region Pre-Flight Checks
# ============================================================================
# Force UTF-8 output so box-drawing characters render correctly
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding            = [System.Text.Encoding]::UTF8

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit
}
# Detect active Windows servicing - do NOT kill it.
# Force-killing TiWorker/DISM mid-servicing corrupts the component store, which
# the repair region would then have to fix. Instead we detect and skip repair.
$ServicingActive = $null -ne (Get-Process -Name "TiWorker", "DISM" -ErrorAction SilentlyContinue)
if ($ServicingActive) {
    Write-Host "[Pre-Flight] Windows servicing active (TiWorker/DISM running). Repair steps will be skipped." -ForegroundColor Yellow
}
# A pending reboot means SoftwareDistribution and upgrade staging are still in use.
$PendingReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"

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
        Write-Host (" " * $script:Width) -NoNewline
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
            } else {
                # Right-align the info so it ends one space left of the 9-character status tag column
                $infoCol = if ($CustomInfo -in @("Saved: 0 MB", "Est: 0 MB")) { "Gray" } elseif ($CustomInfo.StartsWith("Saved:")) { "Red" } elseif ($CustomInfo.StartsWith("Est:")) { "Magenta" } else { "Gray" }
                $currentCol = [Console]::CursorLeft
                $targetCol  = $script:Width - 10 - $CustomInfo.Length
                $pad = [Math]::Max(1, $targetCol - $currentCol)
                Write-Host ((" " * $pad) + $CustomInfo) -NoNewline -ForegroundColor $infoCol
            }
        }

        # Final Success Tag (right-aligned to console width)
        if ($Success) {
            $tag = if ($script:DryRun) { "[EST]" } else { "[SUCCESS]" }
            $currentCol = [Console]::CursorLeft
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor $(if ($script:DryRun) { "Magenta" } else { "Green" })
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
# Dry-run size estimator (read-only; handles missing paths and wildcards)
function Get-PathSize {
    param([string]$Path)
    try {
        $sum = (Get-ChildItem -Path $Path -Recurse -Force -File -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
        if ($null -eq $sum) { return 0 }
        return [int64]$sum
    } catch { return 0 }
}
# Dry-run per-step reporter: prints an estimate tag and accumulates the total
function Write-DryEstimate {
    param([int64]$Bytes)
    if ($Bytes -gt 0) {
        $s = if ($Bytes -ge 1GB) { "{0:N2} GB" -f ($Bytes / 1GB) } else { "{0:N2} MB" -f ($Bytes / 1MB) }
        Write-StepUpdate -Success -CustomInfo "Est: $s"
    } else {
        Write-StepUpdate -Success -CustomInfo "Est: 0 MB"
    }
    if (-not $script:EstYieldBytes) { $script:EstYieldBytes = 0 }
    $script:EstYieldBytes = [int64]$script:EstYieldBytes + [int64]$Bytes
}
# Environment Setup
# Expose DryRun at script scope so the output helpers can see it
$script:DryRun = [bool]$DryRun
$Drive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$StartSpace = $Drive.FreeSpace
$TotalSize = $Drive.Size
# Initialize cumulative yield (bytes)
if (-not $TotalYieldBytes) { $TotalYieldBytes = 0 }
$script:EstYieldBytes = 0
# Ensure TotalSize is valid
$TotalSize = [double]$TotalSize
$script:RegionHistory = @()
if ($TotalSize -le 0) { throw "TotalSize is zero or undefined. Aborting." }

$StartUsagePct = [Math]::Round(((($TotalSize - $StartSpace) / $TotalSize) * 100), 2)
$LastRegionSpace = $Drive.FreeSpace # Rolling baseline for step-by-step reporting
$CS = Get-CimInstance Win32_ComputerSystem
$Vendor = $CS.Manufacturer
$IsVM = ($Vendor -match "QEMU|VMware|Virtual|Hyper-V")
# --- Custom-build architecture display ---
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
$script:Width    = 90
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
$_title = "ADVANCED MAINTENANCE, OPTIMIZATION, & REPAIR TOOL"

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
if ($DryRun) {
    Write-Host "      Mode: DRY RUN - estimate only, no changes will be made" -ForegroundColor Magenta
}
if ($IsVM) {
    Write-Host "      Mode: Virtual Machine" -ForegroundColor Yellow 
}
#endregion

#region Cleanup Helpers
# ============================================================================
# Every cleanup step is built to be safe on a live endpoint with a user signed in.
# Nothing here touches user documents, mail stores, or sync caches. Anything that
# could still be in use is either age-gated or skipped when locked.
function Format-Size {
    param([int64]$Bytes)
    if ($Bytes -ge 1GB) { return "{0:N2} GB" -f ($Bytes / 1GB) }
    return "{0:N2} MB" -f ($Bytes / 1MB)
}
function Get-FreeBytes {
    return [int64](Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace
}
# Closes out a step. Live runs report the measured free-space change on C: since the
# previous step, so every "Saved" figure is real disk space, not an estimate.
function Complete-Step {
    param([int64]$Estimate = 0, [switch]$NoEstimate)
    if ($script:DryRun) {
        if ($NoEstimate) { Write-StepUpdate -Success -CustomInfo "Est: varies" }
        else { Write-DryEstimate $Estimate }
        return
    }
    $Now = Get-FreeBytes
    $Delta = [int64]($Now - $script:LastRegionSpace)
    $script:LastRegionSpace = $Now
    if ($Delta -ge 1MB) { Write-StepUpdate -Success -CustomInfo "Saved: $(Format-Size $Delta)" }
    else { Write-StepUpdate -Success -CustomInfo "Saved: 0 MB" }
}
# Deletes (or sizes, in dry run) files under each path whose last write is older than
# the cutoff. Locked files are skipped. Returns the byte total in dry run.
function Remove-OldFiles {
    param([string[]]$Path, [int]$Days)
    $Cutoff = (Get-Date).AddDays(-$Days)
    $Sum = [int64]0
    foreach ($P in $Path) {
        if (-not (Test-Path $P)) { continue }
        Get-ChildItem -Path $P -Recurse -Force -File -ErrorAction SilentlyContinue |
            Where-Object { $_.LastWriteTime -lt $Cutoff } |
            ForEach-Object {
                if ($script:DryRun) { $Sum += $_.Length }
                else { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
            }
    }
    return $Sum
}
# Deletes the contents of cache folders (wildcards allowed). Locked files are skipped.
function Clear-CachePath {
    param([string[]]$Path)
    $Sum = [int64]0
    foreach ($P in $Path) {
        if (-not (Test-Path $P)) { continue }
        if ($script:DryRun) { $Sum += Get-PathSize $P }
        else { Remove-Item $P -Recurse -Force -ErrorAction SilentlyContinue }
    }
    return $Sum
}
# Removes a whole folder tree, retrying once with ownership taken if rd hits ACL issues.
function Remove-Tree {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return }
    & cmd.exe /c "rd /s /q `"$Path`"" 2>$null | Out-Null
    if (Test-Path -LiteralPath $Path) {
        & takeown /F $Path /R /A /D Y 2>$null | Out-Null
        & icacls $Path /grant "*S-1-5-32-544:F" /T /C /Q 2>$null | Out-Null
        & cmd.exe /c "rd /s /q `"$Path`"" 2>$null | Out-Null
    }
}
$UserProfiles = Get-ChildItem "C:\Users" -Directory -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @("All Users", "Default User") -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) }
$script:Notes = @()
#endregion

#region 1. Snapshots, Upgrade Leftovers & OEM Installers
# ============================================================================
Write-StepUpdate "[01/11] Purging Snapshots, Upgrade Leftovers & OEM Installers"
$RegionEst = [int64]0
# Dell SupportAssist Remediation snapshot purge
# SupportAssist OS Recovery stores system-repair snapshots under
# SARemediation\SystemRepair\{Snapshots,Backup} that can grow to 20-80GB when
# auto-purge stalls. Stop the service, clear the snapshot payload only (not the
# whole tree, so the install keeps working), then restart. This removes the local
# OS-recovery snapshots until SupportAssist rebuilds one.
if ($Vendor -like "*Dell*" -and -not $IsVM) {
    $SARoot = "C:\ProgramData\Dell\SARemediation\SystemRepair"
    if (Test-Path $SARoot) {
        if ($DryRun) {
            foreach ($Sub in @("Snapshots", "Backup")) { $RegionEst += Get-PathSize (Join-Path $SARoot $Sub) }
        } else {
            $SASvcs = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "SupportAssist*" -or $_.DisplayName -like "*SupportAssist*" }
            foreach ($SASvc in $SASvcs) {
                try { Stop-Service $SASvc.Name -Force -ErrorAction Stop -WarningAction SilentlyContinue } catch { }
            }
            foreach ($Sub in @("Snapshots", "Backup")) {
                $SAPath = Join-Path $SARoot $Sub
                if (Test-Path $SAPath) {
                    Start-Process "cmd.exe" -ArgumentList "/c attrib -h -s -r `"$SAPath\*`" /S /D & del /s /f /q `"$SAPath\*`"" -WindowStyle Hidden -Wait
                }
            }
            foreach ($SASvc in $SASvcs) {
                Start-Service $SASvc.Name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            }
        }
    }
}
# Delivery Optimization cache (peer update cache; Windows re-downloads as needed)
$RegionEst += Clear-CachePath @("C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache\*")

# Feature upgrade staging folders. These are left behind after a feature update or a
# failed upgrade attempt. They are only skipped while an upgrade is actually running or
# waiting on its reboot, because that is the one time Windows still needs them.
$UpgradeActive = $null -ne (Get-Process -Name "SetupHost", "SetupPrep", "Windows10UpgraderApp" -ErrorAction SilentlyContinue)
$UpgradeDirs = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS', 'C:\$GetCurrent', 'C:\ESD', 'C:\Windows10Upgrade')
if ($UpgradeActive -or $PendingReboot) {
    # A pending reboot is reported once, on the reboot-pending line before the repair steps.
    if (-not $PendingReboot) { $script:Notes += "Upgrade staging folders were kept because a Windows upgrade is running." }
} else {
    foreach ($D in $UpgradeDirs) {
        if (Test-Path -LiteralPath $D) {
            if ($DryRun) { $RegionEst += Get-PathSize $D } else { Remove-Tree $D }
        }
    }
}
# OEM driver and utility extraction folders. These hold unpacked installers that were
# already run. Only the known OEM extract locations are removed, never C:\Drivers or
# other folders a user or admin may have created on purpose.
$OemDirs = @('C:\SWSetup', 'C:\Dell\Drivers', 'C:\Dell\UpdatePackage', 'C:\AMD', 'C:\NVIDIA')
foreach ($D in $OemDirs) {
    if (Test-Path -LiteralPath $D) {
        if ($DryRun) { $RegionEst += Get-PathSize $D } else { Remove-Tree $D }
    }
}

# --- Windows.old cleanup (no age gate) ---
# This removes Windows.old as soon as it is found, regardless of how old it is. That
# means the "Go back to a previous version of Windows" rollback option is closed off the
# moment this step runs, even if the feature update happened minutes ago. This is an
# intentional tradeoff, not an oversight; if that rollback path ever needs to be
# preserved on a specific machine, skip this run or exclude that machine.
if (Test-Path "C:\Windows.old") {
    if ($DryRun) {
        $RegionEst += Get-PathSize "C:\Windows.old"
    } else {
        # Strip the previous-installation protection so Windows.old is deletable.
        & DISM.exe /Online /Remove-OSUninstall /NoRestart *>&1 | Out-Null

        # cleanmgr is skipped entirely: it ignores -WindowStyle Hidden (spawns an uncontrolled
        # child process), shows its UI in interactive sessions, and silently fails under
        # SYSTEM/LiveConnect where there is no desktop. rd /s /q is an order of magnitude
        # faster than Remove-Item -Recurse for deep trees (minutes vs hours).
        $rdProc = Start-Process "cmd.exe" -ArgumentList "/c rd /s /q `"C:\Windows.old`"" -WindowStyle Hidden -PassThru -ErrorAction SilentlyContinue
        if ($rdProc) {
            $rdProc | Wait-Process -Timeout 1800 -ErrorAction SilentlyContinue
            if (-not $rdProc.HasExited) { $rdProc | Stop-Process -Force -ErrorAction SilentlyContinue }
        }

        # Fallback: if rd hit locked files, take ownership and retry once
        if (Test-Path "C:\Windows.old") {
            & takeown /F "C:\Windows.old" /R /A /D Y 2>$null | Out-Null
            & icacls "C:\Windows.old" /grant "*S-1-5-32-544:F" /T /C /Q 2>$null | Out-Null
            Start-Process "cmd.exe" -ArgumentList "/c rd /s /q `"C:\Windows.old`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
        }

        # Fallback: robocopy mirror against an empty folder. This handles long paths and
        # stale ACL entries far more reliably than rd or Remove-Item.
        if (Test-Path "C:\Windows.old") {
            $EmptyDir = Join-Path $env:TEMP ("amort_empty_" + [guid]::NewGuid().ToString())
            New-Item -ItemType Directory -Path $EmptyDir -Force | Out-Null
            & robocopy.exe $EmptyDir "C:\Windows.old" /MIR /NFL /NDL /NJH /NJS /NC /NS /NP *>&1 | Out-Null
            Remove-Item $EmptyDir -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path "C:\Windows.old") {
                Remove-Item "C:\Windows.old" -Recurse -Force -ErrorAction SilentlyContinue
            }
        }

        # Last resort: whatever is left has an open handle, so schedule it for deletion on
        # the next boot (the same mechanism Windows uses for PendingFileRenameOperations).
        if (Test-Path "C:\Windows.old") {
            try {
                Add-Type -Name Kernel32 -Namespace AmortNative -MemberDefinition @'
[DllImport("kernel32.dll", SetLastError = true, CharSet = CharSet.Auto)]
public static extern bool MoveFileEx(string lpExistingFileName, string lpNewFileName, int dwFlags);
'@ -ErrorAction SilentlyContinue

                Get-ChildItem "C:\Windows.old" -Recurse -Force -ErrorAction SilentlyContinue |
                    Sort-Object { $_.FullName.Length } -Descending |
                    ForEach-Object { [AmortNative.Kernel32]::MoveFileEx($_.FullName, $null, 4) | Out-Null }
                [AmortNative.Kernel32]::MoveFileEx("C:\Windows.old", $null, 4) | Out-Null
                $script:Notes += "Windows.old was partly in use and is scheduled for deletion on the next reboot."
            } catch {
                $script:Notes += "Windows.old is still present after every removal attempt."
            }
        }
    }
}
Complete-Step -Estimate $RegionEst
#endregion

#region 2. Crash Dumps & Error Reports (age-gated)
# ============================================================================
Write-StepUpdate "[02/11] Purging Crash Dumps & Error Reports older than $RetainDays days"
# Recent dumps and reports are kept so there is always evidence for troubleshooting.
$DumpPaths = @(
    "C:\Windows\MEMORY.DMP",
    "C:\Windows\Minidump",
    "C:\Windows\LiveKernelReports",
    "C:\ProgramData\Microsoft\Windows\WER\ReportArchive",
    "C:\ProgramData\Microsoft\Windows\WER\ReportQueue",
    "C:\ProgramData\Microsoft\Windows\WER\Temp"
)
foreach ($U in $UserProfiles) {
    $DumpPaths += "$($U.FullName)\AppData\Local\CrashDumps"
    $DumpPaths += "$($U.FullName)\AppData\Local\Microsoft\Windows\WER"
}
$RegionEst = Remove-OldFiles -Path $DumpPaths -Days $RetainDays
Complete-Step -Estimate $RegionEst
#endregion

#region 3. Temp, Browser, Teams & GPU Caches
# ============================================================================
Write-StepUpdate "[03/11] Purging Temp, Browser, Teams, and GPU Caches"
$RegionEst = [int64]0
# Temp files are age-gated so installers and apps running right now keep their working files.
$TempPaths = @("C:\Windows\Temp", "C:\Windows\SystemTemp")
foreach ($U in $UserProfiles) { $TempPaths += "$($U.FullName)\AppData\Local\Temp" }
$RegionEst += Remove-OldFiles -Path $TempPaths -Days $TempAgeDays

# Cache folders only. Profiles, cookies, sign-ins, history, and settings are never touched.
foreach ($U in $UserProfiles) {
    $UP = $U.FullName
    $CachePaths = @(
        "$UP\AppData\Local\Google\Chrome\User Data\*\Cache\*",
        "$UP\AppData\Local\Google\Chrome\User Data\*\Code Cache\*",
        "$UP\AppData\Local\Microsoft\Edge\User Data\*\Cache\*",
        "$UP\AppData\Local\Microsoft\Edge\User Data\*\Code Cache\*",
        "$UP\AppData\Local\Mozilla\Firefox\Profiles\*\cache2\*",
        "$UP\AppData\Local\BraveSoftware\Brave-Browser\User Data\*\Cache\*",
        "$UP\AppData\Local\Opera Software\Opera Stable\Cache\*",
        # Classic Teams
        "$UP\AppData\Roaming\Microsoft\Teams\Cache\*",
        "$UP\AppData\Roaming\Microsoft\Teams\Code Cache\*",
        "$UP\AppData\Roaming\Microsoft\Teams\GPUCache\*",
        "$UP\AppData\Roaming\Microsoft\Teams\Service Worker\CacheStorage\*",
        # New Teams (WebView2 cache folders only, so the user stays signed in)
        "$UP\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\*\Cache\*",
        "$UP\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\*\Code Cache\*",
        "$UP\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\*\GPUCache\*",
        "$UP\AppData\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\EBWebView\*\Service Worker\CacheStorage\*",
        # GPU shader caches (rebuilt automatically)
        "$UP\AppData\Local\D3DSCache\*",
        "$UP\AppData\Local\AMD\DxCache\*",
        "$UP\AppData\Local\NVIDIA\GLCache\*",
        "$UP\AppData\Local\NVIDIA\DXCache\*"
    )
    $RegionEst += Clear-CachePath $CachePaths
}
Complete-Step -Estimate $RegionEst
#endregion

#region 4. Recycle Bin Purge (age-gated, all users)
# ============================================================================
Write-StepUpdate "[04/11] Emptying Recycle Bin items older than $RetainDays days"
# Each deleted item is a $I metadata file (holding the original size and the deletion
# time) paired with a $R file or folder holding the data. Only pairs deleted more than
# $RetainDays days ago are removed, so anything a user binned recently can still be restored.
$RegionEst = [int64]0
$RBCutoff = (Get-Date).AddDays(-$RetainDays)
$RBRoot = 'C:\$Recycle.Bin'
if (Test-Path -LiteralPath $RBRoot) {
    Get-ChildItem -LiteralPath $RBRoot -Directory -Force -ErrorAction SilentlyContinue | ForEach-Object {
        Get-ChildItem -LiteralPath $_.FullName -Force -File -Filter '$I*' -ErrorAction SilentlyContinue | ForEach-Object {
            $IFile = $_
            $DeletedOn = $null
            $OrigSize = [int64]0
            try {
                $Bytes = [IO.File]::ReadAllBytes($IFile.FullName)
                if ($Bytes.Length -ge 24) {
                    $OrigSize = [BitConverter]::ToInt64($Bytes, 8)
                    $DeletedOn = [DateTime]::FromFileTime([BitConverter]::ToInt64($Bytes, 16))
                }
            } catch { }
            if (-not $DeletedOn) { $DeletedOn = $IFile.LastWriteTime }
            if ($DeletedOn -lt $RBCutoff) {
                $RPath = Join-Path $IFile.DirectoryName ('$R' + $IFile.Name.Substring(2))
                if ($DryRun) {
                    if (Test-Path -LiteralPath $RPath) { $RegionEst += [Math]::Max([int64]0, $OrigSize) }
                } else {
                    if (Test-Path -LiteralPath $RPath) { Remove-Item -LiteralPath $RPath -Recurse -Force -ErrorAction SilentlyContinue }
                    if (-not (Test-Path -LiteralPath $RPath)) { Remove-Item -LiteralPath $IFile.FullName -Force -ErrorAction SilentlyContinue }
                }
            }
        }
    }
}
Complete-Step -Estimate $RegionEst
#endregion

#region 5. Windows Update Download Cache
# ============================================================================
Write-StepUpdate "[05/11] Clearing Windows Update Download Cache"
# Only the Download folder is cleared. DataStore (update history) is kept, and service
# startup types are left at their Windows defaults. BITS is never stopped so the RMM
# and LiveConnect stay connected.
$WUDownload = "C:\Windows\SoftwareDistribution\Download"
if ($PendingReboot) {
    Write-StepUpdate -CustomInfo "[SKIPPED]"
} elseif ($DryRun) {
    Write-DryEstimate (Get-PathSize $WUDownload)
} else {
    $WUWasRunning = (Get-Service wuauserv -ErrorAction SilentlyContinue).Status -eq 'Running'
    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    if (Test-Path $WUDownload) { Remove-Item "$WUDownload\*" -Recurse -Force -ErrorAction SilentlyContinue }
    if ($WUWasRunning) { Start-ServiceSilent wuauserv }
    Complete-Step
}
#endregion

#region 6. Restore Points & Shadow Copies
# ============================================================================
Write-StepUpdate "[06/11] Trimming Restore Points (keeping newest $KeepRestorePoints)"
# Restore points are shadow copies of C:. The newest ones are always kept so the machine
# can still be rolled back. Only older copies beyond that count are deleted.
$RegionEst = [int64]0
$CVol = Get-CimInstance Win32_Volume -Filter "DriveLetter='C:'" -ErrorAction SilentlyContinue
$Shadows = @()
if ($CVol) {
    $Shadows = @(Get-CimInstance Win32_ShadowCopy -ErrorAction SilentlyContinue |
        Where-Object { $_.VolumeName -eq $CVol.DeviceID } |
        Sort-Object InstallDate -Descending)
}
$OldShadows = @($Shadows | Select-Object -Skip $KeepRestorePoints)
if ($OldShadows.Count -gt 0) {
    if ($DryRun) {
        # Shadow storage is shared, so estimate the old copies' share of what is in use.
        $Storage = Get-CimInstance Win32_ShadowStorage -ErrorAction SilentlyContinue |
            Where-Object { $_.Volume.DeviceID -eq $CVol.DeviceID } | Select-Object -First 1
        if ($Storage) { $RegionEst = [int64]([double]$Storage.UsedSpace * $OldShadows.Count / $Shadows.Count) }
    } else {
        foreach ($S in $OldShadows) {
            & vssadmin.exe delete shadows "/shadow=$($S.ID)" /quiet 2>$null | Out-Null
        }
    }
}
Complete-Step -Estimate $RegionEst
#endregion

#region 7. Log Compaction
# ============================================================================
Write-StepUpdate "[07/11] Compressing System Logs (NTFS, nothing deleted)"
# Logs are compressed in place, so every log is still there and readable for audits.
# Files that are open are skipped.
$LogDirs = @("C:\Windows\Logs", "C:\Windows\Panther", "C:\Windows\System32\LogFiles", "C:\ProgramData\Microsoft\Windows\WER")
if ($DryRun) {
    Complete-Step -NoEstimate
} else {
    foreach ($L in $LogDirs) {
        if (Test-Path $L) { & compact.exe /c "/s:$L" /i /q 2>$null | Out-Null }
    }
    Complete-Step
}
#endregion


#region 5. Repair & Integrity
# ============================================================================
if ($IsVM) { [System.GC]::Collect() }
Write-Progress -Activity "Cleaning up" -Completed

# Pre-repair: stability check - warn only, never skip on PendingRename alone
# (PendingFileRenameOperations is routinely re-created by Windows/installers and
# does not block DISM or SFC). Dry run, active servicing, and a pending reboot DO skip.
$SkipRepair = $false
$PendingRename = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
$HasPendingRename = $null -ne $PendingRename
if ($DryRun) {
    Write-Host "        [i] Dry run - DISM and SFC repair steps skipped (no changes)." -ForegroundColor Magenta
}
elseif ($PendingReboot) {
    Write-Host "        [!] Reboot pending, so upgrade folders, WU cache, and DISM/SFC were skipped." -ForegroundColor DarkYellow
}
elseif ($ServicingActive) {
    Write-Host "        [!] Windows servicing active (TiWorker/DISM running) - DISM and SFC repair steps will be skipped." -ForegroundColor DarkYellow
}
elseif ($HasPendingRename) {
    Write-Host "        [!] PendingFileRenameOperations found - skipping repair steps." -ForegroundColor DarkYellow
}
# Ensure TrustedInstaller is available (prevents DISM Error 87 / SFC failures)
if (-not $SkipRepair -and -not $DryRun) {
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
            $width  = $script:Width
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
# --- STEP 5: RestoreHealth ---
            Clear-InputBuffer
            $S72 = "[08/11] DISM RestoreHealth"
            Write-StepUpdate $S72 -CustomInfo "[Press ESC to Skip]"
            $Row72 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($DryRun -or $PendingReboot -or $HasPendingRename -or $ServicingActive) {
                Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[SKIPPED]"
            }
            else {
                $DismSpin = [char[]]@('|','/','-','\')
                $DismTmp1 = [System.IO.Path]::GetTempFileName()

                # Use .NET Process directly; Start-Process -PassThru returns $null ExitCode
                # when combined with -RedirectStandardOutput. Wrap in cmd.exe for file redirection.
                $psi1 = New-Object System.Diagnostics.ProcessStartInfo
                $psi1.FileName               = "cmd.exe"
                $psi1.Arguments              = "/c dism.exe /Online /Cleanup-Image /RestoreHealth /NoRestart > `"$DismTmp1`" 2>&1"
                $psi1.UseShellExecute        = $false
                $psi1.CreateNoWindow         = $true
                $psi1.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Proc1 = [System.Diagnostics.Process]::Start($psi1)

                $Skipped1 = $false
                $DismSpinIdx1 = 0
                $DismTimer1 = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc1.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                try { $Proc1.Kill() } catch {}
                                Stop-DismTree
                                $Skipped1 = $true
                                break
                            }
                        }
                    } catch {}
                    try {
                        [Console]::SetCursorPosition(0, $Row72)
                        $EscHint = "[ESC to skip]"
                        $Width = $script:Width
                        $Left = "$S72 $($DismSpin[$DismSpinIdx1 % 4]) $($DismTimer1.Elapsed.ToString('mm\:ss'))"
                        $Spaces = [Math]::Max(1, $Width - $Left.Length - $EscHint.Length)
                        [Console]::ForegroundColor = [ConsoleColor]::Cyan
                        [Console]::Write($Left + (' ' * $Spaces))
                        [Console]::ForegroundColor = [ConsoleColor]::DarkGray
                        [Console]::Write($EscHint)
                        [Console]::ResetColor()
                    } catch {}
                    $DismSpinIdx1++
                    Start-Sleep -Milliseconds 250
                }
                $DismTimer1.Stop()

                if (-not $Skipped1) {
                    $Proc1.WaitForExit()
                    Stop-TiWorker
                }
                $ExitCode1 = $Proc1.ExitCode
                try { $Proc1.Dispose() } catch {}
                Remove-Item $DismTmp1 -Force -ErrorAction SilentlyContinue

                if ($Skipped1) {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[SKIPPED]"
                }
                elseif ($ExitCode1 -in @(0, 3010)) {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -Success
                }
                else {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[FAILED:0x$($ExitCode1.ToString('X'))]"
                }
            }

# --- STEP 6: DISM ComponentCleanup ---
            $S73 = "[09/11] DISM ComponentCleanup"
            Write-StepUpdate $S73 -CustomInfo "[Press ESC to Skip]"
            $Row73 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($DryRun -or $PendingReboot -or $HasPendingRename -or $ServicingActive) {
                Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[SKIPPED]"
            }
            else {
                # BITS is left running so the RMM and LiveConnect stay connected.
                Stop-Service wuauserv -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Stop-Service TrustedInstaller -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Service TrustedInstaller -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Sleep -Seconds 3

                Clear-InputBuffer

                $DismTmp2 = [System.IO.Path]::GetTempFileName()

                $psi2 = New-Object System.Diagnostics.ProcessStartInfo
                $psi2.FileName               = "cmd.exe"
                $psi2.Arguments              = "/c dism.exe /Online /Cleanup-Image /StartComponentCleanup /NoRestart > `"$DismTmp2`" 2>&1"
                $psi2.UseShellExecute        = $false
                $psi2.CreateNoWindow         = $true
                $psi2.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Proc2 = [System.Diagnostics.Process]::Start($psi2)

                $Skipped2 = $false
                $DismSpinIdx2 = 0
                $DismTimer2 = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc2.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                try { $Proc2.Kill() } catch {}
                                Stop-DismTree
                                $Skipped2 = $true
                                break
                            }
                        }
                    } catch {}
                    try {
                        [Console]::SetCursorPosition(0, $Row73)
                        $EscHint = "[ESC to skip]"
                        $Width = $script:Width
                        $Left = "$S73 $($DismSpin[$DismSpinIdx2 % 4]) $($DismTimer2.Elapsed.ToString('mm\:ss'))"
                        $Spaces = [Math]::Max(1, $Width - $Left.Length - $EscHint.Length)
                        [Console]::ForegroundColor = [ConsoleColor]::Cyan
                        [Console]::Write($Left + (' ' * $Spaces))
                        [Console]::ForegroundColor = [ConsoleColor]::DarkGray
                        [Console]::Write($EscHint)
                        [Console]::ResetColor()
                    } catch {}
                    $DismSpinIdx2++
                    Start-Sleep -Milliseconds 250
                }
                $DismTimer2.Stop()

                if (-not $Skipped2) {
                    $Proc2.WaitForExit()
                    Stop-TiWorker
                }
                $ExitCode2 = $Proc2.ExitCode
                try { $Proc2.Dispose() } catch {}
                Remove-Item $DismTmp2 -Force -ErrorAction SilentlyContinue

                Start-Service wuauserv -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

                if ($Skipped2) {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[SKIPPED]"
                }
                elseif ($ExitCode2 -in @(0, 3010)) {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -Success
                }
                else {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[FAILED:0x$($ExitCode2.ToString('X'))]"
                }
            }

# --- STEP 7: SFC /scannow ---
            $S74 = "[10/11] SFC /scannow"
            Write-StepUpdate $S74 -CustomInfo "[Press ESC to Skip]"
            $Row74 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($DryRun -or $PendingReboot -or $HasPendingRename -or $ServicingActive) {
                Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[SKIPPED]"
            }
            else {
                Clear-InputBuffer
                $SfcTmp = [System.IO.Path]::GetTempFileName()

                # sfc.exe writes Unicode; capture via cmd redirection for reliable exit code
                $psi3 = New-Object System.Diagnostics.ProcessStartInfo
                $psi3.FileName               = "cmd.exe"
                $psi3.Arguments              = "/c sfc.exe /scannow > `"$SfcTmp`" 2>&1"
                $psi3.UseShellExecute        = $false
                $psi3.CreateNoWindow         = $true
                $psi3.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Proc3 = [System.Diagnostics.Process]::Start($psi3)

                $Skipped3 = $false
                $SfcSpinIdx = 0
                $SfcTimer = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc3.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                try { $Proc3.Kill() } catch {}
                                Get-Process -Name "sfc" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
                                $Skipped3 = $true
                                break
                            }
                        }
                    } catch {}
                    try {
                        [Console]::SetCursorPosition(0, $Row74)
                        $EscHint = "[ESC to skip]"
                        $Width = $script:Width
                        $Left = "$S74 $($DismSpin[$SfcSpinIdx % 4]) $($SfcTimer.Elapsed.ToString('mm\:ss'))"
                        $Spaces = [Math]::Max(1, $Width - $Left.Length - $EscHint.Length)
                        [Console]::ForegroundColor = [ConsoleColor]::Cyan
                        [Console]::Write($Left + (' ' * $Spaces))
                        [Console]::ForegroundColor = [ConsoleColor]::DarkGray
                        [Console]::Write($EscHint)
                        [Console]::ResetColor()
                    } catch {}
                    $SfcSpinIdx++
                    Start-Sleep -Milliseconds 250
                }
                $SfcTimer.Stop()

                if (-not $Skipped3) {
                    $Proc3.WaitForExit()
                }
                $ExitCode3 = $Proc3.ExitCode
                try { $Proc3.Dispose() } catch {}
                Remove-Item $SfcTmp -Force -ErrorAction SilentlyContinue

                if ($Skipped3) {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[SKIPPED]"
                }
                elseif ($ExitCode3 -in @(0, 1)) {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -Success
                }
                elseif ($ExitCode3 -eq 2) {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[WARNING]"
                }
                else {
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[FAILED:0x$($ExitCode3.ToString('X'))]"
                }
            }
}
# Component cleanup can free real space, so report it instead of folding it in silently.
if (-not $DryRun) {
    $CurrentSpace = Get-FreeBytes
    $RepairSaved = [int64]($CurrentSpace - $LastRegionSpace)
    if ($RepairSaved -ge 1MB) { Write-Host "        Component store cleanup freed $(Format-Size $RepairSaved)." -ForegroundColor Gray }
    $LastRegionSpace = $CurrentSpace
}
#endregion
#region 8. Final Optimization
# ============================================================================
Write-StepUpdate "[11/11] Disabling Hibernation, Flushing DNS & Running SSD TRIM"
if ($DryRun) {
    $HibBytes = [int64]0
    if (Test-Path "C:\hiberfil.sys") {
        try { $HibBytes = [int64](Get-Item "C:\hiberfil.sys" -Force -ErrorAction SilentlyContinue).Length } catch { $HibBytes = 0 }
    }
    Write-DryEstimate $HibBytes
} else {
    & ipconfig.exe /flushdns | Out-Null
    & powercfg.exe /h off | Out-Null
    try { Optimize-Volume -DriveLetter C -ReTrim -ErrorAction SilentlyContinue | Out-Null } catch { }
    Complete-Step
}
#endregion

#region Final Summary
# ============================================================================
# Before and after come straight from the volume, and "Space Recovered" is simply
# After minus Before, so the three numbers always reconcile on the screenshot.
$FinalFree  = Get-FreeBytes
$StartFree  = [int64]$StartSpace
$TotalGBStr = "{0:N0} GB" -f ($TotalSize / 1GB)
function Get-FreeColor { param([double]$Pct) if ($Pct -lt 10) { "Red" } elseif ($Pct -lt 20) { "DarkYellow" } else { "Green" } }
function Write-FreeLine {
    param([string]$Label, [int64]$Free)
    $Pct = [Math]::Round(($Free / $TotalSize) * 100, 1)
    Write-Host $Label -NoNewline -ForegroundColor $InfoCol
    Write-Host ("{0} free of {1} ({2}% free)" -f (Format-Size $Free), $TotalGBStr, $Pct) -ForegroundColor (Get-FreeColor $Pct)
}

Write-HLine -Style dashed
if ($DryRun) {
    $ProjFree = [int64]($FinalFree + $script:EstYieldBytes)
    Write-Host "Est. Recoverable    : " -NoNewline -ForegroundColor $InfoCol
    Write-Host (Format-Size $script:EstYieldBytes) -ForegroundColor Magenta
    Write-FreeLine "Projected After     : " $ProjFree
} else {
    $Recovered = [int64]($FinalFree - $StartFree)
    Write-FreeLine "Free After          : " $FinalFree
    Write-Host "Space Recovered     : " -NoNewline -ForegroundColor $InfoCol
    if ($Recovered -gt 0) { Write-Host (Format-Size $Recovered) -ForegroundColor Yellow }
    else { Write-Host "0 MB (disk usage grew during the run)" -ForegroundColor DarkYellow }
}
foreach ($N in $script:Notes) { Write-Host "        [!] $N" -ForegroundColor DarkYellow }
# Footer
$_sfx   = "█"
$_ffillW = $script:Width - $_artW - 1 - $_sfx.Length
$_footer = if ($DryRun) { "  DRY RUN COMPLETE" } else { "  MAINTENANCE COMPLETE" }
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
#endregion