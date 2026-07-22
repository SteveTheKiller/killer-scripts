<#
.SYNOPSIS
    Advanced Maintenance, Optimization, and Repair Tool (AMORT) v16.0
    Developed by Steve the Killer | Updated: 2026-07-22
.DESCRIPTION
    Automated Windows 10/11 disk-space reclamation and integrity repair for
    MSP field and remote use. Purges Dell SupportAssist snapshots, browser,
    Office and GPU caches, the Recycle Bin, Delivery Optimization, the installer
    cache, the search index, and aged Windows.old; resets the Windows Update
    database; runs DISM and SFC repair; removes the hibernation file; and
    performs SSD TRIM while reporting disk space recovered at each stage.

    v16.0 scope change: privacy/telemetry hardening (old Region 1), browser
    hardening/uBlock (old Region 2), and OEM/software debloat (old Region 3)
    were removed. Those behaviors now belong to SHADE and DEBLOAT. AMORT is
    cleanup + repair only, safe to run on live, managed endpoints.
.PARAMETER DryRun
    Read-only estimate mode. Makes no changes: every destructive step is skipped
    and each target is sized instead, reporting estimated reclaim per category
    plus a projected free-space total. Use before committing on a disk alert.
#>
param([switch]$DryRun)
$_fver   = "| v16.0"
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
            } elseif ($CustomInfo.StartsWith("(Saved:")) {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Red
            } elseif ($CustomInfo.StartsWith("(Est:")) {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Magenta
            } else {
                Write-Host " $CustomInfo" -NoNewline -ForegroundColor Gray
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
        Write-StepUpdate -Success -CustomInfo "(Est: $s)"
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
if ($DryRun) {
    Write-Host "      Mode: DRY RUN - estimate only, no changes will be made" -ForegroundColor Magenta
}
if ($IsVM) {
    Write-Host "      Mode: Virtual Machine" -ForegroundColor Yellow 
}
#endregion

#region 1. Snapshot & Storage Purge
# ============================================================================
Write-StepUpdate "[01/08] Purging Snapshots, Installer Cache & Search Index..."
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
            # Stop any Dell SupportAssist / remediation services holding the snapshots open
            $SASvcs = Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "SupportAssist*" -or $_.DisplayName -like "*SupportAssist*" }
            foreach ($SASvc in $SASvcs) {
                try { Stop-Service $SASvc.Name -Force -ErrorAction Stop -WarningAction SilentlyContinue } catch { }
            }
            foreach ($Sub in @("Snapshots", "Backup")) {
                $SAPath = Join-Path $SARoot $Sub
                if (Test-Path $SAPath) {
                    # Snapshot files are hidden/system/protected: strip attributes, then delete contents (keep the folder)
                    Start-Process "cmd.exe" -ArgumentList "/c attrib -h -s -r `"$SAPath\*`" /S /D & del /s /f /q `"$SAPath\*`"" -WindowStyle Hidden -Wait
                }
            }
            # Restart the services we stopped
            foreach ($SASvc in $SASvcs) {
                Start-Service $SASvc.Name -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
            }
        }
    }
}
# Delivery Optimization
$DOCache = "C:\Windows\ServiceProfiles\NetworkService\AppData\Local\Microsoft\Windows\DeliveryOptimization\Cache"
if (Test-Path $DOCache) {
    if ($DryRun) {
        $RegionEst += Get-PathSize $DOCache
    } else {
        Remove-Item "$DOCache\*" -Recurse -Force -ErrorAction SilentlyContinue | Out-Null
    }
}

# --- SEARCH INDEX RESET ---
$SearchPath = "C:\ProgramData\Microsoft\Search\Data\Applications\Windows"
if ($DryRun) {
    if (Test-Path $SearchPath) { $RegionEst += Get-PathSize $SearchPath }
} else {
    $SvcName = "WSearch"
    $Svc = Get-Service $SvcName -ErrorAction SilentlyContinue
    if ($Svc -and $Svc.Status -ne 'Stopped') {
        Stop-Service $SvcName -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        $RetryCount = 0
        while ((Get-Service $SvcName).Status -ne 'Stopped' -and $RetryCount -lt 10) {
            $dots = '.' * (($RetryCount % 3) + 1)
            $savedRow = [Console]::CursorTop
            [Console]::SetCursorPosition(0, $script:StepRow)
            Write-Host ("$($script:LastStepMessage) [Stopping WSearch$dots]").PadRight($script:Width) -NoNewline -ForegroundColor Cyan
            [Console]::SetCursorPosition(0, $savedRow)
            Start-Sleep -Seconds 2
            $RetryCount++
        }
        # Restore clean step line (strip the "[Stopping WSearch...]" suffix)
        [Console]::SetCursorPosition(0, $script:StepRow)
        Write-Host $script:LastStepMessage.PadRight($script:Width) -NoNewline -ForegroundColor Cyan
        [Console]::SetCursorPosition(0, $script:StepRow + 1)
        if ((Get-Service $SvcName).Status -ne 'Stopped') {
            Get-Process "SearchIndexer" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        }
    }

    if (Test-Path $SearchPath) { 
        # FIX: Using cmd /c del bypasses the PowerShell ArgumentException 
        # if files disappear during the recursive delete.
        Start-Process "cmd.exe" -ArgumentList "/c del /s /f /q `"$SearchPath\*`"" -WindowStyle Hidden -Wait
    }
    Start-Service $SvcName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
}

# Windows Installer Cache (orphaned packages older than 90 days)
$InstallerPath = "C:\Windows\Installer"
if (Test-Path $InstallerPath) {
    $InstalledProducts = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue | Where-Object { $_.LocalPackage } | Select-Object -ExpandProperty LocalPackage
    $OrphanPkgs = Get-ChildItem $InstallerPath -Filter "*.ms[ip]" -Force -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notin $InstalledProducts -and $_.LastWriteTime -lt (Get-Date).AddDays(-90) }
    if ($DryRun) {
        $OrphanSum = ($OrphanPkgs | Measure-Object -Property Length -Sum).Sum
        if ($OrphanSum) { $RegionEst += [int64]$OrphanSum }
    } else {
        $OrphanPkgs | Remove-Item -Force -ErrorAction SilentlyContinue
    }
}

# --- Windows.old cleanup (age-gated) ---
# Only remove once the feature-update rollback window has closed, so we never kill
# a live rollback path. Read the real window via DISM; fall back to 14 days.
if (Test-Path "C:\Windows.old") {
    $OldAgeDays = ((Get-Date) - (Get-Item "C:\Windows.old").LastWriteTime).TotalDays
    $UninstallWindow = 14
    try {
        $DismWin = & DISM.exe /Online /Get-OSUninstallWindow 2>$null
        $WinLine = $DismWin | Select-String -Pattern 'Uninstall Window\D+(\d+)'
        if ($WinLine) { $UninstallWindow = [int]$WinLine.Matches[0].Groups[1].Value }
    } catch { }

    if ($OldAgeDays -gt $UninstallWindow) {
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
                & icacls "C:\Windows.old" /grant Administrators:F /T /C /Q 2>$null | Out-Null
                Start-Process "cmd.exe" -ArgumentList "/c rd /s /q `"C:\Windows.old`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
            }
        }
    }
}

# --- Region 1: finalize and accumulate ---
if ($DryRun) {
    Write-DryEstimate $RegionEst
} else {
    $CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $CurrentSpace = [int64]$CurrentDrive.FreeSpace
    if ($null -eq $LastRegionSpace) { $LastRegionSpace = $CurrentSpace }
    $RegionSavedBytes = [int64]($CurrentSpace - $LastRegionSpace)
    if ($RegionSavedBytes -gt 0) {
        $SavedStr = if ($RegionSavedBytes -ge 1GB) { "{0:N2} GB" -f ($RegionSavedBytes / 1GB) } else { "{0:N2} MB" -f ($RegionSavedBytes / 1MB) }
        Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
    } else {
        Write-StepUpdate -Success -CustomInfo "Saved: 0 MB"
        $RegionSavedBytes = 0
    }
    if (-not $script:TotalYieldBytes) { $script:TotalYieldBytes = 0 }
    $script:TotalYieldBytes = [int64]$script:TotalYieldBytes
    if ($RegionSavedBytes -gt 0) { $script:TotalYieldBytes += $RegionSavedBytes }
    if (-not $script:RegionHistory) { $script:RegionHistory = @() }
    $script:RegionHistory += [pscustomobject]@{ Region = 'Region1'; Bytes = $RegionSavedBytes; Time = (Get-Date) }
    $LastRegionSpace = $CurrentSpace
}
#endregion

#region 2. Deep Cache Purge
# ============================================================================
Write-StepUpdate "[02/08] Purging Browser, Office, and GPU Caches..."
$RegionEst = [int64]0
$GlobalCaches = @("C:\Windows\Temp\*", "C:\Windows\Prefetch\*", "C:\Windows\SystemTemp\*")
foreach ($P in $GlobalCaches) {
    if (Test-Path $P) {
        if ($DryRun) { $RegionEst += Get-PathSize $P }
        else { Remove-Item $P -Recurse -Force -ErrorAction SilentlyContinue | Out-Null }
    }
}
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
    $OutlookPath = "$UP\AppData\Local\Microsoft\Outlook"
    if ($DryRun) {
        foreach ($OP in $OffPaths) { $RegionEst += Get-PathSize $OP }
        if (Test-Path $OutlookPath) {
            $NstSum = (Get-ChildItem $OutlookPath -Filter "*.nst" -Force -ErrorAction SilentlyContinue | Measure-Object -Property Length -Sum).Sum
            if ($NstSum) { $RegionEst += [int64]$NstSum }
        }
        foreach ($SP in $ShaderPaths) { $RegionEst += Get-PathSize $SP }
        foreach ($T in $TargetDirs) { $RegionEst += Get-PathSize $T }
    } else {
        # Office & Outlook (NST) Cleanup
        foreach ($OP in $OffPaths) { if (Test-Path $OP) { Remove-Item $OP -Recurse -Force -ErrorAction SilentlyContinue } }
        # Outlook NST (Search Index) files
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
}
# --- Region 2: finalize ---
if ($DryRun) {
    Write-DryEstimate $RegionEst
} else {
    $CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
    if ($RegionSaved -gt 0) {
        $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
        Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
    } else {
        Write-StepUpdate -Success
    }
    if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
    $LastRegionSpace = $CurrentDrive.FreeSpace
}
#endregion

#region 3. Recycle Bin Purge
# ============================================================================
Write-StepUpdate "[03/08] Emptying Recycle Bin (all users)..."
# rd on the per-volume store clears every user's Recycle Bin. Clear-RecycleBin only
# empties the current identity's bin, which under SYSTEM/LiveConnect misses the
# logged-in user's deleted files (often the biggest quick win on a disk alert).
# Windows recreates the folder automatically.
if ($DryRun) {
    $RegionEst = Get-PathSize "C:\`$Recycle.Bin"
    Write-DryEstimate $RegionEst
} else {
    if (Test-Path "C:\`$Recycle.Bin") {
        Start-Process "cmd.exe" -ArgumentList "/c rd /s /q `"C:\`$Recycle.Bin`"" -WindowStyle Hidden -Wait -ErrorAction SilentlyContinue
    }
    $CurrentDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
    $RegionSaved = $CurrentDrive.FreeSpace - $LastRegionSpace
    if ($RegionSaved -gt 0) {
        $SavedStr = if ($RegionSaved -gt 1GB) { "$([math]::Round($RegionSaved / 1GB, 2)) GB" } else { "$([math]::Round($RegionSaved / 1MB, 2)) MB" }
        Write-StepUpdate -Success -CustomInfo "(Saved: $SavedStr)"
    } else {
        Write-StepUpdate -Success
    }
    if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
    $LastRegionSpace = $CurrentDrive.FreeSpace
}
#endregion

#region 4. Windows Update Database Reset
# ============================================================================
Write-StepUpdate "[04/08] Resetting Windows Update Database..."
# CHECK: If a reboot is pending, SoftwareDistribution is likely locked. 
# Skip to prevent the script from hanging.
$PendingReboot = Test-Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"

if ($PendingReboot) {
    Write-StepUpdate -CustomInfo "[SKIPPED]"
    Write-Host "        (Reboot pending - SoftwareDistribution locked)" -ForegroundColor DarkYellow
} elseif ($DryRun) {
    $RegionEst = Get-PathSize "C:\Windows\SoftwareDistribution"
    Write-DryEstimate $RegionEst
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
    if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
    $LastRegionSpace = $CurrentDrive.FreeSpace
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
    Write-Host "        [!] Reboot pending - DISM and SFC repair steps will be skipped." -ForegroundColor DarkYellow
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
            $S72 = "[05/08] DISM RestoreHealth..."
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
            $S73 = "[06/08] DISM ComponentCleanup..."
            Write-StepUpdate $S73 -CustomInfo "[Press ESC to Skip]"
            $Row73 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($DryRun -or $PendingReboot -or $HasPendingRename -or $ServicingActive) {
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

                Start-Service bits -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
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
            $S74 = "[07/08] SFC /scannow..."
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
if (-not $DryRun) {
    $CurrentSpace = (Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'").FreeSpace
    $RegionSaved = $CurrentSpace - $LastRegionSpace
    if ($RegionSaved -gt 0) { $TotalYieldBytes += [int64]$RegionSaved }
    $LastRegionSpace = $CurrentSpace
}
#endregion

#region 6. Final Optimization
# ============================================================================
Write-StepUpdate "[08/08] Finalizing Network, Hibernation & SSD TRIM..."
if ($DryRun) {
    # powercfg /h off reclaims hiberfil.sys; estimate its current size
    $HibBytes = [int64]0
    if (Test-Path "C:\hiberfil.sys") {
        try { $HibBytes = [int64](Get-Item "C:\hiberfil.sys" -Force -ErrorAction SilentlyContinue).Length } catch { $HibBytes = 0 }
    }
    Write-DryEstimate $HibBytes
} else {
    & ipconfig.exe /flushdns | Out-Null
    & powercfg.exe /h off | Out-Null
    try { Optimize-Volume -DriveLetter C -ReTrim -ErrorAction SilentlyContinue | Out-Null } catch { }
    Write-StepUpdate -Success
}

# --- FINAL SUMMARY ---
# Ensure disk info and total size are valid
$FinalDrive = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='C:'"
$TotalSize = [double]$TotalSize
if ($TotalSize -le 0) { throw "TotalSize is zero or undefined. Aborting final summary." }

if (-not $script:TotalYieldBytes) { $script:TotalYieldBytes = 0 }
$script:TotalYieldBytes = [int64]$script:TotalYieldBytes
if (-not $script:EstYieldBytes) { $script:EstYieldBytes = 0 }
$script:EstYieldBytes = [int64]$script:EstYieldBytes

$FinalFree = [int64]$FinalDrive.FreeSpace
$FinalUsedPct = [Math]::Round(((($TotalSize - $FinalFree) / $TotalSize) * 100), 2)
$FinalUsedGB = [Math]::Round(($TotalSize - $FinalFree) / 1GB, 2)
$FinalTotalGB = [Math]::Round($TotalSize / 1GB, 0)
$FinalColor = if ($FinalUsedPct -ge 90) { "Red" } elseif ($FinalUsedPct -ge 80) { "DarkYellow" } else { "Green" }

Write-HLine -Style dashed
if ($DryRun) {
    # Projected free space if the estimated reclaim were applied
    $ProjFree = [int64]($FinalFree + $script:EstYieldBytes)
    $ProjUsedPct = [Math]::Round(((($TotalSize - $ProjFree) / $TotalSize) * 100), 2)
    $ProjColor = if ($ProjUsedPct -ge 90) { "Red" } elseif ($ProjUsedPct -ge 80) { "DarkYellow" } else { "Green" }
    if ($script:EstYieldBytes -ge 1GB) { $EstStr = "{0:N2} GB" -f ($script:EstYieldBytes / 1GB) } else { $EstStr = "{0:N2} MB" -f ($script:EstYieldBytes / 1MB) }

    Write-Host "Current Disk Usage  : " -NoNewline -ForegroundColor $InfoCol
    Write-Host "$FinalUsedGB GB used of $FinalTotalGB GB ($FinalUsedPct%)" -ForegroundColor $FinalColor
    Write-Host "Est. Recoverable    : " -NoNewline -ForegroundColor $InfoCol
    Write-Host "$EstStr" -ForegroundColor Magenta
    Write-Host "Projected After Run : " -NoNewline -ForegroundColor $InfoCol
    Write-Host "$([Math]::Round(($TotalSize - $ProjFree) / 1GB, 2)) GB used of $FinalTotalGB GB ($ProjUsedPct%)" -ForegroundColor $ProjColor
} else {
    if ($script:TotalYieldBytes -ge 1GB) { $TotalStr = "{0:N2} GB" -f ($script:TotalYieldBytes / 1GB) } else { $TotalStr = "{0:N2} MB" -f ($script:TotalYieldBytes / 1MB) }
    Write-Host "Final Disk Usage    : " -NoNewline -ForegroundColor $InfoCol
    Write-Host "$FinalUsedGB GB used of $FinalTotalGB GB ($FinalUsedPct%)" -ForegroundColor $FinalColor
    Write-Host "Space Recovered     : " -NoNewline -ForegroundColor $InfoCol
    Write-Host "$TotalStr" -ForegroundColor Yellow
}
# Footer
$_sfx   = "█"
$_ffillW = $script:Width - $_artW - 1 - $_sfx.Length
$_footer = if ($DryRun) { "  DRY RUN COMPLETE" } else { "  MAINTENANCE COMPLETE" }
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
#endregion