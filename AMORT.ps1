<#
.SYNOPSIS
    Advanced Maintenance, Optimization, and Repair Tool (AMORT) v18.0
    Developed by Steve the Killer | Updated: 2026-11-18
.DESCRIPTION
    One-step Windows 10/11 disk-space reclamation and integrity repair for MSP
    field and remote use. Built to be safe on personal and business machines
    while a user is signed in and working: no user documents, mail stores,
    sync caches, or search indexes are touched, and anything that could still
    be needed is age-gated or retained.

    Removes Dell SupportAssist snapshots, feature-upgrade leftovers, OEM driver
    extracts, Delivery Optimization and Windows Update download caches, browser/Teams/GPU caches, and aged temp files, crash dumps, error
    reports, and Recycle Bin items. Windows.old and the previous-version rollback
    data are kept unless -RemoveRollback is given. Trims restore points to the newest few,
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
.PARAMETER RemoveRollback
    Opt in to removing Windows.old and the "Go back to a previous version of Windows"
    option. Off by default. Refused while Windows servicing is active, a reboot is
    pending or a Windows upgrade is running. A console that cannot pass arguments can
    set $RemoveRollback = $true before running the script.
#>
param(
    [switch]$DryRun,
    [ValidateRange(1, 365)][int]$RetainDays = 14,
    [ValidateRange(1, 64)][int]$KeepRestorePoints = 2,
    [switch]$RemoveRollback
)

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# Console and system boundaries. Each one is a small function so a missing console (RMM, SYSTEM,
# redirected output) never throws and a test can replace it.
function Test-AmortElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
function Test-AmortConsole {
    try { $null = [Console]::CursorTop; return $true } catch { return $false }
}
function Get-AmortCursorTop {
    try { return [int][Console]::CursorTop } catch { return -1 }
}
function Get-AmortCursorLeft {
    try { return [int][Console]::CursorLeft } catch { return 0 }
}
function Set-AmortCursor {
    param([int]$Left, [int]$Top)
    try { if ($Top -ge 0) { [Console]::SetCursorPosition($Left, $Top) } } catch { }
}
function New-AmortTempFile {
    [System.IO.Path]::GetTempFileName()
}
function Start-AmortProcess {
    param([System.Diagnostics.ProcessStartInfo]$StartInfo)
    [System.Diagnostics.Process]::Start($StartInfo)
}
# True when Windows is waiting for a reboot. Reads the 64-bit registry view, because a 32-bit host
# would otherwise be redirected to WOW6432Node. Covers servicing and Windows Update.
function Test-AmortRebootPending {
    $keys = @(
        'SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    )
    $base = $null
    try {
        $base = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        foreach ($k in $keys) {
            $sub = $base.OpenSubKey($k)
            if ($sub) { $sub.Close(); return $true }
        }
    } catch { } finally { if ($base) { $base.Close() } }
    return $false
}
# A 32-bit host on 64-bit Windows (some RMM agents launch one) sees a redirected System32.
# Sysnative reaches the real one, so DISM, vssadmin and cmd act on the running system.
function Test-AmortWow64 {
    [Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess
}
function Get-AmortNativePath {
    param([string]$Name)
    if (Test-AmortWow64) { return (Join-Path $env:windir ("Sysnative" + [IO.Path]::DirectorySeparatorChar + $Name)) }
    $Name
}
# Runs a native tool and returns its exit code and output, so callers can check the result.
function Invoke-AmortTool {
    param([string]$FilePath, [string[]]$ArgumentList = @())
    if ($FilePath -notmatch '[\\/]') { $FilePath = Get-AmortNativePath $FilePath }
    $out = & $FilePath @ArgumentList 2>&1
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = (@($out) -join [Environment]::NewLine) }
}

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
$script:Width = 90

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
        $script:StepRow = Get-AmortCursorTop
        & $writeMsg $printMsg $false
        Write-Host "" # Move cursor to next line so WARNINGS have a place to go
    }
    elseif ($isDone) {
        if (-not (Test-AmortConsole)) {
            # No console (RMM pane, SYSTEM, redirected output): one plain line per result.
            $tag = if ($Success) { if ($script:DryRun) { "[EST]" } else { "[SUCCESS]" } } else { "" }
            Write-Host ("        " + (@($CustomInfo, $tag) | Where-Object { $_ }) -join " ")
            return
        }
        # COMPLETING A STEP: Jump back to the saved row to overwrite the Cyan text
        $currentPos = Get-AmortCursorTop
        Set-AmortCursor 0 $script:StepRow

        # Clear the original Cyan line
        Write-Host (" " * $script:Width) -NoNewline
        Set-AmortCursor 0 $script:StepRow

        # Reprint the line in the "Done" (Gray/White) style
        & $writeMsg $printMsg $true

        # Add Custom Info (Saved MB/GB)
        if ($CustomInfo) {
            if ($CustomInfo -eq "[SKIPPED]") {
                $tag = "[SKIPPED]"
                $currentCol = Get-AmortCursorLeft
                $targetCol  = $script:Width - $tag.Length
                if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
                Write-Host $tag -ForegroundColor Yellow
            } else {
                # Right-align the info so it ends one space left of the 9-character status tag column
                $infoCol = if ($CustomInfo -in @("Saved: 0 MB", "Est: 0 MB")) { "Gray" } elseif ($CustomInfo.StartsWith("Saved:")) { "Red" } elseif ($CustomInfo.StartsWith("Est:")) { "Magenta" } else { "Gray" }
                $currentCol = Get-AmortCursorLeft
                $targetCol  = $script:Width - 10 - $CustomInfo.Length
                $pad = [Math]::Max(1, $targetCol - $currentCol)
                Write-Host ((" " * $pad) + $CustomInfo) -NoNewline -ForegroundColor $infoCol
            }
        }

        # Final Success Tag (right-aligned to console width)
        if ($Success) {
            $tag = if ($script:DryRun) { "[EST]" } else { "[SUCCESS]" }
            $currentCol = Get-AmortCursorLeft
            $targetCol  = $script:Width - $tag.Length
            if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
            Write-Host $tag -ForegroundColor $(if ($script:DryRun) { "Magenta" } else { "Green" })
        }

        # Return the cursor to where it was (below any warnings that appeared)
        if ($currentPos -gt $script:StepRow) {
            Set-AmortCursor 0 $currentPos
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
    # .NET walk instead of the pipeline, which crawled on a Temp folder with hundreds of
    # thousands of entries. Links are never followed; unreadable folders and locked files are skipped.
    $Cutoff = (Get-Date).AddDays(-$Days)
    $Sum = [int64]0
    $Reparse = [IO.FileAttributes]::ReparsePoint
    foreach ($P in $Path) {
        if ([string]::IsNullOrEmpty($P)) { continue }
        if ([IO.File]::Exists($P)) {
            $F = New-Object IO.FileInfo $P
            if ($F.LastWriteTime -lt $Cutoff) {
                if ($script:DryRun) { $Sum += $F.Length } else { try { $F.Delete() } catch { } }
            }
            continue
        }
        if (-not [IO.Directory]::Exists($P)) { continue }
        $Stack = New-Object System.Collections.Generic.Stack[string]
        $Stack.Push($P)
        $SubDirs = New-Object System.Collections.Generic.List[string]
        while ($Stack.Count -gt 0) {
            $Dir = New-Object IO.DirectoryInfo $Stack.Pop()
            try { $Entries = $Dir.GetFileSystemInfos() } catch { continue }
            foreach ($E in $Entries) {
                if ($E.Attributes -band [IO.FileAttributes]::Directory) {
                    if (-not ($E.Attributes -band $Reparse)) {
                        $Stack.Push($E.FullName)
                        if ($E.LastWriteTime -lt $Cutoff) { $SubDirs.Add($E.FullName) }
                    }
                    continue
                }
                if ($E.LastWriteTime -ge $Cutoff) { continue }
                if ($script:DryRun) { $Sum += $E.Length; continue }
                try {
                    if ($E.Attributes -band [IO.FileAttributes]::ReadOnly) { $E.Attributes = [IO.FileAttributes]::Normal }
                    $E.Delete()
                } catch { }
            }
        }
        # Old folders left empty are removed deepest first. A non-recursive delete fails on
        # anything still holding files, so only truly empty folders go, never the root itself.
        if (-not $script:DryRun) {
            for ($i = $SubDirs.Count - 1; $i -ge 0; $i--) {
                try { [IO.Directory]::Delete($SubDirs[$i], $false) } catch { }
            }
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

# Helper: clear console lines from $StartRow to current row, then reprint a step result
function Clear-AndReprintStep {
    param([int]$StartRow, [string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($StartRow -ge 0 -and (Test-AmortConsole)) {
        try {
            $endRow = Get-AmortCursorTop
            $width  = $script:Width
            for ($r = $StartRow; $r -le $endRow; $r++) {
                [Console]::SetCursorPosition(0, $r)
                [Console]::Write(' ' * $width)
            }
            [Console]::SetCursorPosition(0, $StartRow)
        } catch {}
    }
    if ($Success) { Write-StepUpdate $Message -Success }
    elseif ($CustomInfo -eq "[SKIPPED]") { Write-StepUpdate $Message -CustomInfo "[SKIPPED]" }
    elseif ($CustomInfo -match '^\[FAILED' -or $CustomInfo -eq "[WARNING]") {
        $color = if ($CustomInfo -eq "[WARNING]") { "Yellow" } else { "Red" }
        if (-not (Test-AmortConsole)) {
            Write-Host ("        " + $CustomInfo) -ForegroundColor $color
            return
        }
        # Print step label in Gray, description in White, result in color
        if ($Message -match '^(\[[\d./]+\])(\s+.+)$') {
            Write-Host $Matches[1] -NoNewline -ForegroundColor DarkGray
            Write-Host $Matches[2] -NoNewline -ForegroundColor White
        } else { Write-Host $Message -NoNewline -ForegroundColor White }
        $currentCol = Get-AmortCursorLeft
        $targetCol  = $script:Width - $CustomInfo.Length
        if ($targetCol -gt $currentCol) { Write-Host (" " * ($targetCol - $currentCol)) -NoNewline }
        Write-Host $CustomInfo -ForegroundColor $color
    }
    elseif ($CustomInfo) { Write-StepUpdate $Message -CustomInfo $CustomInfo }
}
# Helper: flush buffered console keypresses using raw .NET Console API (bypasses PSReadLine)
function Clear-InputBuffer { try { while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null } } catch {} }
# Ids of every process started, directly or not, by the given parent. Only processes this run
# launched are ever stopped. TiWorker and other servicing belong to Windows and are never named.
function Get-AmortDescendantId {
    param([int]$ParentId)
    $all = @(Get-CimInstance -ClassName Win32_Process -ErrorAction SilentlyContinue)
    $found = New-Object System.Collections.ArrayList
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($ParentId)
    while ($queue.Count -gt 0) {
        $cur = $queue.Dequeue()
        foreach ($p in $all) {
            if ([int]$p.ParentProcessId -eq $cur -and [int]$p.ProcessId -ne $ParentId -and -not $found.Contains([int]$p.ProcessId)) {
                [void]$found.Add([int]$p.ProcessId)
                $queue.Enqueue([int]$p.ProcessId)
            }
        }
    }
    @($found)
}
# Stops the process this run started and the tree under it, children first, by id.
function Stop-AmortOwnedTree {
    param([int]$ProcessId)
    $ids = @(Get-AmortDescendantId -ParentId $ProcessId)
    [array]::Reverse($ids)
    foreach ($id in ($ids + $ProcessId)) { Stop-Process -Id $id -Force -ErrorAction SilentlyContinue }
}
function Get-FreeColor { param([double]$Pct) if ($Pct -lt 10) { "Red" } elseif ($Pct -lt 20) { "DarkYellow" } else { "Green" } }
function Write-FreeLine {
    param([string]$Label, [int64]$Free)
    $Pct = [Math]::Round(($Free / $TotalSize) * 100, 1)
    Write-Host $Label -NoNewline -ForegroundColor $InfoCol
    Write-Host ("{0} free of {1} ({2}% free)" -f (Format-Size $Free), $TotalGBStr, $Pct) -ForegroundColor (Get-FreeColor $Pct)
}

# System Restore keeps its points as persistent, client-accessible shadow copies from the Microsoft
# software provider. Backup, hardware-provider, mounted and other-volume copies belong to something
# else, so they are never counted for retention and never deleted.
$script:AmortSystemShadowProvider = '{b5946137-7b9f-4925-af80-51abd60b20d5}'
function Get-AmortShadowInventory {
    param([string]$VolumeDeviceId)
    $all = @(Get-CimInstance -ClassName Win32_ShadowCopy -ErrorAction SilentlyContinue |
        Where-Object { $_.VolumeName -eq $VolumeDeviceId })
    $eligible = @($all | Where-Object {
        $_.ID -and $_.InstallDate -and
        $_.ClientAccessible -eq $true -and $_.Persistent -eq $true -and
        $_.ProviderID -eq $script:AmortSystemShadowProvider -and
        -not $_.ExposedName
    } | Sort-Object InstallDate -Descending)
    [pscustomobject]@{ Eligible = $eligible; SkippedCount = ($all.Count - $eligible.Count) }
}
# Deletes one shadow copy of C: and confirms it is gone. Returns $true only when it really is.
function Remove-AmortShadowCopy {
    param([string]$Id, [string]$VolumeDeviceId)
    $r = Invoke-AmortTool "vssadmin.exe" @("delete", "shadows", "/for=C:", "/shadow=$Id", "/quiet")
    if ($r.ExitCode -ne 0) { return $false }
    $left = @(Get-AmortShadowInventory -VolumeDeviceId $VolumeDeviceId).Eligible | Where-Object { $_.ID -eq $Id }
    return (@($left).Count -eq 0)
}

# Removing the rollback window is opt-in. The switch is the contract; a global $RemoveRollback of
# exactly $true also counts, for consoles that cannot pass arguments.
function Test-AmortRemoveRollbackRequested {
    param([bool]$Switch)
    [bool]($Switch -or ($global:RemoveRollback -is [bool] -and $global:RemoveRollback))
}
# Decides whether Windows.old and the rollback data may be removed right now.
function Get-AmortRollbackDecision {
    param([bool]$Requested, [bool]$ServicingActive, [bool]$PendingReboot, [bool]$UpgradeActive)
    $reason = $null
    if (-not $Requested) { $reason = 'not requested' }
    elseif ($ServicingActive) { $reason = 'Windows servicing is active' }
    elseif ($PendingReboot) { $reason = 'a reboot is pending' }
    elseif ($UpgradeActive) { $reason = 'a Windows upgrade is running' }
    [pscustomobject]@{ Allowed = (-not $reason); Requested = $Requested; Reason = $reason }
}

# Removes C:\Windows.old. Strips the previous-installation protection first, then tries rd,
# an ownership retry, a robocopy mirror and finally a delete-on-reboot. Only called after an
# explicit -RemoveRollback and a clear Get-AmortRollbackDecision.
function Remove-AmortWindowsOld {
# Strip the previous-installation protection so Windows.old is deletable.
$dismResult = Invoke-AmortTool "DISM.exe" @("/Online", "/Remove-OSUninstall", "/NoRestart")
if ($dismResult.ExitCode -ne 0) { $script:Notes += "DISM /Remove-OSUninstall returned $($dismResult.ExitCode); removing Windows.old anyway." }

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
        $script:AmortReboot = $true
    } catch {
        $script:Notes += "Windows.old is still present after every removal attempt."
        $script:AmortFailures += "Windows.old could not be removed"
    }
}
}

function Invoke-AmortCleanup {
    param(
        [switch]$DryRun,
        [int]$RetainDays = 14,
        [int]$KeepRestorePoints = 2,
        [switch]$RemoveRollback
    )
$TempAgeDays = 2
$_fver   = "| v18.0"
#region Pre-Flight Checks
# ============================================================================
# Force UTF-8 output so box-drawing characters render correctly
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$OutputEncoding            = [System.Text.Encoding]::UTF8

# Detect active Windows servicing - do NOT kill it.
# Force-killing TiWorker/DISM mid-servicing corrupts the component store, which
# the repair region would then have to fix. Instead we detect and skip repair.
$ServicingActive = $null -ne (Get-Process -Name "TiWorker", "DISM" -ErrorAction SilentlyContinue)
if ($ServicingActive) {
    Write-Host "[Pre-Flight] Windows servicing active (TiWorker/DISM running). Repair steps will be skipped." -ForegroundColor Yellow
}
# A pending reboot means SoftwareDistribution and upgrade staging are still in use.
$PendingReboot = Test-AmortRebootPending

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
if ($TotalSize -le 0) { throw "PRECONDITION: the C: volume size could not be read." }

$StartUsagePct = [Math]::Round(((($TotalSize - $StartSpace) / $TotalSize) * 100), 2)
$script:LastRegionSpace = $Drive.FreeSpace # Rolling baseline for step-by-step reporting
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
try { Clear-Host } catch { }
$script:Width    = 90
$LineCol   = "DarkCyan"
$MainCol   = "DarkYellow"
$BorderCol = "Cyan"
$ArtCol    = "White"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$InfoCol      = "Cyan"
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
$UserProfiles = Get-ChildItem "C:\Users" -Directory -Force -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin @("All Users", "Default User") -and -not ($_.Attributes -band [IO.FileAttributes]::ReparsePoint) }
$script:Notes = @()
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
# $WINDOWS.~BT and $WINDOWS.~WS can hold the rollback data, so while Windows.old exists they follow the
# same opt-in decision as Windows.old itself.
$RollbackStaging = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS')
$KeepRollbackStaging = $false
if (Test-Path "C:\Windows.old") {
    $StagingDecision = Get-AmortRollbackDecision -Requested (Test-AmortRemoveRollbackRequested -Switch $RemoveRollback) -ServicingActive $ServicingActive -PendingReboot $PendingReboot -UpgradeActive $UpgradeActive
    $KeepRollbackStaging = -not $StagingDecision.Allowed
}
$UpgradeDirs = @('C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS', 'C:\$GetCurrent', 'C:\ESD', 'C:\Windows10Upgrade')
if ($UpgradeActive -or $PendingReboot) {
    # A pending reboot is reported once, on the reboot-pending line before the repair steps.
    if (-not $PendingReboot) { $script:Notes += "Upgrade staging folders were kept because a Windows upgrade is running." }
} else {
    foreach ($D in $UpgradeDirs) {
        if ($KeepRollbackStaging -and $RollbackStaging -contains $D) { continue }
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

# --- Windows.old and the rollback window (opt-in only) ---
# Removing Windows.old closes the "Go back to a previous version of Windows" option for good, so
# it only happens with -RemoveRollback, and never while servicing is active, a reboot is pending
# or an upgrade is running. By default it is kept and the technician is told how to remove it.
if (Test-Path "C:\Windows.old") {
    $Decision = Get-AmortRollbackDecision -Requested (Test-AmortRemoveRollbackRequested -Switch $RemoveRollback) -ServicingActive $ServicingActive -PendingReboot $PendingReboot -UpgradeActive $UpgradeActive
    if ($Decision.Allowed) {
        if ($DryRun) { $RegionEst += Get-PathSize "C:\Windows.old" }
        else { Remove-AmortWindowsOld }
    } elseif ($Decision.Requested) {
        $script:Notes += "Windows.old was not removed: $($Decision.Reason)."
        if (-not $DryRun) { $script:AmortSkips += "Windows.old was not removed: $($Decision.Reason)" }
    } else {
        $script:Notes += "Windows.old was kept because it holds the option to go back to a previous version of Windows. Run with -RemoveRollback to remove it."
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
    $Inventory = Get-AmortShadowInventory -VolumeDeviceId $CVol.DeviceID
    $Shadows = @($Inventory.Eligible)
    if ($Inventory.SkippedCount -gt 0) { $script:Notes += "$($Inventory.SkippedCount) shadow copies of C: were not System Restore points and were left alone." }
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
            if (-not (Remove-AmortShadowCopy -Id $S.ID -VolumeDeviceId $CVol.DeviceID)) {
                $script:AmortFailures += "restore point $($S.ID) could not be deleted"
            }
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
        if (Test-Path $L) { Invoke-AmortTool "compact.exe" @("/c", "/s:$L", "/i", "/q") | Out-Null }
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
$PendingRename = Get-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Session Manager" -Name "PendingFileRenameOperations" -ErrorAction SilentlyContinue
$HasPendingRename = $null -ne $PendingRename
# One decision for every repair step: dry run, active servicing, a pending reboot or a pending
# rename. TrustedInstaller is only touched when repair will actually run.
$SkipRepair = [bool]($DryRun -or $PendingReboot -or $ServicingActive -or $HasPendingRename)
if ($DryRun) {
    Write-Host "        [i] Dry run - DISM and SFC repair steps skipped (no changes)." -ForegroundColor Magenta
}
elseif ($PendingReboot) {
    Write-Host "        [!] Reboot pending, so upgrade folders, WU cache, and DISM/SFC were skipped." -ForegroundColor DarkYellow
    $script:AmortSkips += "reboot pending, so upgrade folders, WU cache, DISM and SFC were skipped"
}
elseif ($ServicingActive) {
    Write-Host "        [!] Windows servicing active (TiWorker/DISM running) - DISM and SFC repair steps will be skipped." -ForegroundColor DarkYellow
    $script:AmortSkips += "Windows servicing was active, so DISM and SFC were skipped"
}
elseif ($HasPendingRename) {
    Write-Host "        [!] PendingFileRenameOperations found - skipping repair steps." -ForegroundColor DarkYellow
    $script:AmortSkips += "PendingFileRenameOperations found, so DISM and SFC were skipped"
}
# Ensure TrustedInstaller is available (prevents DISM Error 87 / SFC failures)
if (-not $SkipRepair -and -not $DryRun) {
    $TI = Get-Service -Name "TrustedInstaller" -ErrorAction SilentlyContinue
    if ($TI) {
        if ($TI.StartType -eq 'Disabled') { Set-Service -Name "TrustedInstaller" -StartupType Manual }
        if ($TI.Status -ne 'Running') { Start-Service -Name "TrustedInstaller" -ErrorAction SilentlyContinue }
    }
}
# --- STEP 5: RestoreHealth ---
            Clear-InputBuffer
            $S72 = "[08/11] DISM RestoreHealth"
            Write-StepUpdate $S72 -CustomInfo "[Press ESC to Skip]"
            $Row72 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($SkipRepair) {
                Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[SKIPPED]"
            }
            else {
                $DismSpin = [char[]]@('|','/','-','\')
                $DismTmp1 = New-AmortTempFile

                # Use .NET Process directly; Start-Process -PassThru returns $null ExitCode
                # when combined with -RedirectStandardOutput. Wrap in cmd.exe for file redirection.
                $psi1 = New-Object System.Diagnostics.ProcessStartInfo
                $psi1.FileName               = Get-AmortNativePath "cmd.exe"
                $psi1.Arguments              = "/c dism.exe /Online /Cleanup-Image /RestoreHealth /NoRestart > `"$DismTmp1`" 2>&1"
                $psi1.UseShellExecute        = $false
                $psi1.CreateNoWindow         = $true
                $psi1.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Proc1 = Start-AmortProcess $psi1

                $Skipped1 = $false
                $DismSpinIdx1 = 0
                $DismTimer1 = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc1.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                Stop-AmortOwnedTree -ProcessId $Proc1.Id
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
                }
                $ExitCode1 = $Proc1.ExitCode
                try { $Proc1.Dispose() } catch {}
                Remove-Item $DismTmp1 -Force -ErrorAction SilentlyContinue

                if ($Skipped1) {
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[SKIPPED]"
                }
                elseif ($ExitCode1 -in @(0, 3010)) {
                    if ($ExitCode1 -eq 3010) { $script:AmortReboot = $true }
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -Success
                }
                else {
                    $script:AmortFailures += "DISM RestoreHealth failed (0x$($ExitCode1.ToString('X')))"
                    Clear-AndReprintStep -StartRow $Row72 -Message $S72 -CustomInfo "[FAILED:0x$($ExitCode1.ToString('X'))]"
                }
            }

# --- STEP 6: DISM ComponentCleanup ---
            $S73 = "[09/11] DISM ComponentCleanup"
            Write-StepUpdate $S73 -CustomInfo "[Press ESC to Skip]"
            $Row73 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($SkipRepair) {
                Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[SKIPPED]"
            }
            else {
                # BITS is left running so the RMM and LiveConnect stay connected.
                Stop-Service wuauserv -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Stop-Service TrustedInstaller -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Service TrustedInstaller -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
                Start-Sleep -Seconds 3

                Clear-InputBuffer

                $DismTmp2 = New-AmortTempFile

                $psi2 = New-Object System.Diagnostics.ProcessStartInfo
                $psi2.FileName               = Get-AmortNativePath "cmd.exe"
                $psi2.Arguments              = "/c dism.exe /Online /Cleanup-Image /StartComponentCleanup /NoRestart > `"$DismTmp2`" 2>&1"
                $psi2.UseShellExecute        = $false
                $psi2.CreateNoWindow         = $true
                $psi2.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Proc2 = Start-AmortProcess $psi2

                $Skipped2 = $false
                $DismSpinIdx2 = 0
                $DismTimer2 = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc2.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                Stop-AmortOwnedTree -ProcessId $Proc2.Id
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
                }
                $ExitCode2 = $Proc2.ExitCode
                try { $Proc2.Dispose() } catch {}
                Remove-Item $DismTmp2 -Force -ErrorAction SilentlyContinue

                Start-Service wuauserv -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

                if ($Skipped2) {
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[SKIPPED]"
                }
                elseif ($ExitCode2 -in @(0, 3010)) {
                    if ($ExitCode2 -eq 3010) { $script:AmortReboot = $true }
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -Success
                }
                else {
                    $script:AmortFailures += "DISM ComponentCleanup failed (0x$($ExitCode2.ToString('X')))"
                    Clear-AndReprintStep -StartRow $Row73 -Message $S73 -CustomInfo "[FAILED:0x$($ExitCode2.ToString('X'))]"
                }
            }

# --- STEP 7: SFC /scannow ---
            $S74 = "[10/11] SFC /scannow"
            Write-StepUpdate $S74 -CustomInfo "[Press ESC to Skip]"
            $Row74 = try { [Console]::CursorTop - 1 } catch { -1 }

            if ($SkipRepair) {
                Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[SKIPPED]"
            }
            else {
                Clear-InputBuffer
                $SfcTmp = New-AmortTempFile

                # sfc.exe writes Unicode; capture via cmd redirection for reliable exit code
                $psi3 = New-Object System.Diagnostics.ProcessStartInfo
                $psi3.FileName               = Get-AmortNativePath "cmd.exe"
                $psi3.Arguments              = "/c sfc.exe /scannow > `"$SfcTmp`" 2>&1"
                $psi3.UseShellExecute        = $false
                $psi3.CreateNoWindow         = $true
                $psi3.WindowStyle            = [System.Diagnostics.ProcessWindowStyle]::Hidden
                $Proc3 = Start-AmortProcess $psi3

                $Skipped3 = $false
                $SfcSpinIdx = 0
                $SfcTimer = [Diagnostics.Stopwatch]::StartNew()

                while (-not $Proc3.HasExited) {
                    try {
                        if ([Console]::KeyAvailable) {
                            $Key = [Console]::ReadKey($true)
                            if ($Key.Key -eq [ConsoleKey]::Escape) {
                                Stop-AmortOwnedTree -ProcessId $Proc3.Id
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
                    $script:AmortFailures += "SFC /scannow reported a warning"
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[WARNING]"
                }
                else {
                    $script:AmortFailures += "SFC /scannow failed (0x$($ExitCode3.ToString('X')))"
                    Clear-AndReprintStep -StartRow $Row74 -Message $S74 -CustomInfo "[FAILED:0x$($ExitCode3.ToString('X'))]"
                }
            }
# Component cleanup can free real space, so report it instead of folding it in silently.
if (-not $DryRun) {
    $CurrentSpace = Get-FreeBytes
    $RepairSaved = [int64]($CurrentSpace - $script:LastRegionSpace)
    if ($RepairSaved -ge 1MB) { Write-Host "        Component store cleanup freed $(Format-Size $RepairSaved)." -ForegroundColor Gray }
    $script:LastRegionSpace = $CurrentSpace
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
    Invoke-AmortTool "ipconfig.exe" @("/flushdns") | Out-Null
    Invoke-AmortTool "powercfg.exe" @("/h", "off") | Out-Null
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
}

function Invoke-AmortMain {
    param(
        [switch]$DryRun,
        [int]$RetainDays = 14,
        [int]$KeepRestorePoints = 2,
        [switch]$RemoveRollback
    )
    if (-not (Test-AmortElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Exit-KillerScript 1
    }
    $script:AmortFailures = @()
    $script:AmortSkips    = @()
    $script:AmortReboot   = $false
    $script:Notes         = @()

    $fatal = $null
    $precondition = $false
    try {
        Invoke-AmortCleanup -DryRun:$DryRun -RetainDays $RetainDays -KeepRestorePoints $KeepRestorePoints -RemoveRollback:$RemoveRollback
    } catch {
        $fatal = $_.Exception.Message
        if ($fatal -like 'PRECONDITION:*') { $precondition = $true; $fatal = $fatal.Substring(13).Trim() }
    }

    if ($fatal) {
        Write-Host $fatal -ForegroundColor Red
        Exit-KillerScript 1
    }
    Exit-KillerScript 0
}

if (-not $global:KillerScriptsTestMode) { Invoke-AmortMain -DryRun:$DryRun -RetainDays $RetainDays -KeepRestorePoints $KeepRestorePoints -RemoveRollback:$RemoveRollback }
