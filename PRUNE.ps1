<#
.SYNOPSIS
    Profile Removal Utility for Neglected Entries (PRUNE) v1.0
    Developed by SteveTheKiller | Updated: 2026-03-20
.DESCRIPTION
    Scans local Windows user profiles and displays them sorted by last-used
    date, flagging stale, orphaned, and disabled accounts for safe removal.
    When all profiles are active, enters investigation mode to identify the
    services or scheduled tasks keeping each hive mounted, with options to
    remove offending tasks and attempt immediate profile deletion on the spot.
    A -Username parameter allows targeting a specific profile directly.
.NOTES
    Parameters:
      -Username    Target a specific profile by username or folder name for direct removal.
      -SkipSizes   Skip folder size calculation for faster scanning.

    Profile Safety Guidance:
      LastUseTime is the most reliable staleness signal — it is OS-managed
      and updated on every profile load/unload, not affected by file-system
      activity. Thresholds used by this tool: >90d = stale, >365d = old.
      Orphaned profiles (SID no longer resolves to any account) are almost
      always safe to remove. Profiles marked [LIVE] are currently loaded
      and cannot be removed without a reboot or logoff.
#>
param(
    [string]$Username,
    [switch]$SkipSizes
)

$EXIT_SUCCESS = 0
$EXIT_CANCEL  = 2
$EXIT_DENIED  = 3

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit $EXIT_DENIED
}

#region [0] - HELPERS & THEME
# ============================================================================
Clear-Host
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$script:Width   = 85
$script:Version = "v1.0"

$LineCol   = "DarkGreen"    # HLine dashes, section labels
$MainCol   = "DarkRed"        # Titles, step text
$ArtCol    = "Red"           # Banner/footer box-drawing art
$BorderCol = "White"         # Banner/footer █ prefix/suffix
$DimCol    = "DarkGray"    # Muted text, completed steps
$WarnCol   = "DarkYellow"  # Warnings, advisory notes
$OkCol     = "Green"       # Success, done indicators
$AccentCol = "Yellow"      # Highlighted values

function Write-HLine {
    param([string]$Style = "dashed", [int]$Width = $script:Width)
    $line = if ($Style -eq "dashed") { ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width) } else { "-" * $Width }
    $colors = @([ConsoleColor]$LineCol, [ConsoleColor]$ArtCol, [ConsoleColor]$DimCol)
    $useConsole = $true
    try { $saved = [Console]::ForegroundColor } catch { $useConsole = $false }
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') { $fg = [ConsoleColor]$DimCol }
        else { $fg = $colors[$i % $colors.Count]; $i++ }
        if ($useConsole) { [Console]::ForegroundColor = $fg; [Console]::Write($char) }
        else { Write-Host $char -NoNewline -ForegroundColor $fg }
    }
    if ($useConsole) { [Console]::ForegroundColor = $saved; [Console]::WriteLine() }
    else { Write-Host "" }
}

$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($Success) {
        $savedTop = [Console]::CursorTop
        if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
            [Console]::SetCursorPosition(0, $script:StepRow)
            $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
            if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
            Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
            [Console]::SetCursorPosition(0, $savedTop)
        }
        $script:StepRow = -1
    } else {
        $script:StepRow = [Console]::CursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $MainCol
    }
}

function Get-ActiveUser {
    $ep = Get-Process -Name "explorer" -IncludeUserName -ErrorAction SilentlyContinue | Select-Object -First 1
    if ($ep) { return $ep.UserName -replace '^.*\\' }
    $qs = quser 2>$null | Select-String -Pattern "console\s+\d+\s+Active" | Select-Object -First 1
    if ($qs) { return ($qs.ToString().Trim() -split '\s+')[0] }
    $cs = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName
    if ($cs) { return $cs -replace '^.*\\' }
    return $null
}

function Format-FileSize {
    param([long]$Bytes)
    if ($Bytes -lt 0)   { return "     ---" }
    if ($Bytes -ge 1GB) { return "{0,6:N1} GB" -f ($Bytes / 1GB) }
    if ($Bytes -ge 1MB) { return "{0,6:N0} MB" -f ($Bytes / 1MB) }
    if ($Bytes -ge 1KB) { return "{0,6:N0} KB" -f ($Bytes / 1KB) }
    return "{0,7} B" -f $Bytes
}

function Get-FolderSizeBytes {
    param([string]$Path)
    if (-not (Test-Path $Path -PathType Container)) { return -1 }
    try {
        $measured = Get-ChildItem -Path $Path -Recurse -Force -ErrorAction SilentlyContinue |
            Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue
        if ($null -ne $measured.Sum) { return [long]$measured.Sum }
        return 0L
    } catch { return 0 }
}

function Resolve-SIDToAccount {
    param([string]$SID)
    try {
        $sidObj = New-Object System.Security.Principal.SecurityIdentifier($SID)
        return ($sidObj.Translate([System.Security.Principal.NTAccount])).Value
    } catch { return $null }
}

function Get-AccountStatus {
    param([string]$SID, [string]$AccountName)
    if (-not $AccountName) { return "Orphaned" }
    $domainPart = if ($AccountName -like "*\*") { ($AccountName -split "\\")[0] } else { $null }
    $userPart   = if ($AccountName -like "*\*") { ($AccountName -split "\\")[1] } else { $AccountName }
    $isLocal    = (-not $domainPart) -or ($domainPart -ieq $env:COMPUTERNAME)
    if ($isLocal) {
        try {
            $lu = Get-LocalUser -Name $userPart -ErrorAction Stop
            return if ($lu.Enabled) { "Local" } else { "Disabled" }
        } catch { return "Local" }
    }
    return "Domain"
}

function Show-Footer {
    $_sfx    = "█"
    $_ftrW   = $_artW + 1
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_footer = "  PRUNE SEQUENCE COMPLETE"
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol  -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
}
#endregion

#region [1] - BANNER
# ============================================================================
$_pfx  = "█  "
$_art1 = "╔═╗ ╦═╗ ╦ ╦ ╔╗╔ ╔══ "
$_art2 = "╠═╝ ╠╦╝ ║ ║ ║║║ ╠═  "
$_art3 = "╩   ╩╚╝ ╚═╝ ╝╚╝ ╚══ "
$_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1  = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "PROFILE REMOVAL UTILITY FOR NEGLECTED ENTRIES"
$_ver   = "| $($script:Version)"
$_pad   = " " * [Math]::Max(0, ($_fillW - $_title.Length - $_ver.Length))
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

Write-Host "[>] Device Name  : " -NoNewline -ForegroundColor $LineCol; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $AccentCol
Write-Host "[>] Logged In As : " -NoNewline -ForegroundColor $LineCol; Write-Host "$($env:USERNAME)" -ForegroundColor $AccentCol
if ($Username) {
    Write-Host "[>] Target       : " -NoNewline -ForegroundColor $LineCol; Write-Host $Username -ForegroundColor $AccentCol
}
if ($SkipSizes) {
    Write-Host "[>] Size Scan    : " -NoNewline -ForegroundColor $LineCol
    Write-Host "Disabled (-SkipSizes)" -ForegroundColor $WarnCol
}
Write-HLine
#endregion

#region [2] - MAIN LOGIC
# ============================================================================

# --- Step 1: Scan Profiles ---
Write-StepUpdate "Scanning user profiles..."

$rawProfiles = @(Get-CimInstance -ClassName Win32_UserProfile |
    Where-Object { -not $_.Special } |
    Sort-Object LastUseTime -Descending)

if ($rawProfiles.Count -eq 0) {
    Write-StepUpdate -Success
    Write-Host ""
    Write-Host "    No user profiles found." -ForegroundColor $WarnCol
    Show-Footer
    exit $EXIT_SUCCESS
}

$profiles   = [System.Collections.Generic.List[object]]::new()
$scanRow    = [Console]::CursorTop
$scanPad    = " " * $script:Width

for ($i = 0; $i -lt $rawProfiles.Count; $i++) {
    $wp = $rawProfiles[$i]
    [Console]::SetCursorPosition(0, $scanRow)
    $label = [System.IO.Path]::GetFileName($wp.LocalPath)
    Write-Host ("        [{0}/{1}] {2,-35}" -f ($i + 1), $rawProfiles.Count, $label) -ForegroundColor $DimCol -NoNewline

    $accountName  = Resolve-SIDToAccount -SID $wp.SID
    $status       = Get-AccountStatus -SID $wp.SID -AccountName $accountName
    $folderExists = Test-Path $wp.LocalPath -PathType Container
    $sizeBytes    = if ($SkipSizes -or -not $folderExists) { -1 } else { Get-FolderSizeBytes -Path $wp.LocalPath }
    $daysSince    = if ($wp.LastUseTime) { [int]((Get-Date) - $wp.LastUseTime.ToLocalTime()).TotalDays } else { -1 }

    $displayName  = if ($accountName) { $accountName } else { $wp.SID }
    if ($displayName.Length -gt 24) { $displayName = $displayName.Substring(0, 21) + "..." }

    $profiles.Add([PSCustomObject]@{
        TableIndex   = 0
        SID          = $wp.SID
        AccountName  = $accountName
        DisplayName  = $displayName
        LocalPath    = $wp.LocalPath
        FolderExists = $folderExists
        LastUsed     = $wp.LastUseTime
        DaysSince    = $daysSince
        SizeBytes    = $sizeBytes
        SizeDisplay  = Format-FileSize -Bytes $sizeBytes
        Status       = $status
        Loaded       = $wp.Loaded
    })
}

# Clear the scan progress line
[Console]::SetCursorPosition(0, $scanRow)
Write-Host $scanPad -NoNewline
[Console]::SetCursorPosition(0, $scanRow)

Write-StepUpdate -Success

# Sort oldest-first (unknown LastUseTime at bottom), assign table indices
$sorted = @($profiles | Sort-Object @{
    Expression = { if ($_.DaysSince -lt 0) { -1 } else { $_.DaysSince } }
    Descending = $true
})
for ($i = 0; $i -lt $sorted.Count; $i++) { $sorted[$i].TableIndex = $i + 1 }

$toDelete = [System.Collections.Generic.List[object]]::new()

if ($Username) {
    # --- Direct Profile Targeting ---
    $match = $sorted | Where-Object {
        ([System.IO.Path]::GetFileName($_.LocalPath) -ieq $Username) -or
        ($_.AccountName -ieq $Username) -or
        ($_.AccountName -like "*\$Username")
    } | Select-Object -First 1

    if (-not $match) {
        Write-Host ""
        Write-Host "    [!] No profile found matching '$Username'." -ForegroundColor $WarnCol
        Write-Host "    Run without -Username to list all profiles on this device." -ForegroundColor $DimCol
        Show-Footer
        exit $EXIT_CANCEL
    }

    if ($match.Loaded) {
        Write-Host ""
        Write-Host "    [!] '$($match.DisplayName)' is an active session [LIVE] and cannot be removed." -ForegroundColor $WarnCol
        Show-Footer
        exit $EXIT_CANCEL
    }

    $dateDisplay = if ($match.LastUsed) { $match.LastUsed.ToLocalTime().ToString("yyyy-MM-dd") } else { "Unknown" }
    $daysDisplay = if ($match.DaysSince -ge 0) { "$($match.DaysSince) days ago" } else { "Unknown" }
    $dateAgeCol  = if ($match.DaysSince -gt 365) { "Red" } elseif ($match.DaysSince -gt 90) { $WarnCol } else { $OkCol }
    $statCol     = switch ($match.Status) {
        "Orphaned" { "Red" } "Disabled" { $WarnCol } "Local" { $MainCol } "Domain" { $AccentCol } default { $DimCol }
    }

    Write-Host ""
    Write-Host "    Target Profile" -ForegroundColor $LineCol
    Write-Host ""
    Write-Host "      Username  : " -NoNewline -ForegroundColor $DimCol; Write-Host $match.DisplayName -ForegroundColor $MainCol
    Write-Host "      Path      : " -NoNewline -ForegroundColor $DimCol; Write-Host $match.LocalPath -ForegroundColor $DimCol
    Write-Host "      Last Used : " -NoNewline -ForegroundColor $DimCol; Write-Host "$dateDisplay ($daysDisplay)" -ForegroundColor $dateAgeCol
    Write-Host "      Size      : " -NoNewline -ForegroundColor $DimCol; Write-Host $match.SizeDisplay -ForegroundColor $AccentCol
    Write-Host "      Status    : " -NoNewline -ForegroundColor $DimCol; Write-Host $match.Status -ForegroundColor $statCol
    Write-Host ""

    $toDelete.Add($match)

} else {
    # --- Interactive Profile Table ---
    Write-Host ""

    Write-Host "    " -NoNewline
    Write-Host " # " -ForegroundColor $LineCol -NoNewline
    Write-Host "  " -NoNewline
    Write-Host ("Username".PadRight(24)) -ForegroundColor $LineCol -NoNewline
    Write-Host "  " -NoNewline
    Write-Host ("Last Used".PadRight(10)) -ForegroundColor $LineCol -NoNewline
    Write-Host "    " -NoNewline
    Write-Host "Days".PadLeft(5) -ForegroundColor $LineCol -NoNewline
    Write-Host "    " -NoNewline
    Write-Host "Size".PadLeft(8) -ForegroundColor $LineCol -NoNewline
    Write-Host "    " -NoNewline
    Write-Host "Status" -ForegroundColor $LineCol
    Write-Host ("    " + ("-" * 3) + "  " + ("-" * 24) + "  " + ("-" * 10) + "    " + ("-" * 5) + "    " + ("-" * 8) + "    " + ("-" * 10)) -ForegroundColor $LineCol

    foreach ($p in $sorted) {
        $idxStr  = if ($p.Loaded) { "[L]" } else { $p.TableIndex.ToString().PadLeft(3) }
        $idxCol  = if ($p.Loaded) { $WarnCol } else { $DimCol }
        $ageCol  = if ($p.Loaded)              { $WarnCol }
                   elseif ($p.DaysSince -lt 0)   { $DimCol  }
                   elseif ($p.DaysSince -gt 365)  { "Red"    }
                   elseif ($p.DaysSince -gt 90)   { $WarnCol }
                   else                           { $OkCol   }
        $statCol = switch ($p.Status) {
            "Orphaned" { "Red" } "Disabled" { $WarnCol } "Local" { $MainCol } "Domain" { $AccentCol } default { $DimCol }
        }
        $dateStr = if ($p.LastUsed) { $p.LastUsed.ToLocalTime().ToString("yyyy-MM-dd") } else { "Unknown   " }
        $daysStr = if ($p.DaysSince -ge 0) { ("$($p.DaysSince)d").PadLeft(5) } else { "  ???" }
        $statStr = $p.Status
        if ($p.Loaded)            { $statStr += " [LIVE]"      }
        if (-not $p.FolderExists) { $statStr += " [NO FOLDER]" }

        Write-Host "    " -NoNewline
        Write-Host $idxStr                        -ForegroundColor $idxCol    -NoNewline
        Write-Host "  "                           -NoNewline
        Write-Host ($p.DisplayName.PadRight(24))  -ForegroundColor $MainCol   -NoNewline
        Write-Host "  "                           -NoNewline
        Write-Host $dateStr                       -ForegroundColor $ageCol    -NoNewline
        Write-Host "    "                         -NoNewline
        Write-Host $daysStr                       -ForegroundColor $ageCol    -NoNewline
        Write-Host "    "                         -NoNewline
        Write-Host $p.SizeDisplay                 -ForegroundColor $AccentCol -NoNewline
        Write-Host "    "                         -NoNewline
        Write-Host $statStr                       -ForegroundColor $statCol
    }

    Write-Host ""
    Write-Host "    " -NoNewline -ForegroundColor $DimCol
    Write-Host "[L]" -NoNewline -ForegroundColor $WarnCol
    Write-Host " = Active session, cannot remove  |  Days: " -NoNewline -ForegroundColor $DimCol
    Write-Host ">365" -NoNewline -ForegroundColor "Red"
    Write-Host " / " -NoNewline -ForegroundColor $DimCol
    Write-Host ">90" -NoNewline -ForegroundColor $WarnCol
    Write-Host " / " -NoNewline -ForegroundColor $DimCol
    Write-Host "<=90" -ForegroundColor $OkCol
    Write-Host ""

    $selectable = @($sorted | Where-Object { -not $_.Loaded })
    if ($selectable.Count -eq 0) {
        Write-Host "    All profiles are currently active. No profiles can be removed." -ForegroundColor $WarnCol
        Write-Host ""

        $liveProfiles = @($sorted | Where-Object { $_.Loaded })
        $dynStart = [Console]::CursorTop

        # Initial render: numbered list
        Write-Host "    Investigate why a profile is loaded:" -ForegroundColor $DimCol
        Write-Host ""
        for ($li = 0; $li -lt $liveProfiles.Count; $li++) {
            $lp = $liveProfiles[$li]
            Write-Host ("      {0,2})  " -f ($li + 1)) -NoNewline -ForegroundColor $AccentCol
            Write-Host $lp.DisplayName.PadRight(26) -NoNewline -ForegroundColor $MainCol
            Write-Host $lp.LocalPath -ForegroundColor $DimCol
        }
        Write-Host ""

        :investigateLoop while ($true) {
            Write-Host "    Profile number to investigate, or " -NoNewline -ForegroundColor $DimCol
            Write-Host "Q" -NoNewline -ForegroundColor $AccentCol
            Write-Host " to quit: " -NoNewline -ForegroundColor $DimCol
            $liveInput = ""
            :inputLoop while ($true) {
                $ik = [Console]::ReadKey($true)
                if ($ik.Key -eq [ConsoleKey]::Enter) {
                    Write-Host ""; break inputLoop
                } elseif ($ik.Key -eq [ConsoleKey]::Escape -or $ik.KeyChar -in @('q','Q')) {
                    Write-Host $ik.KeyChar; $liveInput = "Q"; break inputLoop
                } elseif ($ik.Key -eq [ConsoleKey]::Backspace) {
                    if ($liveInput.Length -gt 0) { $liveInput = $liveInput.Substring(0, $liveInput.Length - 1); [Console]::Write("`b `b") }
                } elseif ($ik.KeyChar -match '\d') {
                    $liveInput += $ik.KeyChar; [Console]::Write($ik.KeyChar)
                }
            }
            Write-Host ""

            if ($liveInput -eq 'Q') { break investigateLoop }

            if ($liveInput -notmatch '^\s*(\d+)\s*$' -or [int]$Matches[1] -lt 1 -or [int]$Matches[1] -gt $liveProfiles.Count) {
                Write-Host "    [!] Invalid selection." -ForegroundColor $WarnCol
                Write-Host ""
                continue
            }

            $target = $liveProfiles[[int]$Matches[1] - 1]
            $uname  = if ($target.AccountName -like "*\*") { ($target.AccountName -split "\\")[1] } else { $target.AccountName }
            if (-not $uname) { $uname = [System.IO.Path]::GetFileName($target.LocalPath) }

            # Erase from $dynStart to current cursor, then jump back up
            $wipeEnd = [Console]::CursorTop
            [Console]::SetCursorPosition(0, $dynStart)
            for ($r = $dynStart; $r -lt $wipeEnd; $r++) { Write-Host (" " * $script:Width) }
            [Console]::SetCursorPosition(0, $dynStart)

            # Print results in place of the list
            Write-Host "    Investigating: " -NoNewline -ForegroundColor $LineCol
            Write-Host $target.DisplayName -ForegroundColor $MainCol
            Write-Host ""

            $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
                Where-Object { $_.StartName -like "*$uname*" })

            Write-Host "    Services running as this account:" -ForegroundColor $DimCol
            if ($services.Count -gt 0) {
                foreach ($svc in $services) {
                    $svcStateCol = if ($svc.State -eq "Running") { $WarnCol } else { $DimCol }
                    Write-Host "      " -NoNewline
                    Write-Host $svc.Name.PadRight(28) -NoNewline -ForegroundColor $AccentCol
                    Write-Host $svc.State.PadRight(10) -NoNewline -ForegroundColor $svcStateCol
                    Write-Host $svc.DisplayName -ForegroundColor $DimCol
                }
            } else {
                Write-Host "      None found." -ForegroundColor $DimCol
            }

            $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
                Where-Object { $_.Principal.UserId -like "*$uname*" })

            Write-Host "    Scheduled tasks configured for this account:" -ForegroundColor $DimCol
            if ($tasks.Count -gt 0) {
                foreach ($task in $tasks) {
                    $taskName     = if ($task.TaskName.Length -gt 48) { $task.TaskName.Substring(0, 45) + "..." } else { $task.TaskName }
                    $taskStateCol = if ($task.State -eq "Running") { $WarnCol } else { $DimCol }
                    Write-Host "      " -NoNewline
                    Write-Host $taskName.PadRight(50) -NoNewline -ForegroundColor $AccentCol
                    Write-Host $task.State -ForegroundColor $taskStateCol
                }
            } else {
                Write-Host "      None found." -ForegroundColor $DimCol
            }
            Write-Host ""

            if ($services.Count -eq 0 -and $tasks.Count -eq 0) {
                Write-Host "    No services or tasks found — hive may be stale. A reboot may release it." -ForegroundColor $WarnCol
                Write-Host ""
            }

            $tasksCleared = $false
            if ($tasks.Count -gt 0) {
                Write-Host "    Remove the $($tasks.Count) scheduled task(s) above? " -NoNewline -ForegroundColor $DimCol
                Write-Host "[Y/N] " -NoNewline -ForegroundColor $AccentCol
                $removeKey = [Console]::ReadKey($true)
                Write-Host $removeKey.KeyChar
                if ($removeKey.Key -eq [ConsoleKey]::Y) {
                    $tasksCleared = $true
                    foreach ($task in $tasks) {
                        try {
                            Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
                            Write-Host "      Removed: $($task.TaskName)" -ForegroundColor $OkCol
                        } catch {
                            Write-Host "      Failed : $($task.TaskName) — $($_.Exception.Message)" -ForegroundColor "Red"
                            $tasksCleared = $false
                        }
                    }
                }
                Write-Host ""
            }

            if ($services.Count -gt 0) {
                Write-Host "    Services must be stopped/disabled manually:" -ForegroundColor $WarnCol
                Write-Host "      Stop-Service -Name <name>; Set-Service -Name <name> -StartupType Disabled" -ForegroundColor $DimCol
                Write-Host ""
            }

            # Only offer removal if there were no tasks, or the user removed them
            if ($tasks.Count -gt 0 -and -not $tasksCleared) {
                $relaunchArgs = @{}
                if ($SkipSizes) { $relaunchArgs['SkipSizes'] = $true }
                & $PSCommandPath @relaunchArgs
                exit $EXIT_SUCCESS
            }

            # Offer immediate profile removal — attempt reg unload to release the hive first
            Write-Host "    Attempt to remove this profile now? " -NoNewline -ForegroundColor $DimCol
            Write-Host "[Y/N] " -NoNewline -ForegroundColor $AccentCol
            $tryRemoveKey = [Console]::ReadKey($true)
            Write-Host $tryRemoveKey.KeyChar
            Write-Host ""

            $profileRemoved = $false
            if ($tryRemoveKey.Key -eq [ConsoleKey]::Y) {
                $recheckProf = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($target.SID)'" -ErrorAction SilentlyContinue
                if (-not $recheckProf) {
                    Write-Host "    Profile entry is already gone." -ForegroundColor $DimCol
                    $profileRemoved = $true
                } elseif ($recheckProf.Loaded) {
                    & reg unload "HKU\$($target.SID)" 2>&1 | Out-Null
                    Start-Sleep -Milliseconds 750
                    $recheckProf = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($target.SID)'" -ErrorAction SilentlyContinue
                    if ($recheckProf -and $recheckProf.Loaded) {
                        Write-Host "    Profile hive is still held open — reboot and re-run PRUNE." -ForegroundColor $WarnCol
                    } else {
                        try {
                            if ($recheckProf) { Remove-CimInstance -InputObject $recheckProf -ErrorAction Stop }
                            Write-Host "    Profile removed successfully." -ForegroundColor $OkCol
                            $profileRemoved = $true
                        } catch {
                            Write-Host "    Removal failed: $($_.Exception.Message)" -ForegroundColor "Red"
                        }
                    }
                } else {
                    try {
                        Remove-CimInstance -InputObject $recheckProf -ErrorAction Stop
                        Write-Host "    Profile removed successfully." -ForegroundColor $OkCol
                        $profileRemoved = $true
                    } catch {
                        Write-Host "    Removal failed: $($_.Exception.Message)" -ForegroundColor "Red"
                    }
                }

            }

            # Whether removal succeeded or was declined, rescan or quit
            $relaunchArgs = @{}
            if ($SkipSizes) { $relaunchArgs['SkipSizes'] = $true }
            if ($profileRemoved) {
                Write-Host ""
                Write-Host "    Press " -NoNewline -ForegroundColor $DimCol
                Write-Host "Enter" -NoNewline -ForegroundColor $AccentCol
                Write-Host " to re-scan profiles, or " -NoNewline -ForegroundColor $DimCol
                Write-Host "Q" -NoNewline -ForegroundColor $AccentCol
                Write-Host " to quit." -ForegroundColor $DimCol
                $reloadKey = [Console]::ReadKey($true)
                Write-Host ""
                if ($reloadKey.Key -ne [ConsoleKey]::Q) { & $PSCommandPath @relaunchArgs }
                else { Show-Footer }
            } else {
                & $PSCommandPath @relaunchArgs
            }
            exit $EXIT_SUCCESS
        }

        Show-Footer
        exit $EXIT_SUCCESS
    }

    Write-Host "    Enter numbers to remove (e.g. " -NoNewline -ForegroundColor $DimCol
    Write-Host "1,3,5" -NoNewline -ForegroundColor $AccentCol
    Write-Host " or " -NoNewline -ForegroundColor $DimCol
    Write-Host "2-4" -NoNewline -ForegroundColor $AccentCol
    Write-Host "), or " -NoNewline -ForegroundColor $DimCol
    Write-Host "Q" -NoNewline -ForegroundColor $AccentCol
    Write-Host " to quit." -ForegroundColor $DimCol
    Write-Host ""
    $rawInput = Read-Host "    Selection"
    Write-Host ""

    if ($rawInput -match '^\s*[Qq]') {
        Write-Host "    Cancelled." -ForegroundColor $DimCol
        Show-Footer
        exit $EXIT_CANCEL
    }

    $selectedIndices = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($part in ($rawInput -split ',')) {
        $part = $part.Trim()
        if ($part -match '^(\d+)-(\d+)$') {
            $lo = [int]$Matches[1]; $hi = [int]$Matches[2]
            [Math]::Min($lo, $hi)..[Math]::Max($lo, $hi) | ForEach-Object { [void]$selectedIndices.Add($_) }
        } elseif ($part -match '^\d+$') {
            [void]$selectedIndices.Add([int]$part)
        }
    }

    $skippedLive = [System.Collections.Generic.List[string]]::new()
    $notFound    = [System.Collections.Generic.List[int]]::new()

    foreach ($idx in ($selectedIndices | Sort-Object)) {
        $match = $sorted | Where-Object { $_.TableIndex -eq $idx } | Select-Object -First 1
        if (-not $match) { [void]$notFound.Add($idx); continue }
        if ($match.Loaded) { [void]$skippedLive.Add($match.DisplayName); continue }
        $toDelete.Add($match)
    }

    if ($notFound.Count -gt 0) {
        Write-Host ("    [!] Invalid number(s): {0}" -f ($notFound -join ', ')) -ForegroundColor $WarnCol
    }
    if ($skippedLive.Count -gt 0) {
        Write-Host ("    [!] Skipped (active session): {0}" -f ($skippedLive -join ', ')) -ForegroundColor $WarnCol
    }

    if ($toDelete.Count -eq 0) {
        Write-Host "    No valid profiles selected. Exiting." -ForegroundColor $DimCol
        Show-Footer
        exit $EXIT_CANCEL
    }
}

# --- Confirmation: Step 1 (Y/N) ---
Write-HLine -Style "solid"
Write-Host "    The following profiles will be PERMANENTLY DELETED:" -ForegroundColor "Red"
Write-Host ""

$totalBytes = 0L
foreach ($p in $toDelete) {
    $flags = @()
    if ($p.Status -eq "Orphaned")      { $flags += "orphaned" }
    if ($p.Status -eq "Disabled")      { $flags += "disabled" }
    if (-not $p.FolderExists)          { $flags += "no folder on disk" }
    $flagStr = if ($flags.Count -gt 0) { "  [$($flags -join ', ')]" } else { "" }

    Write-Host "      - " -NoNewline -ForegroundColor $DimCol
    Write-Host $p.DisplayName -NoNewline -ForegroundColor "Red"
    Write-Host "  " -NoNewline
    Write-Host $p.LocalPath -NoNewline -ForegroundColor $DimCol
    Write-Host "  $($p.SizeDisplay)$flagStr" -ForegroundColor $AccentCol
    if ($p.SizeBytes -gt 0) { $totalBytes += $p.SizeBytes }
}

Write-Host ""
if ($totalBytes -gt 0) {
    Write-Host "    Estimated space to recover: " -NoNewline -ForegroundColor $DimCol
    Write-Host (Format-FileSize -Bytes $totalBytes) -ForegroundColor $OkCol
    Write-Host ""
}

Write-Host "    [!] Profile data will be permanently lost. This cannot be undone." -ForegroundColor $WarnCol
Write-Host ""
$confirm1 = Read-Host "    Delete $($toDelete.Count) profile(s)? [Y/N]"
Write-Host ""

if ($confirm1 -notmatch '^[Yy]$') {
    Write-Host "    Cancelled." -ForegroundColor $DimCol
    Show-Footer
    exit $EXIT_CANCEL
}

# --- Confirmation: Step 2 (type DELETE) ---
Write-Host "    Type " -NoNewline -ForegroundColor $DimCol
Write-Host "DELETE" -NoNewline -ForegroundColor "Red"
Write-Host " (all caps) to confirm, or anything else to abort:" -ForegroundColor $DimCol
Write-Host ""
$confirm2 = Read-Host "    Confirm"
Write-Host ""

if ($confirm2 -cne 'DELETE') {
    Write-Host "    Confirmation text did not match. Cancelled." -ForegroundColor $DimCol
    Show-Footer
    exit $EXIT_CANCEL
}

Write-HLine -Style "solid"

# --- Step 2: Remove Profiles ---
Write-StepUpdate "Removing selected profiles..."

$removedCount = 0
$failedList   = [System.Collections.Generic.List[string]]::new()

foreach ($p in $toDelete) {
    Write-Host ("        Removing: {0,-28} {1}" -f $p.DisplayName, $p.LocalPath) -ForegroundColor $DimCol
    try {
        $wmiProf = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($p.SID)'" -ErrorAction Stop
        if ($wmiProf) {
            # Remove-CimInstance handles both registry cleanup and folder deletion
            Remove-CimInstance -InputObject $wmiProf -ErrorAction Stop
        } elseif ($p.FolderExists) {
            # Fallback: WMI entry already gone, remove the folder manually
            Remove-Item -Path $p.LocalPath -Recurse -Force -ErrorAction Stop
        }
        $removedCount++
    } catch {
        [void]$failedList.Add($p.DisplayName)
        Write-Host ("        [!] Failed: {0} — {1}" -f $p.DisplayName, $_.Exception.Message) -ForegroundColor "Red"
    }
}

Write-StepUpdate -Success

# --- Step 3: Verify ---
Write-StepUpdate "Verifying removal..."

$residual = [System.Collections.Generic.List[string]]::new()
foreach ($p in $toDelete) {
    if ($failedList -contains $p.DisplayName) { continue }
    $wmiCheck = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($p.SID)'" -ErrorAction SilentlyContinue
    if ($wmiCheck)                { [void]$residual.Add("$($p.DisplayName) (WMI entry remains)") }
    elseif (Test-Path $p.LocalPath) { [void]$residual.Add("$($p.DisplayName) (folder remains)")    }
}

Write-StepUpdate -Success

#endregion

#region [3] - RESULTS
# ============================================================================
Write-HLine

if ($failedList.Count -eq 0 -and $residual.Count -eq 0) {
    Write-Host "    STATUS  : " -NoNewline -ForegroundColor $LineCol
    Write-Host "COMPLETE" -ForegroundColor $OkCol
    Write-Host "    Removed : $removedCount profile(s)" -ForegroundColor $DimCol
    if ($totalBytes -gt 0) {
        Write-Host "    Freed   : $(Format-FileSize -Bytes $totalBytes)" -ForegroundColor $OkCol
    }
} else {
    Write-Host "    STATUS  : " -NoNewline -ForegroundColor $LineCol
    Write-Host "COMPLETED WITH WARNINGS" -ForegroundColor $WarnCol
    Write-Host "    Removed : $removedCount profile(s)" -ForegroundColor $DimCol
    if ($failedList.Count -gt 0) {
        Write-Host "    Failed  : $($failedList.Count) — $($failedList -join ', ')" -ForegroundColor "Red"
    }
    if ($residual.Count -gt 0) {
        Write-Host "    Residual: $($residual -join '; ')" -ForegroundColor $WarnCol
    }
}

Write-Host ""
Show-Footer
#endregion

exit $EXIT_SUCCESS
