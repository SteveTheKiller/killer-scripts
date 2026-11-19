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

    Exit codes:
      0     Every selected profile was removed and verified gone.
      1     Failure: nothing was removed.
      2     Not elevated, or a precondition failed (no match, live session, cancelled).
      3     Partial success: some profiles were removed, others failed or left residue.

    Profile Safety Guidance:
      LastUseTime is the most reliable staleness signal - it is OS-managed
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

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

#region [0] - HELPERS & THEME
# ============================================================================
$script:Width   = 85
$script:Version = "v1.0"

$script:LineCol   = "DarkGreen"    # HLine dashes, section labels
$script:MainCol   = "DarkRed"      # Titles, step text
$script:ArtCol    = "Red"          # Banner/footer box-drawing art
$script:BorderCol = "White"        # Banner/footer block prefix/suffix
$script:DimCol    = "DarkGray"     # Muted text, completed steps
$script:WarnCol   = "DarkYellow"   # Warnings, advisory notes
$script:OkCol     = "Green"        # Success, done indicators
$script:AccentCol = "Yellow"       # Highlighted values

function Test-PruneElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Console helpers. An RMM console has no cursor and no keyboard, so every one of these
# degrades quietly instead of throwing or waiting.
function Get-PruneCursorTop {
    try { return [Console]::CursorTop } catch { return -1 }
}

function Set-PruneCursor {
    param([int]$Row)
    try { [Console]::SetCursorPosition(0, $Row); return $true } catch { return $false }
}

function Read-PruneKey {
    # Returns the key pressed, or $null when there is no console to ask. Callers treat $null as "no".
    try { return [Console]::ReadKey($true) } catch { return $null }
}

function Read-PruneLine {
    param([string]$Prompt)
    # Returns $null when there is no console to ask. Callers treat $null as "no".
    try { return (Read-Host $Prompt) } catch { return $null }
}

function Write-HLine {
    param([string]$Style = "dashed", [int]$Width = $script:Width)
    $line = if ($Style -eq "dashed") { ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width) } else { "-" * $Width }
    $colors = @([ConsoleColor]$script:LineCol, [ConsoleColor]$script:ArtCol, [ConsoleColor]$script:DimCol)
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') { $fg = [ConsoleColor]$script:DimCol }
        else { $fg = $colors[$i % $colors.Count]; $i++ }
        Write-Host $char -NoNewline -ForegroundColor $fg
    }
    Write-Host ""
}

$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($Success) {
        $savedTop = Get-PruneCursorTop
        if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop -and (Set-PruneCursor -Row $script:StepRow)) {
            $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $script:DimCol -NoNewline
            if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $script:WarnCol -NoNewline }
            Write-Host "$_pad[SUCCESS]" -ForegroundColor $script:OkCol
            [void](Set-PruneCursor -Row $savedTop)
        } elseif ($script:StepMsg) {
            Write-Host ("{0} [SUCCESS]" -f $script:StepMsg) -ForegroundColor $script:DimCol
        }
        $script:StepRow = -1
    } else {
        $script:StepRow = Get-PruneCursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $script:MainCol
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
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) { return -1 }
    try {
        $measured = Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
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
        # Look the account up by SID, so a similar name can never answer for it. A failed lookup is
        # reported as its own state and is never guessed to be an enabled local account.
        try {
            $sidObj = $null
            try { $sidObj = New-Object System.Security.Principal.SecurityIdentifier($SID) } catch { $sidObj = $null }
            $lu = if ($sidObj) { Get-LocalUser -SID $sidObj -ErrorAction Stop } else { Get-LocalUser -Name $userPart -ErrorAction Stop }
            if ($lu.Enabled) { return "Local" }
            return "Disabled"
        } catch { return "LookupFailed" }
    }
    return "Domain"
}

function Show-Banner {
    param([string]$Username, [switch]$SkipSizes)
    $_pfx  = "█  "
    $script:Art1 = "╔═╗ ╦═╗ ╦ ╦ ╔╗╔ ╔══ "
    $script:Art2 = "╠═╝ ╠╦╝ ║ ║ ║║║ ╠═  "
    $script:Art3 = "╩   ╩╚╝ ╚═╝ ╝╚╝ ╚══ "
    $script:ArtW  = [Math]::Max($script:Art1.Length, [Math]::Max($script:Art2.Length, $script:Art3.Length))
    $script:Art1  = $script:Art1.PadRight($script:ArtW); $script:Art2 = $script:Art2.PadRight($script:ArtW); $script:Art3 = $script:Art3.PadRight($script:ArtW)
    $_fillW = $script:Width - $_pfx.Length - $script:ArtW
    $_title = "PROFILE REMOVAL UTILITY FOR NEGLECTED ENTRIES"
    $script:VerTag = "| $($script:Version)"
    $_pad   = " " * [Math]::Max(0, ($_fillW - $_title.Length - $script:VerTag.Length))
    Write-Host $_pfx -ForegroundColor $script:BorderCol -NoNewline; Write-Host $script:Art1 -ForegroundColor $script:ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $script:LineCol
    Write-Host $_pfx -ForegroundColor $script:BorderCol -NoNewline; Write-Host $script:Art2 -ForegroundColor $script:ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $script:MainCol
    Write-Host $_pfx -ForegroundColor $script:BorderCol -NoNewline; Write-Host $script:Art3 -ForegroundColor $script:ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $script:LineCol

    Write-Host "[>] Device Name  : " -NoNewline -ForegroundColor $script:LineCol; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $script:AccentCol
    Write-Host "[>] Logged In As : " -NoNewline -ForegroundColor $script:LineCol; Write-Host "$($env:USERNAME)" -ForegroundColor $script:AccentCol
    if ($Username) {
        Write-Host "[>] Target       : " -NoNewline -ForegroundColor $script:LineCol; Write-Host $Username -ForegroundColor $script:AccentCol
    }
    if ($SkipSizes) {
        Write-Host "[>] Size Scan    : " -NoNewline -ForegroundColor $script:LineCol
        Write-Host "Disabled (-SkipSizes)" -ForegroundColor $script:WarnCol
    }
    Write-HLine
}

function Show-Footer {
    $_sfx    = "█"
    $_ftrW   = $script:ArtW + 1
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_footer = "  PRUNE SEQUENCE COMPLETE"
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $script:VerTag.Length))
    Write-Host ("-" * $_ffillW) -ForegroundColor $script:LineCol -NoNewline; Write-Host " $($script:Art1)" -ForegroundColor $script:ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $script:BorderCol
    Write-Host "$_footer$_fpad$($script:VerTag)" -ForegroundColor $script:MainCol  -NoNewline; Write-Host " $($script:Art2)" -ForegroundColor $script:ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $script:BorderCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $script:LineCol -NoNewline; Write-Host " $($script:Art3)" -ForegroundColor $script:ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $script:BorderCol
}

function Get-StatusColor {
    param([string]$Status)
    switch ($Status) {
        "Orphaned" { "Red" } "Disabled" { $script:WarnCol } "Local" { $script:MainCol } "Domain" { $script:AccentCol } default { $script:DimCol }
    }
}
#endregion

#region [1] - PROFILE SCAN AND REMOVAL
# ============================================================================
function Get-PruneProfileRecord {
    # Fresh Win32_UserProfile lookup by SID. Returns $null when the record is gone.
    param([string]$SID, [switch]$Quiet)
    Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$SID'" -ErrorAction SilentlyContinue
}

function Invoke-PruneRegUnload {
    param([string]$SID)
    & reg unload "HKU\$SID" 2>&1 | Out-Null
}

function Remove-PruneProfile {
    # Removes one profile. Returns an object with Removed and Message.
    param([Parameter(Mandatory)]$Profile)
    try {
        $wmiProf = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($Profile.SID)'" -ErrorAction Stop
        if ($wmiProf) {
            # Remove-CimInstance handles both registry cleanup and folder deletion
            Remove-CimInstance -InputObject $wmiProf -ErrorAction Stop
        } elseif ($Profile.FolderExists) {
            # Fallback: WMI entry already gone, remove the folder manually
            Remove-Item -Path $Profile.LocalPath -Recurse -Force -ErrorAction Stop
        }
        return [pscustomobject]@{ Removed = $true; Message = '' }
    } catch {
        return [pscustomobject]@{ Removed = $false; Message = $_.Exception.Message }
    }
}

function Get-PruneProfiles {
    param([switch]$SkipSizes)
    $rawProfiles = @(Get-CimInstance -ClassName Win32_UserProfile |
        Where-Object { -not $_.Special } |
        Sort-Object LastUseTime -Descending)

    $profiles = [System.Collections.Generic.List[object]]::new()
    if ($rawProfiles.Count -eq 0) { return ,$profiles }

    $scanRow = Get-PruneCursorTop
    $scanPad = " " * $script:Width

    for ($i = 0; $i -lt $rawProfiles.Count; $i++) {
        $wp = $rawProfiles[$i]
        if ($scanRow -ge 0 -and (Set-PruneCursor -Row $scanRow)) {
            $label = [System.IO.Path]::GetFileName($wp.LocalPath)
            Write-Host ("        [{0}/{1}] {2,-35}" -f ($i + 1), $rawProfiles.Count, $label) -ForegroundColor $script:DimCol -NoNewline
        }

        $accountName  = Resolve-SIDToAccount -SID $wp.SID
        $status       = Get-AccountStatus -SID $wp.SID -AccountName $accountName
        $folderExists = Test-Path -LiteralPath $wp.LocalPath -PathType Container
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
    if ($scanRow -ge 0 -and (Set-PruneCursor -Row $scanRow)) {
        Write-Host $scanPad -NoNewline
        [void](Set-PruneCursor -Row $scanRow)
    }
    return ,$profiles
}

function Resolve-PrunePrincipalSid {
    # Resolves a service or task owner (DOMAIN\name, .\name, a bare name, a UPN or a SID string) to a
    # SID string. Returns $null when the owner cannot be resolved.
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    $n = $Name.Trim()
    try {
        if ($n -match '^S-1-\d+(-\d+)+$') { return (New-Object System.Security.Principal.SecurityIdentifier($n)).Value }
        if ($n.StartsWith('.\')) { $n = "$($env:COMPUTERNAME)\" + $n.Substring(2) }
        return ((New-Object System.Security.Principal.NTAccount($n)).Translate([System.Security.Principal.SecurityIdentifier])).Value
    } catch { return $null }
}

function Find-PruneOwnedItems {
    # Services and scheduled tasks configured to run as exactly the given profile's account. An owner
    # counts only when it resolves to the profile's SID, so a similar name (bob, bobby, CORP\bob) is
    # never listed and can never be removed. An owner that does not resolve is left alone.
    param([Parameter(Mandatory)]$Target)
    $cache = @{}
    $ownedBy = {
        param([string]$Owner)
        if ([string]::IsNullOrWhiteSpace($Owner)) { return $false }
        if (-not $cache.ContainsKey($Owner)) { $cache[$Owner] = Resolve-PrunePrincipalSid -Name $Owner }
        $resolved = $cache[$Owner]
        return [bool]($resolved -and ($resolved -ieq $Target.SID))
    }

    $services = @(Get-CimInstance Win32_Service -ErrorAction SilentlyContinue |
        Where-Object { & $ownedBy $_.StartName })
    $tasks = @(Get-ScheduledTask -ErrorAction SilentlyContinue |
        Where-Object { & $ownedBy $_.Principal.UserId })
    [pscustomobject]@{ Services = $services; Tasks = $tasks }
}

function Invoke-PruneInvestigation {
    # Shows what keeps one loaded profile mounted and offers to remove tasks and the profile.
    # Returns an object: Rescan (start over), Removed (the profile is gone).
    param([Parameter(Mandatory)]$Target)

    Write-Host "    Investigating: " -NoNewline -ForegroundColor $script:LineCol
    Write-Host $Target.DisplayName -ForegroundColor $script:MainCol
    Write-Host ""

    $owned    = Find-PruneOwnedItems -Target $Target
    $services = @($owned.Services)
    $tasks    = @($owned.Tasks)

    Write-Host "    Services running as this account:" -ForegroundColor $script:DimCol
    if ($services.Count -gt 0) {
        foreach ($svc in $services) {
            $svcStateCol = if ($svc.State -eq "Running") { $script:WarnCol } else { $script:DimCol }
            Write-Host "      " -NoNewline
            Write-Host $svc.Name.PadRight(28) -NoNewline -ForegroundColor $script:AccentCol
            Write-Host $svc.State.PadRight(10) -NoNewline -ForegroundColor $svcStateCol
            Write-Host $svc.DisplayName -ForegroundColor $script:DimCol
        }
    } else {
        Write-Host "      None found." -ForegroundColor $script:DimCol
    }

    Write-Host "    Scheduled tasks configured for this account:" -ForegroundColor $script:DimCol
    if ($tasks.Count -gt 0) {
        foreach ($task in $tasks) {
            $taskName     = if ($task.TaskName.Length -gt 48) { $task.TaskName.Substring(0, 45) + "..." } else { $task.TaskName }
            $taskStateCol = if ($task.State -eq "Running") { $script:WarnCol } else { $script:DimCol }
            Write-Host "      " -NoNewline
            Write-Host $taskName.PadRight(50) -NoNewline -ForegroundColor $script:AccentCol
            Write-Host $task.State -ForegroundColor $taskStateCol
        }
    } else {
        Write-Host "      None found." -ForegroundColor $script:DimCol
    }
    Write-Host ""

    if ($services.Count -eq 0 -and $tasks.Count -eq 0) {
        Write-Host "    No services or tasks found - hive may be stale. A reboot may release it." -ForegroundColor $script:WarnCol
        Write-Host ""
    }

    $tasksCleared = $false
    if ($tasks.Count -gt 0) {
        Write-Host "    Remove the $($tasks.Count) scheduled task(s) above? " -NoNewline -ForegroundColor $script:DimCol
        Write-Host "[Y/N] " -NoNewline -ForegroundColor $script:AccentCol
        $removeKey = Read-PruneKey
        if ($removeKey) { Write-Host $removeKey.KeyChar } else { Write-Host "N (no console)" }
        if ($removeKey -and $removeKey.Key -eq [ConsoleKey]::Y) {
            $tasksCleared = $true
            foreach ($task in $tasks) {
                try {
                    Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction Stop
                    Write-Host "      Removed: $($task.TaskName)" -ForegroundColor $script:OkCol
                } catch {
                    Write-Host "      Failed : $($task.TaskName) - $($_.Exception.Message)" -ForegroundColor "Red"
                    $tasksCleared = $false
                }
            }
        }
        Write-Host ""
    }

    if ($services.Count -gt 0) {
        Write-Host "    Services must be stopped/disabled manually:" -ForegroundColor $script:WarnCol
        Write-Host "      Stop-Service -Name <name>; Set-Service -Name <name> -StartupType Disabled" -ForegroundColor $script:DimCol
        Write-Host ""
    }

    # Only offer removal if there were no tasks, or the user removed them
    if ($tasks.Count -gt 0 -and -not $tasksCleared) {
        return [pscustomobject]@{ Rescan = $true; Removed = $false }
    }

    # Offer immediate profile removal - attempt reg unload to release the hive first
    Write-Host "    Attempt to remove this profile now? " -NoNewline -ForegroundColor $script:DimCol
    Write-Host "[Y/N] " -NoNewline -ForegroundColor $script:AccentCol
    $tryRemoveKey = Read-PruneKey
    if ($tryRemoveKey) { Write-Host $tryRemoveKey.KeyChar } else { Write-Host "N (no console)" }
    Write-Host ""

    $profileRemoved = $false
    if ($tryRemoveKey -and $tryRemoveKey.Key -eq [ConsoleKey]::Y) {
        $recheckProf = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($Target.SID)'" -ErrorAction SilentlyContinue
        if (-not $recheckProf) {
            Write-Host "    Profile entry is already gone." -ForegroundColor $script:DimCol
            $profileRemoved = $true
        } elseif ($recheckProf.Loaded) {
            Invoke-PruneRegUnload -SID $Target.SID
            Start-Sleep -Milliseconds 750
            $recheckProf = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($Target.SID)'" -ErrorAction SilentlyContinue
            if ($recheckProf -and $recheckProf.Loaded) {
                Write-Host "    Profile hive is still held open - reboot and re-run PRUNE." -ForegroundColor $script:WarnCol
            } else {
                try {
                    if ($recheckProf) { Remove-CimInstance -InputObject $recheckProf -ErrorAction Stop }
                    Write-Host "    Profile removed successfully." -ForegroundColor $script:OkCol
                    $profileRemoved = $true
                } catch {
                    Write-Host "    Removal failed: $($_.Exception.Message)" -ForegroundColor "Red"
                }
            }
        } else {
            try {
                Remove-CimInstance -InputObject $recheckProf -ErrorAction Stop
                Write-Host "    Profile removed successfully." -ForegroundColor $script:OkCol
                $profileRemoved = $true
            } catch {
                Write-Host "    Removal failed: $($_.Exception.Message)" -ForegroundColor "Red"
            }
        }
    }
    [pscustomobject]@{ Rescan = $true; Removed = $profileRemoved }
}
#endregion

#region [2] - MAIN LOGIC
# ============================================================================
function New-PruneOutcome {
    param([int]$Code, [bool]$Rescan = $false)
    [pscustomobject]@{ Code = $Code; Rescan = $Rescan }
}

function Invoke-PruneSession {
    # One scan-select-remove pass. Returns an outcome with the exit code and whether to scan again.
    param([string]$Username, [switch]$SkipSizes)

    $script:PruneFinal = $null
    Show-Banner -Username $Username -SkipSizes:$SkipSizes

    # --- Step 1: Scan Profiles ---
    Write-StepUpdate "Scanning user profiles..."
    $profiles = Get-PruneProfiles -SkipSizes:$SkipSizes

    if ($profiles.Count -eq 0) {
        Write-StepUpdate -Success
        Write-Host ""
        Write-Host "    No user profiles found." -ForegroundColor $script:WarnCol
        Show-Footer
        $script:PruneFinal = "[PASS] PRUNE found no user profiles, nothing to remove (exit 0)"
        return (New-PruneOutcome -Code 0)
    }

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
            Write-Host "    [!] No profile found matching '$Username'." -ForegroundColor $script:WarnCol
            Write-Host "    Run without -Username to list all profiles on this device." -ForegroundColor $script:DimCol
            Show-Footer
            $script:PruneFinal = "[FAIL] PRUNE found no profile matching '$Username', nothing was changed (exit 2)"
            return (New-PruneOutcome -Code 2)
        }

        if ($match.Loaded) {
            Write-Host ""
            Write-Host "    [!] '$($match.DisplayName)' is an active session [LIVE] and cannot be removed." -ForegroundColor $script:WarnCol
            Show-Footer
            $script:PruneFinal = "[FAIL] PRUNE refused: '$($match.DisplayName)' is an active session, nothing was changed (exit 2)"
            return (New-PruneOutcome -Code 2)
        }

        $dateDisplay = if ($match.LastUsed) { $match.LastUsed.ToLocalTime().ToString("yyyy-MM-dd") } else { "Unknown" }
        $daysDisplay = if ($match.DaysSince -ge 0) { "$($match.DaysSince) days ago" } else { "Unknown" }
        $dateAgeCol  = if ($match.DaysSince -gt 365) { "Red" } elseif ($match.DaysSince -gt 90) { $script:WarnCol } else { $script:OkCol }
        $statCol     = Get-StatusColor -Status $match.Status

        Write-Host ""
        Write-Host "    Target Profile" -ForegroundColor $script:LineCol
        Write-Host ""
        Write-Host "      Username  : " -NoNewline -ForegroundColor $script:DimCol; Write-Host $match.DisplayName -ForegroundColor $script:MainCol
        Write-Host "      Path      : " -NoNewline -ForegroundColor $script:DimCol; Write-Host $match.LocalPath -ForegroundColor $script:DimCol
        Write-Host "      Last Used : " -NoNewline -ForegroundColor $script:DimCol; Write-Host "$dateDisplay ($daysDisplay)" -ForegroundColor $dateAgeCol
        Write-Host "      Size      : " -NoNewline -ForegroundColor $script:DimCol; Write-Host $match.SizeDisplay -ForegroundColor $script:AccentCol
        Write-Host "      Status    : " -NoNewline -ForegroundColor $script:DimCol; Write-Host $match.Status -ForegroundColor $statCol
        Write-Host ""

        $toDelete.Add($match)

    } else {
        # --- Interactive Profile Table ---
        Write-Host ""

        Write-Host "    " -NoNewline
        Write-Host " # " -ForegroundColor $script:LineCol -NoNewline
        Write-Host "  " -NoNewline
        Write-Host ("Username".PadRight(24)) -ForegroundColor $script:LineCol -NoNewline
        Write-Host "  " -NoNewline
        Write-Host ("Last Used".PadRight(10)) -ForegroundColor $script:LineCol -NoNewline
        Write-Host "    " -NoNewline
        Write-Host "Days".PadLeft(5) -ForegroundColor $script:LineCol -NoNewline
        Write-Host "    " -NoNewline
        Write-Host "Size".PadLeft(8) -ForegroundColor $script:LineCol -NoNewline
        Write-Host "    " -NoNewline
        Write-Host "Status" -ForegroundColor $script:LineCol
        Write-Host ("    " + ("-" * 3) + "  " + ("-" * 24) + "  " + ("-" * 10) + "    " + ("-" * 5) + "    " + ("-" * 8) + "    " + ("-" * 10)) -ForegroundColor $script:LineCol

        foreach ($p in $sorted) {
            $idxStr  = if ($p.Loaded) { "[L]" } else { $p.TableIndex.ToString().PadLeft(3) }
            $idxCol  = if ($p.Loaded) { $script:WarnCol } else { $script:DimCol }
            $ageCol  = if ($p.Loaded)              { $script:WarnCol }
                       elseif ($p.DaysSince -lt 0)   { $script:DimCol  }
                       elseif ($p.DaysSince -gt 365)  { "Red"    }
                       elseif ($p.DaysSince -gt 90)   { $script:WarnCol }
                       else                           { $script:OkCol   }
            $statCol = Get-StatusColor -Status $p.Status
            $dateStr = if ($p.LastUsed) { $p.LastUsed.ToLocalTime().ToString("yyyy-MM-dd") } else { "Unknown   " }
            $daysStr = if ($p.DaysSince -ge 0) { ("$($p.DaysSince)d").PadLeft(5) } else { "  ???" }
            $statStr = $p.Status
            if ($p.Loaded)            { $statStr += " [LIVE]"      }
            if (-not $p.FolderExists) { $statStr += " [NO FOLDER]" }

            Write-Host "    " -NoNewline
            Write-Host $idxStr                        -ForegroundColor $idxCol    -NoNewline
            Write-Host "  "                           -NoNewline
            Write-Host ($p.DisplayName.PadRight(24))  -ForegroundColor $script:MainCol   -NoNewline
            Write-Host "  "                           -NoNewline
            Write-Host $dateStr                       -ForegroundColor $ageCol    -NoNewline
            Write-Host "    "                         -NoNewline
            Write-Host $daysStr                       -ForegroundColor $ageCol    -NoNewline
            Write-Host "    "                         -NoNewline
            Write-Host $p.SizeDisplay                 -ForegroundColor $script:AccentCol -NoNewline
            Write-Host "    "                         -NoNewline
            Write-Host $statStr                       -ForegroundColor $statCol
        }

        Write-Host ""
        Write-Host "    " -NoNewline -ForegroundColor $script:DimCol
        Write-Host "[L]" -NoNewline -ForegroundColor $script:WarnCol
        Write-Host " = Active session, cannot remove  |  Days: " -NoNewline -ForegroundColor $script:DimCol
        Write-Host ">365" -NoNewline -ForegroundColor "Red"
        Write-Host " / " -NoNewline -ForegroundColor $script:DimCol
        Write-Host ">90" -NoNewline -ForegroundColor $script:WarnCol
        Write-Host " / " -NoNewline -ForegroundColor $script:DimCol
        Write-Host "<=90" -ForegroundColor $script:OkCol
        Write-Host ""

        $selectable = @($sorted | Where-Object { -not $_.Loaded })
        if ($selectable.Count -eq 0) {
            Write-Host "    All profiles are currently active. No profiles can be removed." -ForegroundColor $script:WarnCol
            Write-Host ""

            $liveProfiles = @($sorted | Where-Object { $_.Loaded })

            Write-Host "    Investigate why a profile is loaded:" -ForegroundColor $script:DimCol
            Write-Host ""
            for ($li = 0; $li -lt $liveProfiles.Count; $li++) {
                $lp = $liveProfiles[$li]
                Write-Host ("      {0,2})  " -f ($li + 1)) -NoNewline -ForegroundColor $script:AccentCol
                Write-Host $lp.DisplayName.PadRight(26) -NoNewline -ForegroundColor $script:MainCol
                Write-Host $lp.LocalPath -ForegroundColor $script:DimCol
            }
            Write-Host ""

            :investigateLoop while ($true) {
                Write-Host "    Profile number to investigate, or " -NoNewline -ForegroundColor $script:DimCol
                Write-Host "Q" -NoNewline -ForegroundColor $script:AccentCol
                Write-Host " to quit: " -NoNewline -ForegroundColor $script:DimCol
                $liveInput = ""
                :inputLoop while ($true) {
                    $ik = Read-PruneKey
                    if (-not $ik) {
                        Write-Host "Q (no console)"; $liveInput = "Q"; break inputLoop
                    } elseif ($ik.Key -eq [ConsoleKey]::Enter) {
                        Write-Host ""; break inputLoop
                    } elseif ($ik.Key -eq [ConsoleKey]::Escape -or $ik.KeyChar -in @('q','Q')) {
                        Write-Host $ik.KeyChar; $liveInput = "Q"; break inputLoop
                    } elseif ($ik.Key -eq [ConsoleKey]::Backspace) {
                        if ($liveInput.Length -gt 0) { $liveInput = $liveInput.Substring(0, $liveInput.Length - 1); Write-Host "`b `b" -NoNewline }
                    } elseif ($ik.KeyChar -match '\d') {
                        $liveInput += $ik.KeyChar; Write-Host $ik.KeyChar -NoNewline
                    }
                }
                Write-Host ""

                if ($liveInput -eq 'Q') { break investigateLoop }

                if ($liveInput -notmatch '^\s*(\d+)\s*$' -or [int]$Matches[1] -lt 1 -or [int]$Matches[1] -gt $liveProfiles.Count) {
                    Write-Host "    [!] Invalid selection." -ForegroundColor $script:WarnCol
                    Write-Host ""
                    continue
                }

                $target = $liveProfiles[[int]$Matches[1] - 1]
                $inv = Invoke-PruneInvestigation -Target $target

                # Whether removal succeeded or was declined, rescan or quit
                if ($inv.Removed) {
                    Write-Host ""
                    Write-Host "    Press " -NoNewline -ForegroundColor $script:DimCol
                    Write-Host "Enter" -NoNewline -ForegroundColor $script:AccentCol
                    Write-Host " to re-scan profiles, or " -NoNewline -ForegroundColor $script:DimCol
                    Write-Host "Q" -NoNewline -ForegroundColor $script:AccentCol
                    Write-Host " to quit." -ForegroundColor $script:DimCol
                    $reloadKey = Read-PruneKey
                    Write-Host ""
                    if ($reloadKey -and $reloadKey.Key -ne [ConsoleKey]::Q) { return (New-PruneOutcome -Code 0 -Rescan $true) }
                    Show-Footer
                    $script:PruneFinal = "[PASS] PRUNE removed the investigated profile (exit 0)"
                    return (New-PruneOutcome -Code 0)
                }
                if ($inv.Rescan) { return (New-PruneOutcome -Code 0 -Rescan $true) }
                Show-Footer
                $script:PruneFinal = "[WARN] PRUNE did not remove the investigated profile (exit 2)"
                return (New-PruneOutcome -Code 2)
            }

            Show-Footer
            $script:PruneFinal = "[WARN] PRUNE found only active sessions, nothing was removed (exit 2)"
            return (New-PruneOutcome -Code 2)
        }

        Write-Host "    Enter numbers to remove (e.g. " -NoNewline -ForegroundColor $script:DimCol
        Write-Host "1,3,5" -NoNewline -ForegroundColor $script:AccentCol
        Write-Host " or " -NoNewline -ForegroundColor $script:DimCol
        Write-Host "2-4" -NoNewline -ForegroundColor $script:AccentCol
        Write-Host "), or " -NoNewline -ForegroundColor $script:DimCol
        Write-Host "Q" -NoNewline -ForegroundColor $script:AccentCol
        Write-Host " to quit." -ForegroundColor $script:DimCol
        Write-Host ""
        $rawInput = Read-PruneLine -Prompt "    Selection"
        Write-Host ""

        if ($null -eq $rawInput -or $rawInput -match '^\s*[Qq]') {
            Write-Host "    Cancelled." -ForegroundColor $script:DimCol
            Show-Footer
            $script:PruneFinal = "[FAIL] PRUNE cancelled, nothing was changed (exit 2)"
            return (New-PruneOutcome -Code 2)
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
            Write-Host ("    [!] Invalid number(s): {0}" -f ($notFound -join ', ')) -ForegroundColor $script:WarnCol
        }
        if ($skippedLive.Count -gt 0) {
            Write-Host ("    [!] Skipped (active session): {0}" -f ($skippedLive -join ', ')) -ForegroundColor $script:WarnCol
        }

        if ($toDelete.Count -eq 0) {
            Write-Host "    No valid profiles selected. Exiting." -ForegroundColor $script:DimCol
            Show-Footer
            $script:PruneFinal = "[FAIL] PRUNE had no valid profile selected, nothing was changed (exit 2)"
            return (New-PruneOutcome -Code 2)
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
        if ($p.Status -eq "LookupFailed")  { $flags += "account lookup failed" }
        if (-not $p.FolderExists)          { $flags += "no folder on disk" }
        $flagStr = if ($flags.Count -gt 0) { "  [$($flags -join ', ')]" } else { "" }

        Write-Host "      - " -NoNewline -ForegroundColor $script:DimCol
        Write-Host $p.DisplayName -NoNewline -ForegroundColor "Red"
        Write-Host "  " -NoNewline
        Write-Host $p.LocalPath -NoNewline -ForegroundColor $script:DimCol
        Write-Host "  $($p.SizeDisplay)$flagStr" -ForegroundColor $script:AccentCol
        if ($p.SizeBytes -gt 0) { $totalBytes += $p.SizeBytes }
    }

    Write-Host ""
    if ($totalBytes -gt 0) {
        Write-Host "    Estimated space to recover: " -NoNewline -ForegroundColor $script:DimCol
        Write-Host (Format-FileSize -Bytes $totalBytes) -ForegroundColor $script:OkCol
        Write-Host ""
    }

    Write-Host "    [!] Profile data will be permanently lost. This cannot be undone." -ForegroundColor $script:WarnCol
    Write-Host ""
    $confirm1 = Read-PruneLine -Prompt "    Delete $($toDelete.Count) profile(s)? [Y/N]"
    Write-Host ""

    if ($confirm1 -notmatch '^[Yy]$') {
        Write-Host "    Cancelled." -ForegroundColor $script:DimCol
        Show-Footer
        $script:PruneFinal = "[FAIL] PRUNE cancelled, nothing was changed (exit 2)"
        return (New-PruneOutcome -Code 2)
    }

    # --- Confirmation: Step 2 (type DELETE) ---
    Write-Host "    Type " -NoNewline -ForegroundColor $script:DimCol
    Write-Host "DELETE" -NoNewline -ForegroundColor "Red"
    Write-Host " (all caps) to confirm, or anything else to abort:" -ForegroundColor $script:DimCol
    Write-Host ""
    $confirm2 = Read-PruneLine -Prompt "    Confirm"
    Write-Host ""

    if ($confirm2 -cne 'DELETE') {
        Write-Host "    Confirmation text did not match. Cancelled." -ForegroundColor $script:DimCol
        Show-Footer
        $script:PruneFinal = "[FAIL] PRUNE confirmation text did not match, nothing was changed (exit 2)"
        return (New-PruneOutcome -Code 2)
    }

    Write-HLine -Style "solid"

    # --- Step 2: Remove Profiles ---
    Write-StepUpdate "Removing selected profiles..."

    $removedCount = 0
    $failedList   = [System.Collections.Generic.List[string]]::new()

    foreach ($p in $toDelete) {
        Write-Host ("        Removing: {0,-28} {1}" -f $p.DisplayName, $p.LocalPath) -ForegroundColor $script:DimCol
        $r = Remove-PruneProfile -Profile $p
        if ($r.Removed) {
            $removedCount++
        } else {
            [void]$failedList.Add($p.DisplayName)
            Write-Host ("        [!] Failed: {0} - {1}" -f $p.DisplayName, $r.Message) -ForegroundColor "Red"
        }
    }

    Write-StepUpdate -Success

    # --- Step 3: Verify ---
    Write-StepUpdate "Verifying removal..."

    $residual = [System.Collections.Generic.List[string]]::new()
    foreach ($p in $toDelete) {
        if ($failedList -contains $p.DisplayName) { continue }
        $wmiCheck = Get-CimInstance -ClassName Win32_UserProfile -Filter "SID='$($p.SID)'" -ErrorAction SilentlyContinue
        if ($wmiCheck)                                { [void]$residual.Add("$($p.DisplayName) (WMI entry remains)") }
        elseif (Test-Path -LiteralPath $p.LocalPath)  { [void]$residual.Add("$($p.DisplayName) (folder remains)")    }
    }

    Write-StepUpdate -Success

    # --- Results ---
    Write-HLine

    # A profile that still has residue is not counted as removed.
    $cleanRemoved = [Math]::Max(0, $removedCount - $residual.Count)
    $code = 0
    if ($failedList.Count -eq 0 -and $residual.Count -eq 0) {
        Write-Host "    STATUS  : " -NoNewline -ForegroundColor $script:LineCol
        Write-Host "COMPLETE" -ForegroundColor $script:OkCol
        Write-Host "    Removed : $cleanRemoved profile(s)" -ForegroundColor $script:DimCol
        if ($totalBytes -gt 0) {
            Write-Host "    Freed   : $(Format-FileSize -Bytes $totalBytes)" -ForegroundColor $script:OkCol
        }
        $finalLine = "[PASS] PRUNE complete: $cleanRemoved profile(s) removed and verified (exit 0)"
    } else {
        $code = if ($cleanRemoved -gt 0) { 3 } else { 1 }
        $label = if ($code -eq 3) { "PARTIAL" } else { "FAILED" }
        Write-Host "    STATUS  : " -NoNewline -ForegroundColor $script:LineCol
        Write-Host $label -ForegroundColor $script:WarnCol
        Write-Host "    Removed : $cleanRemoved profile(s)" -ForegroundColor $script:DimCol
        if ($failedList.Count -gt 0) {
            Write-Host "    Failed  : $($failedList.Count) - $($failedList -join ', ')" -ForegroundColor "Red"
        }
        if ($residual.Count -gt 0) {
            Write-Host "    Residual: $($residual -join '; ')" -ForegroundColor $script:WarnCol
        }
        $tag = if ($code -eq 3) { "[WARN]" } else { "[FAIL]" }
        $finalLine = "$tag PRUNE $($label.ToLower()): $cleanRemoved removed, $($failedList.Count) failed, $($residual.Count) with residue (exit $code)"
    }

    Write-Host ""
    Show-Footer
    $script:PruneFinal = $finalLine
    return (New-PruneOutcome -Code $code)
}

function Invoke-PruneMain {
    param([string]$Username, [switch]$SkipSizes)

    if (-not (Test-PruneElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Write-Output "[FAIL] PRUNE not elevated, nothing was changed (exit 2)"
        Exit-KillerScript 2
    }

    try { Clear-Host } catch { }
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

    $outcome = $null
    do {
        $outcome = Invoke-PruneSession -Username $Username -SkipSizes:$SkipSizes
    } while ($outcome.Rescan)

    # The final result line goes to the output stream, where an RMM captures it.
    if ($script:PruneFinal) { Write-Output $script:PruneFinal }

    Exit-KillerScript $outcome.Code
}
#endregion

if (-not $global:KillerScriptsTestMode) { Invoke-PruneMain -Username $Username -SkipSizes:$SkipSizes }
