<#
.SYNOPSIS
    Scheduled Task Administration & Routine Executor (S.T.A.R.E.) v1.1
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    A clean, interactive terminal wizard that makes Windows Scheduled
    Task management fast and reliable for technicians. Supports Daily,
    Weekly, Monthly, and Startup triggers; run a command or browse the
    filesystem to select a script, with a progressive UI that keeps
    your selections in view at every step.
#>

#region [0] - EXIT, SYSTEM BOUNDARIES
# ============================================================================

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

function Exit-StareRun {
    # One exit point. The final line goes to the success stream so an RMM captures it, and it
    # always agrees with the exit code.
    param([int]$Code, [string]$Text)
    Write-Output "STARE RESULT: exit code $Code - $Text"
    Exit-KillerScript $Code
}

# System boundaries. Tests mock these.
function Test-StareElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

function Test-StareConsole {
    # True when a person can answer key prompts. False for RMM consoles, redirected input and no console.
    try {
        if ([Console]::IsInputRedirected) { return $false }
        $null = [Console]::KeyAvailable
        return $true
    } catch { return $false }
}

function Read-StareKey {
    [Console]::ReadKey($true)
}

function Get-StareCursor {
    # Current cursor position, or $null when the host has no cursor to move.
    try { $p = $Host.UI.RawUI.CursorPosition; return [pscustomobject]@{ X = [int]$p.X; Y = [int]$p.Y } } catch { return $null }
}

function Move-StareCursor {
    param([int]$X, [int]$Y)
    try { $Host.UI.RawUI.CursorPosition = New-Object System.Management.Automation.Host.Coordinates($X, $Y) } catch { }
}

function Get-StareScheduleService {
    $svc = New-Object -ComObject "Schedule.Service"
    $svc.Connect()
    $svc
}
#endregion

#region [1] - BANNER & DISPLAY HELPERS
# ============================================================================

$script:Acronym    = "STARE"
$script:AcronymFmt = ($script:Acronym.ToCharArray() -join '.') + '.'
$script:ScriptName = "Scheduled Task Administration & Routine Executor"
$script:Version    = "v1.1"

$script:Width = 85

$LineCol = "Magenta"
$MainCol = "Gray"
$script:SubtitleFG = "Red"
$script:LabelFG    = "DarkRed"
$script:SubLabelFG = "Magenta"
$script:MainTextFG = "Gray"
$script:ValueFG    = "Yellow"
$script:MutedFG    = "DarkGray"
$script:DoneFG     = "Green"
$script:WarnFG     = "DarkYellow"
$script:LineColors = @(
    [ConsoleColor]::DarkRed,
    [ConsoleColor]::Magenta,
    [ConsoleColor]::DarkMagenta,
    [ConsoleColor]::DarkGray,
    [ConsoleColor]::Red
)

function Write-HLine {
    param([string]$Style = "dashed", [int]$Width = $script:Width)
    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "━" * $Width
    }
    $colors = $script:LineColors
    $useConsole = $true
    try { $saved = [Console]::ForegroundColor } catch { $useConsole = $false }
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') {
            $fg = [ConsoleColor]$script:MutedFG
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

function Write-Banner {
    $_pfx  = "█  "
    $_art1 = "╔═╗ ╔╦╗ ╔═╗ ╦═╗ ╔═╗ "
    $_art2 = "╚═╗  ║  ╠═╣ ╠╦╝ ║╣  "
    $_art3 = "╚═╝  ╩  ╩ ╩ ╩╚═ ╚═╝ "
    $_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1  = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
    $_fillW = $script:Width - $_pfx.Length - $_artW
    $_title = $script:ScriptName.ToUpper()
    $_ver   = "| $($script:Version)"
    $_pad   = " " * [Math]::Max(0, ($_fillW - $_title.Length - $_ver.Length))
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $SubtitleFG -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $SubtitleFG -NoNewline; Write-Host "$_title$_pad$_ver" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $SubtitleFG -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
}

function Write-StatusHeader {
    param(
        [string]  $CurrentTaskName = "",
        [string]  $ScriptPath      = "",
        [string[]]$WizardLines     = @()
    )

    Write-Host "Device Name  : " -ForegroundColor $script:LabelFG -NoNewline
    Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $script:ValueFG

    if ($CurrentTaskName) {
        Write-Host "Task Name    : " -ForegroundColor $script:LabelFG -NoNewline
        Write-Host "$CurrentTaskName" -ForegroundColor $script:ValueFG
    }

    if ($ScriptPath) {
        $leaf = Split-Path $ScriptPath -Leaf
        Write-Host "Script       : " -ForegroundColor $script:LabelFG -NoNewline
        Write-Host "$leaf" -ForegroundColor $script:ValueFG
    }

    foreach ($wl in $WizardLines) {
        if ($wl -match '^(.+?) : (.*)$') {
            Write-Host "$($Matches[1]) : " -ForegroundColor $script:LabelFG -NoNewline
            Write-Host "$($Matches[2])" -ForegroundColor $script:ValueFG
        } else {
            Write-Host $wl -ForegroundColor $script:ValueFG
        }
    }

    Write-HLine -Style dashed
}

function Write-DashedSegment {
    param([int]$Length)
    for ($j = 0; $j -lt $Length; $j++) {
        if ($j % 2 -eq 0) {
            Write-Host "-" -NoNewline -ForegroundColor $script:LineColors[$script:colorIndex % $script:LineColors.Count]
            $script:colorIndex++
        } else {
            Write-Host " " -NoNewline
        }
    }
}

function Read-Esc {
    param([string]$Prompt)
    Write-Host "$Prompt" -ForegroundColor $script:MainTextFG -NoNewline
    $inputValue = ""
    while ($true) {
        if ([Console]::KeyAvailable) {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Escape) { return $null }
            if ($key.Key -eq [ConsoleKey]::Enter)  { Write-Host ""; return $inputValue }
            if ($key.Key -eq [ConsoleKey]::Backspace) {
                if ($inputValue.Length -gt 0) {
                    $inputValue = $inputValue.Substring(0, $inputValue.Length - 1)
                    Write-Host "`b `b" -NoNewline
                }
            } else {
                $inputValue += $key.KeyChar
                Write-Host $key.KeyChar -NoNewline -ForegroundColor $script:ValueFG
            }
        }
        Start-Sleep -Milliseconds 10
    }
}
#endregion
#region [2] - TASK CREATION WIZARD
# ============================================================================

function Invoke-TaskCreationWizard {
    # This wizard-style function encapsulates the entire task creation process
    # to prevent the complex loop control issues experienced previously. Each
    # step is a self-contained prompt that either returns a value or $null if cancelled.

    # --- Step 1: Get Task Name ---
    $taskName = ""
    while ([string]::IsNullOrWhiteSpace($taskName)) {
        Clear-Host; Write-Banner; Write-StatusHeader -ShowCreatorHeader $true
        Write-Host "Enter New Task Name " -ForegroundColor $script:MainTextFG -NoNewline
        Write-Host "(ESC = Back)" -ForegroundColor $script:MutedFG -NoNewline
        Write-Host ":" -ForegroundColor $script:MainTextFG
        Write-Host " > " -ForegroundColor $script:LabelFG -NoNewline
        $taskName = Read-Esc ""
        if ($null -eq $taskName) { return } # Exit wizard
        if ([string]::IsNullOrWhiteSpace($taskName)) {
            Write-Host "Task name cannot be empty!" -ForegroundColor $script:WarnFG; Start-Sleep -Seconds 1
        }
    }

    # --- Step 2: Get Action (Command or Script) ---
    $action = $null
    $scriptPathForHeader = ""
    $commandDisplay = ""
    while ($null -eq $action) {
        Clear-Host; Write-Banner; Write-StatusHeader -CurrentTaskName $taskName -ShowCreatorHeader $true
        Write-Host "What type of task?" -ForegroundColor $script:MainTextFG
        Write-Host "  [1] " -ForegroundColor $script:MainTextFG -NoNewline
        Write-Host "Run a command" -ForegroundColor $script:ValueFG
        Write-Host "  [2] " -ForegroundColor $script:MainTextFG -NoNewline
        Write-Host "Run a script" -ForegroundColor $script:ValueFG
        Write-Host "Selection: " -ForegroundColor $script:LabelFG -NoNewline

        $keyInfo = Read-StareKey
        if ($keyInfo.Key -eq [ConsoleKey]::Escape) { return } # Exit wizard
        $choice = $keyInfo.KeyChar.ToString()

        if ($choice -eq "1") {
            Clear-Host; Write-Banner; Write-StatusHeader -CurrentTaskName $taskName -ShowCreatorHeader $true
            Write-Host "Enter command to run " -ForegroundColor $script:MainTextFG -NoNewline
            Write-Host "(args included, e.g. 'shutdown.exe /s /t 60')" -ForegroundColor $script:MutedFG -NoNewline
            Write-Host ":" -ForegroundColor $script:MainTextFG
            Write-Host " > " -ForegroundColor $script:LabelFG -NoNewline
            $fullInput = Read-Esc ""
            if ($null -eq $fullInput) { continue }
            $fullInput = $fullInput.Trim()
            # Split on first whitespace - everything before is the exe, everything after is args
            $parts     = $fullInput -split '\s+', 2
            $execName  = $parts[0]
            $execArgs  = if ($parts.Count -gt 1) { $parts[1] } else { "" }
            $commandInfo = Get-Command $execName -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
            if (-not $commandInfo) {
                Write-Host "`nCommand not found or is not an executable: '$execName'." -ForegroundColor $script:WarnFG
                Start-Sleep -Seconds 3
                continue
            }
            $fullExecPath = $commandInfo.Source
            try {
                if ([string]::IsNullOrWhiteSpace($execArgs)) {
                    $action = New-ScheduledTaskAction -Execute $fullExecPath -ErrorAction Stop
                } else {
                    $action = New-ScheduledTaskAction -Execute $fullExecPath -Argument $execArgs -ErrorAction Stop
                }
                $commandDisplay = $fullInput
            } catch {
                Write-Host "Error creating action: $($_.Exception.Message)" -ForegroundColor Red; Start-Sleep -Seconds 2
            }
        }
        elseif ($choice -eq "2") {
            $currentDir = if ($PSScriptRoot) { $PSScriptRoot } else { "C:\" }
            while ($null -eq $action) {
                Clear-Host; Write-Banner; Write-StatusHeader -CurrentTaskName $taskName -ShowCreatorHeader $true
                Write-Host "Browsing: $currentDir" -ForegroundColor $script:ValueFG
                Write-Host "(Type 'cd [path]' to jump, ESC to go back)" -ForegroundColor $script:MutedFG
                Write-HLine -Style dashed

                $items = Get-ChildItem -Path $currentDir | Where-Object { $_.PSIsContainer -or $_.Extension -eq ".ps1" }
                # Aligned [0] row with correct colors
                $idxText0 = "[0]".PadLeft(4)
                Write-Host "  $idxText0 " -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host ("{0,-40} " -f ".. (Up One Level)") -ForegroundColor $script:MutedFG -NoNewline
                Write-Host "[DIR]" -ForegroundColor $script:MutedFG

                $i = 1; $map = @{}
                foreach ($item in $items) {
                    $typeTag   = if ($item.PSIsContainer) { "[DIR]" } else { "[FILE]" }
                    $nameColor = if ($item.PSIsContainer) { $script:SubLabelFG } else { $script:ValueFG }
                    $displayName = if ($item.Name.Length -gt 40) { $item.Name.Substring(0, 37) + "..." } else { $item.Name }

                    # Right-align index and apply correct colors to each column
                    $idxText   = "[{0}]" -f $i
                    $idxPadded = $idxText.PadLeft(4)
                    Write-Host "  $idxPadded " -ForegroundColor $script:MainTextFG -NoNewline
                    Write-Host ("{0,-40} " -f $displayName) -ForegroundColor $nameColor -NoNewline
                    Write-Host $typeTag -ForegroundColor $script:MutedFG

                    $map[$i] = $item; $i++
                }

                Write-Host "`nSelection: " -ForegroundColor $script:LabelFG -NoNewline
                $nav = Read-Esc ""
                if ($null -eq $nav) { break } # Exit file browser, back to command/script choice

                if ($nav -like "cd *") {
                    $inputPath = $nav.Substring(3).Trim()
                    try { $currentDir = (Resolve-Path -Path (Join-Path -Path $currentDir -ChildPath $inputPath)).ProviderPath }
                    catch { Write-Host "Invalid Path!" -ForegroundColor $script:WarnFG; Start-Sleep -Seconds 1 }
                }
                elseif ($nav -eq "0") { $currentDir = Split-Path $currentDir -Parent }
                elseif ($map.ContainsKey([int]$nav)) {
                    $selected = $map[[int]$nav]
                    if ($selected.PSIsContainer) {
                        $currentDir = $selected.FullName
                    } else {
                        $scriptPathForHeader = $selected.FullName
                        $execPath = "powershell.exe"
                        $execArgs = "-ExecutionPolicy Bypass -File `"$($selected.FullName)`""
                        $action = New-ScheduledTaskAction -Execute $execPath -Argument $execArgs
                    }
                }
            }
        }
    }

    # --- Step 3: Get Trigger ---
    $trigger = $null
    $taskRegistered = $false
    $finalWizardLines  = [string[]]@()
    $initialWizardLines = [string[]]@()
    if ($commandDisplay) { $initialWizardLines += "Command      : $commandDisplay" }
    while ($null -eq $trigger) {

        Clear-Host
        Write-Banner
        Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
            -WizardLines $initialWizardLines

        Write-Host "Schedule:" -ForegroundColor $script:LabelFG
        Write-Host "  [1] " -ForegroundColor $script:MainTextFG -NoNewline; Write-Host "Daily"   -ForegroundColor $script:ValueFG
        Write-Host "  [2] " -ForegroundColor $script:MainTextFG -NoNewline; Write-Host "Weekly"  -ForegroundColor $script:ValueFG
        Write-Host "  [3] " -ForegroundColor $script:MainTextFG -NoNewline; Write-Host "Monthly" -ForegroundColor $script:ValueFG
        Write-Host "  [4] " -ForegroundColor $script:MainTextFG -NoNewline; Write-Host "Startup" -ForegroundColor $script:ValueFG
        Write-Host "Selection: " -ForegroundColor $script:LabelFG -NoNewline

        $key = Read-StareKey
        if ($key.Key -eq [ConsoleKey]::Escape) { return }
        $sched = $key.KeyChar

        switch ($sched) {

            '1' {
                Clear-Host; Write-Banner
                Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
                    -WizardLines ($initialWizardLines + @("Schedule     : Daily"))
                Write-Host " Time " -ForegroundColor $script:LabelFG -NoNewline
                Write-Host "(HH:mm, default 00:00)" -ForegroundColor $script:MutedFG -NoNewline
                Write-Host ": " -ForegroundColor $script:LabelFG -NoNewline
                $t = Read-Esc ""
                if ($null -eq $t) { continue }
                if ([string]::IsNullOrWhiteSpace($t)) { $t = "00:00" }
                $finalWizardLines = $initialWizardLines + [string[]]@("Schedule     : Daily", "Time         : $t")
                $trigger = New-ScheduledTaskTrigger -Daily -At $t
            }

            '2' {
                # Step A - Days
                Clear-Host; Write-Banner
                Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
                    -WizardLines ($initialWizardLines + @("Schedule     : Weekly"))
                Write-Host " Days " -ForegroundColor $script:LabelFG -NoNewline
                Write-Host "(Mon,Tue,Wed,Thu,Fri,Sat,Sun, e.g. Mon,Wed)" -ForegroundColor $script:MutedFG -NoNewline
                Write-Host ": " -ForegroundColor $script:LabelFG -NoNewline
                $d = Read-Esc ""
                if ($null -eq $d) { continue }

                # Step B - Time
                Clear-Host; Write-Banner
                Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
                    -WizardLines ($initialWizardLines + @("Schedule     : Weekly", "Days         : $d"))
                Write-Host " Time " -ForegroundColor $script:LabelFG -NoNewline
                Write-Host "(HH:mm, default 00:00)" -ForegroundColor $script:MutedFG -NoNewline
                Write-Host ": " -ForegroundColor $script:LabelFG -NoNewline
                $t = Read-Esc ""
                if ($null -eq $t) { continue }
                if ([string]::IsNullOrWhiteSpace($t)) { $t = "00:00" }
                $daysArray = $d.Split(',').Trim()
                $finalWizardLines = $initialWizardLines + [string[]]@("Schedule     : Weekly", "Days         : $d", "Time         : $t")
                $trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek $daysArray -At $t
            }

            '3' {
                while ($true) {
                    # Step A - Day of Month
                    Clear-Host; Write-Banner
                    Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
                        -WizardLines ($initialWizardLines + @("Schedule     : Monthly"))
                    $currentDay = (Get-Date).Day
                    Write-Host " Day of Month " -ForegroundColor $script:LabelFG -NoNewline
                    Write-Host "(1-31, 0=last, default $currentDay)" -ForegroundColor $script:MutedFG -NoNewline
                    Write-Host ": " -ForegroundColor $script:LabelFG -NoNewline
                    $dom = Read-Esc ""
                    if ($null -eq $dom) { break }
                    if ([string]::IsNullOrWhiteSpace($dom)) { $dom = "$currentDay" }
                    $dom = $dom.Trim()
                    $domDisplay = if ($dom -eq "0") { "Last" } else { $dom }

                    # Validate day entries before advancing to Time
                    $daysOfMonthArray = @()
                    $invalid = $false
                    $runOnLastDay = $false
                    if ($dom -eq "0") {
                        $runOnLastDay = $true
                    } else {
                        foreach ($item in $dom.Split(',')) {
                            $clean = $item.Trim()
                            if ($clean -eq "") { continue }
                            $dayInt = 0
                            if (-not [int]::TryParse($clean, [ref]$dayInt)) { $invalid = $true; break }
                            if ($dayInt -lt 1 -or $dayInt -gt 31) { $invalid = $true; break }
                            $daysOfMonthArray += $dayInt
                        }
                    }
                    if ($invalid -or (-not $runOnLastDay -and $daysOfMonthArray.Count -eq 0)) {
                        Write-Host " Invalid day-of-month entry!" -ForegroundColor $script:WarnFG
                        Start-Sleep -Seconds 1
                        continue
                    }

                    # Step B - Time
                    Clear-Host; Write-Banner
                    Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
                        -WizardLines ($initialWizardLines + @("Schedule     : Monthly", "Day of Month : $domDisplay"))
                    Write-Host " Time " -ForegroundColor $script:LabelFG -NoNewline
                    Write-Host "(HH:mm, default 00:00)" -ForegroundColor $script:MutedFG -NoNewline
                    Write-Host ": " -ForegroundColor $script:LabelFG -NoNewline
                    $t = Read-Esc ""
                    if ($null -eq $t) { break }
                    if ([string]::IsNullOrWhiteSpace($t)) { $t = "00:00" }
                    $t = $t.Trim()

                    # Register via COM - New-ScheduledTaskTrigger -Monthly unavailable in PS7
                    try {
                        $startH = [int]($t.Split(':')[0])
                        $startM = [int]($t.Split(':')[1])
                        $startBoundary = (Get-Date -Hour $startH -Minute $startM -Second 0).ToString("yyyy-MM-ddTHH:mm:ss")

                        $comSched = Get-StareScheduleService
                        $rootFolder = $comSched.GetFolder("\")
                        $taskDef = $comSched.NewTask(0)

                        $taskDef.Settings.Enabled = $true
                        $taskDef.Settings.Hidden = $false
                        $taskDef.Settings.MultipleInstances = 2  # TASK_INSTANCES_IGNORE_NEW
                        $taskDef.Settings.ExecutionTimeLimit = "PT72H"

                        $taskDef.Principal.UserId = "S-1-5-18"
                        $taskDef.Principal.RunLevel = 1  # TASK_RUNLEVEL_HIGHEST

                        $comTrigger = $taskDef.Triggers.Create(4)  # TASK_TRIGGER_MONTHLY
                        $comTrigger.StartBoundary = $startBoundary
                        $comTrigger.Enabled = $true
                        $comTrigger.MonthsOfYear = 4095  # all 12 months

                        if ($runOnLastDay) {
                            $comTrigger.RunOnLastDayOfMonth = $true
                        } else {
                            $daysMask = 0
                            foreach ($day in $daysOfMonthArray) { $daysMask = $daysMask -bor (1 -shl ($day - 1)) }
                            $comTrigger.DaysOfMonth = $daysMask
                        }

                        $comAction = $taskDef.Actions.Create(0)  # TASK_ACTION_EXEC
                        $comAction.Path = $action.Execute
                        if ($action.Arguments) { $comAction.Arguments = $action.Arguments }

                        try { $rootFolder.DeleteTask($taskName, 0) } catch {}
                        $rootFolder.RegisterTaskDefinition($taskName, $taskDef, 6, "SYSTEM", $null, 5) | Out-Null

                        $finalWizardLines = $initialWizardLines + [string[]]@("Schedule     : Monthly", "Day of Month : $domDisplay", "Time         : $t")
                        $taskRegistered = $true
                        $trigger = $true  # sentinel to exit schedule loop
                    } catch {
                        $script:StareFailures++
                        Write-Host "`n[ERR] $($_.Exception.GetType().Name): $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "Press any key to go back..." -ForegroundColor DarkGray
                        $null = Read-StareKey
                    }
                    break
                }
            }

            '4' {
                $finalWizardLines = $initialWizardLines + [string[]]@("Schedule     : Startup")
                $trigger = New-ScheduledTaskTrigger -AtStartup
            }
        }
    }
    # --- Step 4: Register Task ---
    Clear-Host
    Write-Banner
    Write-StatusHeader -CurrentTaskName $taskName -ScriptPath $scriptPathForHeader -ShowCreatorHeader $true `
        -WizardLines $finalWizardLines

    if ($taskRegistered) {
        $script:StareChanges++
        Write-Host "[!] Task '$taskName' Created Successfully!" -ForegroundColor $script:DoneFG
    } else {
        try {
            if (Get-ScheduledTask -TaskName $taskName -ErrorAction SilentlyContinue) {
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false
            }
            Register-ScheduledTask -TaskName $taskName -Action $action -Trigger @($trigger) -User "SYSTEM" -RunLevel Highest -Force -ErrorAction Stop | Out-Null
            $script:StareChanges++
            Write-Host "[!] Task '$taskName' Created Successfully!" -ForegroundColor $script:DoneFG
        } catch {
            $script:StareFailures++
            Write-Host "[!] Error creating task: $($_.Exception.Message)" -ForegroundColor Red
        }
    }

    Write-Host "Press any key to return to the list..." -ForegroundColor $script:MutedFG
    $null = Read-StareKey
}
#endregion

#region [3] - MAIN LOOP
# ============================================================================

function Invoke-StareMain {
    $script:StareChanges  = 0
    $script:StareFailures = 0
    $script:MenuPos       = $null

    if (-not (Test-StareElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Exit-StareRun 2 "not elevated, nothing was changed"
    }

    # Force UTF-8 output so box-drawing characters render correctly. Hosts with no console reject this.
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
    try { $global:OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

    if (-not (Test-StareConsole)) {
        Exit-StareRun 2 "STARE is an interactive wizard and needs a console to answer its prompts, nothing was changed"
    }

    $script:TaskAction      = "List"
    $script:mainLoop        = $true
    $script:RefreshRequired = $true

    while ($script:mainLoop) {

        if ($script:TaskAction -eq "List") {

            if ($script:RefreshRequired) {
                Clear-Host
                Write-Banner
                Write-StatusHeader

                $tasks = @(Get-ScheduledTask |
                         Where-Object { $_.TaskPath -notlike "\Microsoft\*" } |
                         Sort-Object { $_.TaskName })

                $nameColWidth = $script:Width - 34
                Write-Host ("{0,-4}"                    -f "#")          -ForegroundColor $script:SubLabelFG -NoNewline
                Write-Host (" {0,-$nameColWidth} "      -f "Task Name") -ForegroundColor $script:SubLabelFG -NoNewline
                Write-Host ("{0,-8} "                   -f "State")     -ForegroundColor $script:SubLabelFG -NoNewline
                Write-Host "Next Run"                                    -ForegroundColor $script:SubLabelFG

                $i = 0
                $script:colorIndex = 0

                Write-HLine -Style dashed

                for ($i = 0; $i -lt $tasks.Count; $i++) {
                    $t = $tasks[$i]
                    $taskInfo = Get-ScheduledTaskInfo -TaskName $t.TaskName -TaskPath $t.TaskPath -ErrorAction SilentlyContinue
                    $nextRun = if ($taskInfo.NextRunTime -and $taskInfo.NextRunTime -ne [datetime]::MinValue) {
                        $taskInfo.NextRunTime.ToString("yyyy-MM-dd HH:mm:ss")
                    } else { "N/A" }

                    $displayName  = $t.TaskName
                    if ($displayName.Length -gt ($nameColWidth - 3)) { $displayName = $displayName.Substring(0, ($nameColWidth - 3)) + "..." }
                    $stateColor = if ($t.State -eq 'Ready') { $script:DoneFG } else { "Red" }

                    Write-Host (" {0,2} "   -f ($i + 1))      -ForegroundColor $script:MainTextFG -NoNewline
                    Write-Host (" {0,-$nameColWidth} " -f $displayName) -ForegroundColor $script:ValueFG    -NoNewline
                    Write-Host ("{0,-8} "   -f $t.State)     -ForegroundColor $stateColor        -NoNewline
                    Write-Host $nextRun -ForegroundColor $script:MutedFG
                }

                # Anchor the top of the footer (the dashed line) ONLY the first time
                if ($null -eq $script:MenuPos) {
                    $script:MenuPos = Get-StareCursor
                    if ($null -eq $script:MenuPos) { $script:MenuPos = [pscustomobject]@{ X = 0; Y = 0 } }
                }

                # Move to the locked anchor and draw the dashed line, then the options line below it
                Move-StareCursor $script:MenuPos.X $script:MenuPos.Y
                Write-HLine -Style dashed
                Move-StareCursor $script:MenuPos.X ($script:MenuPos.Y + 1)

                $script:RefreshRequired = $false

                Write-Host " " -NoNewline
                Write-Host "[" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host "C" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host "]" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host " Create  " -ForegroundColor $script:ValueFG -NoNewline

                Write-Host "[" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host "D" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host "]" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host " Delete  " -ForegroundColor $script:ValueFG -NoNewline

                Write-Host "[" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host "Q" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host "]" -ForegroundColor $script:MainTextFG -NoNewline
                Write-Host " Quit (or ESC)" -ForegroundColor $script:ValueFG

                Write-Host "Selection: " -ForegroundColor $script:LabelFG -NoNewline
            }

            $keyInfo = Read-StareKey
            $choice  = if ($keyInfo.Key -eq [ConsoleKey]::Escape) { "Q" } else { $keyInfo.KeyChar.ToString() }

            if ($choice.ToUpper() -eq "Q") {
                $script:mainLoop = $false
            }
            elseif ($choice.ToUpper() -eq "C") {
                $script:TaskAction      = "Create"
                $script:RefreshRequired = $true
            }
            elseif ($choice.ToUpper() -eq "D") {
                $pos = $script:MenuPos

                # Clear the dashed separator, [C][D][Q] line, and "Selection: " line
                foreach ($row in 0..2) {
                    Move-StareCursor $pos.X ($pos.Y + $row)
                    Write-Host (" " * $script:Width) -NoNewline
                }

                $waitingForDelete = $true
                while ($waitingForDelete) {
                    Move-StareCursor $pos.X ($pos.Y + 1)
                    Write-Host (" " * $script:Width) -NoNewline
                    Move-StareCursor $pos.X $pos.Y
                    Write-HLine -Style dashed

                    Write-Host "Task Number to Delete: " -ForegroundColor $script:LabelFG -NoNewline
                    $idx = Read-Esc ""

                    if ($null -eq $idx) {
                        $waitingForDelete = $false
                        $script:RefreshRequired = $false # Bypass the slow list refresh
                    }
                    elseif ($idx -match '^\d+$' -and $tasks[([int]$idx) - 1]) {
                        $delTaskName = $tasks[([int]$idx) - 1].TaskName
                        Write-Host " Delete '$delTaskName'? [Y/N]: " -ForegroundColor $script:WarnFG -NoNewline
                        $confirmKey = Read-StareKey
                        if ($confirmKey.Key -eq [ConsoleKey]::Y) {
                            Unregister-ScheduledTask -TaskName $delTaskName -Confirm:$false -ErrorAction SilentlyContinue
                            Write-Host "Deleted." -ForegroundColor $script:DoneFG
                            Start-Sleep -Seconds 1
                            $script:RefreshRequired = $true
                            $waitingForDelete = $false
                        } else {
                            Write-Host "Cancelled." -ForegroundColor $script:MutedFG
                            Start-Sleep -Milliseconds 600
                            $waitingForDelete = $false
                            $script:RefreshRequired = $false
                        }
                    }
                    else {
                        Write-Host " Invalid Selection!" -ForegroundColor $script:WarnFG -NoNewline
                        Start-Sleep -Milliseconds 800
                    }
                    # Scrub the 3rd line (where the delete prompt was)
                    Move-StareCursor $script:MenuPos.X ($script:MenuPos.Y + 2)
                    Write-Host (" " * $script:Width) -NoNewline
                }

                # Final Cleanup: Clear the dashed line and "Task Number" line before returning
                Move-StareCursor $pos.X $pos.Y
                Write-Host (" " * $script:Width) -NoNewline
                Move-StareCursor $pos.X ($pos.Y + 1)
                Write-Host (" " * $script:Width) -NoNewline
                # Put the cursor back at the anchor for the next redraw
                Move-StareCursor $script:MenuPos.X $script:MenuPos.Y
                $script:RefreshRequired = $true
            }
        }
        elseif ($script:TaskAction -eq "Create") {
            Invoke-TaskCreationWizard
            $script:TaskAction      = "List"
            $script:RefreshRequired = $true
            # Reset MenuPos so it's recalculated based on the new list length
            $script:MenuPos = $null
        }
    }

    # Erase the menu (dashed separator, options and prompt lines) before the footer
    if ($script:MenuPos) {
        foreach ($row in 0..2) {
            Move-StareCursor $script:MenuPos.X ($script:MenuPos.Y + $row)
            Write-Host (" " * $script:Width) -NoNewline
        }
        Move-StareCursor $script:MenuPos.X $script:MenuPos.Y
    }

    # Footer
    $_sfx   = "█"
    $_ftr1  = " ╔═╗ ╔╦╗ ╔═╗ ╦═╗ ╔═╗ "
    $_ftr2  = " ╚═╗  ║  ╠═╣ ╠╦╝ ║╣  "
    $_ftr3  = " ╚═╝  ╩  ╩ ╩ ╩╚═ ╚═╝ "
    $_ftrW  = [Math]::Max($_ftr1.Length, [Math]::Max($_ftr2.Length, $_ftr3.Length))
    $_ftr1  = $_ftr1.PadRight($_ftrW); $_ftr2 = $_ftr2.PadRight($_ftrW); $_ftr3 = $_ftr3.PadRight($_ftrW)
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_footer = "  TASK ADMINISTRATION SEQUENCE COMPLETE"
    $_fver   = "| $($script:Version)"
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr1 -ForegroundColor $SubtitleFG -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host $_ftr2 -ForegroundColor $SubtitleFG -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr3 -ForegroundColor $SubtitleFG -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol

    # Effective outcome. The final line always agrees with the exit code.
    $ok = $script:StareChanges
    $bad = $script:StareFailures
    if ($bad -gt 0 -and $ok -gt 0) {
        Exit-StareRun 3 "partial: $ok change(s) applied, $bad failed"
    }
    if ($bad -gt 0) {
        Exit-StareRun 1 "$bad change(s) failed, nothing was applied"
    }
    Exit-StareRun 0 "session ended, $ok change(s) applied"
}
#endregion

if (-not $global:KillerScriptsTestMode) { Invoke-StareMain }
