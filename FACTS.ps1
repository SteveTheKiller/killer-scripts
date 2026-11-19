<#
.SYNOPSIS
    Foxit Audit and Control Task Script (FACTS) v3.0
    Developed by SteveTheKiller | Updated: 2026-11-19
.DESCRIPTION
    Audits for Foxit PDF installations and permanently blocks
    auto-updates via registry hardening, service suppression, and
    scheduled task enforcement. Creates a self-healing hourly
    maintenance task when Foxit is detected.
.NOTES
    Exit Codes:
      0  - Success
      1  - Failure
      2  - Not elevated, or a precondition failed
      3  - Partial success
#>

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-FactsElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
# Console access is guarded. With output redirected (RMM, Kaseya LiveConnect) the console handle
# is invalid, so every call falls back to plain line output instead of throwing.
function Get-FactsCursorTop { try { [int][Console]::CursorTop } catch { -1 } }
function Set-FactsCursor {
    param([int]$Left, [int]$Top)
    try { [Console]::SetCursorPosition($Left, $Top); $true } catch { $false }
}
function Clear-FactsScreen { try { Clear-Host -ErrorAction Stop } catch { } }

# --- 64-bit registry helpers ---
# These read and write HKLM through the 64-bit view, so a 32-bit host (some RMM agents launch one)
# sees the same Foxit state as a 64-bit host instead of the redirected WOW6432Node view. They are
# also copied into the hourly task, so they must not depend on anything else in this file.
function Open-FactsHklm64 {
    [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
}
function Get-FactsFoxitApps {
    # Ok is false when the software list could not be read, so a read failure is never taken for "no Foxit".
    $names = New-Object System.Collections.ArrayList
    $ok = $true
    $base = $null
    try {
        $base = Open-FactsHklm64
        foreach ($p in @('SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall', 'SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) {
            $u = $base.OpenSubKey($p)
            if (-not $u) { continue }
            foreach ($n in $u.GetSubKeyNames()) {
                $k = $u.OpenSubKey($n)
                if (-not $k) { continue }
                $d = $k.GetValue('DisplayName')
                if ($d -and ([string]$d -like '*Foxit*')) { [void]$names.Add([string]$d) }
                $k.Close()
            }
            $u.Close()
        }
    } catch { $ok = $false } finally { if ($base) { $base.Close() } }
    [pscustomobject]@{ Ok = $ok; Names = @($names.ToArray()) }
}
function Add-FactsUpdaterKeys {
    param($Base, [string]$Path, $Out)
    $k = $Base.OpenSubKey($Path)
    if (-not $k) { return }
    try {
        foreach ($n in $k.GetSubKeyNames()) {
            $child = "$Path\$n"
            if ($n -eq 'Updater') { [void]$Out.Add($child) }
            Add-FactsUpdaterKeys -Base $Base -Path $child -Out $Out
        }
    } finally { $k.Close() }
}
function Get-FactsUpdaterKeys {
    # Every Updater key under the Foxit Software hives, as a path below HKLM in the 64-bit view.
    $out = New-Object System.Collections.ArrayList
    $base = $null
    try {
        $base = Open-FactsHklm64
        foreach ($hive in @('SOFTWARE\Foxit Software', 'SOFTWARE\WOW6432Node\Foxit Software')) {
            Add-FactsUpdaterKeys -Base $base -Path $hive -Out $out
        }
    } catch { } finally { if ($base) { $base.Close() } }
    @($out.ToArray())
}
function Set-FactsUpdaterValues {
    # Writes the two blocking values and reads them back. True only when both are in place.
    param([string]$Path)
    $base = $null
    try {
        $base = Open-FactsHklm64
        $k = $base.OpenSubKey($Path, $true)
        if (-not $k) { return $false }
        $k.SetValue('UpdateMode', '0', [Microsoft.Win32.RegistryValueKind]::String)
        $k.SetValue('bNoUpdate', 1, [Microsoft.Win32.RegistryValueKind]::DWord)
        $ok = ([string]$k.GetValue('UpdateMode') -eq '0') -and ([int]$k.GetValue('bNoUpdate') -eq 1)
        $k.Close()
        return $ok
    } catch { return $false } finally { if ($base) { $base.Close() } }
}

function Disable-FactsUpdateServicesAndTasks {
    # Disables every Foxit update service and scheduled task (in any task folder) and reads each one back.
    # Returns one line per problem, nothing when everything is in place.
    $issues = New-Object System.Collections.ArrayList
    foreach ($svc in @(Get-Service -ErrorAction SilentlyContinue | Where-Object { $_.Name -like '*Foxit*Update*' })) {
        try {
            Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction Stop
            Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
            $now = Get-Service -Name $svc.Name -ErrorAction Stop
            if ("$($now.StartType)" -ne 'Disabled') { [void]$issues.Add("service $($svc.Name) is still $($now.StartType)") }
        } catch { [void]$issues.Add("service $($svc.Name) could not be disabled: $($_.Exception.Message)") }
    }
    foreach ($t in @(Get-ScheduledTask -TaskName '*Foxit*' -ErrorAction SilentlyContinue | Where-Object { $_.TaskName -like '*Update*' })) {
        $full = "$($t.TaskPath)$($t.TaskName)"
        try {
            Stop-ScheduledTask -TaskPath $t.TaskPath -TaskName $t.TaskName -ErrorAction SilentlyContinue
            Disable-ScheduledTask -TaskPath $t.TaskPath -TaskName $t.TaskName -ErrorAction Stop | Out-Null
            $now = Get-ScheduledTask -TaskPath $t.TaskPath -TaskName $t.TaskName -ErrorAction Stop
            if ("$($now.State)" -ne 'Disabled') { [void]$issues.Add("task $full is still $($now.State)") }
        } catch { [void]$issues.Add("task $full could not be disabled: $($_.Exception.Message)") }
    }
    @($issues.ToArray())
}

# The hourly task runs the same functions the interactive pass uses, so the two cannot drift apart.
# The text is captured here, when the file loads, so it is the real function bodies.
$script:FactsLibraryNames = @('Open-FactsHklm64', 'Get-FactsFoxitApps', 'Add-FactsUpdaterKeys', 'Get-FactsUpdaterKeys', 'Set-FactsUpdaterValues', 'Disable-FactsUpdateServicesAndTasks')
$script:FactsLibrary = (($script:FactsLibraryNames | ForEach-Object {
    "function $_ {`n" + (Get-Command $_ -CommandType Function).ScriptBlock.ToString() + "`n}"
}) -join "`n")
function Get-FactsMaintenanceScript {
    $script:FactsLibrary + @'

$bad = 0
$audit = Get-FactsFoxitApps
if ($audit.Ok -and @($audit.Names).Count -gt 0) {
    foreach ($p in @(Get-FactsUpdaterKeys)) { if (-not (Set-FactsUpdaterValues -Path $p)) { $bad++ } }
    $bad += @(Disable-FactsUpdateServicesAndTasks).Count
}
if ($bad -gt 0) { exit 3 }
exit 0
'@
}

$script:StepRow = -1
$script:StepMsg = ""
$script:Width   = 85

$LineCol   = "DarkMagenta"
$WarnCol   = "DarkYellow"
$BorderCol = "Magenta"
$ArtCol    = "Yellow"
$AccentCol = "Green"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"
$_ver      = "| v3.0"

function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)

    if ($Success) {
        # Complete: go back to saved row, rewrite in DimCol + WarnCol CustomInfo + right-aligned [SUCCESS]
        $combined = if ($CustomInfo) { "$($script:StepMsg) $CustomInfo" } else { $script:StepMsg }
        $_pad = " " * [math]::Max(1, $script:Width - $combined.Length - "[SUCCESS]".Length)
        $savedTop = Get-FactsCursorTop
        $inPlace = ($savedTop -ge 0) -and ($script:StepRow -ge 0) -and ($script:StepRow -lt $savedTop)
        if ($inPlace) { $inPlace = Set-FactsCursor 0 $script:StepRow }
        Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
        if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
        Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
        if ($inPlace) { [void](Set-FactsCursor 0 $savedTop) }
        $script:StepRow = -1
    } elseif ($Message -and -not $CustomInfo) {
        # Announcement: save row, print in InfoCol
        $script:StepRow = Get-FactsCursorTop
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

$TaskName = "FACTS_Universal_Enforcement"

function Invoke-FactsMain {
    #region Pre-Flight Checks
    # ============================================================================
    if (-not (Test-FactsElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Write-Output "[FAIL] FACTS not elevated, nothing was changed (exit 2)"
        Exit-KillerScript 2
    }

    Clear-FactsScreen
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $BorderCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    #endregion

    #region 1. Audit & Discovery
    # ============================================================================
    Write-StepUpdate "[1/4] Auditing Foxit Installations..."
    $Audit = Get-FactsFoxitApps
    if (-not $Audit.Ok) {
        Write-Host " [FAILED]" -ForegroundColor Red
        Write-Output "[FAIL] FACTS could not read the installed software list, nothing was changed (exit 1)"
        Exit-KillerScript 1
    }
    $FoundApps = @($Audit.Names)
    if ($FoundApps.Count -gt 0) {
        Write-StepUpdate -Success -CustomInfo "($($FoundApps.Count) Found)"
    } else {
        Write-StepUpdate -Success -CustomInfo "(None Detected)"
    }
    #endregion

    $Problems = New-Object System.Collections.ArrayList

    #region 2. Universal Registry Hardening
    # ============================================================================
    if ($FoundApps.Count -gt 0) {
        Write-StepUpdate "[2/4] Applying Recursive Registry Blocks..."
        foreach ($Key in @(Get-FactsUpdaterKeys)) {
            if (-not (Set-FactsUpdaterValues -Path $Key)) { [void]$Problems.Add("registry value did not read back: HKLM\$Key") }
        }
        Write-StepUpdate -Success
    } else {
        Write-StepUpdate "[2/4] Skipping Registry Hardening..." -CustomInfo "(Not Required)"
    }
    #endregion

    #region 3. Service & Task Enforcement
    # ============================================================================
    if ($FoundApps.Count -gt 0) {
        Write-StepUpdate "[3/4] Disabling Native Services and Tasks..."
        foreach ($Issue in @(Disable-FactsUpdateServicesAndTasks)) { [void]$Problems.Add($Issue) }
        Write-StepUpdate -Success
    } else {
        Write-StepUpdate "[3/4] Skipping Service/Task Block..." -CustomInfo "(Not Required)"
    }
    #endregion

    #region 4. Hourly Self-Healing Task
    # ============================================================================
    if ($FoundApps.Count -gt 0) {
        Write-StepUpdate "[4/4] Configuring Persistence (Hourly Task)..."
        $EncodedCommand = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes((Get-FactsMaintenanceScript)))
        try {
            $Action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-NoProfile -WindowStyle Hidden -EncodedCommand $EncodedCommand"
            $Trigger = New-ScheduledTaskTrigger -Once -At (Get-Date) -RepetitionInterval (New-TimeSpan -Hours 1)
            # -Force replaces an existing task in one step, so a failed update keeps the installed task.
            Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -User "SYSTEM" -RunLevel Highest -Force -ErrorAction Stop | Out-Null
            $Back = Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue
            if (-not $Back) {
                [void]$Problems.Add("maintenance task $TaskName was registered but could not be read back")
            } elseif ("$($Back.State)" -eq 'Disabled') {
                [void]$Problems.Add("maintenance task $TaskName is registered but disabled")
            }
            Write-StepUpdate -Success
        }
        catch {
            Write-Host " [FAILED]" -ForegroundColor Red
            Write-Output "[FAIL] FACTS could not register the maintenance task: $($_.Exception.Message) (exit 1)"
            Exit-KillerScript 1
        }
    } else {
        Write-StepUpdate "[4/4] Removing existing Maintenance Tasks..."
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) {
            try { Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop }
            catch { [void]$Problems.Add("maintenance task $TaskName could not be removed: $($_.Exception.Message)") }
        }
        Write-StepUpdate -Success -CustomInfo "(No Foxit Detected)"
    }
    # Footer (reuses header art - art on right, fill on left)
    $_ffillW = $script:Width - $_artW - 2  # 1 leading space + 1 sfx char
    $_footer = if ($Problems.Count -gt 0) { "  FOXIT MAINTENANCE ENFORCEMENT INCOMPLETE" } else { "  FOXIT MAINTENANCE ENFORCEMENT COMPLETE" }
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $BorderCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol

    if ($Problems.Count -gt 0) {
        foreach ($Issue in $Problems) { Write-Output "[FAILED] $Issue" }
        Write-Output ("[WARN] FACTS partial: {0} problem(s), see above (exit 3)" -f $Problems.Count)
        Exit-KillerScript 3
    }
    if ($FoundApps.Count -gt 0) {
        Write-Output ("[PASS] FACTS complete: {0} Foxit install(s) blocked, hourly task in place (exit 0)" -f $FoundApps.Count)
    } else {
        Write-Output "[PASS] FACTS complete: no Foxit detected, maintenance task removed (exit 0)"
    }
    Exit-KillerScript 0
    #endregion
}

if (-not $global:KillerScriptsTestMode) { Invoke-FactsMain }
