<#
.SYNOPSIS
    Deployment & Endpoint Provisioning Operations Tool (D.E.P.O.T.) v2.0
    Developed by SteveTheKiller | Updated: 2026-11-18
.DESCRIPTION
    Full new-machine provisioning tool that silently deploys M365,
    Teams, OneDrive, Chrome, Acrobat Reader, Zoom, and 7-Zip with
    'ESC-to-skip' support. Also runs Windows Update, applies privacy/UI
    hardening across all profiles, and self-deletes on completion.
#>

# Captured at script level: $MyInvocation.MyCommand.Path is empty inside a function.
$script:DepotScriptPath = $MyInvocation.MyCommand.Path

$script:Width     = 85
$script:DepotTemp = 'C:\Windows\Temp'
$script:OneDriveVerifySeconds = 30
$script:LineCol   = "DarkRed"
$script:MainCol   = "Green"
$script:AccentCol = "Yellow"
$script:DimCol    = "DarkGray"

# Header art, shared by the banner and the footer
$script:ArtPfx = "█  "
$script:Art1   = "╔╦╗ ╔═╗ ╔═╗ ╔═╗ ╔╦╗ "
$script:Art2   = " ║║ ║╣  ╠═╝ ║ ║  ║  "
$script:Art3   = "╩╩╝ ╚═╝ ╩   ╚═╝  ╩  "
$script:ArtW   = [Math]::Max($script:Art1.Length, [Math]::Max($script:Art2.Length, $script:Art3.Length))
$script:Art1   = $script:Art1.PadRight($script:ArtW)
$script:Art2   = $script:Art2.PadRight($script:ArtW)
$script:Art3   = $script:Art3.PadRight($script:ArtW)

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# Well-known SID of the local Administrators group. Group names are localized, so use the SID.
$script:AdminsGroupSid = 'S-1-5-32-544'

function Reset-DepotState {
    $script:RebootRequired = $false
    $script:UserHives      = @()
    $script:DepotResults   = @()
}

function Add-DepotResult {
    # Records one step outcome. Status is Done, AlreadyInstalled, Skipped or Failed.
    param([string]$Step, [ValidateSet('Done', 'AlreadyInstalled', 'Skipped', 'Failed')][string]$Status, [string]$Detail = '')
    $script:DepotResults = @($script:DepotResults) + [pscustomobject]@{ Step = $Step; Status = $Status; Detail = $Detail }
}

function Get-DepotCurrentSid {
    try { return [Security.Principal.WindowsIdentity]::GetCurrent().User.Value } catch { return $null }
}

function Write-StatusTag {
    param([string]$Text, [string]$Color)
    try {
        $cur = [Console]::CursorLeft
        $pad = $script:Width - $cur - $Text.Length
        if ($pad -gt 0) { Write-Host (" " * $pad) -NoNewline }
    } catch { Write-Host " " -NoNewline }
    Write-Host $Text -ForegroundColor $Color
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
        [ConsoleColor]$script:LineCol,
        [ConsoleColor]$script:AccentCol,
        [ConsoleColor]$script:AccentCol,
        [ConsoleColor]$script:DimCol
    )
    $useConsole = $true
    try { $saved = [Console]::ForegroundColor } catch { $useConsole = $false }
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') {
            $fg = [ConsoleColor]$script:DimCol
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

function Get-Installer {
    param([string]$Url, [string]$Dest)
    try {
        $wc = New-Object System.Net.WebClient
        $wc.Headers.Add("User-Agent","Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120.0.0.0 Safari/537.36")
        $wc.DownloadFile($Url, $Dest)
        $item = Get-Item $Dest -ErrorAction SilentlyContinue
        if (-not $item -or $item.Length -lt 1MB) { return $false }
        # Validate first 2 bytes - MZ = EXE/MSI, D0CF = legacy MSI/CAB
        $bytes = [System.IO.File]::ReadAllBytes($Dest)
        $isMZ   = ($bytes[0] -eq 0x4D -and $bytes[1] -eq 0x5A)
        $isCF   = ($bytes[0] -eq 0xD0 -and $bytes[1] -eq 0xCF)
        if ($isMZ -or $isCF) {
            Write-Host " ($([math]::Round($item.Length/1MB,1)) MB)" -NoNewline -ForegroundColor DarkGray
            return $true
        }
        return $false
    } catch { return $false }
}
# Helper: check if app is already installed by display name pattern
function Test-AppInstalled {
    param([string]$Pattern)
    $paths = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $found = Get-ItemProperty $paths -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match $Pattern }
    return ($null -ne $found)
}
function Stop-DepotProcessTree {
    # Stops a process and everything it started, children first. Walks Win32_Process parent links
    # from the given id, so it never touches a process outside that tree.
    param([int]$RootId)
    $rows  = @(Get-DepotProcessInfo)
    $order = New-Object System.Collections.ArrayList
    $queue = @($RootId)
    while ($queue.Count -gt 0) {
        $cur = [int]$queue[0]
        $queue = @($queue | Select-Object -Skip 1)
        if ($order.Contains($cur)) { continue }
        [void]$order.Add($cur)
        foreach ($r in $rows) {
            if ([int]$r.ParentProcessId -eq $cur -and [int]$r.ProcessId -ne $cur) { $queue += [int]$r.ProcessId }
        }
    }
    for ($i = $order.Count - 1; $i -ge 0; $i--) {
        Stop-Process -Id $order[$i] -Force -ErrorAction SilentlyContinue
    }
}

function Test-DepotEscPressed {
    # True when the technician pressed ESC. False with no console, where there is no key to read.
    try {
        if ($Host.UI.RawUI.KeyAvailable) {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ($key.VirtualKeyCode -eq 27) { return $true }
        }
    } catch { }
    return $false
}

function Get-DepotProcessInfo {
    # Win32_Process rows (ProcessId, ParentProcessId, Name, ExecutablePath, CommandLine). Empty on failure.
    try {
        return @(Get-CimInstance -ClassName Win32_Process -Property ProcessId, ParentProcessId, Name, ExecutablePath, CommandLine -ErrorAction Stop)
    } catch { return @() }
}

# Helper: run an installer with ESC-to-cancel and a timeout. Returns the child's own exit code,
# $null when the technician cancelled with ESC, -1 on timeout and -2 when the exit code could not be read.
# The code comes from the process object we waited on, never from $LASTEXITCODE, which a
# Start-Process launch does not set.
function Start-Installer {
    param(
        [string]$FilePath,
        [string]$ArgumentList,
        [double]$TimeoutMinutes = 20
    )
    $parentDir = Split-Path $FilePath -Parent -ErrorAction SilentlyContinue
    $workDir   = if ($parentDir -and (Test-Path $parentDir)) { $parentDir } else { $script:DepotTemp }
    $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -WindowStyle Hidden -WorkingDirectory $workDir
    if (-not $proc) { return -2 }
    # Reading Handle keeps the process handle open so ExitCode is still readable after exit.
    try { $null = $proc.Handle } catch { }
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $start    = Get-Date
    $spinner  = @('|','/','-','\')
    $spin     = 0
    while (-not $proc.WaitForExit(500)) {
        if ((Get-Date) -gt $deadline) {
            # Only the process this helper launched and its descendants. Other installers and the
            # technician's own browsers or archivers on the machine are left alone.
            Stop-DepotProcessTree -RootId $proc.Id
            Write-Host " [TIMEOUT after $TimeoutMinutes min]" -ForegroundColor Red
            return -1
        }
        if (Test-DepotEscPressed) {
            Stop-DepotProcessTree -RootId $proc.Id
            return $null
        }
        $elapsed  = [int]((Get-Date) - $start).TotalSeconds
        $spinChar = $spinner[$spin % 4]; $spin++
        Write-Host "`r      $spinChar $([math]::Floor($elapsed/60))m$($elapsed % 60)s" -NoNewline -ForegroundColor DarkGray
    }
    Write-Host "`r                    `r" -NoNewline
    if ($null -eq $proc.ExitCode) { return -2 }
    return [int]$proc.ExitCode
}
function Write-DepotStatus {
    # Prints the status tag and records the same outcome for the current step, so the final exit code,
    # the footer and the run log all come from what actually happened.
    param([string]$Status, [string]$Tag, [string]$Color, [string]$Detail = '')
    Write-StatusTag $Tag $Color
    $step = if ($script:DepotStep) { $script:DepotStep } else { 'Unnamed step' }
    if (-not $Detail) { $Detail = $Tag.Trim('[', ']') }
    Add-DepotResult $step $Status $Detail
}

function Show-Result {
    param($ExitCode)
    if ($null -eq $ExitCode)                          { Write-DepotStatus Skipped "[SKIPPED]" "Yellow" "skipped by technician" }
    elseif ($ExitCode -in 0, 3010)                    {
        Write-DepotStatus Done "[DONE]" "Green" "exit $ExitCode"
        if ($ExitCode -eq 3010) { $script:RebootRequired = $true }
    }
    elseif ($ExitCode -in 1638, 1641, -1978335189)    {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan" "exit $ExitCode"
        if ($ExitCode -eq 1641) { $script:RebootRequired = $true }
    }
    elseif ($ExitCode -eq -2)                         { Write-DepotStatus Failed "[FAILED - exit code unavailable]" "Red" "exit code unavailable" }
    elseif ($ExitCode -eq -1)                         { Write-DepotStatus Failed "[FAILED - exit -1]" "Red" "timed out or killed" }
    else                                              { Write-DepotStatus Failed "[FAILED - exit $ExitCode]" "Red" "exit $ExitCode" }
}

function Write-Installing {
    Write-Host " installing..." -NoNewline -ForegroundColor DarkGray
}

function Set-RegValue {
    param([string]$Path, [string]$Name, $Value, [string]$Type = "DWord")
    try {
        if (-not (Test-Path $Path)) { New-Item -Path $Path -Force -ErrorAction Stop | Out-Null }
        Set-ItemProperty -Path $Path -Name $Name -Value $Value -Type $Type -ErrorAction SilentlyContinue
    } catch { <# Key creation failed in mounted hive - skip silently #> }
}
function Set-AllUsers {
    param([string]$SubKey, [string]$Name, $Value, [string]$Type = "DWord")
    foreach ($Hive in $script:UserHives) {
        Set-RegValue -Path "$($Hive.Mount)\$SubKey" -Name $Name -Value $Value -Type $Type
    }
    # Also set for current user in case hive loading missed it
    Set-RegValue -Path "HKCU:\$SubKey" -Name $Name -Value $Value -Type $Type
}

function Test-DepotElevated {
    $Identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
    $Principal = [Security.Principal.WindowsPrincipal]$Identity
    if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) { return $false }
    # Confirm token is fully elevated (handles domain UAC split-token scenarios)
    $TokenElevation = $Identity.Claims |
        Where-Object { $_.Type -eq "http://schemas.microsoft.com/ws/2008/06/identity/claims/groupsid" -and $_.Value -eq "S-1-5-32-544" }
    return [bool]$TokenElevation
}

function Invoke-DepotElevationRelaunch {
    if (-not $script:DepotScriptPath) {
        Write-Host "[!] Not running as Administrator and there is no script file to relaunch. Run from an elevated session." -ForegroundColor Red
        Exit-KillerScript 2
    }
    Write-Host "[!] Not running as Administrator. Relaunching elevated..." -ForegroundColor Yellow
    $relaunched = $false
    try {
        Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$script:DepotScriptPath`"" -Verb RunAs -ErrorAction Stop
        $relaunched = $true
    } catch {
        Write-Host "[!] Could not relaunch elevated: $($_.Exception.Message)" -ForegroundColor Red
    }
    if (-not $relaunched) { Exit-KillerScript 2 }
    Write-Host "[>] Elevated copy launched in a new window. This window can be closed." -ForegroundColor Yellow
    Exit-KillerScript 0
}

function Invoke-DepotReg {
    # Thin wrapper over reg.exe so the registry-hive calls can be mocked. Returns the exit code and output.
    param([string[]]$Arguments)
    $out = & reg @Arguments 2>&1
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = @($out | ForEach-Object { "$_" }) }
}

function Test-DepotOwnedProcess {
    # True only for a process started from a file DEPOT itself downloaded: the staged installer, the
    # ODT folder, or an msiexec whose command line names one of the staged files.
    param($Info)
    if ([int]$Info.ProcessId -eq $PID) { return $false }
    $temp = $script:DepotTemp.TrimEnd('\') + '\'
    $staged = @('ODT.exe', 'TeamsSetup.exe', 'ChromeEnterprise.msi', 'AcrobatReader.exe', '7zip.exe', 'OneDriveSetup.exe', 'ZoomInstaller.msi')
    $exe = [string]$Info.ExecutablePath
    if ($exe) {
        if ($exe.StartsWith($temp + 'ODT\', [StringComparison]::OrdinalIgnoreCase)) { return $true }
        foreach ($n in $staged) { if ($exe.Equals($temp + $n, [StringComparison]::OrdinalIgnoreCase)) { return $true } }
    }
    $cmd = [string]$Info.CommandLine
    if ($cmd) {
        foreach ($n in $staged) { if ($cmd.IndexOf($temp + $n, [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true } }
        if ($cmd.IndexOf($temp + 'ODT\', [StringComparison]::OrdinalIgnoreCase) -ge 0) { return $true }
    }
    return $false
}

function Stop-DepotLeftoverProcess {
    # Stop leftovers of a previous cancelled or crashed run, and only those: processes running from
    # DEPOT's own staged files. Processes that merely share a name (chrome, setup, msiexec, Teams...) are
    # the technician's or Windows' and are left alone.
    $killed = @()
    foreach ($info in (Get-DepotProcessInfo)) {
        if (Test-DepotOwnedProcess $info) {
            Stop-Process -Id ([int]$info.ProcessId) -Force -ErrorAction SilentlyContinue
            $killed += $info.Name
        }
    }
    if ($killed.Count -gt 0) {
        Write-Host "    [!] Killed: $($killed -join ', ')" -ForegroundColor Yellow
    } else {
        Write-Host "    > No stale processes found." -ForegroundColor DarkGray
    }
}

function Remove-DepotOrphanHive {
    # Unload any APT registry hives left mounted from a previous run
    $mountedHives = (Invoke-DepotReg @('query', 'HKU')).Output | Where-Object { $_ -match "APT_" }
    foreach ($hive in $mountedHives) {
        $hiveName = ($hive -split "\\")[-1].Trim()
        $null = Invoke-DepotReg @('unload', "HKU\$hiveName")
        Write-Host "    [!] Unloaded orphaned hive: $hiveName" -ForegroundColor Yellow
    }
}

function Remove-DepotStaleFile {
    # Wipe all temp files this script may have created
    $staleFiles = @(
        "C:\Windows\Temp\ODT.exe",
        "C:\Windows\Temp\ODT",
        "C:\Windows\Temp\TeamsSetup.exe",
        "C:\Windows\Temp\ChromeEnterprise.msi",
        "C:\Windows\Temp\AcrobatReader.exe",
        "C:\Windows\Temp\7zip.exe",
        "C:\Windows\Temp\WU_IDs.txt",
        "C:\Windows\Temp\WU_Install.ps1",
        "C:\Windows\Temp\WU_Result.txt",
        "C:\Windows\Temp\WG_*.ps1",
        "C:\Windows\Temp\WG_*.txt",
        "C:\Windows\Temp\Winget.msix*"
    )
    foreach ($f in $staleFiles) {
        Remove-Item $f -Recurse -Force -ErrorAction SilentlyContinue
    }
}

function Invoke-DepotPreflight {
    # --- STARTUP CLEANUP: kill leftovers from any previous cancelled/crashed run ---
    Write-Host "[*] Checking for leftover processes from previous run..." -ForegroundColor DarkGray
    Stop-DepotLeftoverProcess
    # Stop Windows Update service if left running
    $wuSvc = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
    if ($wuSvc -and $wuSvc.Status -eq 'Running') {
        Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
        Write-Host "    [!] Stopped leftover wuauserv." -ForegroundColor Yellow
    }
    Remove-DepotOrphanHive
    Remove-DepotStaleFile
}

function Test-DepotInteractive {
    # True only when a person can answer a prompt: a real console host with keyboard input. RMM consoles
    # (LiveConnect), -EncodedCommand runs, redirected input and SYSTEM sessions are not interactive.
    try {
        if (-not [Environment]::UserInteractive) { return $false }
        if ([Console]::IsInputRedirected) { return $false }
        return ($Host.Name -eq 'ConsoleHost')
    } catch { return $false }
}

function Clear-DepotScreen {
    try { Clear-Host } catch { }
}

function Read-DepotChoice {
    param([string]$Prompt)
    if (-not (Test-DepotInteractive)) {
        # Every prompt defaults to N, so an unattended run takes the safe answer.
        Write-Host "$Prompt[no console, defaulting to N]" -ForegroundColor DarkGray
        return 'N'
    }
    Write-Host $Prompt -ForegroundColor Black -BackgroundColor Red -NoNewline
    $_k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    $choice = $_k.Character.ToString()
    Write-Host $choice
    return $choice
}

function Invoke-DepotRename {
    $script:DepotStep = 'Rename'
    $CurrentName = $env:COMPUTERNAME
    $NewName = ("$(Read-Host "   > New Device Name")").Trim()
    # NetBIOS rules: 1 to 15 letters, digits or hyphens, not all digits, no leading or trailing hyphen.
    $valid = ($NewName -match '^[A-Za-z0-9]([A-Za-z0-9-]{0,13}[A-Za-z0-9])?$') -and ($NewName -notmatch '^\d+$')
    if ($valid -and $NewName -ne $CurrentName) {
        Write-Host "[>] Renaming computer from $CurrentName to $NewName..." -NoNewline
        try {
            Rename-Computer -NewName $NewName -Force -ErrorAction Stop
            Write-Host " Done." -ForegroundColor Green
            Write-Host "[!] Note: Name change requires a reboot to take effect." -ForegroundColor Red
            $script:RebootRequired = $true
            Add-DepotResult 'Rename' 'Done' "renamed to $NewName, reboot needed"
        } catch {
            Write-Host " Failed: $($_.Exception.Message)" -ForegroundColor Red
            Add-DepotResult 'Rename' 'Failed' "rename to $NewName failed: $($_.Exception.Message)"
        }
    } else {
        Write-Host "[!] Invalid name or same as current. Skipping." -ForegroundColor Gray
    }
}

function New-DepotLocalAdmin {
    # Returns the new account's SID once it exists, is enabled and is in Administrators. Otherwise $null.
    $NewAdminName = Read-Host "   > Local Admin Username"
    $Password = Read-Host "   > Local Admin Password" -AsSecureString
    if ($NewAdminName -and $Password -and $Password.Length -gt 0) {
        Write-Host "[>] Creating local account: $NewAdminName..." -NoNewline
        try {
            $created = New-LocalUser -Name $NewAdminName -Password $Password -Description "Local Admin created" -ErrorAction Stop
            Add-LocalGroupMember -SID $script:AdminsGroupSid -Member $NewAdminName -ErrorAction Stop
            $check = Get-LocalUser -SID $created.SID -ErrorAction Stop
            if (-not $check -or -not $check.Enabled) { throw "The new account is not enabled." }
            Write-Host " Done." -ForegroundColor Green
            Add-DepotResult "Create local admin $NewAdminName" 'Done'
            return $created.SID
        } catch {
            Write-Host " Failed." -ForegroundColor Red
            Write-Warning $_.Exception.Message
            Add-DepotResult "Create local admin $NewAdminName" 'Failed' $_.Exception.Message
            return $null
        }
    }
    Write-Host "[!] Username or Password cannot be empty. Skipping." -ForegroundColor Gray
    return $null
}

function Invoke-DepotAdminPurge {
    # Disables the built-in Administrator and deletes other local administrators, all by SID, and only
    # when the account created earlier in this run is verified as an enabled administrator.
    param($SurvivorSid)
    $step = "Purge administrators"
    Write-Host "[>] Verifying a usable administrator will remain..." -NoNewline
    try {
        $users   = @(Get-LocalUser -ErrorAction Stop)
        $members = @(Get-LocalGroupMember -SID $script:AdminsGroupSid -ErrorAction Stop)
    } catch {
        Write-Host " Failed." -ForegroundColor Red
        Write-Host "[!] Could not read local accounts, so nothing was disabled or deleted: $($_.Exception.Message)" -ForegroundColor Red
        Add-DepotResult $step 'Failed' "Could not read local accounts: $($_.Exception.Message)"
        return
    }
    $builtIn = $users | Where-Object { $_.SID.Value -match '^S-1-5-21-[\d-]+-500$' } | Select-Object -First 1
    $survivor = $null
    if ($SurvivorSid) {
        $survivor = $users | Where-Object { $_.SID.Value -eq "$SurvivorSid" -and $_.Enabled } | Select-Object -First 1
        if ($survivor -and -not ($members | Where-Object { $_.SID.Value -eq "$SurvivorSid" })) { $survivor = $null }
    }
    if (-not $builtIn -or -not $survivor -or $survivor.SID.Value -eq $builtIn.SID.Value) {
        Write-Host " Refused." -ForegroundColor Red
        Write-Host "[!] No new administrator was created and verified in this run, so the built-in Administrator and other admins were left alone." -ForegroundColor Red
        Add-DepotResult $step 'Failed' 'No verified administrator would remain'
        return
    }
    Write-Host " OK ($($survivor.Name))." -ForegroundColor Green

    $failed = 0
    Write-Host "[>] Disabling built-in Administrator account ($($builtIn.Name))..." -NoNewline
    try {
        if ($builtIn.Enabled) {
            Disable-LocalUser -SID $builtIn.SID -ErrorAction Stop
            $after = Get-LocalUser -SID $builtIn.SID -ErrorAction Stop
            if ($after.Enabled) { throw "The account is still enabled." }
        }
        Write-Host " Done." -ForegroundColor Green
        Write-Host "    Recovery: Enable-LocalUser -SID $($builtIn.SID.Value)" -ForegroundColor DarkGray
    } catch {
        Write-Host " Failed." -ForegroundColor Red
        Write-Warning $_.Exception.Message
        $failed++
        Add-DepotResult "Disable built-in Administrator" 'Failed' $_.Exception.Message
    }

    Write-Host "[>] Scanning for other local administrators to delete..." -NoNewline
    $machinePrefix = ($builtIn.SID.Value -replace '-500$', '') + '-'
    $currentSid = Get-DepotCurrentSid
    $AdminsToPurge = @($members | Where-Object {
        $_.ObjectClass -eq "User" -and
        $_.SID.Value.StartsWith($machinePrefix) -and
        $_.SID.Value -ne $builtIn.SID.Value -and
        $_.SID.Value -ne $survivor.SID.Value -and
        $_.SID.Value -ne $currentSid
    })
    if ($AdminsToPurge.Count -gt 0) {
        Write-Host " Found $($AdminsToPurge.Count)." -ForegroundColor Cyan
        foreach ($Admin in $AdminsToPurge) {
            Write-Host "   [!] Deleting account: $($Admin.Name)..." -NoNewline
            try {
                Remove-LocalUser -SID $Admin.SID -ErrorAction Stop
                if (Get-LocalUser -SID $Admin.SID -ErrorAction SilentlyContinue) { throw "The account still exists." }
                Write-Host " Done." -ForegroundColor Green
            } catch {
                Write-Host " Failed." -ForegroundColor Red
                Write-Warning $_.Exception.Message
                $failed++
                Add-DepotResult "Delete account $($Admin.Name)" 'Failed' $_.Exception.Message
            }
        }
    } else {
        Write-Host " None found." -ForegroundColor Green
    }
    if ($failed -eq 0) { Add-DepotResult $step 'Done' }
}

function Invoke-DepotInteractiveSetup {
    Clear-DepotScreen
    Write-Host "[*] Interactive Pre-Deployment Setup..." -ForegroundColor DarkRed -BackgroundColor Black
    Write-Host ""
    Write-Host "    Note: All prompts default to 'N' if Enter is pressed." -ForegroundColor DarkRed -BackgroundColor Black
    Write-Host ""
    # 1.1: Machine Renaming
    $RenameChoice = Read-DepotChoice "1) Do you want to rename the machine? (y/N) "
    if ($RenameChoice -eq "Y" -or $RenameChoice -eq "y") { Invoke-DepotRename }
    # 1.2: Local Admin Creation
    Write-Host ""
    $AdminChoice = Read-DepotChoice "2) Do you want to make a local admin account? (y/N) "
    $NewAdminSid = $null
    if ($AdminChoice -eq "Y" -or $AdminChoice -eq "y") { $NewAdminSid = New-DepotLocalAdmin }
    # 1.3: Purge & Secure Admin Accounts
    Write-Host ""
    $PurgeChoice = Read-DepotChoice "3) Do you want to purge other admins and disable built-in Administrator? (y/N) "
    if ($PurgeChoice -eq "Y" -or $PurgeChoice -eq "y") { Invoke-DepotAdminPurge -SurvivorSid $NewAdminSid }
}

function Show-DepotBanner {
    Clear-DepotScreen
    $_fillW = $script:Width - $script:ArtPfx.Length - $script:ArtW
    $_title = "DEPLOYMENT & ENDPOINT PROVISIONING OPERATIONS TOOL"
    $_tpad  = " " * ($_fillW - $_title.Length)
    Write-Host $script:ArtPfx -ForegroundColor $script:LineCol -NoNewline; Write-Host $script:Art1 -ForegroundColor $script:AccentCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $script:LineCol
    Write-Host $script:ArtPfx -ForegroundColor $script:LineCol -NoNewline; Write-Host $script:Art2 -ForegroundColor $script:AccentCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $script:MainCol
    Write-Host $script:ArtPfx -ForegroundColor $script:LineCol -NoNewline; Write-Host $script:Art3 -ForegroundColor $script:AccentCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $script:LineCol
    Write-Host "[>] Device Name    : " -NoNewline -ForegroundColor $script:LineCol
    Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $script:AccentCol
    # By SID: the group name is localized. A host without the LocalAccounts module (32-bit) shows no list.
    try {
        $Admins = Get-LocalGroupMember -SID $script:AdminsGroupSid -ErrorAction Stop | Where-Object { $_.ObjectClass -eq "User" }
    } catch {
        $Admins = $null
    }
    $AdminPrefix = "[>] Local Admins   : "
    Write-Host $AdminPrefix -NoNewline -ForegroundColor $script:LineCol
    $AdminList = $Admins.Name
    if ($AdminList -is [string]) { $AdminList = @($AdminList) }
    $CurrentLine = ""; $MaxLen = $script:Width - $AdminPrefix.Length
    foreach ($Name in $AdminList) {
        $Item = if ($Name -eq $AdminList[-1]) { $Name } else { "$Name, " }
        if (($CurrentLine.Length + $Item.Length) -gt $MaxLen) {
            Write-Host $CurrentLine -ForegroundColor $script:AccentCol
            Write-Host (" " * $AdminPrefix.Length) -NoNewline; $CurrentLine = $Item
        } else { $CurrentLine += $Item }
    }
    if ($CurrentLine) { Write-Host $CurrentLine -ForegroundColor $script:AccentCol }
    Write-HLine -Style "dashed"
}

function Invoke-DepotWindowsUpdate {
    Write-Host "[*] Deploying Windows & Microsoft Updates..." -ForegroundColor DarkRed
    $script:DepotStep = "Windows Update"

    try {
        # Reset WU service state first to clear any stuck session
        Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
        Start-Service wuauserv -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3

        # Kick off the full WU cycle via Update Session Orchestrator - fire and forget
        Start-Process "UsoClient.exe" -ArgumentList "StartScan"     -WindowStyle Hidden -ErrorAction Stop
        Start-Sleep -Seconds 2
        Start-Process "UsoClient.exe" -ArgumentList "StartDownload" -WindowStyle Hidden -ErrorAction Stop
        Start-Sleep -Seconds 2
        Start-Process "UsoClient.exe" -ArgumentList "StartInstall"  -WindowStyle Hidden -ErrorAction Stop
        $script:RebootRequired = $true
        Add-DepotResult $script:DepotStep "Done" "triggered"

        Write-Host "      > Windows Update triggered in background." -ForegroundColor Gray
        Write-Host "      > Updates will continue in background after script exit." -ForegroundColor Gray
    } catch {
        Write-Host "      > Failed to trigger Windows Update: $_" -ForegroundColor Red
        Add-DepotResult $script:DepotStep "Failed" "UsoClient trigger failed: $_"
        Write-Host "      > Run Windows Update manually after reboot." -ForegroundColor Yellow
    }
}

function Get-DepotWebText {
    param([string]$Url)
    $wc = New-Object System.Net.WebClient
    $wc.Headers.Add("User-Agent","Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
    return $wc.DownloadString($Url)
}

function Test-DepotMsiOffice {
    # True when a Windows Installer (MSI) Office product is installed. The Click-to-Run support
    # components are MSI packages too, so they are excluded by name.
    $uninstall = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )
    $msi = Get-ItemProperty $uninstall -ErrorAction SilentlyContinue | Where-Object {
        $_.WindowsInstaller -eq 1 -and
        $_.DisplayName -match '^Microsoft (Office|Visio|Project|Word|Excel|PowerPoint|Outlook|Access|OneNote|Publisher|InfoPath|SharePoint Designer|Skype for Business|Lync)' -and
        $_.DisplayName -notmatch 'Click-to-Run'
    }
    if ($msi) { return $true }
    # MSI Office keeps its files in Office11 to Office16, outside the Click-to-Run "root" folder.
    foreach ($pf in @($env:ProgramFiles, ${env:ProgramFiles(x86)})) {
        if ($pf -and (Test-Path (Join-Path $pf 'Microsoft Office\Office1?\WINWORD.EXE'))) { return $true }
    }
    return $false
}

function Install-DepotOffice {
    $script:DepotStep = 'Office'
    # --- Microsoft 365 via Office Deployment Tool ---
    Write-Host "      > Installing: Microsoft 365..." -NoNewline -ForegroundColor Gray
    $OfficeInstalled = (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") -or
                       (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE") -or
                       (Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue)
    if ($OfficeInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $ODTDir  = Join-Path $script:DepotTemp "ODT"
        $ODTExe  = Join-Path $script:DepotTemp "ODT.exe"
        # Try multiple ODT sources - scrape MS page first, then fall back to known URLs
        $downloaded = $false
        $ODTUrls = @()
        # Attempt to scrape current URL from Microsoft download page
        try {
            $html = Get-DepotWebText "https://www.microsoft.com/en-us/download/details.aspx?id=49117"
            $match = [regex]::Match($html, '"url"\s*:\s*"(https://download\.microsoft\.com/download/[^"]+officedeploymenttool[^"]+\.exe)"')
            if ($match.Success) { $ODTUrls += $match.Groups[1].Value }
        } catch { }
        # Always include known fallback URLs
        $ODTUrls += "https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_17928-20114.exe"
        foreach ($url in $ODTUrls) {
            Write-Host " downloading ODT..." -NoNewline -ForegroundColor DarkGray
            if (Get-Installer -Url $url -Dest $ODTExe) { $downloaded = $true; break }
        }
        if (-not $downloaded) {
            Write-DepotStatus Failed "[FAILED - download]" "Red"
        } else {
            if (-not (Test-Path $ODTDir)) { New-Item -ItemType Directory -Path $ODTDir | Out-Null }
            # Extract ODT - spawns a child process internally so we poll for setup.exe rather than trusting -Wait
            Write-Host " extracting..." -NoNewline -ForegroundColor DarkGray
            $extractProc = Start-Process $ODTExe -ArgumentList "/quiet /extract:`"$ODTDir`"" -PassThru -WindowStyle Hidden
            $extractDeadline = (Get-Date).AddSeconds(30)
            $setup = "$ODTDir\setup.exe"
            while (-not (Test-Path $setup) -and (Get-Date) -lt $extractDeadline) {
                Start-Sleep -Milliseconds 500
            }
            if ($extractProc -and -not $extractProc.HasExited) {
                Stop-DepotProcessTree -RootId $extractProc.Id
            }
            if (-not (Test-Path $setup)) {
                Write-DepotStatus Failed "[FAILED - ODT extract]" "Red"
            } else {
                @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="Current">
    <Product ID="O365BusinessRetail">
      <Language ID="en-us"/>
      <ExcludeApp ID="Groove"/>
      <ExcludeApp ID="Lync"/>
    </Product>
  </Add>
  <Display Level="None" AcceptEULA="TRUE"/>
  <Property Name="AUTOACTIVATE" Value="1"/>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
  <Property Name="SharedComputerLicensing" Value="0"/>
  <Updates Enabled="TRUE"/>
  <RemoveMSI/>
</Configuration>
"@ | Set-Content "$ODTDir\config.xml"
                Unblock-File $setup -ErrorAction SilentlyContinue
                Unblock-File "$ODTDir\config.xml" -ErrorAction SilentlyContinue
                if (Test-DepotMsiOffice) {
                    # The config above carries <RemoveMSI/>, so ODT owns removal of MSI Office. Deleting its
                    # files and keys here would break the existing install if ODT then fails.
                    Write-Host " MSI Office found, ODT will remove it..." -NoNewline -ForegroundColor DarkGray
                } else {
                    # Scrub leftover C2R registry keys and folders before install
                    $C2RKeys = @(
                        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun",
                        "HKLM:\SOFTWARE\Microsoft\Office\16.0",
                        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0",
                        "HKLM:\SOFTWARE\Microsoft\AppVISV"
                    )
                    foreach ($key in $C2RKeys) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
                    @(
                        "C:\Program Files\Microsoft Office",
                        "C:\Program Files (x86)\Microsoft Office",
                        "C:\Program Files\Common Files\Microsoft Shared\ClickToRun",
                        "C:\ProgramData\Microsoft\ClickToRun"
                    ) | ForEach-Object { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }
                }
                Write-Installing
                $ExitCode = Start-Installer -FilePath $setup -ArgumentList "/configure `"$ODTDir\config.xml`"" -TimeoutMinutes 90
                Show-Result $ExitCode
            }
            # ODTDir may still be locked by OfficeClickToRun - delete async so we don't block
            Start-Process "cmd.exe" -ArgumentList "/c timeout /t 120 /nobreak >nul && rd /s /q `"$ODTDir`" && del /f /q `"$ODTExe`"" -WindowStyle Hidden
        }
    }
}

function Wait-DepotOneDrive {
    # OneDrive's installer can exit 0 and still leave nothing behind. Look for the real OneDrive.exe, giving the
    # per-machine installer a short window to finish writing it.
    $deadline = (Get-Date).AddSeconds($script:OneDriveVerifySeconds)
    while ($true) {
        # A 32-bit host sees ProgramFiles as the (x86) folder, so ask for the 64-bit one by name too.
        $roots = @($env:ProgramFiles, $env:ProgramW6432, ${env:ProgramFiles(x86)}) | Where-Object { $_ } | Select-Object -Unique
        foreach ($root in $roots) {
            if (Test-Path "$root\Microsoft OneDrive\OneDrive.exe") { return $true }
        }
        if ((Get-Date) -ge $deadline) { return $false }
        Start-Sleep -Seconds 2
    }
}

function Install-DepotOneDrive {
    $script:DepotStep = 'OneDrive'
    Write-Host "      > Installing: OneDrive..." -NoNewline -ForegroundColor Gray
    $OneDriveInstalled = (Test-Path "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe") -or
                         (Test-Path "$env:ProgramFiles (x86)\Microsoft OneDrive\OneDrive.exe") -or
                         (Test-Path "$env:SystemRoot\SysWOW64\OneDriveSetup.exe")
    if ($OneDriveInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $dest = Join-Path $script:DepotTemp "OneDriveSetup.exe"
        Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
        if (Get-Installer -Url "https://go.microsoft.com/fwlink/?linkid=844652" -Dest $dest) {
            Write-Installing
            $ExitCode = Start-Installer -FilePath $dest -ArgumentList "/allusers /silent"
            if ($ExitCode -in 0, 3010) {
                if (Wait-DepotOneDrive) { Show-Result $ExitCode }
                else { Write-DepotStatus Failed "[FAILED - OneDrive.exe not found after install]" "Red" "installer exited $ExitCode but OneDrive.exe is missing" }
            } else { Show-Result $ExitCode }
        } else { Write-DepotStatus Failed "[FAILED - download]" "Red" }
        Remove-Item $dest -Force -ErrorAction SilentlyContinue
    }
}

function Install-DepotTeams {
    $script:DepotStep = 'Teams'
    Write-Host "      > Installing: Microsoft Teams..." -NoNewline -ForegroundColor Gray
    $TeamsInstalled = (Test-Path "$env:ProgramFiles\WindowsApps\MSTeams*\ms-teams.exe") -or
                      (Test-Path "$env:ProgramFiles (x86)\Microsoft\Teams\current\Teams.exe") -or
                      (Test-Path "$env:ProgramFiles\Microsoft\Teams\current\Teams.exe")
    if ($TeamsInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $dest = Join-Path $script:DepotTemp "TeamsSetup.exe"
        Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
        if (Get-Installer -Url "https://go.microsoft.com/fwlink/?linkid=2196106" -Dest $dest) {
            Write-Installing
            $ExitCode = Start-Installer -FilePath $dest -ArgumentList "/silent"
            Show-Result $ExitCode
        } else { Write-DepotStatus Failed "[FAILED - download]" "Red" }
        Remove-Item $dest -Force -ErrorAction SilentlyContinue
    }
}

function Install-DepotChrome {
    $script:DepotStep = 'Chrome'
    Write-Host "      > Installing: Google Chrome..." -NoNewline -ForegroundColor Gray
    $ChromeInstalled = (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") -or
                       (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
    if ($ChromeInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $dest = Join-Path $script:DepotTemp "ChromeEnterprise.msi"
        $ChromeUrls = @(
            "https://dl.google.com/dl/chrome/install/googlechromestandaloneenterprise64.msi",
            "https://dl.google.com/chrome/install/googlechromestandaloneenterprise64.msi",
            "https://dl.google.com/edgedl/chrome/install/googlechromestandaloneenterprise64.msi"
        )
        $ChromeDL = $false
        foreach ($cu in $ChromeUrls) {
            Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
            if (Get-Installer -Url $cu -Dest $dest) { $ChromeDL = $true; break }
        }
        if ($ChromeDL) {
            Write-Installing
            $ExitCode = Start-Installer "msiexec.exe" -ArgumentList "/i `"$dest`" /qn /norestart ALLUSERS=1"
            Show-Result $ExitCode
        } else { Write-DepotStatus Failed "[FAILED - download]" "Red" }
        Remove-Item $dest -Force -ErrorAction SilentlyContinue
    }
}

function Install-DepotAcrobat {
    $script:DepotStep = 'Acrobat'
    Write-Host "      > Installing: Adobe Acrobat Reader..." -NoNewline -ForegroundColor Gray
    $AcrobatInstalled = (Test-Path "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe") -or
                        (Test-Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe") -or
                        (Test-Path "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe")
    if ($AcrobatInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $dest = Join-Path $script:DepotTemp "AcrobatReader.exe"
        $AcrobatUrls = @(
            "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2300820555/AcroRdrDC2300820555_en_US.exe",
            "https://get.adobe.com/reader/download/?installer=Reader_DC_2300820555_English_for_Windows&os=Windows%2011&browser_type=KHTML&browser_dist=Chrome&d=true&d_pkg=exe"
        )
        $downloaded = $false
        foreach ($url in $AcrobatUrls) {
            if (Get-Installer -Url $url -Dest $dest) { $downloaded = $true; break }
        }
        if ($downloaded) {
            Write-Installing
            $ExitCode = Start-Installer -FilePath $dest -ArgumentList "/sAll /rs /msi EULA_ACCEPT=YES /qn REBOOT=ReallySuppress"
            Show-Result $ExitCode
        } else { Write-DepotStatus Failed "[FAILED - download]" "Red" }
        Remove-Item $dest -Force -ErrorAction SilentlyContinue
    }
}

function Install-DepotZoom {
    $script:DepotStep = 'Zoom'
    Write-Host "      > Installing: Zoom..." -NoNewline -ForegroundColor Gray
    $ZoomInstalled = (Test-Path "C:\Program Files\Zoom\bin\Zoom.exe") -or
                     (Test-Path "C:\Program Files (x86)\Zoom\bin\Zoom.exe")
    if ($ZoomInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $dest = Join-Path $script:DepotTemp "ZoomInstaller.msi"
        Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
        if (Get-Installer -Url "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64" -Dest $dest) {
            Write-Installing
            $ExitCode = Start-Installer "msiexec.exe" -ArgumentList "/i `"$dest`" /qn /norestart ALLUSERS=1 REBOOT=ReallySuppress"
            Show-Result $ExitCode
        } else { Write-DepotStatus Failed "[FAILED - download]" "Red" }
        Remove-Item $dest -Force -ErrorAction SilentlyContinue
    }
}

function Install-Depot7Zip {
    $script:DepotStep = '7-Zip'
    Write-Host "      > Installing: 7-Zip..." -NoNewline -ForegroundColor Gray
    $7ZipInstalled = (Test-Path "C:\Program Files\7-Zip\7z.exe") -or
                     (Test-Path "C:\Program Files (x86)\7-Zip\7z.exe")
    if ($7ZipInstalled) {
        Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
    } else {
        $dest = Join-Path $script:DepotTemp "7zip.exe"
        Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
        if (Get-Installer -Url "https://www.7-zip.org/a/7z2407-x64.exe" -Dest $dest) {
            Write-Installing
            $ExitCode = Start-Installer -FilePath $dest -ArgumentList "/S"
            Show-Result $ExitCode
        } else { Write-DepotStatus Failed "[FAILED - download]" "Red" }
        Remove-Item $dest -Force -ErrorAction SilentlyContinue
    }
}

function Install-DepotBrowserExtensions {
    $script:DepotStep = 'BrowserExtensions'
    # --- Browser Extension Policy Keys ---
    Write-Host "      > Adding uBlock adblocker extensions..." -NoNewline -ForegroundColor Gray
    $IsElevated = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")
    if ($IsElevated) {
        try {
            $EdgePref   = "HKLM:\SOFTWARE\Policies\Microsoft\Edge\ExtensionInstallForcelist"
            $ChromePref = "HKLM:\SOFTWARE\Policies\Google\Chrome\ExtensionInstallForcelist"
            $EdgeInstalled   = (Test-Path $EdgePref) -and ((Get-Item $EdgePref -ErrorAction SilentlyContinue).Property | ForEach-Object { (Get-ItemProperty $EdgePref).$_ }) -match "odfafepnkmbhccpbejgmiehpchacaeak"
            $ChromeInstalled = (Test-Path $ChromePref) -and ((Get-Item $ChromePref -ErrorAction SilentlyContinue).Property | ForEach-Object { (Get-ItemProperty $ChromePref).$_ }) -match "ddmmnachommbenpoecpkehpconingflb"
            if ($EdgeInstalled -and $ChromeInstalled) {
                Write-DepotStatus AlreadyInstalled "[ALREADY INSTALLED]" "Cyan"
            } else {
                # Edge - uBlock Origin
                if (!(Test-Path $EdgePref)) { New-Item -Path $EdgePref -Force | Out-Null }
                if (-not $EdgeInstalled) {
                    $EdgeCount = ((Get-Item $EdgePref).Property | Measure-Object).Count
                    Set-ItemProperty -Path $EdgePref -Name ($EdgeCount + 1).ToString() -Value "odfafepnkmbhccpbejgmiehpchacaeak;https://edge.microsoft.com/extensionwebstorebase/v1/crx" -ErrorAction Stop
                }
                # Chrome - uBlock Origin Lite (MV3)
                if (!(Test-Path $ChromePref)) { New-Item -Path $ChromePref -Force | Out-Null }
                if (-not $ChromeInstalled) {
                    $ChromeCount = ((Get-Item $ChromePref).Property | Measure-Object).Count
                    Set-ItemProperty -Path $ChromePref -Name ($ChromeCount + 1).ToString() -Value "ddmmnachommbenpoecpkehpconingflb;https://clients2.google.com/service/update2/crx" -ErrorAction Stop
                }
                Write-DepotStatus Done "[DONE]" "Green"
            }
        } catch { Write-DepotStatus Failed "[FAILED]" "Red" }
    } else {
        Write-DepotStatus Skipped "[SKIPPED - not elevated]" "Yellow"
    }
}

function Install-DepotApps {
    Write-Host "[*] Executing Application Provisioning..." -ForegroundColor DarkRed
    Write-Host "    (Press 'ESC' at any time to skip current install)" -ForegroundColor DarkGray
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Install-DepotOffice
    Install-DepotOneDrive
    Install-DepotTeams
    Install-DepotChrome
    Install-DepotAcrobat
    Install-DepotZoom
    Install-Depot7Zip
    Install-DepotBrowserExtensions
}

function Invoke-DepotCleanup {
    Write-Host "[*] Commencing Deep System Clean..." -ForegroundColor DarkRed

    # The update cache is not cleared here. UsoClient started a background download and install earlier in
    # this run, and stopping wuauserv or deleting SoftwareDistribution\Download now would pull files out from
    # under it. The cache is reset before USO starts instead.
    Write-Host "      > Update Cache..." -NoNewline -ForegroundColor Gray
    Write-StatusTag "[LEFT IN PLACE - updates running]" "Yellow"
    Write-Host "      > Purging Temp Folders..." -NoNewline -ForegroundColor Gray
    $script:DepotStep = 'Cleanup'
    $TempPaths = @()
    if (Test-DepotAnyFailed) {
        # Installer logs and leftovers in C:\Windows\Temp are the evidence for the failed step, keep them.
        Write-Host " (a step failed, keeping C:\Windows\Temp for troubleshooting)" -NoNewline -ForegroundColor Yellow
    } else {
        $TempPaths += "C:\Windows\Temp\*"
    }
    $TempPaths += @(
        "C:\Windows\Prefetch\*",
        "$env:LocalAppData\Microsoft\Windows\Explorer\thumbcache_*.db"
    )
    foreach ($Path in $TempPaths) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction SilentlyContinue
    }
    Write-StatusTag "[DONE]" "Green"
    Write-Host "      > Finalizing Disk Space..." -NoNewline -ForegroundColor Gray
    Clear-RecycleBin -Confirm:$false -ErrorAction SilentlyContinue
    Write-StatusTag "[DONE]" "Green"
}

function Get-DepotProfileHive {
    # Local (S-1-5-21) and Entra ID (S-1-12-1) user profiles that have an NTUSER.DAT. The pattern is anchored
    # so service accounts, .bak profiles and other look-alikes are never touched.
    $sidPattern = '^S-1-(5-21|12-1)-\d+-\d+-\d+-\d+$'
    $found = @()
    $rows = @(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -ErrorAction SilentlyContinue)
    foreach ($r in $rows) {
        if ("$($r.PSChildName)" -notmatch $sidPattern) { continue }
        $path = [Environment]::ExpandEnvironmentVariables("$($r.ProfileImagePath)")
        if (-not $path) { continue }
        $dat = "$path\NTUSER.DAT"
        if (-not (Test-Path $dat)) { continue }
        $found += [pscustomobject]@{ SID = "$($r.PSChildName)"; Path = $dat }
    }
    return $found
}

function Mount-DepotUserHive {
    # Fills $script:UserHives with the hives that are really writable. A hive that is already loaded (the user
    # is logged on) is written in place and never unloaded. A hive DEPOT loads is verified before it counts.
    $script:UserHives = @()
    $script:DepotProfilesSeen = 0
    foreach ($p in (Get-DepotProfileHive)) {
        $script:DepotProfilesSeen++
        $live = "Registry::HKEY_USERS\$($p.SID)"
        if (Test-Path $live) {
            $script:UserHives += [pscustomobject]@{ Mount = $live; SID = $p.SID; Loaded = $false; Name = $null }
            continue
        }
        $name = "APT_$($p.SID)"
        $r = Invoke-DepotReg @('load', "HKU\$name", $p.Path)
        $mount = "Registry::HKEY_USERS\$name"
        if ($r.ExitCode -eq 0 -and (Test-Path $mount)) {
            $script:UserHives += [pscustomobject]@{ Mount = $mount; SID = $p.SID; Loaded = $true; Name = $name }
        } else {
            Write-Host " [profile $($p.SID) could not be loaded]" -NoNewline -ForegroundColor Yellow
            Add-DepotResult 'Privacy' 'Failed' "profile $($p.SID) hive could not be loaded"
        }
    }
}

function Dismount-DepotUserHive {
    # Unloads only the hives this run loaded, and says so when one will not let go.
    [GC]::Collect()
    Start-Sleep -Milliseconds 500
    foreach ($Hive in @($script:UserHives | Where-Object { $_.Loaded })) {
        $r = Invoke-DepotReg @('unload', "HKU\$($Hive.Name)")
        if ($r.ExitCode -ne 0) {
            Start-Sleep -Milliseconds 1000
            [GC]::Collect()
            $r = Invoke-DepotReg @('unload', "HKU\$($Hive.Name)")
        }
        if ($r.ExitCode -ne 0) {
            Add-DepotResult 'Privacy' 'Failed' "could not unload hive $($Hive.Name); it is cleaned up on the next run"
        }
    }
}

function Invoke-DepotPrivacy {
    Write-Host "[*] Applying Privacy & UI Settings..." -ForegroundColor DarkRed
    $script:DepotStep = 'Privacy'
    # --- Load all user hives so per-user keys can be set for every account ---
    Mount-DepotUserHive
    # ---- MICROSOFT EDGE ---------------------------------------------------------
    Write-Host "      > Configuring Microsoft Edge..." -NoNewline -ForegroundColor Gray
    $EdgePolicy = "HKLM:\SOFTWARE\Policies\Microsoft\Edge"
    if (-not (Test-Path $EdgePolicy)) { New-Item -Path $EdgePolicy -Force | Out-Null }
    # Disable Copilot / Edge AI sidebar button in toolbar
    Set-RegValue "$EdgePolicy"                      "HubsSidebarEnabled"               0
    Set-RegValue "$EdgePolicy"                      "CopilotPageContext"                0
    Set-RegValue "$EdgePolicy"                      "EdgeEntraCopilotPageContext"       0
    # Remove sidebar entirely
    Set-RegValue "$EdgePolicy"                      "EdgeSidebarEnabled"               0
    # Disable background running when browser is closed
    Set-RegValue "$EdgePolicy"                      "BackgroundModeEnabled"            0
    # Disable startup boost (pre-loads Edge on Windows start)
    Set-RegValue "$EdgePolicy"                      "StartupBoostEnabled"              0
    # Additional toolbar/UI cleanup
    Set-RegValue "$EdgePolicy"                      "ShowMicrosoftRewards"             0
    Set-RegValue "$EdgePolicy"                      "EdgeShoppingAssistantEnabled"     0
    Set-RegValue "$EdgePolicy"                      "PersonalizationReportingEnabled"  0
    Set-RegValue "$EdgePolicy"                      "EdgeFollowEnabled"                0
    Set-RegValue "$EdgePolicy"                      "ShowRecommendationsEnabled"       0
    Set-RegValue "$EdgePolicy"                      "EdgeWorkspacesEnabled"            0
    Set-RegValue "$EdgePolicy"                      "DiscoverPageContextEnabled"       0
    # Also set per-user preferences as fallback (policy keys take precedence but belt+suspenders)
    Set-AllUsers "SOFTWARE\Policies\Microsoft\Edge"            "HubsSidebarEnabled" 0
    Set-AllUsers "SOFTWARE\Policies\Microsoft\Edge"            "BackgroundModeEnabled" 0
    Write-StatusTag "[DONE]" "Green"
    # ---- TELEMETRY & DIAGNOSTICS ------------------------------------------------
    Write-Host "      > Disabling Telemetry & Diagnostics..." -NoNewline -ForegroundColor Gray
    # Telemetry level: 0 = Security only (Enterprise), 1 = Basic (minimum for Home/Pro)
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"     "AllowTelemetry"                  1
    Set-RegValue "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" "AllowTelemetry" 0
    Set-RegValue "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\DataCollection" "MaxTelemetryAllowed" 0
    # Disable feedback / diagnostic prompts
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection"     "DoNotShowFeedbackNotifications"  1
    Set-AllUsers  "SOFTWARE\Microsoft\Siuf\Rules"                                  "NumberOfSIUFInPeriod"            0
    Set-AllUsers  "SOFTWARE\Microsoft\Siuf\Rules"                                  "PeriodInNanoSeconds"             0
    # Disable activity history / timeline
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"             "EnableActivityFeed"              0
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"             "PublishUserActivities"           0
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\System"             "UploadUserActivities"            0
    # Disable advertising ID
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\AdvertisingInfo"     "Enabled"                        0
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo"    "DisabledByGroupPolicy"           1
    # Disable app launch tracking
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "Start_TrackProgs"               0
    # Disable Windows Error Reporting
    Set-RegValue "HKLM:\SOFTWARE\Microsoft\Windows\Windows Error Reporting"      "Disabled"                        1
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Error Reporting" "Disabled"                   1
    # Disable Customer Experience Improvement Program
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\SQMClient\Windows"          "CEIPEnable"                      0
    # Stop and disable telemetry services
    $telemetrySvcs = @("DiagTrack", "dmwappushservice")
    foreach ($Svc in $telemetrySvcs) {
        if (Get-Service -Name $Svc -ErrorAction SilentlyContinue) {
            & sc.exe stop $Svc >$null 2>&1
            & sc.exe config $Svc start= disabled >$null 2>&1
            Set-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Services\$Svc" -Name "Start" -Value 4 -ErrorAction SilentlyContinue
        }
    }
    # Disable tailored experiences and language opt-out
    Set-AllUsers "SOFTWARE\Microsoft\Windows\CurrentVersion\Privacy"   "TailoredExperiencesWithDiagnosticDataEnabled" 0
    Set-AllUsers "Control Panel\International\User Profile"            "HttpAcceptLanguageOptOut"                     1
    # OEM content injection prevention
    Set-RegValue "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Device Metadata" "PreventDeviceMetadataFromNetwork" 1
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"          "DisableWindowsConsumerFeatures"   1
    # AI / Recall
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsAI"             "DisableAIDataAnalysis"            1
    Write-StatusTag "[DONE]" "Green"
    # ---- TASKBAR ----------------------------------------------------------------
    Write-Host "      > Configuring Taskbar..." -NoNewline -ForegroundColor Gray
    # Align taskbar to left (0 = Left, 1 = Center)
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "TaskbarAl"                      0
    # Disable Bing web results in search
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Search"              "BingSearchEnabled"               0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\SearchSettings"      "IsDynamicSearchBoxPresent"       0
    Set-AllUsers  "SOFTWARE\Policies\Microsoft\Windows\Explorer"                  "DisableSearchBoxSuggestions"     1
    # Disable Task View button
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "ShowTaskViewButton"              0
    # Disable Widgets
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Dsh"                         "AllowNewsAndInterests"           0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "TaskbarDa"                      0
    # Disable Chat / Teams taskbar icon
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "TaskbarMn"                      0
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Chat"        "ChatIcon"                        3
    # Disable News and Interests
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Feeds"      "EnableFeeds"                     0
    Write-StatusTag "[DONE]" "Green"
    # ---- START MENU -------------------------------------------------------------
    Write-Host "      > Configuring Start Menu..." -NoNewline -ForegroundColor Gray
    # Disable suggested apps / recommendations in Start
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"           "HideRecentlyAddedApps"           1
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Explorer"           "HideRecommendedSection"          1
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "Start_IrisRecommendations"      0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Advanced"  "Start_AccountNotifications"     0
    # Disable tips / suggestions / spotlight
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338388Enabled" 0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338389Enabled" 0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-353694Enabled" 0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-353696Enabled" 0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SystemPaneSuggestionsEnabled"    0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SoftLandingEnabled"              0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "RotatingLockScreenEnabled"       0
    # Disable lock screen tips / spotlight
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"       "DisableWindowsSpotlightFeatures" 1
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CloudContent"       "DisableSoftLanding"              1
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "RotatingLockScreenOverlayEnabled"  0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338387Enabled"   0
    Write-StatusTag "[DONE]" "Green"
    # ---- NOTIFICATION / ACTION CENTER -------------------------------------------------------------
    Write-Host "      > Disabling Notification Suggestions..." -NoNewline -ForegroundColor Gray
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\Notifications\Settings\Windows.SystemToast.Suggested" "Enabled" 0
    Set-RegValue "HKLM:\SOFTWARE\Policies\Microsoft\Windows\CurrentVersion\PushNotifications" "NoToastApplicationNotification" 0
    # Disable "Get tips and suggestions when using Windows"
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-338393Enabled" 0
    Set-AllUsers  "SOFTWARE\Microsoft\Windows\CurrentVersion\ContentDeliveryManager" "SubscribedContent-310093Enabled" 0
    Write-StatusTag "[DONE]" "Green"
    # ---- UNLOAD USER HIVES -------------------------------------------------------------
    $reached = @($script:UserHives).Count
    Dismount-DepotUserHive
    Write-Host "    Privacy & UI settings applied to $reached of $($script:DepotProfilesSeen) user profile(s)." -ForegroundColor Cyan
}

function Test-DepotAnyFailed {
    return (@($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count -gt 0)
}

function Get-DepotExitCode {
    # 0 clean, 3010 clean with a reboot needed, 3 some steps failed but others worked, 1 nothing worked.
    # A failure outranks a pending reboot: the technician must see that something went wrong.
    $failed = @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count
    $worked = @($script:DepotResults | Where-Object { $_.Status -in 'Done', 'AlreadyInstalled' }).Count
    if ($failed -gt 0) { if ($worked -gt 0) { return 3 } else { return 1 } }
    if ($script:RebootRequired) { return 3010 }
    return 0
}

function Write-DepotRunLog {
    # Keeps the per-step results where a technician can read them after the window is gone.
    param([int]$Code)
    try {
        $dir = Join-Path $env:ProgramData 'DEPOT'
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        $lines = @("DEPOT run $(Get-Date -Format 's') on $($env:COMPUTERNAME), exit code $Code")
        foreach ($r in $script:DepotResults) { $lines += ("{0,-18} {1,-16} {2}" -f $r.Step, $r.Status, $r.Detail) }
        Set-Content -LiteralPath (Join-Path $dir 'LastRun.txt') -Value $lines -Encoding UTF8 -ErrorAction Stop
    } catch { }
}

function Complete-Depot {
    $Code      = Get-DepotExitCode
    $Failures  = @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' })
    $RebootPending = [bool]$script:RebootRequired
    if ($Failures.Count -gt 0)  { $RebootStatus = "COMPLETED WITH FAILURES"; $StatusColor = "Red" }
    elseif ($RebootPending)     { $RebootStatus = "REBOOT REQUIRED";         $StatusColor = "Red" }
    else                        { $RebootStatus = "SYSTEM READY";            $StatusColor = "Green" }
    Write-HLine -Style "dashed"
    Write-Host "[>] Status  : " -NoNewline -ForegroundColor $script:LineCol
    Write-Host $RebootStatus -ForegroundColor $StatusColor -BackgroundColor Black
    foreach ($f in $Failures) {
        Write-Host "    [!] $($f.Step): $($f.Detail)" -ForegroundColor Red
    }
    if ($Failures.Count -gt 0 -and $RebootPending) {
        Write-Host "    [!] A reboot is also pending." -ForegroundColor Yellow
    }
    # Footer
    $_sfx    = "█"
    $_ftrW   = $script:ArtW + 1
    $_ffillW = [Math]::Max(0, $script:Width - $_ftrW - $_sfx.Length)
    $_footer = if ($Failures.Count -gt 0) { "  ENDPOINT PROVISIONING SEQUENCE FINISHED WITH FAILURES" } else { "  ENDPOINT PROVISIONING SEQUENCE COMPLETE" }
    $_ver    = "| v2.0"
    $_fpad   = " " * [Math]::Max(0, $_ffillW - $_footer.Length - $_ver.Length)
    Write-Host ("-" * $_ffillW) -ForegroundColor $script:LineCol -NoNewline; Write-Host " $($script:Art1)" -ForegroundColor $script:AccentCol -NoNewline; Write-Host $_sfx -ForegroundColor $script:LineCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $script:MainCol -NoNewline; Write-Host " $($script:Art2)" -ForegroundColor $script:AccentCol -NoNewline; Write-Host $_sfx -ForegroundColor $script:LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $script:LineCol -NoNewline; Write-Host " $($script:Art3)" -ForegroundColor $script:AccentCol -NoNewline; Write-Host $_sfx -ForegroundColor $script:LineCol

    # One plain line for an RMM console or a log, always in agreement with the exit code.
    $FinalLine = "DEPOT RESULT: exit code $Code - $RebootStatus"
    if ($Failures.Count -gt 0) { $FinalLine += " (" + (($Failures | ForEach-Object { "$($_.Step): $($_.Detail)" }) -join '; ') + ")" }
    Write-Output $FinalLine
    Write-DepotRunLog -Code $Code

    # Only a person at a console is asked anything. Unattended, there is no prompt and no reboot.
    $Choice = ''
    if (Test-DepotInteractive) {
        $Choice = Read-Host -Prompt "`nPress ENTER to exit, or type 'RESTART' to reboot"
    }
    # Self-delete only after a good run (0 or 3010). After a failure the script stays so it can be rerun.
    $selfPath = $script:DepotScriptPath
    if (($Code -eq 0 -or $Code -eq 3010) -and $selfPath -and (Test-Path $selfPath)) {
        Start-Process powershell -ArgumentList "-NoProfile -Command `"Start-Sleep 5; Remove-Item -Path '$selfPath' -Force`"" -WindowStyle Hidden
    }
    if ($Choice -eq "RESTART") { Restart-Computer -Force }
    Exit-KillerScript $Code
}

function Invoke-DepotMain {
    Reset-DepotState
    if (-not (Test-DepotElevated)) { Invoke-DepotElevationRelaunch }
    Invoke-DepotPreflight
    Invoke-DepotInteractiveSetup
    Show-DepotBanner
    Invoke-DepotWindowsUpdate
    Install-DepotApps
    Invoke-DepotCleanup
    Invoke-DepotPrivacy
    Complete-Depot
}

if (-not $global:KillerScriptsTestMode) { Invoke-DepotMain }
