<#
.SYNOPSIS
    Outlook Repair & Configuration Assistant (ORCA) v2.1
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Resets broken Outlook installations for one or all user profiles.
    Supports New Outlook (Microsoft Store), Classic Outlook (Office/M365),
    or both simultaneously. Clears cached data, registry hives, and
    authentication tokens. Optionally removes OST files and purges 3rd-party
    extensions for Classic Outlook. New Outlook is reinstalled clean from
    the Microsoft Store after reset.
.NOTES
    Can be run as Administrator (for all profiles) or Standard User (current profile only).
    NTUSER.DAT is loaded for per-user registry cleanup on non-current accounts
    (requires the target user to be logged out).
    .\ORCA.ps1
        Prompts for user, Outlook type, and OST/ext. deletion preference.
    .\ORCA.ps1 -Silent
        Non-interactive - All Users, Both Outlook types, no OST/ext. deletion.
#>
param(
    [switch]$Silent
)
$EXIT_SUCCESS = 0   # Completed successfully
$EXIT_ERROR   = 1   # General / unexpected error
$EXIT_CANCEL  = 2   # User cancelled

$script:IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
$script:DeletedOSTs = @()
$script:PurgedExtensions = New-Object System.Collections.Generic.List[string]

# Suppress standard progress bars (fixes "Deployment operation progress" popups from Appx cmdlets)
$ProgressPreference = 'SilentlyContinue'
# Ensure TLS 1.2 is enabled for PowerShell Gallery downloads (critical for PS 5.1)
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

#region [1] - BANNER & DISPLAY HELPERS
# ============================================================================
Clear-Host
$script:Width  = 85
$_ver    = "| v2.1"
$LineCol   = "DarkCyan"
$MainCol   = "Cyan"
$WarnCol   = "DarkYellow"
$BorderCol = "Cyan"
$ArtCol    = "White"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"

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

# Art (shared by header and footer)
$_pfx  = "█  "
$_art1 = "╔═╗ ╦═╗ ╔═╗ ╔═╗ "
$_art2 = "║ ║ ╠╦╝ ║   ╠═╣ "
$_art3 = "╚═╝ ╩╚═ ╚═╝ ╩ ╩ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "OUTLOOK REPAIR & CONFIGURATION ASSISTANT"
$_tpad  = " " * ($_fillW - $_title.Length)

function Write-Header {
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("━" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("━" * $_fillW) -ForegroundColor $LineCol
}

function Write-Step {
    param([string]$StepNum, [string]$Message)
    Write-Host "[$StepNum] $Message" -ForegroundColor $InfoCol -NoNewline
}
function Complete-Step {
    param([string]$StepNum, [string]$Message)
    $pad = " " * [math]::Max(1, $script:Width - "[$StepNum] $Message".Length - "[DONE]".Length)
    Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
    Write-Host "[$StepNum]" -ForegroundColor $DimCol -NoNewline
    Write-Host " $Message$pad" -ForegroundColor $DimCol -NoNewline
    Write-Host "[DONE]" -ForegroundColor $OkCol
}

function Get-ActiveUser {
    # Method 1: Get the owner of the explorer.exe process. This is the most reliable way to find the interactive user.
    # Requires elevation.
    if ($script:IsAdmin) {
        $explorerProc = Get-Process -Name "explorer" -IncludeUserName -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($explorerProc) {
            return $explorerProc.UserName -replace '^.*\\'
        }
    }

    # Method 2: Use 'quser' to find the active console session. Good fallback for RDP/RMM scenarios.
    $quserOutput = quser 2>$null
    if ($quserOutput) {
        $activeSession = $quserOutput | Select-String -Pattern "console\s+\d+\s+Active" | Select-Object -First 1
        if ($activeSession) {
            return ($activeSession.ToString().Trim() -split '\s+')[0].TrimStart('>')
        }
    }

    # Method 3: Fallback to Win32_ComputerSystem for basic scenarios.
    $csUser = (Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue).UserName
    if ($csUser) {
        return $csUser -replace '^.*\\'
    }

    # Method 4: Final fallback for non-elevated context is the environment variable.
    if (-not $script:IsAdmin -and $env:USERNAME) {
        return $env:USERNAME
    }

    return $null
}

# ---- Cleanup helper functions ------------------------------------------------

function Invoke-PurgeExtensions {
    param([string]$UserName)
    if (-not $script:PurgeExtensions) { return }

    try {
        $userSID = (New-Object System.Security.Principal.NTAccount($UserName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch {
        return
    }
    $hiveName   = "ORCA_Ext_$UserName"
    $hiveLoaded = $false
    $root       = $null

    # Check if user's hive is already loaded
    if (Test-Path "Registry::HKU\$userSID") {
        $root = "Registry::HKU\$userSID"
    } else {
        $ntUser = "C:\Users\$UserName\NTUSER.DAT"
        if (-not (Test-Path $ntUser)) { return } # Silently skip if no hive
        reg load "HKU\$hiveName" $ntUser *>$null
        if ($LASTEXITCODE -eq 0) {
            $root = "Registry::HKU\$hiveName"; $hiveLoaded = $true
        } else { return } # Silently skip if hive is locked
    }

    # Remove user-specific add-ins
    $userPaths = @("$root\Software\Microsoft\Office\Outlook\Addins", "$root\Software\WOW6432Node\Microsoft\Office\Outlook\Addins")
    foreach ($path in $userPaths) {
        if (Test-Path $path) {
            Get-ChildItem $path -ErrorAction SilentlyContinue | ForEach-Object {
                $extName = "$($_.PSChildName) [$UserName]"
                if (-not $script:PurgedExtensions.Contains($extName)) { $script:PurgedExtensions.Add($extName) }
            }
            Remove-Item $path -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    if ($hiveLoaded) { [gc]::Collect(); [gc]::WaitForPendingFinalizers(); reg unload "HKU\$hiveName" *>$null }

    # Remove machine-wide add-ins (only if admin)
    if ($script:IsAdmin) {
        "HKLM:\Software\Microsoft\Office\Outlook\Addins", "HKLM:\Software\WOW6432Node\Microsoft\Office\Outlook\Addins" |
            Where-Object { Test-Path $_ } | ForEach-Object {
                $rootPath = $_
                Get-ChildItem $rootPath -ErrorAction SilentlyContinue | ForEach-Object {
                    $extName = "$($_.PSChildName) [Machine]"
                    if (-not $script:PurgedExtensions.Contains($extName)) { $script:PurgedExtensions.Add($extName) }
                }
                Remove-Item $rootPath -Recurse -Force -ErrorAction SilentlyContinue
            }
    }
}

function Invoke-TerminateOutlook {
    param([string]$Mode)
    if ($Mode -ne "Classic") {
        Get-Process -Name "olk", "WebViewHost", "msedgewebview2" -ErrorAction SilentlyContinue | Stop-Process -Force
    }
    if ($Mode -ne "New") {
        Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue | Stop-Process -Force
    }
}

function Invoke-ClearNewOutlookData {
    param([string]$UserName)
    $local = "C:\Users\$UserName\AppData\Local"

    if ($script:IsAdmin) {
        try {
            $userSID = (New-Object System.Security.Principal.NTAccount($UserName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
            $pkg = Get-AppxPackage -Name "Microsoft.OutlookForWindows" -User $userSID -ErrorAction SilentlyContinue
            if ($pkg) { Remove-AppxPackage -Package $pkg.PackageFullName -User $userSID -ErrorAction SilentlyContinue }
        } catch {}
    } else {
        # Non-admin can only remove for self; the -User parameter requires elevation.
        Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
    }
    @(
        "$local\Microsoft\Olk",
        "$local\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe"
    ) | Where-Object { Test-Path $_ } | ForEach-Object { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }
}

function Invoke-ClearClassicOutlookData {
    param([string]$UserName)
    $local = "C:\Users\$UserName\AppData\Local"
    @(
        "$local\Microsoft\Outlook\RoamCache",
        "$local\Microsoft\Outlook\Offline Address Books",
        "$local\Temp\Outlook Logging"
    ) | Where-Object { Test-Path $_ } | ForEach-Object { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }

    if ($script:DeleteOST) {
        $ostFiles = Get-ChildItem "$local\Microsoft\Outlook" -Filter "*.ost" -ErrorAction SilentlyContinue
        foreach ($ostFile in $ostFiles) {
            $script:DeletedOSTs += $ostFile.FullName
            Remove-Item -Path $ostFile.FullName -Force -ErrorAction SilentlyContinue
        }
    }
}

function Invoke-RegistryCleanup {
    param([string]$UserName, [string]$Mode)
    try {
        $userSID = (New-Object System.Security.Principal.NTAccount($UserName)).Translate([System.Security.Principal.SecurityIdentifier]).Value
    } catch {
        return
    }
    $hiveName   = "ORCA_$UserName"
    $hiveLoaded = $false
    $root       = $null

    # Check if user's hive is already loaded (i.e., they are logged in)
    if (Test-Path "Registry::HKU\$userSID") {
        $root = "Registry::HKU\$userSID"
    } else {
        # User is not logged in, so we try to load their hive manually
        $ntUser = "C:\Users\$UserName\NTUSER.DAT"
        if (-not (Test-Path $ntUser)) {
            Write-Host " (NTUSER.DAT not found - skipping registry)" -ForegroundColor $WarnCol -NoNewline
            return
        }
        reg load "HKU\$hiveName" $ntUser *>$null
        if ($LASTEXITCODE -eq 0) {
            $root = "Registry::HKU\$hiveName"
            $hiveLoaded = $true
        } else {
            Write-Host " (hive locked - skipping registry)" -ForegroundColor $WarnCol -NoNewline
            return
        }
    }

    $keys = @()
    if ($Mode -ne "Classic") { $keys += "$root\Software\Microsoft\Olk" }
    if ($Mode -ne "New") {
        $keys += "$root\Software\Microsoft\Office\16.0\Outlook\AutoDiscover"
        $keys += "$root\Software\Microsoft\Office\16.0\Outlook\Options\General"
    }
    $keys | Where-Object { Test-Path $_ } | ForEach-Object { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }

    if ($hiveLoaded) {
        [gc]::Collect()
        [gc]::WaitForPendingFinalizers()
        reg unload "HKU\$hiveName" *>$null
    }
}

function Invoke-AuthFlush {
    param([string]$UserName)
    $local = "C:\Users\$UserName\AppData\Local"
    @(
        "$local\Microsoft\IdentityCache",
        "$local\Microsoft\Edge\User Data\Default\EBWebView",
        "$local\Microsoft\OneAuth",
        "$local\Microsoft\TokenBroker\Cache"
    ) | Where-Object { Test-Path $_ } | ForEach-Object { Remove-Item $_ -Recurse -Force -ErrorAction SilentlyContinue }

    cmdkey /list | ForEach-Object {
        if ($_ -like "*Target:*Outlook*" -or $_ -like "*Target:*Office*") {
            $target = ($_ -split "Target: ")[1].Trim()
            cmdkey /delete:$target
        }
    }
    Get-ChildItem -Path "$local\Packages\Microsoft.AAD.BrokerPlugin_*" -ErrorAction SilentlyContinue |
        Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 15
}
#endregion

#region [2] - CONFIGURATION & SETUP
# ============================================================================
Write-Header
if (-not $script:IsAdmin) {
    Write-Host "  Standard User Mode (Current Profile Only)" -ForegroundColor $WarnCol
}

# === USER SELECTION ===
$ExcludedProfiles = @("Public", "Default", "All Users", "LocalService", "NetworkService", "SYSTEM")
$UserProfiles = Get-ChildItem "C:\Users" | Where-Object { $_.PSIsContainer -and $_.Name -notin $ExcludedProfiles }

if ($script:IsAdmin) {
    # Detect the active interactive user (works even when script is elevated)
    $_activeUser   = Get-ActiveUser

    # Reorder profiles: Active user first (Index 0), then others
    if ($null -ne $_activeUser -and $UserProfiles.Name -contains $_activeUser) {
        $activeObj = $UserProfiles | Where-Object { $_.Name -eq $_activeUser }
        $otherObjs = $UserProfiles | Where-Object { $_.Name -ne $_activeUser }
        $UserProfiles = @($activeObj) + @($otherObjs)
        $_defaultLabel = "0"
        $_defaultValue = "0"
    } else {
        $_defaultLabel = "A"
        $_defaultValue = "A"
    }

    if (-not $Silent) {
        Write-Host ""
        Write-Host "Select user profile:" -ForegroundColor $LineCol
        for ($i = 0; $i -lt $UserProfiles.Count; $i++) {
            $uName = $UserProfiles[$i].Name
            Write-Host "  [$i] " -ForegroundColor $DimCol -NoNewline
            Write-Host "$uName" -ForegroundColor $AccentCol -NoNewline
            if ($uName -eq $_activeUser) {
                Write-Host " (Current User)" -ForegroundColor $WarnCol
            } else {
                Write-Host ""
            }
        }
        Write-Host "  [A] " -ForegroundColor $DimCol -NoNewline; Write-Host "All Users" -ForegroundColor $AccentCol
        Write-Host ""
        Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
        Write-Host "[0-$($UserProfiles.Count - 1)/A]  Enter = $_defaultLabel  Esc = quit: " -ForegroundColor $DimCol -NoNewline
        try {
            $key = [Console]::ReadKey($true)
            Write-Host $key.KeyChar
            if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
            $userChoice = if ($key.Key -eq [ConsoleKey]::Enter) { $_defaultValue } else { $key.KeyChar.ToString() }
        } catch {
            $raw = (Read-Host).Trim()
            $userChoice = if ([string]::IsNullOrWhiteSpace($raw)) { $_defaultValue } else { $raw }
        }
        $script:AllUsers = ($userChoice -eq 'A' -or $userChoice -eq 'a')
        if ($script:AllUsers) {
            $TargetUsers = $UserProfiles.Name
        } elseif ($null -ne $UserProfiles[$userChoice]) {
            $TargetUsers = @($UserProfiles[$userChoice].Name)
        } else {
            Write-Host "Invalid selection. Exiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL
        }
    } else {
        $script:AllUsers = $true
        $TargetUsers = $UserProfiles.Name
    }
} else {
    # Non-admin: force single-user mode for the current user and skip user selection prompt
    $script:AllUsers = $false
    $_activeUser = Get-ActiveUser
    if ($null -eq $_activeUser) {
        Write-Error "Could not determine the current user. Please run as Administrator to target other profiles."
        Exit $EXIT_ERROR
    }
    $TargetUsers = @($_activeUser)
}

# === OUTLOOK MODE SELECTION ===
if (-not $Silent) {
    Write-Host ""
    Write-Host "Select Outlook version to reset:" -ForegroundColor $LineCol
    Write-Host "  [1] " -ForegroundColor $DimCol -NoNewline; Write-Host "New Outlook" -ForegroundColor $AccentCol
    Write-Host "  [2] " -ForegroundColor $DimCol -NoNewline; Write-Host "Classic Outlook" -ForegroundColor $AccentCol
    Write-Host "  [3] " -ForegroundColor $DimCol -NoNewline; Write-Host "Both" -ForegroundColor $AccentCol
    Write-Host ""
    while ($true) {
        Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
        Write-Host "[1-3]  Enter = 3  Esc = quit: " -ForegroundColor $DimCol -NoNewline
        try {
            $key = [Console]::ReadKey($true)
            if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
            if ($key.Key -eq [ConsoleKey]::Enter) { Write-Host "3"; $modeChoice = "3" }
            else { Write-Host $key.KeyChar; $modeChoice = $key.KeyChar.ToString() }
        } catch {
            $raw = (Read-Host).Trim()
            if ([string]::IsNullOrWhiteSpace($raw)) { $modeChoice = "3" }
            elseif ($raw -eq 'q' -or $raw -eq 'Q') { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
            else { $modeChoice = $raw }
        }
        $script:OutlookMode = switch ($modeChoice) {
            "1" { "New" }
            "2" { "Classic" }
            "3" { "Both" }
            default { $null }
        }
        if ($script:OutlookMode) { break }
        Write-Host "Invalid selection, try again." -ForegroundColor $WarnCol
    }
} else {
    $script:OutlookMode = "Both"
}

# === OST DELETION (Classic / Both only) ===
$script:DeleteOST = $false
if (-not $Silent -and $script:OutlookMode -ne "New") {
    Write-Host ""
    Write-Host "Delete OST files?" -ForegroundColor $LineCol
    Write-Host "  (removes local email cache - mailbox will re-sync on next login)" -ForegroundColor $WarnCol
    Write-Host ""
    Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
    Write-Host "[Y/N]  Enter = N  Esc = quit: " -ForegroundColor $DimCol -NoNewline
    try {
        $key = [Console]::ReadKey($true)
        Write-Host $key.KeyChar
        if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
        $ostChoice = if ($key.Key -eq [ConsoleKey]::Enter) { "N" } else { $key.KeyChar.ToString() }
        $script:DeleteOST = ($ostChoice -eq 'Y' -or $ostChoice -eq 'y')
    } catch {
        $raw = (Read-Host).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { $raw = "N" }
        $script:DeleteOST = ($raw -eq 'Y' -or $raw -eq 'y')
    }
}

# === EXTENSION PURGE (Classic / Both only) ===
$script:PurgeExtensions = $false
if (-not $Silent -and $script:OutlookMode -ne "New") {
    Write-Host ""
    Write-Host "Purge all Outlook extensions?" -ForegroundColor $LineCol
    Write-Host "  (removes all 3rd-party add-ins like Salesforce, etc.)" -ForegroundColor $WarnCol
    Write-Host ""
    Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
    Write-Host "[Y/N]  Enter = N  Esc = quit: " -ForegroundColor $DimCol -NoNewline
    try {
        $key = [Console]::ReadKey($true)
        Write-Host $key.KeyChar
        if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
        $extChoice = if ($key.Key -eq [ConsoleKey]::Enter) { "N" } else { $key.KeyChar.ToString() }
        $script:PurgeExtensions = ($extChoice -eq 'Y' -or $extChoice -eq 'y')
    } catch {
        $raw = (Read-Host).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { $raw = "N" }
        $script:PurgeExtensions = ($raw -eq 'Y' -or $raw -eq 'y')
    }
}

# Redraw with confirmed selections
Clear-Host
Write-Header
Write-Host "Device Name     : " -ForegroundColor $LineCol -NoNewline; Write-Host $env:COMPUTERNAME -ForegroundColor $AccentCol
Write-Host "Target Profile  : " -ForegroundColor $LineCol -NoNewline
if ($script:AllUsers) {
    Write-Host "All Users " -ForegroundColor $AccentCol -NoNewline
    Write-Host "($($TargetUsers.Count) profiles)" -ForegroundColor $WarnCol
} else {
    Write-Host $TargetUsers[0] -ForegroundColor $AccentCol
}
Write-Host "Mode            : " -ForegroundColor $LineCol -NoNewline
if ($script:OutlookMode -eq "Both") {
    Write-Host "Both " -ForegroundColor $AccentCol -NoNewline; Write-Host "(New and Classic Outlook)" -ForegroundColor $WarnCol
} else {
    Write-Host "$($script:OutlookMode) Outlook" -ForegroundColor $AccentCol
}
if ($script:OutlookMode -ne "New") {
    Write-Host "Delete OSTs     : " -ForegroundColor $LineCol -NoNewline
    if ($script:DeleteOST) { Write-Host "Yes" -ForegroundColor Red } else { Write-Host "No" -ForegroundColor $DimCol }
    Write-Host "Purge Extensions: " -ForegroundColor $LineCol -NoNewline
    if ($script:PurgeExtensions) { Write-Host "Yes" -ForegroundColor Red } else { Write-Host "No" -ForegroundColor $DimCol }
}
Write-HLine -Style dashed
#endregion

#region [3] - MAIN LOGIC
# ============================================================================
trap { Write-Error $_.Exception.Message; Exit $EXIT_ERROR }
$script:skipReinstall = $false

# Step counts per mode (WinGet is a pre-step outside the user loop)
# New:     4 steps - [1] Kill/remove  [2] Registry  [3] Auth  [4] Reinstall
# Classic: 3 steps - [1] Kill/clear   [2] Registry  [3] Auth
# Both:    5 steps - [1] Kill/Classic clear  [2] New remove  [3] Registry  [4] Auth  [5] Reinstall
$cleanupSteps = switch ($script:OutlookMode) { "New" { 3 } "Classic" { 3 } "Both" { 4 } }
$reinstallSteps = if ($script:OutlookMode -ne "Classic") { 1 } else { 0 }
$T = $cleanupSteps + $reinstallSteps
$isSingleUser = ($TargetUsers.Count -eq 1)
$stepOffset   = 0

# If single user and using New/Both, WinGet becomes Step 1
if ($isSingleUser -and $script:OutlookMode -ne "Classic") {
    $T++
    $stepOffset = 1
}

$loopTotalSteps = $T

# --- WinGet bootstrap (New / Both only, runs once before user loop) ---
if ($script:OutlookMode -ne "Classic") {
    $wgLabel = if ($isSingleUser) { "1/$T" } else { "0/$loopTotalSteps" }
    Write-Step $wgLabel "Validating WinGet environment..."
    $WinGetTemp = Join-Path $env:TEMP "WinGet"
    if (Test-Path $WinGetTemp) { Remove-Item $WinGetTemp -Recurse -Force -ErrorAction SilentlyContinue }
    if (-not (Get-Command "winget" -ErrorAction SilentlyContinue)) {
        if (-not (Get-Module -ListAvailable Microsoft.WinGet.Client)) {
            Install-PackageProvider -Name NuGet -Force | Out-Null
            Install-Module -Name Microsoft.WinGet.Client -Force -Repository PSGallery -Scope CurrentUser -AllowClobber | Out-Null
        }
        Import-Module Microsoft.WinGet.Client
        Start-Process powershell.exe -ArgumentList "-NoProfile -Command `"Import-Module Microsoft.WinGet.Client; Repair-WinGetPackageManager -Latest`"" -WindowStyle Hidden -Wait
    } else {
        Start-Process -FilePath "cmd.exe" -ArgumentList "/c winget source reset --force && winget source update --silent" -WindowStyle Hidden -Wait
    }
    Complete-Step $wgLabel "Validating WinGet environment..."
    if ($script:AllUsers) {
        Write-HLine -Style dashed
    }
}

# --- Per-user loop ---
foreach ($UserName in $TargetUsers) {
    if ($script:AllUsers) {
        Write-Host "  User: " -ForegroundColor $LineCol -NoNewline
        Write-Host $UserName -ForegroundColor $AccentCol
    }

    switch ($script:OutlookMode) {

        "New" {
            Write-Step "$($stepOffset + 1)/$loopTotalSteps" "Killing processes and removing package..."
            Invoke-TerminateOutlook -Mode "New"
            Invoke-ClearNewOutlookData -UserName $UserName
            Complete-Step "$($stepOffset + 1)/$loopTotalSteps" "Killing processes and removing package..."

            Write-Step "$($stepOffset + 2)/$loopTotalSteps" "Cleaning registry hives..."
            Invoke-RegistryCleanup -UserName $UserName -Mode "New"
            # Extension purge not applicable to New Outlook
            Complete-Step "$($stepOffset + 2)/$loopTotalSteps" "Cleaning registry hives..."

            Write-Step "$($stepOffset + 3)/$loopTotalSteps" "Flushing authentication tokens..."
            Invoke-AuthFlush -UserName $UserName
            Complete-Step "$($stepOffset + 3)/$loopTotalSteps" "Flushing authentication tokens..."

            # Reinstall is handled after the user loop
        }

        "Classic" {
            Write-Step "1/$loopTotalSteps" "Killing processes and clearing cache..."
            Invoke-TerminateOutlook -Mode "Classic"
            Invoke-ClearClassicOutlookData -UserName $UserName
            Complete-Step "1/$loopTotalSteps" "Killing processes and clearing cache..."

            Write-Step "2/$loopTotalSteps" "Cleaning registry hives..."
            Invoke-RegistryCleanup -UserName $UserName -Mode "Classic"
            Invoke-PurgeExtensions -UserName $UserName
            Complete-Step "2/$loopTotalSteps" "Cleaning registry hives..."

            Write-Step "3/$loopTotalSteps" "Flushing authentication tokens..."
            Invoke-AuthFlush -UserName $UserName
            Complete-Step "3/$loopTotalSteps" "Flushing authentication tokens..."
        }

        "Both" {
            Write-Step "$($stepOffset + 1)/$loopTotalSteps" "Killing processes and clearing cache..."
            Invoke-TerminateOutlook -Mode "Both"
            Invoke-ClearClassicOutlookData -UserName $UserName
            Complete-Step "$($stepOffset + 1)/$loopTotalSteps" "Killing processes and clearing cache..."

            Write-Step "$($stepOffset + 2)/$loopTotalSteps" "Removing New Outlook package and data..."
            Invoke-ClearNewOutlookData -UserName $UserName
            Complete-Step "$($stepOffset + 2)/$loopTotalSteps" "Removing New Outlook package and data..."

            Write-Step "$($stepOffset + 3)/$loopTotalSteps" "Cleaning registry hives..."
            Invoke-RegistryCleanup -UserName $UserName -Mode "Both"
            Invoke-PurgeExtensions -UserName $UserName
            Complete-Step "$($stepOffset + 3)/$loopTotalSteps" "Cleaning registry hives..."

            Write-Step "$($stepOffset + 4)/$loopTotalSteps" "Flushing authentication tokens..."
            Invoke-AuthFlush -UserName $UserName
            Complete-Step "$($stepOffset + 4)/$loopTotalSteps" "Flushing authentication tokens..."

            # Reinstall is handled after the user loop
        }
    }
    if ($script:AllUsers) {
        Write-HLine -Style dashed
    }
}

# --- System-Wide Reinstall (if applicable) ---
if ($script:OutlookMode -ne "Classic") {
    if ($script:AllUsers) {
        Write-Host "  System: " -ForegroundColor $LineCol -NoNewline
        Write-Host "Global Provisioning" -ForegroundColor $AccentCol
    }

    $reinstallStepNum = $stepOffset + $cleanupSteps + 1
    $stepMsg = "Reinstalling New Outlook..."
    if ($script:AllUsers) { $stepMsg = "Provisioning New Outlook for all users..." }
    Write-Step "$reinstallStepNum/$T" $stepMsg

    $installed = $false

    # If running as SYSTEM and targeting a single, non-SYSTEM user, skip winget install
    if (-not $script:AllUsers -and ($env:USERNAME -eq "SYSTEM") -and ($TargetUsers[0] -ne "SYSTEM")) {
        $script:skipReinstall = $true
        $status = "[SKIPPED]"
        $pad = " " * [math]::Max(1, $script:Width - "[$reinstallStepNum/$T] $stepMsg".Length - $status.Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host "[$reinstallStepNum/$T]" -ForegroundColor $DimCol -NoNewline
        Write-Host " $stepMsg$pad" -ForegroundColor $DimCol -NoNewline
        Write-Host $status -ForegroundColor $WarnCol
        Write-Host "Cannot install New Outlook for '$($TargetUsers[0])' from SYSTEM context." -ForegroundColor $WarnCol
        Write-Host "Please run as the target user or select 'All Users' for provisioning." -ForegroundColor $WarnCol
    }
    
    if (-not $script:skipReinstall) {
        if ($script:AllUsers) {
            $dlDir = Join-Path $env:TEMP "ORCA_DL"
            New-Item -Path $dlDir -ItemType Directory -Force | Out-Null
            $wingetDLArgs = "download --id 9NRX63209R7B --source msstore --accept-package-agreements --silent --download-directory `"$dlDir`""
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c winget $wingetDLArgs" -WindowStyle Hidden -Wait
            
            $packagePath = Get-ChildItem -Path $dlDir -Filter "*.msix" -Recurse | Select-Object -First 1
            if ($packagePath) {
                Add-AppxProvisionedPackage -Online -PackagePath $packagePath.FullName -SkipLicense | Out-Null
            }
            Remove-Item $dlDir -Recurse -Force -ErrorAction SilentlyContinue
        } else {
            # Single-user install logic
            $wingetArgs = 'install -e --id 9NRX63209R7B --source msstore --accept-package-agreements --accept-source-agreements --silent --force --disable-interactivity'
            Start-Process -FilePath "cmd.exe" -ArgumentList "/c echo Y | winget $wingetArgs" -WindowStyle Hidden -Wait
        }

        # Verification
        $Timer = 0; $Timeout = 60
        while (-not $installed -and $Timer -lt $Timeout) {
            if ($script:AllUsers) {
                $installed = Get-AppxProvisionedPackage -Online | Where-Object { $_.DisplayName -eq "Microsoft.OutlookForWindows" }
            } else {
                $installed = Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue
            }
            if (-not $installed) { Start-Sleep -Seconds 2; $Timer += 2 }
        }
    }


    if ($installed) {
        Complete-Step "$reinstallStepNum/$T" $stepMsg
    } elseif (-not $script:skipReinstall) { # Only report FAILED if it wasn't explicitly skipped
        $status = "[FAILED]"
        $pad = " " * [math]::Max(1, $script:Width - "[$reinstallStepNum/$T] $stepMsg".Length - $status.Length)
        Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
        Write-Host "[$reinstallStepNum/$T]" -ForegroundColor $DimCol -NoNewline
        Write-Host " $stepMsg$pad" -ForegroundColor $DimCol -NoNewline
        Write-Host $status -ForegroundColor "Red"
        exit $EXIT_ERROR
        Write-Host "      > Winget/Provisioning failed. Try manual install from Store." -ForegroundColor $WarnCol
    }
}
#endregion

#region [4] - RESULTS
# ============================================================================
Write-HLine -Style dashed
if (-not $script:skipReinstall) {
    $_status = "System ready for fresh login."
    $_padL = " " * [math]::Floor(($script:Width - $_status.Length) / 2)
    Write-Host "$_padL$_status" -ForegroundColor $OkCol
}

if ($script:DeletedOSTs.Count -gt 0) {
    Write-Host "Deleted OSTs :" -ForegroundColor $LineCol
    foreach ($ostPath in $script:DeletedOSTs) {
        Write-Host "  - $ostPath" -ForegroundColor $DimCol
    }
}

if ($script:PurgedExtensions.Count -gt 0) {
    Write-Host "Purged Extensions :" -ForegroundColor $LineCol
    foreach ($ext in $script:PurgedExtensions) {
        Write-Host "  - $ext" -ForegroundColor $DimCol
    }
}

if (-not $script:AllUsers) {
    if (-not $script:skipReinstall) {
        if ([Security.Principal.WindowsIdentity]::GetCurrent().IsSystem) {
            # This case may not be hit if IsSystem check is unreliable, but keep for robustness
            Write-Host ""
            Write-Host "NOTE: Script running as 'SYSTEM', target is '$($TargetUsers[0])'." -ForegroundColor $WarnCol
        } elseif ($TargetUsers[0] -ne $env:USERNAME) {
            Write-Host ""
            Write-Host "NOTE: Script running as '$env:USERNAME', target is '$($TargetUsers[0])'." -ForegroundColor $WarnCol
            Write-Host "Please launch Outlook manually as the target user." -ForegroundColor $DimCol
        } else {
            $launchChoice = $null
            Write-Host ""
            if ($script:OutlookMode -eq "New") {
                Write-Host "Launch New Outlook?" -ForegroundColor $InfoCol
                Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
                Write-Host "[Y/N]  Enter = Yes  Esc = quit: " -ForegroundColor $DimCol -NoNewline
                try {
                    $key = [Console]::ReadKey($true)
                    Write-Host $key.KeyChar
                    if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
                    $yn = if ($key.Key -eq [ConsoleKey]::Enter) { "Y" } else { $key.KeyChar.ToString() }
                    if ($yn -eq "Y" -or $yn -eq "y") { $launchChoice = "2" }
                } catch {
                    $raw = (Read-Host).Trim(); if ($raw -eq "" -or $raw -eq "Y" -or $raw -eq "y") { $launchChoice = "2" }
                }
            } elseif ($script:OutlookMode -eq "Classic") {
                Write-Host "Launch Classic Outlook?" -ForegroundColor $InfoCol
                Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
                Write-Host "[Y/N]  Enter = Yes  Esc = quit: " -ForegroundColor $DimCol -NoNewline
                try {
                    $key = [Console]::ReadKey($true)
                    Write-Host $key.KeyChar
                    if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
                    $yn = if ($key.Key -eq [ConsoleKey]::Enter) { "Y" } else { $key.KeyChar.ToString() }
                    if ($yn -eq "Y" -or $yn -eq "y") { $launchChoice = "1" }
                } catch {
                    $raw = (Read-Host).Trim(); if ($raw -eq "" -or $raw -eq "Y" -or $raw -eq "y") { $launchChoice = "1" }
                }
            } else {
                Write-Host "Launch Outlook?" -ForegroundColor $LineCol
                Write-Host "  [1] " -ForegroundColor $DimCol -NoNewline; Write-Host "Classic Outlook" -ForegroundColor $AccentCol
                Write-Host "  [2] " -ForegroundColor $DimCol -NoNewline; Write-Host "New Outlook" -ForegroundColor $AccentCol
                Write-Host ""
                Write-Host "Choice " -ForegroundColor $LineCol -NoNewline
                Write-Host "[1-2]  Enter = No  Esc = quit: " -ForegroundColor $DimCol -NoNewline
                try {
                    $key = [Console]::ReadKey($true)
                    Write-Host $key.KeyChar
                    if ($key.Key -eq [ConsoleKey]::Escape) { Write-Host "`nExiting." -ForegroundColor $DimCol; Exit $EXIT_CANCEL }
                    $launchChoice = if ($key.Key -eq [ConsoleKey]::Enter) { "N" } else { $key.KeyChar.ToString() }
                } catch {
                    $raw = (Read-Host).Trim()
                    if ([string]::IsNullOrWhiteSpace($raw)) { $launchChoice = "N" } else { $raw }
                }
            }

            if ($launchChoice -eq "1") {
                $classicPath = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue).'(default)'
                if ($classicPath -and (Test-Path $classicPath)) {
                    Start-Process $classicPath -ErrorAction SilentlyContinue
                } else { Write-Warning "Classic Outlook (OUTLOOK.EXE) not found." }
            } elseif ($launchChoice -eq "2") {
                if (Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue) {
                    try {
                        Start-Process "shell:appsfolder\Microsoft.OutlookForWindows_8wekyb3d8bbwe!olk" -ErrorAction Stop
                    } catch {
                        $alias = "$env:LOCALAPPDATA\Microsoft\WindowsApps\olk.exe"
                        if (Test-Path $alias) {
                            Start-Process $alias -ErrorAction SilentlyContinue
                        } else {
                            Write-Warning "New Outlook is installed but could not be launched. Please open it manually."
                        }
                    }
                } else {
                    Write-Warning "New Outlook installation was not detected."
                }
            }
        }
    }
}

# Footer
$_sfx    = "█"
$_ftrW   = $_artW + 1
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  REPAIR AND CONFIGURATION SEQUENCE COMPLETE"
$_fpad   = " " * ($_ffillW - $_footer.Length - $_ver.Length)
Write-Host ("━" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("━" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol

exit $EXIT_SUCCESS
#endregion
