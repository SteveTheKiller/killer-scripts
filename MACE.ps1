<#
.SYNOPSIS
    Microsoft Application Cleanse & Eradication (MACE) v1.1
    Developed by SteveTheKiller | Updated: 2026-03-06
.DESCRIPTION
    Completely removes OneDrive, New Outlook, Office/M365, Microsoft
    Project, and Microsoft Teams from a Windows endpoint. Clears all
    associated registry keys, cached credentials, profile data, and
    temp folders. Repairs OneDrive shell folder redirects and path
    pollution. Supports an interactive per-user or system-wide scope
    selection with ESC-to-abort.
#>

#region 0: INITIALIZATION
# ============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit 1
}

$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo, [switch]$Pending)
    if ($Success) {
        $savedTop = [Console]::CursorTop
        if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
            [Console]::SetCursorPosition(0, $script:StepRow)
            $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
            Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
            [Console]::SetCursorPosition(0, $savedTop)
        }
        $script:StepRow = -1
    } else {
        $script:StepRow = [Console]::CursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $InfoCol
    }
}

function Write-SubResult {
    param([string]$Tag, [string]$Color)
    $_pad = " " * [math]::Max(1, $script:Width - [Console]::CursorLeft - $Tag.Length)
    Write-Host "$_pad$Tag" -ForegroundColor $Color
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
function Write-Banner {
    param([string]$Text, [switch]$NoNewline)
    $esc = [char]27
    if ($script:BannerBG -eq "Black") {
        Write-Host "${esc}[40m$Text${esc}[0m" -NoNewline:$NoNewline -ForegroundColor $script:BannerFG
    } else {
        Write-Host $Text -NoNewline:$NoNewline -ForegroundColor $script:BannerFG -BackgroundColor $script:BannerBG
    }
}
Clear-Host
$script:Width    = 85

$LineCol   = "Red"
$MainCol   = "Gray"
$WarnCol   = "DarkYellow"
$BorderCol = "Red"
$ArtCol    = "White"
$AccentCol = "Red"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"
$_ver    = "| v1.2"

# Header Art & Logic
$_pfx  = "█  "
$_art1 = "╔╦╗ ╔═╗ ╔═╗ ╔═╗ "
$_art2 = "║║║ ╠═╣ ║   ║╣  "
$_art3 = "╩ ╩ ╩ ╩ ╚═╝ ╚═╝ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "MICROSOFT APPLICATION CLEANSE & ERADICATION"

# Safety padding to prevent negative multiplication error
$_tpad  = " " * [Math]::Max(0, ($_fillW - $_title.Length))

Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
#endregion

#region 1: SCOPE SELECTION
# ============================================================================
$IsInteractive = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected
$ExcludedProfiles = @("Public", "Default", "Default User", "All Users")
$DetectedProfiles = Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -notin $ExcludedProfiles }

Write-Host "[>] Device  : " -NoNewline -ForegroundColor $LineCol
Write-Host $env:COMPUTERNAME -ForegroundColor $MainCol

$menuRow = [Console]::CursorTop
if ($IsInteractive) {
    Write-Host "[?] Select target user profile:" -ForegroundColor Yellow
    Write-Host ""
    for ($i = 0; $i -lt $DetectedProfiles.Count; $i++) {
        Write-Host "    " -NoNewline
        Write-Host "[$i]" -NoNewline -ForegroundColor $AccentCol
        Write-Host " $($DetectedProfiles[$i].Name)" -ForegroundColor White
    }
    Write-Host "    " -NoNewline
    Write-Host '[A]' -NoNewline -ForegroundColor $AccentCol
    Write-Host " All Users (system-wide) " -NoNewline -ForegroundColor White
    Write-Host '[Default]' -ForegroundColor $DimCol
    Write-Host "    " -NoNewline
    Write-Host '[ESC]' -NoNewline -ForegroundColor $DimCol
    Write-Host " Exit" -ForegroundColor $DimCol
    Write-Host ""
    Write-Host "Enter choice: " -NoNewline -ForegroundColor $AccentCol
    $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    if ($key.VirtualKeyCode -eq 27) {
        Write-Host ""
        Write-Host "[!] Aborted." -ForegroundColor $DimCol
        Exit 0
    }
    $ScopeChoice = $key.Character.ToString()

    if ($ScopeChoice -match '^\d+$' -and [int]$ScopeChoice -lt $DetectedProfiles.Count) {
        $SelectedProfile = $DetectedProfiles[[int]$ScopeChoice]
        $TargetProfiles  = @($SelectedProfile.FullName)
        $AllUsers        = $false
        $ScopeLabel      = "$($SelectedProfile.Name)  ($($SelectedProfile.FullName))"
    } else {
        $TargetProfiles = $DetectedProfiles | Select-Object -ExpandProperty FullName
        $AllUsers       = $true
        $ScopeLabel     = "All Users (System-Wide)"
    }
} else {
    $TargetProfiles = $DetectedProfiles | Select-Object -ExpandProperty FullName
    $AllUsers       = $true
    $ScopeLabel     = "All Users (System-Wide)"
    Write-Host "[>] Non-interactive session: defaulting to All Users scope." -ForegroundColor Yellow
}

if ($IsInteractive) {
    $endRow = [Console]::CursorTop
    for ($r = $menuRow; $r -le $endRow; $r++) {
        [Console]::SetCursorPosition(0, $r)
        [Console]::Write(" " * $script:Width)
    }
    [Console]::SetCursorPosition(0, $menuRow)
}
Write-Host "[>] Scope   : " -NoNewline -ForegroundColor Red
Write-Host $ScopeLabel -ForegroundColor Yellow
Write-HLine -Style dashed
#endregion

#region 2: PROCESS TERMINATION
# ============================================================================
Write-StepUpdate "[1/5] Terminating Active Processes..." -Pending
$Processes = @(
    "OneDrive", "OneDriveSetup",
    "olk", "Outlook",
    "Teams", "ms-teams", "Update",
    "WINWORD", "EXCEL", "POWERPNT", "ONENOTE", "MSACCESS", "MSPUB", "WINPROJ",
    "OfficeClickToRun", "OfficeC2RClient",
    "msedgewebview2", "WebViewHost"
)
foreach ($proc in $Processes) {
    Get-Process -Name $proc -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
}
Start-Sleep -Seconds 3
Write-StepUpdate -Success
#endregion

#region 3: APPLICATION UNINSTALL
# ============================================================================
Write-StepUpdate "[2/5] Uninstalling Applications..."
$winget = Get-Command winget -ErrorAction SilentlyContinue

# Determine if the single target profile (if any) is the currently logged-in user
$TargetIsSelf = (-not $AllUsers) -and ($TargetProfiles[0] -eq $env:USERPROFILE)

# --- OneDrive ---
Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " OneDrive..." -NoNewline -ForegroundColor Gray
$ODSetups = @(
    "$env:SYSTEMROOT\System32\OneDriveSetup.exe",
    "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe",
    "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe"
)
# System32/SysWOW64 OneDriveSetup.exe ships with Windows itself - not an install indicator
$ODFound = (Test-Path "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe") -or
           (Test-Path "HKLM:\SOFTWARE\Microsoft\OneDrive") -or
           ($null -ne ($TargetProfiles | Where-Object { Test-Path "$_\AppData\Local\Microsoft\OneDrive" } | Select-Object -First 1))
if ($ODFound) {
    $ODArgs = if ($AllUsers) { "/uninstall /allusers" } else { "/uninstall" }
    foreach ($setup in $ODSetups) {
        if (Test-Path $setup) {
            Start-Process $setup -ArgumentList $ODArgs -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
            break
        }
    }
    if ($winget) { & winget uninstall --id Microsoft.OneDrive --silent --accept-source-agreements 2>$null | Out-Null }
    Write-SubResult "[DONE]" $OkCol
} else {
    Write-SubResult "[NOT-INSTALLED]" $InfoCol
}

# --- New Outlook ---
Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " New Outlook (olk)..." -NoNewline -ForegroundColor Gray
$OlkFound = $null -ne (Get-AppxPackage -AllUsers -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue)
if ($OlkFound) {
    if ($AllUsers) {
        Get-AppxPackage -AllUsers -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*OutlookForWindows*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    } elseif ($TargetIsSelf) {
        Get-AppxPackage -Name "Microsoft.OutlookForWindows" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
    } else {
        Write-Host " (AppX skipped - target user not active session)" -NoNewline -ForegroundColor DarkGray
    }
    Write-SubResult "[DONE]" $OkCol
} else {
    Write-SubResult "[NOT-INSTALLED]" $InfoCol
}

# --- Office / M365 ---
Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " Office / M365 / Project..." -NoNewline -ForegroundColor Gray
$C2RSetup = @(
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\setup64.exe",
    "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\setup64.exe"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
$OfficeMSIPkgs = @(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -match "Microsoft (Office|365|Project)" -and $_.UninstallString -match "msiexec" })
$OfficeFound = ($null -ne $C2RSetup) -or ($OfficeMSIPkgs.Count -gt 0)

if ($OfficeFound) {
    if ($C2RSetup) {
        Write-Host " (ClickToRun found...)" -NoNewline -ForegroundColor DarkGray
        $XMLPath = "$env:TEMP\MACE_OfficeRemove.xml"
        @"
<Configuration>
  <Remove All="TRUE"/>
  <Display Level="None" AcceptEULA="TRUE"/>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
</Configuration>
"@ | Set-Content $XMLPath -Encoding UTF8
        $proc = Start-Process $C2RSetup -ArgumentList "/configure `"$XMLPath`"" -PassThru -WindowStyle Hidden
        $proc | Wait-Process -Timeout 900 -ErrorAction SilentlyContinue
        if (-not $proc.HasExited) { $proc | Stop-Process -Force -ErrorAction SilentlyContinue }
        Remove-Item $XMLPath -Force -ErrorAction SilentlyContinue
    } else {
        Write-Host " (no local C2R, using MSI fallback)" -NoNewline -ForegroundColor DarkGray
    }
    # MSI fallback for any remaining installs
    $OfficePkgs = @(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.DisplayName -match "Microsoft (Office|365|Project)" -and $_.UninstallString -match "msiexec" })
    foreach ($pkg in $OfficePkgs) {
        $guid = [regex]::Match($pkg.UninstallString, '\{[A-Z0-9-]+\}').Value
        if ($guid) { Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue }
    }
    Write-SubResult "[DONE]" $OkCol
} else {
    Write-SubResult "[NOT-INSTALLED]" $InfoCol
}

# --- Teams ---
Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " Microsoft Teams..." -NoNewline -ForegroundColor Gray
$TeamsPkgs = @(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*","HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue |
    Where-Object { $_.DisplayName -like "*Teams Machine-Wide*" -or $_.DisplayName -eq "Microsoft Teams" })
$TeamsAppx = @(Get-AppxPackage -AllUsers -Name "MSTeams"        -ErrorAction SilentlyContinue) +
             @(Get-AppxPackage -AllUsers -Name "MicrosoftTeams" -ErrorAction SilentlyContinue)
$TeamsFound = ($TeamsPkgs.Count -gt 0) -or ($TeamsAppx.Count -gt 0)

if ($TeamsFound) {
    foreach ($pkg in $TeamsPkgs) {
        if ($pkg.UninstallString -match "msiexec") {
            $guid = [regex]::Match($pkg.UninstallString, '\{[A-Z0-9-]+\}').Value
            if ($guid) { Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue }
        }
    }
    if ($AllUsers) {
        Get-AppxPackage -AllUsers -Name "MSTeams"        -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxPackage -AllUsers -Name "MicrosoftTeams" -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
        Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like "*Teams*" } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
    } elseif ($TargetIsSelf) {
        Get-AppxPackage -Name "MSTeams"        -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
        Get-AppxPackage -Name "MicrosoftTeams" -ErrorAction SilentlyContinue | Remove-AppxPackage -ErrorAction SilentlyContinue
    } else {
        Write-Host " (AppX skipped - target user not active session)" -NoNewline -ForegroundColor DarkGray
    }
    if ($winget) {
        & winget uninstall --id Microsoft.Teams    --silent --accept-source-agreements 2>$null | Out-Null
        & winget uninstall --name "Microsoft Teams" --silent --accept-source-agreements 2>$null | Out-Null
    }
    Write-SubResult "[DONE]" $OkCol
} else {
    Write-SubResult "[NOT-INSTALLED]" $InfoCol
}
#endregion

Write-StepUpdate -Success
#region 4: FILESYSTEM PURGE
# ============================================================================
Write-StepUpdate "[3/5] Purging Residual Files & Folders..." -Pending
$SystemFolders = @(
    "$env:PROGRAMDATA\Microsoft OneDrive",
    "C:\Program Files\Microsoft OneDrive",
    "C:\Program Files\Microsoft Office",
    "C:\Program Files (x86)\Microsoft Office",
    "C:\Program Files\Common Files\Microsoft Shared\ClickToRun",
    "C:\ProgramData\Microsoft\ClickToRun",
    "C:\Program Files\Microsoft\Teams",
    "C:\Program Files (x86)\Microsoft\Teams",
    "C:\Program Files (x86)\Microsoft\TeamsMachineInstaller"
)
$RelUserPaths = @(
    "OneDrive",
    "AppData\Local\Microsoft\OneDrive",
    "AppData\Local\Microsoft\Olk",
    "AppData\Local\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe",
    "AppData\Local\Microsoft\Outlook\HubAppCache",
    "AppData\Local\Microsoft\Office",
    "AppData\Roaming\Microsoft\Office",
    "AppData\Roaming\Microsoft\Outlook",
    "AppData\Local\Microsoft\Teams",
    "AppData\Roaming\Microsoft\Teams",
    "AppData\Local\Microsoft\IdentityCache",
    "AppData\Local\Microsoft\OneAuth",
    "AppData\Local\Microsoft\TokenBroker\Cache",
    "AppData\Local\Microsoft\Edge\User Data\Default\EBWebView",
    "AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy"
)
foreach ($folder in $SystemFolders) {
    if (Test-Path $folder) { Remove-Item $folder -Recurse -Force -ErrorAction SilentlyContinue }
}
foreach ($profilePath in $TargetProfiles) {
    foreach ($rel in $RelUserPaths) {
        $full = Join-Path $profilePath $rel
        if (Test-Path $full) { Remove-Item $full -Recurse -Force -ErrorAction SilentlyContinue }
    }
}
Write-StepUpdate -Success
#endregion

#region 5: REGISTRY, PATHS & SHELL FOLDER REPAIR
# ============================================================================
Write-StepUpdate "[4/5] Registry, Paths & Shell Folder Repair..." -Pending

# Machine-level keys
$MachineKeys = @(
    "HKLM:\SOFTWARE\Microsoft\OneDrive",
    "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun",
    "HKLM:\SOFTWARE\Microsoft\Office\16.0",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0",
    "HKLM:\SOFTWARE\Microsoft\AppVISV",
    "HKLM:\SOFTWARE\Microsoft\Teams",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams",
    "HKLM:\SOFTWARE\WOW6432Node\Microsoft\MicrosoftTeams"
)
foreach ($key in $MachineKeys) {
    if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
}

$UserRegKeys   = @(
    "Software\Microsoft\OneDrive",
    "Software\Microsoft\Office",
    "Software\Microsoft\Teams",
    "Software\Microsoft\Olk",
    "Software\Microsoft\Office\16.0\Outlook"
)
$UserRunValues = @("OneDrive", "Teams", "com.squirrel.Teams.Teams")
$ShellFixes    = @(
    @{ Name = "Personal";    Default = "%USERPROFILE%\Documents"; Dir = "Documents" },
    @{ Name = "Desktop";     Default = "%USERPROFILE%\Desktop";   Dir = "Desktop"   },
    @{ Name = "My Pictures"; Default = "%USERPROFILE%\Pictures";  Dir = "Pictures"  },
    @{ Name = "My Music";    Default = "%USERPROFILE%\Music";     Dir = "Music"     },
    @{ Name = "My Video";    Default = "%USERPROFILE%\Videos";    Dir = "Videos"    }
)

function Invoke-PerUserRegistry {
    param([string]$HiveRoot, [string]$ProfilePath)
    foreach ($key in $UserRegKeys) {
        $full = "$HiveRoot\$key"
        if (Test-Path $full) { Remove-Item $full -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $runPath = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Run"
    foreach ($val in $UserRunValues) {
        if (Get-ItemProperty -Path $runPath -Name $val -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path $runPath -Name $val -ErrorAction SilentlyContinue
        }
    }
    $USFPath = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $SFPath  = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    foreach ($fix in $ShellFixes) {
        $current = (Get-ItemProperty -Path $USFPath -Name $fix.Name -ErrorAction SilentlyContinue).$($fix.Name)
        if ($current -like "*OneDrive*") {
            Set-ItemProperty -Path $USFPath -Name $fix.Name -Value $fix.Default -ErrorAction SilentlyContinue
            if ($ProfilePath) {
                $absPath = Join-Path $ProfilePath $fix.Dir
                Set-ItemProperty -Path $SFPath -Name $fix.Name -Value $absPath -ErrorAction SilentlyContinue
                New-Item -ItemType Directory -Path $absPath -Force -ErrorAction SilentlyContinue | Out-Null
            }
        }
    }
    $navKey = "$HiveRoot\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    if (Test-Path $navKey) {
        Set-ItemProperty -Path $navKey -Name "System.IsPinnedToNameSpaceTree" -Value 0 -ErrorAction SilentlyContinue
    }
}

# Match selected profiles to ProfileList entries for hive loading
$ProfileList = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -ErrorAction SilentlyContinue |
    Where-Object {
        $_.PSChildName -match "S-1-5-21" -and
        (Test-Path "$($_.ProfileImagePath)\NTUSER.DAT") -and
        ($TargetProfiles -contains $_.ProfileImagePath)
    }

foreach ($prof in $ProfileList) {
    $HiveName   = "MACE_$($prof.PSChildName)"
    reg load "HKU\$HiveName" "$($prof.ProfileImagePath)\NTUSER.DAT" 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        # Hive loaded cleanly (user not logged in)
        Invoke-PerUserRegistry -HiveRoot "Registry::HKU\$HiveName" -ProfilePath $prof.ProfileImagePath
        [GC]::Collect()
        Start-Sleep -Milliseconds 500
        reg unload "HKU\$HiveName" 2>$null | Out-Null
    } elseif (Test-Path "Registry::HKU\$($prof.PSChildName)") {
        # User is currently logged in - their hive is already mounted under their SID
        Invoke-PerUserRegistry -HiveRoot "Registry::HKU\$($prof.PSChildName)" -ProfilePath $prof.ProfileImagePath
    }
    # If neither succeeded, the profile is locked by another process - skip silently
}

# Repair PSModulePath - strip dead/OneDrive-redirected entries
$env:PSModulePath = ($env:PSModulePath -split ';' | Where-Object { $_ -and (Test-Path $_) }) -join ';'

Write-StepUpdate -Success
#endregion

#region 6: CREDENTIAL FLUSH
# ============================================================================
Write-StepUpdate "[5/5] Flushing Cached Credentials..." -Pending
cmdkey /list 2>$null | ForEach-Object {
    if ($_ -match "Target:\s*(.+)") {
        $target = $Matches[1].Trim()
        if ($target -match "Office|Outlook|Teams|OneDrive|MicrosoftOffice|MicrosoftTeams|msteams|lync") {
            cmdkey /delete:$target 2>$null | Out-Null
        }
    }
}
Write-StepUpdate -Success
#endregion

# Footer (reuses header art - art on right, fill on left)
$_ffillW = $script:Width - $_artW - 2  # 1 leading space + 1 sfx char
$_footer = "  MACE SEQUENCE COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
Exit 0
