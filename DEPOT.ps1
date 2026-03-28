<#
.SYNOPSIS
    Deployment & Endpoint Provisioning Operations Tool (D.E.P.O.T.) v1.2
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Full new-machine provisioning tool that silently deploys M365,
    Teams, OneDrive, Chrome, Acrobat Reader, Zoom, and 7-Zip with
    'ESC-to-skip' support. Also runs Windows Update, applies privacy/UI
    hardening across all profiles, and self-deletes on completion.
#>

#region 0: PRE-FLIGHT CHECKS
# --- SELF-ELEVATION: Re-launch as full administrator if not already elevated ---
$Identity  = [Security.Principal.WindowsIdentity]::GetCurrent()
$Principal = [Security.Principal.WindowsPrincipal]$Identity
if (-not $Principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "[!] Not running as Administrator. Relaunching elevated..." -ForegroundColor Yellow
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}
# Confirm token is fully elevated (handles domain UAC split-token scenarios)
$TokenElevation = [System.Security.Principal.WindowsIdentity]::GetCurrent().Claims |
    Where-Object { $_.Type -eq "http://schemas.microsoft.com/ws/2008/06/identity/claims/groupsid" -and $_.Value -eq "S-1-5-32-544" }
if (-not $TokenElevation) {
    # Force a new elevated process using explicit RunAs to get a full admin token
    $scriptPath = $MyInvocation.MyCommand.Path
    Start-Process powershell.exe -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"$scriptPath`"" -Verb RunAs
    exit
}
# --- STARTUP CLEANUP: kill leftovers from any previous cancelled/crashed run ---
Write-Host "[*] Checking for leftover processes from previous run..." -ForegroundColor DarkGray
# Kill any stale installer/maintenance processes
$StaleProcs = @(
    "winget", "WinGet",
    "OfficeClickToRun", "OfficeC2RClient", "setup",
    "msiexec",
    "TrustedInstaller", "wusa", "wuauclt",
    "dism", "DismHost",
    "sfc",
    "chrome", "ChromeSetup",
    "AcroRd32", "AcroRdrDC", "Reader_sl",
    "7zG", "7zFM",
    "Teams", "Update",
    "Zoom", "ZoomInstaller"
)
$killed = @()
foreach ($name in $StaleProcs) {
    $procs = Get-Process -Name $name -ErrorAction SilentlyContinue |
        Where-Object { $_.Path -notmatch "System32|SysWOW64|WindowsApps" -or $name -match "msiexec|dism|sfc|TrustedInstaller|wusa" }
    foreach ($p in $procs) {
        $p | Stop-Process -Force -ErrorAction SilentlyContinue
        $killed += $p.Name
    }
}
if ($killed.Count -gt 0) {
    Write-Host "    [!] Killed: $($killed -join ', ')" -ForegroundColor Yellow
} else {
    Write-Host "    > No stale processes found." -ForegroundColor DarkGray
}
# Stop Windows Update service if left running
$wuSvc = Get-Service -Name wuauserv -ErrorAction SilentlyContinue
if ($wuSvc -and $wuSvc.Status -eq 'Running') {
    Stop-Service -Name wuauserv -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
    Write-Host "    [!] Stopped leftover wuauserv." -ForegroundColor Yellow
}
# Unload any APT registry hives left mounted from a previous run
$mountedHives = (reg query HKU 2>$null) -split "`n" | Where-Object { $_ -match "APT_" }
foreach ($hive in $mountedHives) {
    $hiveName = ($hive -split "\\")[-1].Trim()
    reg unload "HKU\$hiveName" 2>$null | Out-Null
    Write-Host "    [!] Unloaded orphaned hive: $hiveName" -ForegroundColor Yellow
}
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
#endregion

#region 1: INTERACTIVE DEVICE SETUP
# ------------------------------------------------------------------------------
Clear-Host
Write-Host "[*] Interactive Pre-Deployment Setup..." -ForegroundColor DarkRed -BackgroundColor Black
Write-Host ""
Write-Host "    Note: All prompts default to 'N' if Enter is pressed." -ForegroundColor DarkRed -BackgroundColor Black
Write-Host ""
# 1.1: Machine Renaming
Write-Host "1) Do you want to rename the machine? (y/N) " -ForegroundColor Black -BackgroundColor Red -NoNewline
$_k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); $RenameChoice = $_k.Character.ToString(); Write-Host $RenameChoice
if ($RenameChoice -eq "Y" -or $RenameChoice -eq "y") {
    $CurrentName = $env:COMPUTERNAME
    $NewName = Read-Host "   > New Device Name"
    if ($NewName -and $NewName -ne $CurrentName) {
        Write-Host "[>] Renaming computer from $CurrentName to $NewName..." -NoNewline
        Rename-Computer -NewName $NewName -Force -ErrorAction SilentlyContinue
        Write-Host " Done." -ForegroundColor Green
        Write-Host "[!] Note: Name change requires a reboot to take effect." -ForegroundColor Red
    } else {
        Write-Host "[!] Invalid name or same as current. Skipping." -ForegroundColor Gray
    }
}
# 1.2: Local Admin Creation
Write-Host ""
Write-Host "2) Do you want to make a local admin account? (y/N) " -ForegroundColor Black -BackgroundColor Red -NoNewline
$_k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); $AdminChoice = $_k.Character.ToString(); Write-Host $AdminChoice
if ($AdminChoice -eq "Y" -or $AdminChoice -eq "y") {
    $NewAdminName = Read-Host "   > Local Admin Username"
    $Password = Read-Host "   > Local Admin Password" -AsSecureString
    if ($NewAdminName -and $Password) {
        Write-Host "[>] Creating local account: $NewAdminName..." -NoNewline
        try {
            New-LocalUser -Name $NewAdminName -Password $Password -Description "Local Admin created" -ErrorAction Stop
            Add-LocalGroupMember -Group "Administrators" -Member $NewAdminName -ErrorAction Stop
            Write-Host " Done." -ForegroundColor Green
        } catch {
            Write-Host " Failed." -ForegroundColor Red
            Write-Warning $_.Exception.Message
        }
    } else {
        Write-Host "[!] Username or Password cannot be empty. Skipping." -ForegroundColor Gray
    }
}
# 1.3: Purge & Secure Admin Accounts
Write-Host ""
Write-Host "3) Do you want to purge other admins and disable built-in Administrator? (y/N) " -ForegroundColor Black -BackgroundColor Red -NoNewline
$_k = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown"); $PurgeChoice = $_k.Character.ToString(); Write-Host $PurgeChoice
if ($PurgeChoice -eq "Y" -or $PurgeChoice -eq "y") {
    Write-Host "[>] Disabling built-in Administrator account..." -NoNewline
    Disable-LocalUser -Name "Administrator" -ErrorAction SilentlyContinue
    Write-Host " Done." -ForegroundColor Green

    Write-Host "[>] Scanning for other local administrators to delete..." -NoNewline
    $AdminsToPurge = Get-LocalGroupMember -Group "Administrators" | Where-Object { 
        $_.Name -notlike "*\Administrator" -and 
        $_.Name -notlike "*\$NewAdminName" -and 
        $_.ObjectClass -eq "User" 
    }
    if ($AdminsToPurge) {
        Write-Host " Found $($AdminsToPurge.Count)." -ForegroundColor Cyan
        foreach ($Admin in $AdminsToPurge) {
            Write-Host "   [!] Deleting account: $($Admin.Name)..." -NoNewline
            Remove-LocalUser -Name $Admin.Name -ErrorAction SilentlyContinue
            Write-Host " Done." -ForegroundColor Green
        }
    } else {
        Write-Host " None found." -ForegroundColor Green
    }
}
#endregion

#region 2: WELCOME BANNER
# ------------------------------------------------------------------------------
Clear-Host
$script:Width  = 85

$LineCol   = "DarkRed"
$MainCol   = "Green"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"

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
        [ConsoleColor]$LineCol,
        [ConsoleColor]$AccentCol,
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

# Header
$_pfx  = "█  "
$_art1 = "╔╦╗ ╔═╗ ╔═╗ ╔═╗ ╔╦╗ "
$_art2 = " ║║ ║╣  ╠═╝ ║ ║  ║  "
$_art3 = "╩╩╝ ╚═╝ ╩   ╚═╝  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "DEPLOYMENT & ENDPOINT PROVISIONING OPERATIONS TOOL"
$_tpad  = " " * ($_fillW - $_title.Length)
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $AccentCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $AccentCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $AccentCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host "[>] Device Name    : " -NoNewline -ForegroundColor $LineCol
Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $AccentCol
$Admins = Get-LocalGroupMember -Group "Administrators" | Where-Object { $_.ObjectClass -eq "User" }
$AdminPrefix = "[>] Local Admins   : "
Write-Host $AdminPrefix -NoNewline -ForegroundColor $LineCol
$AdminList = $Admins.Name
if ($AdminList -is [string]) { $AdminList = @($AdminList) }
$CurrentLine = ""; $MaxLen = $script:Width - $AdminPrefix.Length
foreach ($Name in $AdminList) {
    $Item = if ($Name -eq $AdminList[-1]) { $Name } else { "$Name, " }
    if (($CurrentLine.Length + $Item.Length) -gt $MaxLen) {
        Write-Host $CurrentLine -ForegroundColor $AccentCol
        Write-Host (" " * $AdminPrefix.Length) -NoNewline; $CurrentLine = $Item
    } else { $CurrentLine += $Item }
}
if ($CurrentLine) { Write-Host $CurrentLine -ForegroundColor $AccentCol }
Write-HLine -Style "dashed"
#endregion

#region 3: WINDOWS UPDATE AUTOMATION
# ------------------------------------------------------------------------------
# WU COM API deadlocks on fresh VMs regardless of approach (direct COM, PSWindowsUpdate, etc.)
# UsoClient triggers the built-in Update Orchestrator in the background and returns immediately.
# Updates will download and install on their own; a reboot is almost always required after.
Write-Host "[*] Deploying Windows & Microsoft Updates..." -ForegroundColor DarkRed
$RebootRequired = $true

try {
    # Reset WU service state first to clear any stuck session
    Stop-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Remove-Item "C:\Windows\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
    Start-Service wuauserv -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3

    # Kick off the full WU cycle via Update Session Orchestrator - fire and forget
    Start-Process "UsoClient.exe" -ArgumentList "StartScan"    -WindowStyle Hidden
    Start-Sleep -Seconds 2
    Start-Process "UsoClient.exe" -ArgumentList "StartDownload" -WindowStyle Hidden
    Start-Sleep -Seconds 2
    Start-Process "UsoClient.exe" -ArgumentList "StartInstall"  -WindowStyle Hidden

    Write-Host "      > Windows Update triggered in background." -ForegroundColor Gray
    Write-Host "      > Updates will continue in background after script exit." -ForegroundColor Gray
} catch {
    Write-Host "      > Failed to trigger Windows Update: $_" -ForegroundColor Red
    Write-Host "      > Run Windows Update manually after reboot." -ForegroundColor Yellow
}
#endregion

#region 4: APP PROVISIONING
# ------------------------------------------------------------------------------
Write-Host "[*] Executing Application Provisioning..." -ForegroundColor DarkRed
Write-Host "    (Press 'ESC' at any time to skip current install)" -ForegroundColor DarkGray
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
# Helper: download a file silently with retry
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
# Helper: run an installer with ESC-to-cancel and timeout, return exit code
function Start-Installer {
    param(
        [string]$FilePath,
        [string]$ArgumentList,
        [int]$TimeoutMinutes = 20
    )
    # Always wrap in cmd /c - this gives us a single reliable process handle.
    # Direct Start-Process -PassThru hangs on WaitForExit/HasExited when the
    # launched EXE spawns children that inherit the parent handle (ODT, Acrobat, etc.)
    $parentDir = Split-Path $FilePath -Parent -ErrorAction SilentlyContinue
    $workDir   = if ($parentDir -and (Test-Path $parentDir)) { $parentDir } else { "C:\Windows\Temp" }
    if ($FilePath -match "msiexec") {
        $proc = Start-Process "cmd.exe" -ArgumentList "/c msiexec $ArgumentList" -PassThru -WindowStyle Hidden -WorkingDirectory $workDir
    } else {
        $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru -WindowStyle Hidden -WorkingDirectory $workDir
    }
    $deadline = (Get-Date).AddMinutes($TimeoutMinutes)
    $start    = Get-Date
    $spinner  = @('|','/','-','\')
    $spin     = 0
    while (-not $proc.WaitForExit(500)) {
        if ((Get-Date) -gt $deadline) {
            $proc | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Host " [TIMEOUT after $TimeoutMinutes min]" -ForegroundColor Red
            return -1
        }
        try {
        if ($Host.UI.RawUI.KeyAvailable) {
            $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            if ($key.VirtualKeyCode -eq 27) {
                $proc | Stop-Process -Force -ErrorAction SilentlyContinue
                Get-Process | Where-Object { $_.Name -match "msiexec|setup|OfficeClickToRun|chrome|acro|7z|Zoom" } |
                    Stop-Process -Force -ErrorAction SilentlyContinue
                return $null
            }
        }
        } catch { }
        $elapsed  = [int]((Get-Date) - $start).TotalSeconds
        $spinChar = $spinner[$spin % 4]; $spin++
        Write-Host "`r      $spinChar $([math]::Floor($elapsed/60))m$($elapsed % 60)s" -NoNewline -ForegroundColor DarkGray
    }
    Write-Host "`r                    `r" -NoNewline
    return $proc.ExitCode
}
function Show-Result {
    param($ExitCode)
    if ($null -eq $ExitCode)                          { Write-StatusTag "[SKIPPED]" "Yellow" }
    elseif ($ExitCode -in 0, 3010)                    { Write-StatusTag "[DONE]" "Green" }
    elseif ($ExitCode -in 1638, 1641, -1978335189)    { Write-StatusTag "[ALREADY INSTALLED]" "Cyan" }
    else                                              { Write-StatusTag "[FAILED - exit $ExitCode]" "Red" }
}

function Write-Installing {
    Write-Host " installing..." -NoNewline -ForegroundColor DarkGray
}
# --- Microsoft 365 via Office Deployment Tool ---
Write-Host "      > Installing: Microsoft 365..." -NoNewline -ForegroundColor Gray
# Kill any lingering Office/C2R processes from previous runs
Get-Process -Name "OfficeClickToRun","OfficeC2RClient","setup" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$OfficeInstalled = (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") -or
                   (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE") -or
                   (Get-Service -Name "ClickToRunSvc" -ErrorAction SilentlyContinue)
if ($OfficeInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $ODTDir  = "C:\Windows\Temp\ODT"
    $ODTExe  = "C:\Windows\Temp\ODT.exe"
    # Try multiple ODT sources - scrape MS page first, then fall back to known URLs
    $downloaded = $false
    $ODTUrls = @()
    # Attempt to scrape current URL from Microsoft download page
    try {
        $wc2 = New-Object System.Net.WebClient
        $wc2.Headers.Add("User-Agent","Mozilla/5.0 (Windows NT 10.0; Win64; x64)")
        $html = $wc2.DownloadString("https://www.microsoft.com/en-us/download/details.aspx?id=49117")
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
        Write-StatusTag "[FAILED - download]" "Red"
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
            $extractProc | Stop-Process -Force -ErrorAction SilentlyContinue
        }
        if (-not (Test-Path $setup)) {
            Write-StatusTag "[FAILED - ODT extract]" "Red"
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
            Write-Installing
            & "$setup" /configure "$ODTDir\config.xml"
            Show-Result $LASTEXITCODE
        }
        # ODTDir may still be locked by OfficeClickToRun - delete async so we don't block
        Start-Process "cmd.exe" -ArgumentList "/c timeout /t 120 /nobreak >nul && rd /s /q `"$ODTDir`" && del /f /q `"$ODTExe`"" -WindowStyle Hidden
    }
}
# --- Microsoft Teams ---
# --- OneDrive & Teams (included with M365) ---
Write-Host "      > Installing: OneDrive..." -NoNewline -ForegroundColor Gray
$OneDriveInstalled = (Test-Path "$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe") -or
                     (Test-Path "$env:ProgramFiles (x86)\Microsoft OneDrive\OneDrive.exe") -or
                     (Test-Path "$env:SystemRoot\SysWOW64\OneDriveSetup.exe")
if ($OneDriveInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $dest = "C:\Windows\Temp\OneDriveSetup.exe"
    Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
    if (Get-Installer -Url "https://go.microsoft.com/fwlink/?linkid=844652" -Dest $dest) {
        Write-Installing
        Write-Installing
        Start-Process -FilePath $dest -ArgumentList "/allusers /silent" -WindowStyle Hidden
        # OneDrive spawns a child and exits - wait for the real setup process to finish
        Start-Sleep -Seconds 3
        $deadline = (Get-Date).AddMinutes(5)
        while ((Get-Process -Name "OneDriveSetup" -ErrorAction SilentlyContinue) -and (Get-Date) -lt $deadline) {
            Start-Sleep -Seconds 2
        }
        Write-StatusTag "[DONE]" "Green"
    } else { Write-StatusTag "[FAILED - download]" "Red" }
    Remove-Item $dest -Force -ErrorAction SilentlyContinue
}
Write-Host "      > Installing: Microsoft Teams..." -NoNewline -ForegroundColor Gray
$TeamsInstalled = (Test-Path "$env:ProgramFiles\WindowsApps\MSTeams*\ms-teams.exe") -or
                  (Test-Path "$env:ProgramFiles (x86)\Microsoft\Teams\current\Teams.exe") -or
                  (Test-Path "$env:ProgramFiles\Microsoft\Teams\current\Teams.exe")
if ($TeamsInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $dest = "C:\Windows\Temp\TeamsSetup.exe"
    Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
    if (Get-Installer -Url "https://go.microsoft.com/fwlink/?linkid=2196106" -Dest $dest) {
        Write-Installing
        Start-Process -FilePath $dest -ArgumentList "/silent" -Wait -NoNewWindow
        Show-Result $LASTEXITCODE
    } else { Write-StatusTag "[FAILED - download]" "Red" }
    Remove-Item $dest -Force -ErrorAction SilentlyContinue
}
# --- Google Chrome ---
Write-Host "      > Installing: Google Chrome..." -NoNewline -ForegroundColor Gray
# Kill any lingering msiexec from previous installs
Get-Process -Name "msiexec" -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
$ChromeInstalled = (Test-Path "C:\Program Files\Google\Chrome\Application\chrome.exe") -or
                   (Test-Path "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe")
if ($ChromeInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $dest = "C:\Windows\Temp\ChromeEnterprise.msi"
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
        Start-Process "msiexec.exe" -ArgumentList "/i `"$dest`" /qn /norestart ALLUSERS=1" -Wait -NoNewWindow
        Show-Result $LASTEXITCODE
    } else { Write-StatusTag "[FAILED - download]" "Red" }
    Remove-Item $dest -Force -ErrorAction SilentlyContinue
}
# --- Adobe Acrobat Reader DC ---
Write-Host "      > Installing: Adobe Acrobat Reader..." -NoNewline -ForegroundColor Gray
$AcrobatInstalled = (Test-Path "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe") -or
                    (Test-Path "C:\Program Files (x86)\Adobe\Acrobat Reader DC\Reader\AcroRd32.exe") -or
                    (Test-Path "C:\Program Files (x86)\Adobe\Acrobat DC\Acrobat\Acrobat.exe")
if ($AcrobatInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $dest = "C:\Windows\Temp\AcrobatReader.exe"
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
        Start-Process -FilePath $dest -ArgumentList "/sAll /rs /msi EULA_ACCEPT=YES /qn REBOOT=ReallySuppress" -Wait -NoNewWindow
        Show-Result $LASTEXITCODE
    } else { Write-StatusTag "[FAILED - download]" "Red" }
    Remove-Item $dest -Force -ErrorAction SilentlyContinue
}
# --- Zoom ---
Write-Host "      > Installing: Zoom..." -NoNewline -ForegroundColor Gray
$ZoomInstalled = (Test-Path "C:\Program Files\Zoom\bin\Zoom.exe") -or
                 (Test-Path "C:\Program Files (x86)\Zoom\bin\Zoom.exe")
if ($ZoomInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $dest = "C:\Windows\Temp\ZoomInstaller.msi"
    Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
    if (Get-Installer -Url "https://zoom.us/client/latest/ZoomInstallerFull.msi?archType=x64" -Dest $dest) {
        Write-Installing
        Start-Process "msiexec.exe" -ArgumentList "/i `"$dest`" /qn /norestart ALLUSERS=1 REBOOT=ReallySuppress" -Wait -NoNewWindow
        Show-Result $LASTEXITCODE
    } else { Write-StatusTag "[FAILED - download]" "Red" }
    Remove-Item $dest -Force -ErrorAction SilentlyContinue
}
# --- 7-Zip ---
Write-Host "      > Installing: 7-Zip..." -NoNewline -ForegroundColor Gray
$7ZipInstalled = (Test-Path "C:\Program Files\7-Zip\7z.exe") -or
                 (Test-Path "C:\Program Files (x86)\7-Zip\7z.exe")
if ($7ZipInstalled) {
    Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
} else {
    $dest = "C:\Windows\Temp\7zip.exe"
    Write-Host " downloading..." -NoNewline -ForegroundColor DarkGray
    if (Get-Installer -Url "https://www.7-zip.org/a/7z2407-x64.exe" -Dest $dest) {
        Write-Installing
        Start-Process -FilePath $dest -ArgumentList "/S" -Wait -NoNewWindow
        Show-Result $LASTEXITCODE
    } else { Write-StatusTag "[FAILED - download]" "Red" }
    Remove-Item $dest -Force -ErrorAction SilentlyContinue
}
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
            Write-StatusTag "[ALREADY INSTALLED]" "Cyan"
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
            Write-StatusTag "[DONE]" "Green"
        }
    } catch { Write-StatusTag "[FAILED]" "Red" }
} else {
    Write-StatusTag "[SKIPPED - not elevated]" "Yellow"
}
#endregion

#region 5: DISK CLEANUP & SYSTEM HYGIENE
# ------------------------------------------------------------------------------
Write-Host "[*] Commencing Deep System Clean..." -ForegroundColor DarkRed

Write-Host "      > Clearing Update Cache..." -NoNewline -ForegroundColor Gray
Stop-Service -Name "wuauserv" -Force -ErrorAction SilentlyContinue -WarningAction SilentlyContinue
Remove-Item -Path "$env:SystemRoot\SoftwareDistribution\Download\*" -Recurse -Force -ErrorAction SilentlyContinue
Start-Service -Name "wuauserv" -ErrorAction SilentlyContinue
Write-StatusTag "[DONE]" "Green"
Write-Host "      > Purging Temp Folders..." -NoNewline -ForegroundColor Gray
$TempPaths = @(
    "C:\Windows\Temp\*", 
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
#endregion

#region 6: PRIVACY, TELEMETRY & UI SETTINGS
# ------------------------------------------------------------------------------
Write-Host "[*] Applying Privacy & UI Settings..." -ForegroundColor DarkRed
# --- Load all user hives so per-user keys can be set for every account ---
$UserHives = @()
$ProfileList = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" |
    Where-Object { $_.PSChildName -match "S-1-5-21" -and (Test-Path "$($_.ProfileImagePath)\NTUSER.DAT") }
foreach ($UserProfile in $ProfileList) {
    $HivePath  = "$($UserProfile.ProfileImagePath)\NTUSER.DAT"
    $HiveMount = "HKU\APT_$($UserProfile.PSChildName)"
    reg load $HiveMount $HivePath 2>$null | Out-Null
    $UserHives += @{ Mount = "Registry::$HiveMount"; SID = $UserProfile.PSChildName }
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
    foreach ($Hive in $UserHives) {
        Set-RegValue -Path "$($Hive.Mount)\$SubKey" -Name $Name -Value $Value -Type $Type
    }
    # Also set for current user in case hive loading missed it
    Set-RegValue -Path "HKCU:\$SubKey" -Name $Name -Value $Value -Type $Type
}
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
[GC]::Collect()
Start-Sleep -Milliseconds 500
foreach ($Hive in $UserHives) {
    reg unload "HKU\APT_$($Hive.SID)" 2>$null | Out-Null
}
Write-Host "    Privacy & UI settings applied to $($UserHives.Count) user profile(s)." -ForegroundColor Cyan
#endregion

#region 7: FINALIZATION
# ------------------------------------------------------------------------------
$RebootPending = if ($RebootRequired) { $true } else { $false }
$RebootStatus = if ($RebootPending) { "REBOOT REQUIRED" } else { "SYSTEM READY" }
$StatusColor = if ($RebootPending) { "Red" } else { "Green" }
Write-HLine -Style "dashed"
Write-Host "[>] Status  : " -NoNewline -ForegroundColor $LineCol
Write-Host $RebootStatus -ForegroundColor $StatusColor -BackgroundColor Black
# Footer
$_sfx    = "█"
$_ftrW   = $_artW + 1
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  ENDPOINT PROVISIONING SEQUENCE COMPLETE"
$_ver    = "| v1.2"
$_fpad   = " " * ($_ffillW - $_footer.Length - $_ver.Length)
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $AccentCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $AccentCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $AccentCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
$Choice = Read-Host -Prompt "`nPress ENTER to exit, or type 'RESTART' to reboot"
# Self-delete always runs regardless of reboot choice
$selfPath = $MyInvocation.MyCommand.Path
if ($selfPath -and (Test-Path $selfPath)) {
    Start-Process powershell -ArgumentList "-NoProfile -Command `"Start-Sleep 5; Remove-Item -Path '$selfPath' -Force`"" -WindowStyle Hidden
}
if ($Choice -eq "RESTART") { Restart-Computer -Force }
#endregion