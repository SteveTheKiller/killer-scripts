<#
.SYNOPSIS
    Universal Rename Tool (U.R.T.) v1.6
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Renames local or domain-joined computers from an admin shell with
    no GUI popups. Collects domain credentials inline if needed,
    preserves AD trust relationships, and offers an optional immediate
    reboot on completion.
#>

#region Pre-Flight Checks
# ============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit 1
}

# Standardized Console Output
$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($Success) {
        $savedTop = [Console]::CursorTop
        if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
            [Console]::SetCursorPosition(0, $script:StepRow)
            $combined = if ($CustomInfo) { "$($script:StepMsg) $CustomInfo" } else { $script:StepMsg }
            $_pad = " " * [math]::Max(1, $script:Width - $combined.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
            if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
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

$script:HasRawUI = $false
try { $null = $Host.UI.RawUI.KeyAvailable; $script:HasRawUI = $true } catch {}

function Read-InputLine {
    param([string]$Prompt, [switch]$Secure)
    if (-not $script:HasRawUI) {
        if ($Secure) { return Read-Host -Prompt $Prompt -AsSecureString }
        return Read-Host -Prompt $Prompt
    }
    Write-Host "${Prompt}: " -NoNewline
    $chars = ""
    $done = $false
    while (-not $done) {
        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        switch ($key.VirtualKeyCode) {
            27 {  # ESC
                Write-Host ""
                Write-Host "`n[!] Cancelled by user." -ForegroundColor Yellow
                exit 0
            }
            13 {  # Enter
                Write-Host ""
                $done = $true
            }
            8  {  # Backspace
                if ($chars.Length -gt 0) {
                    $chars = $chars.Substring(0, $chars.Length - 1)
                    Write-Host "`b `b" -NoNewline
                }
            }
            default {
                if ($key.Character -ge 32) {
                    $chars += $key.Character
                    Write-Host $(if ($Secure) { "*" } else { $key.Character.ToString() }) -NoNewline
                }
            }
        }
    }
    if ($Secure) { return ConvertTo-SecureString $chars -AsPlainText -Force }
    return $chars
}

Clear-Host
$script:Width  = 85
$_ver    = "| v1.6"
$LineCol   = "Green"
$MainCol   = "White"
$WarnCol   = "Yellow"
$BorderCol = "Green"
$ArtCol    = "DarkCyan"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"

$CS = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
if ($null -eq $CS) { $CS = Get-WmiObject Win32_ComputerSystem }
$IsDomain = ($CS.DomainRole -eq 1) -or ($CS.DomainRole -eq 3)
$DomainDisplay = if ($IsDomain) { $CS.Domain } else { "No" }

function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "-" * $Width
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

# Header
$_pfx  = "█  "
$_art1 = "╦ ╦ ╦═╗ ╔╦╗ "
$_art2 = "║ ║ ╠╦╝  ║  "
$_art3 = "╚═╝ ╩╚═  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "UNIVERSAL RENAME TOOL"
$_tpad  = " " * ($_fillW - $_title.Length)
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

Write-Host "`n[>] Current Device Name  : " -NoNewline -ForegroundColor $LineCol
Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $AccentCol
Write-Host "[>] Domain Joined        : " -NoNewline -ForegroundColor $LineCol
Write-Host "$DomainDisplay" -ForegroundColor $(if ($IsDomain) { $AccentCol } else { "Red" })
Write-HLine -Style dashed
#endregion

#region 1. Input & Validation
# ============================================================================
Write-Host "[1/3] Awaiting Input..." -ForegroundColor Green
$NewName = Read-InputLine "      Enter New Computer Name"
if ([string]::IsNullOrWhiteSpace($NewName)) { 
    Write-Warning "No name entered. Aborting."
    exit 0 
}
if ($NewName -eq $env:COMPUTERNAME) {
    Write-Host "      New name matches current name. Skipping..." -ForegroundColor Gray
    exit 0
}
#endregion

#region 2. Environment Discovery
# ============================================================================
Write-StepUpdate "[2/3] Analyzing Network Environment..."
$EnvType = if ($IsDomain) { "Active Directory Domain" } else { "Workgroup / Entra ID" }
Write-StepUpdate -Success -CustomInfo "($EnvType)"
#endregion

#region 3. Execution
# ============================================================================
Write-Host "[3/3] Commencing Rename Procedure..." -ForegroundColor Green

try {
    if ($IsDomain) {
        Write-Host "      [!] Domain Detected. Enter Domain Admin Credentials below:" -ForegroundColor Yellow
        $DomainUser = Read-InputLine "      Username (DOMAIN\user)"
        $DomainPass = Read-InputLine "      Password" -Secure
        $Creds = New-Object System.Management.Automation.PSCredential($DomainUser, $DomainPass)
        Write-Host "      [+] Authenticated. Renaming in AD and locally..." -ForegroundColor Gray
        Rename-Computer -NewName $NewName -DomainCredential $Creds -Force -ErrorAction Stop
    } 
    else {
        Write-Host "      [+] Local Environment detected. Renaming..." -ForegroundColor Gray
        Rename-Computer -NewName $NewName -Force -ErrorAction Stop
    }
    Write-HLine -Style dashed
    Write-Host "        Current Name : $($env:COMPUTERNAME)" -ForegroundColor Gray
    Write-Host "        Future Name  : $NewName" -ForegroundColor Yellow
    # Footer
    $_sfx    = "█"
    $_ftrW   = $_artW + 1
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_footer = "  RENAME OPERATIONS SEQUENCE COMPLETE"

    $_fpad   = " " * ($_ffillW - $_footer.Length - $_ver.Length)
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    $Choice = Read-InputLine "Press ENTER to exit, or type 'RESTART' to reboot now"
    if ($Choice -eq "RESTART") { 
        Write-Host "System rebooting..." -ForegroundColor Red
        Restart-Computer -Force
    }
    exit 0
} 
catch {
    Write-Host "`n[ERROR] Rename failed: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host -Prompt "Press Enter to exit"
    exit 1
}
#endregion