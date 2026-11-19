<#
.SYNOPSIS
    Universal Rename Tool (U.R.T.) v2.0
    Developed by SteveTheKiller | Updated: 2026-11-19
.DESCRIPTION
    Renames local or domain-joined computers from an admin shell with
    no GUI popups. Collects domain credentials inline if needed,
    preserves AD trust relationships, and offers an optional immediate
    reboot on completion.
    Domain controllers (DomainRole 4 and 5) are refused: renaming a DC
    needs netdom computername, not this tool.
    URT asks for the name at a console. Without one it exits 2 and
    changes nothing. An unanswered restart question means no restart.
    Exit codes: 0 done or cancelled, 1 rename failed, 2 not elevated, no
    console or a domain controller, 3010 renamed and a restart is needed.
#>

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-UrtElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
function Test-UrtInteractive {
    try { [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected } catch { $false }
}
function Clear-UrtScreen { try { Clear-Host -ErrorAction Stop } catch { } }
function Get-UrtComputerSystem {
    try { Get-CimInstance Win32_ComputerSystem -ErrorAction Stop } catch {
        try { Get-WmiObject Win32_ComputerSystem -ErrorAction Stop } catch { $null }
    }
}
function Get-UrtCursorTop { try { [Console]::CursorTop } catch { -1 } }
function Set-UrtCursor {
    param([int]$Left, [int]$Top)
    try { [Console]::SetCursorPosition($Left, $Top) } catch { }
}
# Reads one line without needing a console window. Returns Status ok, cancel (Esc) or none
# (nothing could answer), plus Value, so no caller can hang or exit from inside a prompt.
function Read-UrtInput {
    param([string]$Prompt, [switch]$Secure)
    $hasRaw = $false
    try { $null = $Host.UI.RawUI.KeyAvailable; $hasRaw = $true } catch { }
    if (-not $hasRaw) {
        try {
            if ($Secure) { return [pscustomobject]@{ Status = 'ok'; Value = (Read-Host -Prompt $Prompt -AsSecureString) } }
            return [pscustomobject]@{ Status = 'ok'; Value = (Read-Host -Prompt $Prompt) }
        } catch { return [pscustomobject]@{ Status = 'none'; Value = $null } }
    }
    Write-Host "${Prompt}: " -NoNewline
    $chars = ""
    $done = $false
    while (-not $done) {
        try { $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { return [pscustomobject]@{ Status = 'none'; Value = $null } }
        switch ($key.VirtualKeyCode) {
            27 { Write-Host ""; return [pscustomobject]@{ Status = 'cancel'; Value = $null } }
            13 { Write-Host ""; $done = $true }
            8  {
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
    if ($Secure) { return [pscustomobject]@{ Status = 'ok'; Value = (ConvertTo-SecureString $chars -AsPlainText -Force) } }
    [pscustomobject]@{ Status = 'ok'; Value = $chars }
}

# Win32_ComputerSystem.DomainRole: 0 standalone workstation, 1 member workstation, 2 standalone server,
# 3 member server, 4 backup domain controller, 5 primary domain controller.
function Get-UrtRoleKind {
    param($Role)
    if ($null -eq $Role) { return 'Unknown' }
    switch ([int]$Role) {
        0 { 'Workgroup' }
        2 { 'Workgroup' }
        1 { 'Domain' }
        3 { 'Domain' }
        4 { 'DomainController' }
        5 { 'DomainController' }
        default { 'Unknown' }
    }
}

# Standardized Console Output
$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($Success) {
        $savedTop = Get-UrtCursorTop
        if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
            Set-UrtCursor 0 $script:StepRow
            $combined = if ($CustomInfo) { "$($script:StepMsg) $CustomInfo" } else { $script:StepMsg }
            $_pad = " " * [math]::Max(1, $script:Width - $combined.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
            if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
            Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
            Set-UrtCursor 0 $savedTop
        }
        $script:StepRow = -1
    } else {
        $script:StepRow = Get-UrtCursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $InfoCol
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

$script:Width  = 85
$_ver    = "| v2.0"
$LineCol   = "Green"
$MainCol   = "White"
$WarnCol   = "Yellow"
$BorderCol = "Green"
$ArtCol    = "DarkCyan"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"
$_pfx  = "█  "
$_art1 = "╦ ╦ ╦═╗ ╔╦╗ "
$_art2 = "║ ║ ╠╦╝  ║  "
$_art3 = "╚═╝ ╩╚═  ╩  "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW)
$_art2 = $_art2.PadRight($_artW)
$_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "UNIVERSAL RENAME TOOL"
$_tpad  = " " * ($_fillW - $_title.Length)

function Invoke-UrtMain {
    #region Pre-Flight Checks
    # ============================================================================
    if (-not (Test-UrtElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Exit-KillerScript 1
    }
    if (-not (Test-UrtInteractive)) {
        Exit-KillerScript 1
    }
    $CS = Get-UrtComputerSystem
    $RoleKind = if ($null -ne $CS) { Get-UrtRoleKind $CS.DomainRole } else { 'Unknown' }
    if ($RoleKind -eq 'DomainController') {
        Exit-KillerScript 1
    }
    if ($RoleKind -eq 'Unknown') {
        Exit-KillerScript 1
    }
    $IsDomain = ($RoleKind -eq 'Domain')
    $DomainDisplay = if ($IsDomain) { $CS.Domain } else { "No" }

    Clear-UrtScreen
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
    $in = Read-UrtInput "      Enter New Computer Name"
    if ($in.Status -eq 'none') {
        Exit-KillerScript 1
    }
    if ($in.Status -eq 'cancel') {
        Write-Host "`n[!] Cancelled by user." -ForegroundColor Yellow
        Exit-KillerScript 0
    }
    $NewName = "$($in.Value)".Trim()
    if ([string]::IsNullOrWhiteSpace($NewName)) {
        Write-Warning "No name entered. Aborting."
        Exit-KillerScript 0
    }
    if ($NewName -eq $env:COMPUTERNAME) {
        Write-Host "      New name matches current name. Skipping..." -ForegroundColor Gray
        Exit-KillerScript 0
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
            $u = Read-UrtInput "      Username (DOMAIN\user)"
            if ($u.Status -ne 'ok' -or [string]::IsNullOrWhiteSpace("$($u.Value)")) {
                if ($u.Status -eq 'cancel') {
                    Exit-KillerScript 0
                }
                Exit-KillerScript 1
            }
            $p = Read-UrtInput "      Password" -Secure
            if ($p.Status -eq 'cancel') {
                Exit-KillerScript 0
            }
            if ($p.Status -ne 'ok' -or $null -eq $p.Value) {
                Exit-KillerScript 1
            }
            $Creds = New-Object System.Management.Automation.PSCredential("$($u.Value)".Trim(), $p.Value)
            Write-Host "      [+] Renaming in AD and locally..." -ForegroundColor Gray
            Rename-Computer -NewName $NewName -DomainCredential $Creds -Force -ErrorAction Stop
        }
        else {
            Write-Host "      [+] Local Environment detected. Renaming..." -ForegroundColor Gray
            Rename-Computer -NewName $NewName -Force -ErrorAction Stop
        }
    }
    catch {
        if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT') { throw }
        Write-Host "`n[ERROR] Rename failed: $($_.Exception.Message)" -ForegroundColor Red
        Exit-KillerScript 1
    }
    Write-HLine -Style dashed
    Write-Host "        Current Name : $($env:COMPUTERNAME)" -ForegroundColor Gray
    Write-Host "        Future Name  : $NewName" -ForegroundColor Yellow
    # Footer
    $_sfx    = "█"
    $_ftrW   = $_artW + 1
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_footer = "  RENAME OPERATIONS SEQUENCE COMPLETE"

    $_fpad   = " " * [math]::Max(1, $_ffillW - $_footer.Length - $_ver.Length)
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    # An unanswered question means no restart.
    $Choice = Read-UrtInput "Press ENTER to exit, or type 'RESTART' to reboot now"
    if ($Choice.Status -eq 'ok' -and "$($Choice.Value)".Trim() -eq "RESTART") {
        Write-Host "System rebooting..." -ForegroundColor Red
        Restart-Computer -Force
        Exit-KillerScript 0
    }
    Exit-KillerScript 0
    #endregion
}

if (-not $global:KillerScriptsTestMode) { Invoke-UrtMain }
