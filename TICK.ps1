<#
.SYNOPSIS
    Trigger Immediate Clock Kickstart (TICK) v1.1
    Developed by SteveTheKiller | Updated: 2026-03-12
.DESCRIPTION
    Resets and resyncs the Windows Time service against a chosen NTP peer,
    then reports before/after timestamps with a plain-English summary of
    sync accuracy, stratum, and clock health. Domain-aware: detects if the
    machine is domain-joined or a DC and defaults accordingly.
.NOTES
    Script can be run with parameters:
    .\TICK.ps1 -Silent
        Non-interactive - uses the default NTP source for this machine.
        (Domain Controller if domain-joined, time.cloudflare.com if not.)
    .\TICK.ps1 -Silent -NtpServer time.google.com
        Non-interactive - syncs against the specified NTP server.
#>
param(
    [switch]$Silent,
    [string]$NtpServer
)
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit
}

#region 1) Banner & Display Helpers
# ============================================================================
Clear-Host

$script:Width      = 85
$script:LabelFG    = "DarkGreen"
$script:SubLabelFG = "DarkMagenta"

$BorderCol = "DarkMagenta"
$ArtCol    = "White"
$AccentCol = "DarkGreen"
$DimCol    = "DarkGray"

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
        [ConsoleColor]$AccentCol,
        [ConsoleColor]$BorderCol,
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
    # Header Line 1
    Write-Host "█ " -ForegroundColor $BorderCol -NoNewline
    Write-Host "╔╦╗ ╦ ╔═╗ ╦╔═  " -ForegroundColor $ArtCol -NoNewline
    Write-Host ("-" * 65) -ForegroundColor $BorderCol
    # Header Line 2
    Write-Host "█  " -ForegroundColor $BorderCol -NoNewline
    Write-Host "║  ║ ║   ╠╩╗  " -ForegroundColor $ArtCol -NoNewline
    Write-Host "T R I G G E R   I M M E D I A T E   C L O C K   K I C K S T A R T" -ForegroundColor $AccentCol
    # Header Line 3
    Write-Host "█  " -ForegroundColor $BorderCol -NoNewline
    Write-Host "╩  ╩ ╚═╝ ╩ ╩  " -ForegroundColor $ArtCol -NoNewline
    Write-Host ("-" * 65) -ForegroundColor $BorderCol
    }
function Write-Step {
    param([string]$StepNum, [string]$Message)
    Write-Host "[$StepNum] $Message" -ForegroundColor Magenta -NoNewline
}
function Complete-Step {
    param([string]$StepNum, [string]$Message)
    $pad = " " * [math]::Max(1, $script:Width - "[$StepNum] $Message".Length - "[DONE]".Length)
    Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
    Write-Host "[$StepNum]" -ForegroundColor DarkGray -NoNewline
    Write-Host " $Message$pad" -ForegroundColor White -NoNewline
    Write-Host "[DONE]" -ForegroundColor Green
}
Write-Banner
#endregion

#region 3) NTP Server Selection
# ============================================================================
$NtpServers = @(
    "time.cloudflare.com",
    "time.nist.gov",
    "pool.ntp.org",
    "time.windows.com",
    "time.google.com",
    "time.apple.com",
    "ntp.ubuntu.com"
)
$otherNum = $NtpServers.Count + 1
$UseDomainSync = $false

# Domain detection
$_cs = Get-CimInstance Win32_ComputerSystem
$isDomainJoined = $_cs.PartOfDomain
$isDC = $_cs.DomainRole -ge 4
$domainName = if ($isDomainJoined) { $_cs.Domain } else { $null }
$domainPDC = $null
if ($isDomainJoined) {
    try { $domainPDC = [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().PdcRoleOwner.Name }
    catch { $domainPDC = $null }
}
$choice = $null
if (-not $Silent -and -not $NtpServer) {
    Write-Host ""
    if ($isDC) {
        Write-Host "  DC detected" -ForegroundColor DarkGreen -NoNewline
        Write-Host " - " -ForegroundColor DarkGray -NoNewline
        Write-Host "Use an external NTP server" -ForegroundColor Magenta
        Write-Host ""
    }
    Write-Host "Select NTP Server:" -ForegroundColor $script:LabelFG
    if ($isDomainJoined -and -not $isDC) {
        Write-Host "  [D] " -ForegroundColor DarkGray -NoNewline
        Write-Host "Domain Controller" -ForegroundColor Yellow -NoNewline
        if ($domainPDC) { Write-Host "  ($domainPDC)" -ForegroundColor DarkGray }
        else { Write-Host "" }
    }
    for ($i = 0; $i -lt $NtpServers.Count; $i++) {
        Write-Host "  [$($i + 1)] " -ForegroundColor DarkGray -NoNewline
        Write-Host $NtpServers[$i] -ForegroundColor Yellow
    }
    Write-Host "  [$otherNum] " -ForegroundColor DarkGray -NoNewline
    Write-Host "Other (enter manually)" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Choice " -ForegroundColor $script:LabelFG -NoNewline
    $hint = if ($isDomainJoined -and -not $isDC) { "[D/1-$otherNum]  Enter = D  Esc = quit: " } else { "[1-$otherNum]  Enter = 1  Esc = quit: " }
    Write-Host $hint -ForegroundColor DarkGray -NoNewline
    try {
        $key = [Console]::ReadKey($true)
        Write-Host $key.KeyChar
        if ($key.Key -eq [ConsoleKey]::Escape) {
            Write-Host "`nExiting." -ForegroundColor DarkGray
            Exit
        }
        $choice = if ($key.Key -eq [ConsoleKey]::Enter) { if ($isDomainJoined -and -not $isDC) { "D" } else { "1" } } else { $key.KeyChar.ToString() }
    } catch {
        $raw = (Read-Host).Trim()
        if ([string]::IsNullOrWhiteSpace($raw)) { $choice = if ($isDomainJoined -and -not $isDC) { "D" } else { "1" } }
        elseif ($raw -eq 'q' -or $raw -eq 'Q') { Write-Host "`nExiting." -ForegroundColor DarkGray; Exit }
        else { $choice = $raw }
    }
}
if (-not $NtpServer) {
    if (-not $choice) { $choice = if ($isDomainJoined -and -not $isDC) { "D" } else { "1" } }
    $NtpServer = if (($choice -eq 'D' -or $choice -eq 'd') -and $isDomainJoined) {
        $UseDomainSync = $true
        if ($domainPDC) { $domainPDC } else { "Domain Controller" }
    } elseif ($choice -match '^\d+$' -and [int]$choice -ge 1 -and [int]$choice -le $NtpServers.Count) {
        $NtpServers[[int]$choice - 1]
    } elseif ($choice -eq "$otherNum") {
        Write-Host "Enter NTP server address: " -ForegroundColor $script:LabelFG -NoNewline
        Read-Host
    } else {
        $NtpServers[0]
    }
}
Clear-Host
Write-Banner
Write-Host "Device Name  : " -ForegroundColor $script:LabelFG -NoNewline; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor Yellow
Write-Host "Domain       : " -ForegroundColor $script:LabelFG -NoNewline
if ($isDC) {
    Write-Host $domainName -ForegroundColor Yellow -NoNewline
    Write-Host "  (Domain Controller)" -ForegroundColor DarkGreen
} elseif ($isDomainJoined) {
    Write-Host $domainName -ForegroundColor Yellow
} else {
    Write-Host "No" -ForegroundColor DarkGray
}
if (-not $isDomainJoined) {
    Write-Host "NTP Peer     : " -ForegroundColor $script:LabelFG -NoNewline; Write-Host $NtpServer -ForegroundColor Yellow
}
Write-Host "Time (Before): " -ForegroundColor $script:LabelFG -NoNewline; Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
Write-HLine -Style dashed
#endregion

#region 3) Sync Steps
# ============================================================================
Write-Step "1/4" "Registering and stopping w32time..."
w32tm /register | Out-Null
Stop-Service w32time -Force
Complete-Step "1/4" "Registering and stopping w32time..."
Write-Step "2/4" "Configuring NTP peer..."
if ($UseDomainSync) {
    w32tm /config /syncfromflags:domhier /update | Out-Null
} else {
    w32tm /config /manualpeerlist:"$NtpServer,0x8" /syncfromflags:manual /reliable:yes /update | Out-Null
}
Complete-Step "2/4" "Configuring NTP peer..."
Write-Step "3/4" "Starting w32time service..."
Start-Service w32time
Complete-Step "3/4" "Starting w32time service..."
Write-Step "4/4" "Forcing resync..."
w32tm /resync /force | Out-Null
Complete-Step "4/4" "Forcing resync..."
#endregion

#region 7) Results
# ============================================================================
Write-HLine -Style dashed
Write-Host "Time (After) : " -ForegroundColor $script:LabelFG -NoNewline
Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
# Parse w32tm /query /status into a lookup table
$statusRaw = w32tm /query /status
$s = @{}
foreach ($line in $statusRaw) {
    if ($line -match '^(.+?):\s+(.+)$') {
        $s[$Matches[1].Trim()] = $Matches[2].Trim()
    }
}
# Extract values
$source      = ($s['Source'] -replace ',0x\w+','').Trim()
$refId       = $s['ReferenceId']
$serverIP    = if ($refId -match 'source IP:\s+([\d.]+)') { $Matches[1] } else { $refId }
$lastSync    = $s['Last Successful Sync Time']
$stratum     = $s['Stratum'] -replace '\s*\(.*\)',''
$dispersion  = $s['Root Dispersion']
$leapRaw     = $s['Leap Indicator']
$leap        = if ($leapRaw -match '^0') { "OK - No warnings" } else { "WARNING: $leapRaw" }
$leapColor   = if ($leapRaw -match '^0') { "Green" } else { "Red" }
$dispVal     = [double]($dispersion -replace 's','')
$dispColor   = if ($dispVal -gt 5) { "Yellow" } elseif ($dispVal -gt 1) { "DarkYellow" } else { "Green" }
$dispNote    = if ($dispVal -gt 5) { " (normal after a forced resync)" } elseif ($dispVal -gt 1) { " (slightly elevated)" } else { "" }
$stratumNum  = [int]($stratum.Trim())
$stratumDisplay = switch ($stratumNum) {
    1           { "1 (atomic clock)"; break }
    {$_ -le 5}  { "$stratumNum hops from atomic"; break }
    {$_ -le 10} { "$stratumNum hops from atomic - acceptable"; break }
    default     { "$stratumNum hops from atomic - too far" }
}
Write-Host "  Source         : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $source -ForegroundColor Yellow
Write-Host "  Server IP      : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $serverIP -ForegroundColor Yellow
Write-Host "  Last Sync      : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $lastSync -ForegroundColor Yellow
Write-Host "  Stratum        : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $stratumDisplay -ForegroundColor Yellow
Write-Host "  Max Clock Error: " -ForegroundColor $script:SubLabelFG -NoNewline
Write-Host $dispersion -ForegroundColor $dispColor -NoNewline
if ($dispNote) { Write-Host $dispNote -ForegroundColor DarkYellow } else { Write-Host "" }
Write-Host "  Clock Status   : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $leap -ForegroundColor $leapColor
if ($UseDomainSync) {
    Write-Host " This computer is now synced to the Domain Controller." -ForegroundColor DarkGreen
}
# Footer Line 1 (The Top Bar of the Footer)
Write-Host ("-" * 65) -ForegroundColor $BorderCol -NoNewline
Write-Host " ╔╦╗ ╦ ╔═╗ ╦╔═ " -ForegroundColor $ArtCol -NoNewline
Write-Host "█" -ForegroundColor $BorderCol
# Footer Line 2 (The Text Line)
$FooterText = "  C H R O N O M E T R I C   S E Q U E N C E   C O M P L E T E"
$Padding = " " * (65 - $FooterText.Length)
Write-Host "$FooterText$Padding" -ForegroundColor $AccentCol -NoNewline
Write-Host "  ║  ║ ║   ╠╩╗ " -ForegroundColor $ArtCol -NoNewline
Write-Host "█" -ForegroundColor $BorderCol
# Footer Line 3 (The Bottom Bar of the Footer)
Write-Host ("-" * 65) -ForegroundColor $BorderCol -NoNewline
Write-Host "  ╩  ╩ ╚═╝ ╩ ╩ " -ForegroundColor $ArtCol -NoNewline
Write-Host "█" -ForegroundColor $BorderCol
#endregion