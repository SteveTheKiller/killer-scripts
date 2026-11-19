<#
.SYNOPSIS
    Trigger Immediate Clock Kickstart (TICK) v2.0
    Developed by Steve the Killer | Updated: 2026-11-19
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
    Without a console the menu is skipped and TICK runs as -Silent.
    Exit codes: 0 synced and verified, 1 time service not configured,
    2 not elevated or no usable NTP server, 3 configured but the sync is
    not confirmed (resync failed, wrong source, or last sync unknown/old).
#>
param(
    [switch]$Silent,
    [string]$NtpServer
)
function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-TickElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
function Test-TickInteractive {
    try { [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected } catch { $false }
}
function Read-TickKey {
    # Returns $null when there is no console, so callers fall back to a line prompt or a default.
    try { [Console]::ReadKey($true) } catch { $null }
}
function Read-TickLine {
    try { Read-Host } catch { $null }
}
function Clear-TickScreen { try { Clear-Host -ErrorAction Stop } catch { } }
function Get-TickComputerSystem { Get-CimInstance Win32_ComputerSystem }
function Get-TickPdcName {
    try { [System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().PdcRoleOwner.Name } catch { $null }
}
# Runs w32tm and returns its exit code and output lines, so every step can be checked.
function Invoke-TickW32tm {
    param([string[]]$Arguments)
    $exe = 'w32tm.exe'
    if ([Environment]::Is64BitOperatingSystem -and -not [Environment]::Is64BitProcess) {
        $exe = Join-Path $env:windir 'Sysnative\w32tm.exe'
    }
    $lines = @()
    $code = 1
    try {
        $lines = @(& $exe @Arguments 2>&1 | ForEach-Object { "$_" })
        $code = $LASTEXITCODE
    } catch { $lines = @("$($_.Exception.Message)"); $code = 1 }
    [pscustomobject]@{ ExitCode = $code; Output = $lines }
}

#region 1) Banner & Display Helpers
# ============================================================================
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
    Write-Host ("-" * 68) -ForegroundColor $BorderCol
    # Header Line 2
    Write-Host "█  " -ForegroundColor $BorderCol -NoNewline
    Write-Host "║  ║ ║   ╠╩╗  " -ForegroundColor $ArtCol -NoNewline
    Write-Host "T R I G G E R   I M M E D I A T E   C L O C K   K I C K S T A R T" -ForegroundColor $AccentCol
    # Header Line 3
    Write-Host "█  " -ForegroundColor $BorderCol -NoNewline
    Write-Host "╩  ╩ ╚═╝ ╩ ╩  " -ForegroundColor $ArtCol -NoNewline
    Write-Host ("-" * 68) -ForegroundColor $BorderCol
    }
function Write-Step {
    param([string]$StepNum, [string]$Message)
    Write-Host "[$StepNum] $Message" -ForegroundColor Magenta -NoNewline
}
function Complete-Step {
    param([string]$StepNum, [string]$Message, [switch]$Failed)
    $tag = if ($Failed) { "[FAIL]" } else { "[DONE]" }
    $pad = " " * [math]::Max(1, $script:Width - "[$StepNum] $Message".Length - $tag.Length)
    Write-Host ("`r" + (" " * 100) + "`r") -NoNewline
    Write-Host "[$StepNum]" -ForegroundColor DarkGray -NoNewline
    Write-Host " $Message$pad" -ForegroundColor White -NoNewline
    Write-Host $tag -ForegroundColor $(if ($Failed) { "Red" } else { "Green" })
}

# True when the source w32tm reports is the peer that was asked for. Compares the source name
# only, never localized text. For the domain route any domain controller under the domain counts.
function Test-TickSourceMatches {
    param([string]$Source, [string]$Expected, [bool]$IsDomainRoute, [string]$DomainName)
    $src = ("$Source" -replace ',0x\w+', '').Trim().ToLowerInvariant()
    if (-not $src) { return $false }
    $exp = "$Expected".Trim().ToLowerInvariant()
    if ($IsDomainRoute) {
        if ($exp -and $exp -ne 'domain controller' -and ($src -eq $exp -or $src.StartsWith("$exp."))) { return $true }
        if ($DomainName) {
            $dom = $DomainName.Trim().ToLowerInvariant()
            if ($src.EndsWith(".$dom") -or $src -eq $dom) { return $true }
        }
        return $false
    }
    return ($src -eq $exp)
}

function Invoke-TickMain {
    param([switch]$Silent, [string]$NtpServer)
    if (-not (Test-TickElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Write-Output "[FAIL] TICK not elevated, nothing was changed (exit 2)"
        Exit-KillerScript 2
    }
    Clear-TickScreen
    Write-Banner
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
    $_cs = Get-TickComputerSystem
    $isDomainJoined = $_cs.PartOfDomain
    $isDC = $_cs.DomainRole -ge 4
    $domainName = if ($isDomainJoined) { $_cs.Domain } else { $null }
    $domainPDC = $null
    if ($isDomainJoined) {
        $domainPDC = Get-TickPdcName
    }
    $choice = $null
    # With no console (RMM, LiveConnect) nothing can answer the menu, so run as -Silent.
    $interactive = (-not $Silent) -and (-not $NtpServer) -and (Test-TickInteractive)
    if ($interactive) {
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
        $key = Read-TickKey
        if ($key) {
            Write-Host $key.KeyChar
            if ($key.Key -eq [ConsoleKey]::Escape) {
                Write-Host "`nExiting." -ForegroundColor DarkGray
                Exit-KillerScript 0
            }
            $choice = if ($key.Key -eq [ConsoleKey]::Enter) { if ($isDomainJoined -and -not $isDC) { "D" } else { "1" } } else { $key.KeyChar.ToString() }
        } else {
            $raw = "$(Read-TickLine)".Trim()
            if ([string]::IsNullOrWhiteSpace($raw)) { $choice = if ($isDomainJoined -and -not $isDC) { "D" } else { "1" } }
            elseif ($raw -eq 'q' -or $raw -eq 'Q') { Write-Host "`nExiting." -ForegroundColor DarkGray; Exit-KillerScript 0 }
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
            "$(Read-TickLine)".Trim()
        } else {
            $NtpServers[0]
        }
    }
    if (-not $UseDomainSync) {
        if ([string]::IsNullOrWhiteSpace($NtpServer)) {
            Write-Output "[FAIL] TICK needs an NTP server address, nothing was changed (exit 2)"
            Exit-KillerScript 2
        }
        if ($NtpServer -notmatch '^[A-Za-z0-9]([A-Za-z0-9._:-]*[A-Za-z0-9])?$') {
            Write-Output "[FAIL] TICK was given an NTP server that is not a plain host name or address: '$NtpServer'. Nothing was changed (exit 2)"
            Exit-KillerScript 2
        }
    }
    Clear-TickScreen
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
    # Configuration problems (exit 1) are kept apart from a sync that is not confirmed (exit 3).
    $ConfigProblems = New-Object System.Collections.ArrayList
    $StepFailed = $false
    Write-Step "1/4" "Registering and stopping w32time..."
    $res = Invoke-TickW32tm @("/register")
    if ($res.ExitCode -ne 0) { $StepFailed = $true; [void]$ConfigProblems.Add("w32tm /register failed (code $($res.ExitCode))") }
    try { Stop-Service w32time -Force -ErrorAction Stop } catch { $StepFailed = $true; [void]$ConfigProblems.Add("the time service could not be stopped: $($_.Exception.Message)") }
    Complete-Step "1/4" "Registering and stopping w32time..." -Failed:$StepFailed
    Write-Step "2/4" "Configuring NTP peer..."
    if ($UseDomainSync) {
        $res = Invoke-TickW32tm @("/config", "/syncfromflags:domhier", "/update")
    } else {
        $res = Invoke-TickW32tm @("/config", "/manualpeerlist:$NtpServer,0x8", "/syncfromflags:manual", "/reliable:yes", "/update")
    }
    if ($res.ExitCode -ne 0) { [void]$ConfigProblems.Add("w32tm /config failed (code $($res.ExitCode))") }
    Complete-Step "2/4" "Configuring NTP peer..." -Failed:($res.ExitCode -ne 0)
    Write-Step "3/4" "Starting w32time service..."
    $StartOk = $true
    try { Start-Service w32time -ErrorAction Stop } catch { $StartOk = $false; [void]$ConfigProblems.Add("the time service could not be started: $($_.Exception.Message)") }
    $svc = Get-Service w32time -ErrorAction SilentlyContinue
    if ($StartOk -and (-not $svc -or "$($svc.Status)" -ne 'Running')) { $StartOk = $false; [void]$ConfigProblems.Add("the time service is not running after the restart") }
    Complete-Step "3/4" "Starting w32time service..." -Failed:(-not $StartOk)
    Write-Step "4/4" "Forcing resync..."
    $ResyncOk = $false
    if ($StartOk) {
        $res = Invoke-TickW32tm @("/resync", "/force")
        $ResyncOk = ($res.ExitCode -eq 0)
    }
    Complete-Step "4/4" "Forcing resync..." -Failed:(-not $ResyncOk)
    #endregion

    #region 7) Results
    # ============================================================================
    Write-HLine -Style dashed
    Write-Host "Time (After) : " -ForegroundColor $script:LabelFG -NoNewline
    Write-Host "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor Yellow
    # Parse w32tm /query /status into a lookup table. The keys are English, so a localized
    # Windows leaves values missing; those show as UNKNOWN instead of stopping the script.
    $statusRes = Invoke-TickW32tm @("/query", "/status")
    $s = @{}
    foreach ($line in @($statusRes.Output)) {
        if ($line -match '^(.+?):\s+(.+)$') {
            $s[$Matches[1].Trim()] = $Matches[2].Trim()
        }
    }
    # The source name does not depend on the language, so it comes from /query /source.
    $srcRes = Invoke-TickW32tm @("/query", "/source")
    $srcLine = @($srcRes.Output | Where-Object { $_ -and "$_".Trim() } | Select-Object -First 1)
    $reportedSource = if ($srcRes.ExitCode -eq 0 -and $srcLine.Count -gt 0) { ("$($srcLine[0])" -replace ',0x\w+', '').Trim() } else { $null }
    $source      = if ($reportedSource) { $reportedSource } elseif ($s['Source']) { ($s['Source'] -replace ',0x\w+', '').Trim() } else { "UNKNOWN" }
    $refId       = $s['ReferenceId']
    $serverIP    = if (-not $refId) { "UNKNOWN" } elseif ($refId -match 'source IP:\s+([\d.]+)') { $Matches[1] } else { $refId }
    $lastSync    = if ($s['Last Successful Sync Time']) { $s['Last Successful Sync Time'] } else { "UNKNOWN" }
    $stratumRaw  = if ($s['Stratum']) { ($s['Stratum'] -replace '\s*\(.*\)', '').Trim() } else { $null }
    $dispersion  = if ($s['Root Dispersion']) { $s['Root Dispersion'] } else { "UNKNOWN" }
    $leapRaw     = $s['Leap Indicator']
    $leap        = if (-not $leapRaw) { "UNKNOWN" } elseif ($leapRaw -match '^0') { "OK - No warnings" } else { "WARNING: $leapRaw" }
    $leapColor   = if (-not $leapRaw) { "DarkYellow" } elseif ($leapRaw -match '^0') { "Green" } else { "Red" }
    $dispVal     = $null
    $tmpD = 0.0
    if ($s['Root Dispersion'] -and [double]::TryParse((($s['Root Dispersion'] -replace 's', '' -replace ',', '.').Trim()), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$tmpD)) { $dispVal = $tmpD }
    $dispColor   = if ($null -eq $dispVal) { "DarkYellow" } elseif ($dispVal -gt 5) { "Yellow" } elseif ($dispVal -gt 1) { "DarkYellow" } else { "Green" }
    $dispNote    = if ($null -eq $dispVal) { "" } elseif ($dispVal -gt 5) { " (normal after a forced resync)" } elseif ($dispVal -gt 1) { " (slightly elevated)" } else { "" }
    $stratumNum  = $null
    $tmpS = 0
    if ($stratumRaw -and [int]::TryParse($stratumRaw, [ref]$tmpS)) { $stratumNum = $tmpS }
    $stratumDisplay = if ($null -eq $stratumNum) { "UNKNOWN" } else {
        switch ($stratumNum) {
            1           { "1 (atomic clock)"; break }
            {$_ -le 5}  { "$stratumNum hops from atomic"; break }
            {$_ -le 10} { "$stratumNum hops from atomic - acceptable"; break }
            default     { "$stratumNum hops from atomic - too far" }
        }
    }
    Write-Host "  Source         : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $source -ForegroundColor Yellow
    Write-Host "  Server IP      : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $serverIP -ForegroundColor Yellow
    Write-Host "  Last Sync      : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $lastSync -ForegroundColor Yellow
    Write-Host "  Stratum        : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $stratumDisplay -ForegroundColor Yellow
    Write-Host "  Max Clock Error: " -ForegroundColor $script:SubLabelFG -NoNewline
    Write-Host $dispersion -ForegroundColor $dispColor -NoNewline
    if ($dispNote) { Write-Host $dispNote -ForegroundColor DarkYellow } else { Write-Host "" }
    Write-Host "  Clock Status   : " -ForegroundColor $script:SubLabelFG -NoNewline; Write-Host $leap -ForegroundColor $leapColor

    # Verdict. Configured, synced and unknown are kept apart, and a claim is only made once checked.
    $SourceOk = Test-TickSourceMatches -Source $reportedSource -Expected $NtpServer -IsDomainRoute $UseDomainSync -DomainName $domainName
    $LastSyncTime = $null
    $tmpT = [datetime]::MinValue
    if ($s['Last Successful Sync Time'] -and [datetime]::TryParse($s['Last Successful Sync Time'], [Globalization.CultureInfo]::CurrentCulture, [Globalization.DateTimeStyles]::None, [ref]$tmpT)) { $LastSyncTime = $tmpT }
    $Problem = $null
    $ExitCode = 0
    if ($ConfigProblems.Count -gt 0) {
        $ExitCode = 1
        $Problem = "the time service was not configured cleanly: " + ($ConfigProblems -join "; ")
    } elseif (-not $ResyncOk) {
        $ExitCode = 3
        $Problem = "the time source was configured but the resync failed, so the sync is not confirmed"
    } elseif (-not $SourceOk) {
        $ExitCode = 3
        $Problem = "the sync is not confirmed, the time source reports '$source' instead of '$NtpServer'"
    } elseif ($null -eq $LastSyncTime) {
        $ExitCode = 3
        $Problem = "the sync is not confirmed, the last successful sync time is UNKNOWN (status output not recognised)"
    } elseif ([Math]::Abs(((Get-Date) - $LastSyncTime).TotalMinutes) -gt 10) {
        $ExitCode = 3
        $Problem = "the sync is not confirmed, the last successful sync was $lastSync"
    }
    if ($ExitCode -eq 0 -and $UseDomainSync) {
        Write-Host " This computer is now synced to the Domain Controller." -ForegroundColor DarkGreen
    }
    # Footer Line 1 (The Top Bar of the Footer)
    Write-Host ("-" * 69) -ForegroundColor $BorderCol -NoNewline
    Write-Host " ╔╦╗ ╦ ╔═╗ ╦╔═ " -ForegroundColor $ArtCol -NoNewline
    Write-Host "█" -ForegroundColor $BorderCol
    # Footer Line 2 (The Text Line)
    $FooterText = "  C H R O N O M E T R I C   S E Q U E N C E   C O M P L E T E"
    $Padding = " " * (69 - $FooterText.Length)
    Write-Host "$FooterText$Padding" -ForegroundColor $AccentCol -NoNewline
    Write-Host "  ║  ║ ║   ╠╩╗ " -ForegroundColor $ArtCol -NoNewline
    Write-Host "█" -ForegroundColor $BorderCol
    # Footer Line 3 (The Bottom Bar of the Footer)
    Write-Host ("-" * 69) -ForegroundColor $BorderCol -NoNewline
    Write-Host "  ╩  ╩ ╚═╝ ╩ ╩ " -ForegroundColor $ArtCol -NoNewline
    Write-Host "█" -ForegroundColor $BorderCol
    #endregion
    if ($ExitCode -eq 1) {
        Write-Output "[FAIL] TICK $Problem (exit 1)"
        Exit-KillerScript 1
    }
    if ($ExitCode -eq 3) {
        Write-Output "[WARN] TICK partial: $Problem (exit 3)"
        Exit-KillerScript 3
    }
    Write-Output "[PASS] TICK synced to $source, last successful sync $lastSync (exit 0)"
    Exit-KillerScript 0
}

if (-not $global:KillerScriptsTestMode) { Invoke-TickMain -Silent:$Silent -NtpServer $NtpServer }
