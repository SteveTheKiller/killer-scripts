<#
.SYNOPSIS
    Bulk Endpoint App Mover (BEAM) v2.00
    Created by Steve the Killer | Updated: 2026-11-18
.DESCRIPTION
    Packages any local or hosted .exe / .msi as an Intune Win32 app and deploys it to
    one or more Entra-joined devices as SYSTEM. No LAN, no stored credentials. For
    machines reachable only through the cloud (Intune check-in), where a local install
    or a remote-control repair is not practical.
.NOTES
    Intune Win32 app deployment only (SYSTEM context).
    Run from your own workstation: enter an installer source (URL, or blank to browse
    for a local file), optional install switches, the target hostnames, and a Global
    Admin email each run. First use in a new tenant needs one-time admin consent to the
    Graph permissions, or Graph returns Forbidden.

    Each run has a unique run ID. The install runs your package and stamps the ID into a
    marker only when the installer exits 0 or a reboot-required 3010/1641, so detection
    matches only a successful install. Any other installer exit code is passed through to
    Intune and leaves no marker. Each app has its own marker. Re-running installs again.
    Detection is marker-based and does not inspect the app itself, so it works for any
    installer.

    Each run creates its own Entra group (BEAM-<app>-<run ID>), so members left over from an
    earlier run are never assigned the app. A target is reported as queued only when its
    device matched exactly one Intune/Entra record, was verified in the group, and the
    assignment succeeded. The package tool is pinned to a release and its Microsoft signature
    is verified before it runs. The installer's source URL, SHA256 and signer are shown
    before packaging, and a link scraped from a web page needs an explicit yes.

    Exit codes: 0 success, 1 failure, 2 a precondition failed (missing input or no console),
    3 partial success. The last line of output states the result and agrees with the code.
#>

#region UI Initialization
# RMM agents often run 32-bit PowerShell, where System32 and HKLM\SOFTWARE are redirected to their
# 32-bit copies. Relaunch in 64-bit PowerShell with the same arguments so tools and registry writes land.
if ([IntPtr]::Size -eq 4 -and $env:PROCESSOR_ARCHITEW6432 -and $PSCommandPath -and -not $global:KillerScriptsTestMode) {
    $ksArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', $PSCommandPath)
    foreach ($ksKey in $PSBoundParameters.Keys) {
        $ksVal = $PSBoundParameters[$ksKey]
        if ($ksVal -is [System.Management.Automation.SwitchParameter]) { if ($ksVal.IsPresent) { $ksArgs += "-$ksKey" } }
        else { $ksArgs += "-$ksKey"; $ksArgs += [string]$ksVal }
    }
    $ksArgs += $args
    $ksPs = "$env:windir\Sysnative\WindowsPowerShell\v1.0\powershell.exe"
    $ksNoInput = $true
    try { $ksNoInput = [Console]::IsInputRedirected -or -not [Environment]::UserInteractive } catch { }
    # With no real keyboard, hand the child a closed stdin so its prompts fail closed instead of hanging.
    if ($ksNoInput) { $null | & $ksPs @ksArgs } else { & $ksPs @ksArgs }
    exit $LASTEXITCODE
}
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
$OutputEncoding            = [System.Text.Encoding]::UTF8
$ProgressPreference        = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

$script:Width = 85
$_ver         = "2.00"
$_fver        = "| v$_ver"

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

function Exit-BeamRun {
    # One exit point.
    param([int]$Code, [string]$Text)
    Exit-KillerScript $Code
}

# ESC handling requires a true interactive console. Detect it without touching the input buffer
# (accessing [Console]::KeyAvailable here breaks Read-Host in relayed consoles). In a
# redirected/relayed console ESC can't be read at all, so Ctrl+C is the cancel there.
$script:IsInteractive = $false
try { $script:IsInteractive = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected } catch { }

$LineCol   = "DarkCyan"
$MainCol   = "Cyan"
$BorderCol = "DarkCyan"
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
        $line = ([char]0x2501).ToString() * $Width
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

function Write-StepUpdate {
    param([string]$Message, [string]$Status = "INFO")
    $sc = @{ OK=$OkCol; FAIL="Red"; INFO=$InfoCol; WARN=$AccentCol; SKIP=$DimCol; TRY=$BorderCol }
    $c  = if ($sc.ContainsKey($Status)) { $sc[$Status] } else { $ArtCol }
    Write-Host "  [" -NoNewline -ForegroundColor $DimCol
    Write-Host $Status -ForegroundColor $c -NoNewline
    Write-Host "] " -NoNewline -ForegroundColor $DimCol
    Write-Host $Message -ForegroundColor $ArtCol
}

function Write-Detail {
    # Aligned label : value line, indented under a "--- host ---" header. Value color carries
    # the status (OK/WARN/FAIL) so the log stays scannable without a per-line status tag.
    param([string]$Label, [string]$Value, [string]$Status = "DATA")
    $vc = switch ($Status) {
        "OK"   { $OkCol }
        "WARN" { $AccentCol }
        "FAIL" { "Red" }
        default { $ArtCol }
    }
    Write-Host ("         " + $Label.PadRight(17) + ": ") -NoNewline -ForegroundColor $DimCol
    Write-Host $Value -ForegroundColor $vc
}

function Test-EscPressed {
    # Drains the key buffer; returns $true if ESC was pressed. Non-interactive-safe.
    if (-not $script:IsInteractive) { return $false }
    $pressed = $false
    try {
        while ([Console]::KeyAvailable) {
            if ([Console]::ReadKey($true).Key -eq [ConsoleKey]::Escape) { $pressed = $true }
        }
    } catch { }
    return $pressed
}

function Get-CsvTargets {
    # Flatten a CSV/list file into a unique target array. Splits on comma/semicolon/whitespace,
    # trims quotes, and drops common header words so a "Hostname" header row isn't treated as a host.
    param([string]$Path)
    $out  = @()
    $skip = @('hostname','host','computername','computer','name','ip','ipaddress','address','targets','target','device','devicename','machine','machinename')
    try { $raw = Get-Content -LiteralPath $Path -ErrorAction Stop } catch { return @() }
    foreach ($line in $raw) {
        foreach ($tok in ($line -split '[,;\s]+')) {
            $t = $tok.Trim().Trim('"').Trim("'")
            if ($t -eq "") { continue }
            if ($skip -contains $t.ToLower()) { continue }
            $out += $t
        }
    }
    return @($out | Select-Object -Unique)
}

function Select-FileBrowser {
    # Mini interactive browser. Up/Down move, Enter opens a folder or picks a matching file,
    # Backspace goes up, Esc cancels. Redraws in place each keystroke and wipes its whole region
    # on exit, leaving the cursor where it started so the surrounding output is untouched.
    # $Extensions filters which files are shown; $Label names the browser in the header.
    param(
        [string]$StartDir     = (Get-Location).Path,
        [string[]]$Extensions = @('.csv', '.txt'),
        [string]$Label        = "File browser"
    )
    $e        = [char]27
    $dir      = $StartDir
    $sel      = 0
    $startRow = [Console]::CursorTop
    $result   = $null
    $window   = 15

    while ($true) {
        $entries = @()
        $parent  = Split-Path $dir -Parent
        if ($parent) { $entries += [PSCustomObject]@{ Name = ".."; Path = $parent; Type = "up" } }
        try {
            Get-ChildItem -LiteralPath $dir -Directory -ErrorAction SilentlyContinue | Sort-Object Name | ForEach-Object {
                $entries += [PSCustomObject]@{ Name = "$($_.Name)\"; Path = $_.FullName; Type = "dir" }
            }
            Get-ChildItem -LiteralPath $dir -File -ErrorAction SilentlyContinue | Where-Object { $_.Extension -in $Extensions } | Sort-Object Name | ForEach-Object {
                $entries += [PSCustomObject]@{ Name = $_.Name; Path = $_.FullName; Type = "file" }
            }
        } catch { }
        if ($sel -ge $entries.Count) { $sel = $entries.Count - 1 }
        if ($sel -lt 0) { $sel = 0 }

        [Console]::SetCursorPosition(0, $startRow)
        Write-Host "${e}[0J" -NoNewline
        Write-HLine dashed
        Write-Host "  $Label  " -NoNewline -ForegroundColor $InfoCol
        Write-Host $dir -ForegroundColor $ArtCol
        Write-Host "  Up/Down move   Enter open/select   Backspace up   Esc cancel" -ForegroundColor $DimCol
        Write-Host ""
        if ($entries.Count -eq 0) {
            Write-Host "    (no folders or matching files here)" -ForegroundColor $DimCol
        } else {
            $start = [Math]::Max(0, [Math]::Min($sel - [int]($window / 2), $entries.Count - $window))
            if ($start -lt 0) { $start = 0 }
            $end = [Math]::Min($entries.Count - 1, $start + $window - 1)
            for ($i = $start; $i -le $end; $i++) {
                $en  = $entries[$i]
                $col = switch ($en.Type) { "file" { $AccentCol } "up" { $DimCol } default { $InfoCol } }
                if ($i -eq $sel) {
                    Write-Host "  > " -NoNewline -ForegroundColor $BorderCol
                    Write-Host $en.Name -ForegroundColor $col
                } else {
                    Write-Host "    " -NoNewline
                    Write-Host $en.Name -ForegroundColor $col
                }
            }
        }

        $k = [Console]::ReadKey($true)
        $done = $false
        switch ($k.Key) {
            "UpArrow"   { if ($sel -gt 0) { $sel-- } }
            "DownArrow" { if ($sel -lt $entries.Count - 1) { $sel++ } }
            "Backspace" { $p = Split-Path $dir -Parent; if ($p) { $dir = $p; $sel = 0 } }
            "Escape"    { $result = $null; $done = $true }
            "Enter"     {
                if ($entries.Count -gt 0) {
                    $en = $entries[$sel]
                    if ($en.Type -eq "file") { $result = $en.Path; $done = $true }
                    else { $dir = $en.Path; $sel = 0 }
                }
            }
        }
        if ($done) { break }
    }

    [Console]::SetCursorPosition(0, $startRow)
    Write-Host "${e}[0J" -NoNewline
    return $result
}

function Read-TargetsField {
    # Read-Host for the targets field, plus an [F] shortcut that opens the CSV browser. Falls back
    # to plain Read-Host on non-interactive/relayed consoles where key reading isn't available.
    # A paste arrives as a burst of keystrokes, so a hostname list starting with F (FUNDPHX-...)
    # would otherwise fire the file browser and the rest of the paste would drive it. F only counts
    # as the shortcut when nothing else is queued behind it (a human keypress); F2 always works.
    param([string]$Prompt)
    if (-not $script:IsInteractive) { return (Read-BeamValue $Prompt) }

    $disp = "$Prompt  [F = load CSV/TXT]"
    Write-Host "${disp}: " -NoNewline
    $buf = ""
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::Enter) { Write-Host ""; return $buf }

        $pending = $false
        try { $pending = [Console]::KeyAvailable } catch { }
        $wantsCsv = ($k.Key -eq [ConsoleKey]::F2) -or
                    ($buf -eq "" -and -not $pending -and ($k.KeyChar -eq 'f' -or $k.KeyChar -eq 'F'))

        if ($wantsCsv) {
            Write-Host ""
            $csv = Select-FileBrowser -Extensions @('.csv', '.txt') -Label "CSV/TXT browser"
            if ($csv) {
                $tgs = Get-CsvTargets $csv
                if ($tgs.Count -gt 0) {
                    Write-StepUpdate "Loaded $($tgs.Count) target(s) from $(Split-Path $csv -Leaf)" OK
                    return ($tgs -join ",")
                }
                Write-StepUpdate "No targets found in $(Split-Path $csv -Leaf): type them instead" WARN
            }
            Write-Host "${disp}: " -NoNewline
            continue
        }
        if ($k.Key -eq [ConsoleKey]::Backspace) {
            if ($buf.Length -gt 0) { $buf = $buf.Substring(0, $buf.Length - 1); Write-Host "`b `b" -NoNewline }
            continue
        }
        if ($k.KeyChar -and -not [char]::IsControl($k.KeyChar)) { $buf += $k.KeyChar; Write-Host $k.KeyChar -NoNewline }
    }
}

function Read-BeamValue {
    # Read-Host that fails closed. With no console to read from it throws BEAM_NOCONSOLE, which
    # the entry point turns into exit 1, instead of an unhandled host error.
    param([string]$Prompt)
    try {
        if ($Prompt) { $v = Read-Host $Prompt } else { $v = Read-Host }
    } catch { throw 'BEAM_NOCONSOLE' }
    if ($null -eq $v) { return '' }
    return [string]$v
}

function Read-BeamConfirm {
    # A yes/no question that defaults to no. No console, a blank answer or anything but y/yes is no.
    param([string]$Prompt)
    try { $a = Read-BeamValue "$Prompt [y/N]" } catch { return $false }
    return ($a.Trim() -match '^(?i)y(es)?$')
}

function Show-BeamInstallerIdentity {
    # Shows what is about to be packaged and deployed as SYSTEM: where it came from, its SHA256 and
    # its Authenticode signer. Returns the signature status so a caller can react to an unsigned file.
    param([string]$Path, [string]$SourceUrl)
    if ($SourceUrl) { Write-Detail "Source URL" $SourceUrl }
    $hash = ''
    try { $hash = (Get-FileHash -LiteralPath $Path -Algorithm SHA256 -ErrorAction Stop).Hash } catch { $hash = "unavailable: $($_.Exception.Message)" }
    Write-Detail "SHA256" $hash
    $status = 'Unknown'; $signer = ''
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        $status = "$($sig.Status)"
        if ($sig.SignerCertificate) { $signer = "$($sig.SignerCertificate.Subject)" }
    } catch { $status = "unreadable ($($_.Exception.Message))" }
    if ($status -eq 'Valid') { Write-Detail "Signer" $signer OK }
    elseif ($signer)         { Write-Detail "Signer" "$signer (signature status: $status)" WARN }
    else                     { Write-Detail "Signer" "none (signature status: $status)" WARN }
    return $status
}

function Show-BeamBanner {
    $_pfx  = "█  "
    $_art1 = "╔╗  ╔═╗ ╔═╗ ╔╦╗ "
    $_art2 = "╠╩╗ ║╣  ╠═╣ ║║║ "
    $_art3 = "╚═╝ ╚═╝ ╩ ╩ ╩ ╩ "
    $_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1  = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
    $_fillW = $script:Width - $_pfx.Length - $_artW
    $_title = "BULK ENDPOINT APP MOVER"
    $_tpad  = " " * [Math]::Max(0, ($_fillW - $_title.Length))

    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host ""
    Write-HLine dashed
}

function Show-BeamFooter {
    param([string]$Title = "  DEPLOYMENT COMPLETE")
    $_art1 = "╔╗  ╔═╗ ╔═╗ ╔╦╗ "
    $_art2 = "╠╩╗ ║╣  ╠═╣ ║║║ "
    $_art3 = "╚═╝ ╚═╝ ╩ ╩ ╩ ╩ "
    $_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1  = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
    $_sfx    = "█"
    $_ffillW = $script:Width - $_artW - 1 - $_sfx.Length
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $Title.Length - $_fver.Length))

    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host "$Title$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
}

function New-BeamAbort {
    param([int]$Code, [string]$Text)
    [pscustomobject]@{ Abort = $true; Code = $Code; Text = $Text }
}

function Read-BeamInputs {
    # Collects every operator input. Returns an Abort object (Code, Text) when a precondition is
    # not met, otherwise the settings the deployment needs.
    Write-StepUpdate "Collecting deployment parameters..." INFO
    Write-Host ""

    # 1. Installer source. Type a direct download URL, or press Enter to browse for a local
    #    .exe / .msi. The file type drives the install command (msiexec for MSI, direct run for EXE).
    $installerUrl   = $null
    $localInstaller = $null
    $setupFile      = $null
    $isMsi          = $false

    while (-not $setupFile) {
        Write-Host "  Installer URL  [blank = browse for a local .exe/.msi]: " -NoNewline
        $srcInput = (Read-BeamValue '').Trim().Trim('"')
        Write-Host ""

        if ($srcInput -eq "") {
            if (-not $script:IsInteractive) {
                Write-StepUpdate "The file picker needs an interactive console. Provide a URL instead: aborting" FAIL
                return (New-BeamAbort 1 "the file picker needs an interactive console. Provide a URL instead. Nothing was changed.")
            }
            $picked = Select-FileBrowser -Extensions @('.exe', '.msi') -Label "Installer browser (.exe/.msi)"
            if (-not $picked) {
                Write-StepUpdate "No file selected. Enter a URL or press F/Enter to browse (Ctrl+C aborts)." WARN
                Write-Host ""
                continue
            }
            $localInstaller = $picked
            $setupFile      = Split-Path $picked -Leaf
            $isMsi          = ([IO.Path]::GetExtension($setupFile).ToLower() -eq '.msi')
            Write-StepUpdate "Selected: $setupFile" OK
        } else {
            $installerUrl = $srcInput
            $pathPart     = ($srcInput -split '\?')[0]
            $ext          = [IO.Path]::GetExtension($pathPart).ToLower()
            $leaf         = Split-Path $pathPart -Leaf
            if ($ext -eq '.msi') {
                $isMsi = $true
                $setupFile = $leaf
            } elseif ($ext -eq '.exe') {
                $isMsi = $false
                $setupFile = $leaf
            } else {
                # URL gives no usable extension. Ask the type so the install command is built correctly.
                $ans   = (Read-BeamValue "  Couldn't read a file type from the URL. Is it an MSI? (y/N)").Trim()
                $isMsi = ($ans -match '^(y|yes)$')
                $setupFile = if ($isMsi) { "installer.msi" } else { "installer.exe" }
            }
            Write-StepUpdate "Source URL accepted ($setupFile)" OK
        }
    }
    Write-Host ""

    # 2. App name shown in Intune. Blank derives it from the installer file name.
    $defaultName = [IO.Path]::GetFileNameWithoutExtension($setupFile)
    $appName     = (Read-BeamValue "  App name  [blank = '$defaultName']").Trim()
    if ($appName -eq "") { $appName = $defaultName }

    # 3. Publisher (required by Intune, cosmetic). Blank defaults to BEAM.
    $publisher = (Read-BeamValue "  Publisher  [blank = 'BEAM']").Trim()
    if ($publisher -eq "") { $publisher = "BEAM" }

    # 4. Install switches. Optional. MSI defaults to /qn when blank; EXE runs bare when blank.
    $phSwitch    = if ($isMsi) { "MSI: blank uses /qn" } else { "EXE: e.g. /S  /silent  /qn  -s" }
    $installArgs = (Read-BeamValue "  Install switches (optional)  [$phSwitch]").Trim()
    if ($isMsi -and $installArgs -eq "") { $installArgs = "/qn" }
    Write-Host ""

    # 5. Targets. Intune matches by hostname, so IP entries are dropped.
    $targetInput  = Read-TargetsField "  Targets (comma-separated hostnames, IPs can't match Intune)"
    $targets      = @($targetInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" })
    $ipTargets    = @($targets | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' })
    if ($ipTargets.Count -gt 0) {
        Write-StepUpdate "Dropping IP target(s), Intune matches by hostname only: $($ipTargets -join ', ')" WARN
        $targets = @($targets | Where-Object { $_ -notmatch '^\d{1,3}(\.\d{1,3}){3}$' })
    }
    if ($targets.Count -eq 0) {
        Write-StepUpdate "No hostname targets to deploy: aborting" FAIL
        return (New-BeamAbort 1 "no hostname targets to deploy. Nothing was changed.")
    }
    Write-Host ""

    # 6. Global Admin for Intune. The domain resolves the tenant ID for the sign-in.
    $adminUpn = Read-BeamValue "  Global admin email for Intune (required)"
    if (-not $adminUpn) {
        Write-StepUpdate "A Global Admin / Intune Admin sign-in is required for this tool: aborting" FAIL
        return (New-BeamAbort 1 "an admin sign-in email is required. Nothing was changed.")
    }
    $adminDomain = ($adminUpn -split "@")[-1]
    if (-not $adminDomain -or $adminDomain -eq $adminUpn) {
        Write-StepUpdate "Could not parse a domain from '$adminUpn': aborting" FAIL
        return (New-BeamAbort 1 "could not parse a domain from the admin email. Nothing was changed.")
    }
    Write-StepUpdate "Resolving tenant for domain: $adminDomain" INFO
    $tenantId = $null
    try {
        $oidc     = Invoke-RestMethod "https://login.microsoftonline.com/$adminDomain/v2.0/.well-known/openid-configuration" -ErrorAction Stop
        $tenantId = ([regex]::Match($oidc.issuer, '(?i)[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}')).Value
    } catch { }
    if (-not $tenantId) {
        Write-StepUpdate "Tenant lookup failed for '$adminDomain': aborting" FAIL
        return (New-BeamAbort 1 "tenant lookup failed for '$adminDomain'. Nothing was changed.")
    }
    Write-StepUpdate "Tenant ID: $tenantId" OK

    Write-Host ""
    Write-HLine dashed
    Write-Host ""
    [pscustomobject]@{
        Abort = $false; InstallerUrl = $installerUrl; LocalInstaller = $localInstaller; SetupFile = $setupFile
        IsMsi = $isMsi; AppName = $appName; Publisher = $publisher; InstallArgs = $installArgs
        Targets = @($targets); TenantId = $tenantId
    }
}

function Get-BeamWorkRoot {
    $base = $env:LOCALAPPDATA
    if (-not $base) { $base = $env:TEMP }
    Join-Path $base "BEAM"
}

# IntuneWinAppUtil.exe is fetched from a fixed release tag, never a moving branch, and is only ever
# executed after its Authenticode signature is valid and names Microsoft Corporation as the signer.
$script:BeamToolTag = 'v1.8.7'

function Test-BeamToolSignature {
    # True only for a valid Authenticode signature whose subject has O=Microsoft Corporation.
    param([string]$Path)
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        if ("$($sig.Status)" -ne 'Valid' -or -not $sig.SignerCertificate) { return $false }
        return ("$($sig.SignerCertificate.Subject)" -match '(^|,\s*)O=Microsoft Corporation(,|$)')
    } catch { return $false }
}

function Initialize-BeamTool {
    # Returns the path of a verified IntuneWinAppUtil.exe, downloading it when missing or when the
    # cached copy no longer verifies, or $null on failure. A file that fails verification is deleted
    # and never run.
    $workRoot = Get-BeamWorkRoot
    $toolPath = Join-Path $workRoot "IntuneWinAppUtil.exe"
    try { if (-not (Test-Path -LiteralPath $workRoot)) { New-Item -ItemType Directory -Path $workRoot -Force -ErrorAction Stop | Out-Null } }
    catch { Write-StepUpdate "Could not create ${workRoot}: $($_.Exception.Message)" FAIL; return $null }

    if (Test-Path -LiteralPath $toolPath) {
        if (Test-BeamToolSignature $toolPath) {
            Write-StepUpdate "IntuneWinAppUtil found at $toolPath (Microsoft signature verified)" OK
            Write-Host ""
            return $toolPath
        }
        Write-StepUpdate "The cached IntuneWinAppUtil does not carry a valid Microsoft signature: removing it." WARN
        try { Remove-Item -LiteralPath $toolPath -Force -ErrorAction Stop }
        catch { Write-StepUpdate "Could not remove the unverified copy: $($_.Exception.Message)" FAIL; return $null }
    }

    $url = "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/$($script:BeamToolTag)/IntuneWinAppUtil.exe"
    $tmp = Join-Path $workRoot ("IntuneWinAppUtil-" + [guid]::NewGuid().ToString('N').Substring(0, 8) + ".download")
    Write-StepUpdate "Downloading IntuneWinAppUtil $($script:BeamToolTag)..." INFO
    try {
        Invoke-WebRequest -Uri $url -OutFile $tmp -ErrorAction Stop
        if (-not (Test-BeamToolSignature $tmp)) {
            Write-StepUpdate "The downloaded IntuneWinAppUtil does not carry a valid Microsoft signature: discarded, nothing was run." FAIL
            return $null
        }
        Move-Item -LiteralPath $tmp -Destination $toolPath -Force -ErrorAction Stop
        Write-StepUpdate "IntuneWinAppUtil ready at $toolPath (Microsoft signature verified)" OK
    } catch {
        Write-StepUpdate "IntuneWinAppUtil download failed: cannot continue" FAIL
        return $null
    } finally {
        if (Test-Path -LiteralPath $tmp) { try { Remove-Item -LiteralPath $tmp -Force -ErrorAction Stop } catch { } }
    }
    Write-Host ""
    return $toolPath
}
function Get-BeamUtcNow { [DateTime]::UtcNow }

function Invoke-BeamPackager {
    # Runs IntuneWinAppUtil.exe. Its output is not needed, only the .intunewin it writes.
    param([string]$Tool, [string]$SourceDir, [string]$SetupFile, [string]$OutDir)
    & $Tool -c $SourceDir -s $SetupFile -o $OutDir -q 2>&1 | Out-Null
}

function Open-BeamListener {
    # Loopback listener on a free port. "localhost" prefix needs no URL ACL / elevation.
    $listener = New-Object System.Net.HttpListener
    for ($p = 8400; $p -le 8420; $p++) {
        try {
            $listener.Prefixes.Clear(); $listener.Prefixes.Add("http://localhost:$p/"); $listener.Start()
            return [pscustomobject]@{ Listener = $listener; Port = $p }
        } catch { }
    }
    return $null
}

function Start-BeamBrowser {
    param([string]$Url)
    try { Start-Process $Url } catch { }
}

function Receive-BeamRedirect {
    # Waits for the loopback redirect, with a timeout so it never hangs forever. Returns
    # Cancelled = $true when the operator pressed ESC, otherwise the query string or $null.
    param($Listener, [int]$TimeoutSec = 180)
    $query    = $null
    $ctxTask  = $Listener.GetContextAsync()
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while (-not $ctxTask.IsCompleted -and (Get-Date) -lt $deadline) {
        if (Test-EscPressed) { try { $Listener.Stop() } catch { }; return [pscustomobject]@{ Cancelled = $true; Query = $null } }
        Start-Sleep -Milliseconds 300
    }
    if ($ctxTask.IsCompleted) {
        try {
            $ctx  = $ctxTask.Result
            $query = $ctx.Request.Url.Query
            $html = "<html><body style='font-family:Segoe UI;background:#111;color:#00CED1;text-align:center;padding-top:60px'><h2>BEAM authenticated. You can close this tab.</h2></body></html>"
            $buf  = [System.Text.Encoding]::UTF8.GetBytes($html)
            $ctx.Response.ContentType = "text/html"; $ctx.Response.StatusCode = 200
            $ctx.Response.OutputStream.Write($buf, 0, $buf.Length)
            $ctx.Response.OutputStream.Close()
        } catch { }
    }
    try { $Listener.Stop() } catch { }
    [pscustomobject]@{ Cancelled = $false; Query = $query }
}

function Get-BeamClipboardText {
    $t = $null
    try { $t = (Get-Clipboard -Raw) } catch { try { $t = (Get-Clipboard) -join '' } catch { $t = $null } }
    return $t
}

function Get-GraphToken {
    param([string]$TenantId)
    $clientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    $scope    = "offline_access DeviceManagementApps.ReadWrite.All DeviceManagementManagedDevices.ReadWrite.All DeviceManagementManagedDevices.PrivilegedOperations.All Group.ReadWrite.All Device.Read.All"

    # Authorization-code + PKCE with a loopback listener that captures the redirect
    # automatically, no pasting. This is why BEAM should be run on your own workstation
    # (where the sign-in browser lives), not proxied through a remote console: the listener
    # and the browser must be on the same machine. Device code flow is CA-blocked; this is a
    # normal browser sign-in (MFA / compliant-device checks still apply).

    # PKCE pair (base64url, no padding).
    $rng      = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $vbytes   = New-Object byte[] 32
    $rng.GetBytes($vbytes)
    $verifier = (([Convert]::ToBase64String($vbytes)) -replace '\+','-' -replace '/','_' -replace '=','')
    $chalRaw  = [System.Security.Cryptography.SHA256]::Create().ComputeHash([System.Text.Encoding]::ASCII.GetBytes($verifier))
    $challenge = (([Convert]::ToBase64String($chalRaw)) -replace '\+','-' -replace '/','_' -replace '=','')
    $state    = [Guid]::NewGuid().ToString("N")

    $lb = Open-BeamListener
    if (-not $lb) { Write-StepUpdate "Could not bind a loopback port (8400-8420) for auth" FAIL; return $null }
    $port        = $lb.Port
    $redirectUri = "http://localhost:$port/"

    $authUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/authorize?" +
        "client_id=$clientId&response_type=code" +
        "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))&response_mode=query" +
        "&scope=$([Uri]::EscapeDataString($scope))&state=$state" +
        "&code_challenge=$challenge&code_challenge_method=S256&prompt=select_account"

    Write-Host ""
    Write-StepUpdate "A browser sign-in is required. This works best when BEAM runs on your own" INFO
    Write-StepUpdate "workstation (not through a remote relayed console)." INFO
    Write-Host ""
    Write-StepUpdate "Opening browser... if it doesn't open, paste this URL into your browser:" WARN
    Write-Host "  $authUrl" -ForegroundColor $InfoCol
    Write-Host ""
    Start-BeamBrowser $authUrl

    Write-StepUpdate "Waiting for sign-in to complete (up to 180s, ESC to cancel)..." INFO
    $redirect = Receive-BeamRedirect -Listener $lb.Listener
    if ($redirect.Cancelled) { Write-StepUpdate "Sign-in cancelled by operator" WARN; return $null }
    $query = $redirect.Query

    # Last resort (e.g. browser on another machine): read the redirect URL from the clipboard.
    if (-not ($query -and $query -match 'code=')) {
        Write-Host ""
        Write-StepUpdate "No local redirect captured (browser likely on another machine)." WARN
        Write-StepUpdate "Copy the FULL address-bar URL from the browser, then press Enter here." INFO
        Read-BeamValue "  (press Enter after copying the URL)" | Out-Null
        $query = Get-BeamClipboardText
    }
    if (-not $query) { Write-StepUpdate "No authorization response obtained: aborting auth" FAIL; return $null }

    # Extract code (guard UnescapeDataString so a stray % can't throw). [?&] anchor skips session_state.
    # The returned state is mandatory and must be exactly the one BEAM issued, exactly once. A missing,
    # empty, repeated or different state aborts before anything is sent to the token endpoint. Checked
    # for both the loopback redirect and a pasted clipboard URL, and before an error is reported.
    $stateVals = @([regex]::Matches($query, '[?&]state=([^&\s]*)') | ForEach-Object { $_.Groups[1].Value })
    $stateOk = $false
    if ($stateVals.Count -eq 1) {
        $got = $stateVals[0]
        try { $got = [Uri]::UnescapeDataString($got) } catch { }
        $stateOk = ($got -ceq $state)
    }
    if (-not $stateOk) {
        Write-StepUpdate "OAuth state missing or not the one BEAM issued: aborting auth" FAIL; return $null
    }

    $code = $null
    if ($query -match '[?&]code=([^&\s]+)') {
        try { $code = [Uri]::UnescapeDataString($matches[1]) } catch { $code = $matches[1] }
    } elseif ($query -match 'error=([^&\s]+)') {
        Write-StepUpdate "Authorization failed: $([Uri]::UnescapeDataString($matches[1]))" FAIL; return $null
    }
    if (-not $code) { Write-StepUpdate "No authorization code found in response" FAIL; return $null }

    $tokenBody = "client_id=$clientId&grant_type=authorization_code" +
        "&code=$([Uri]::EscapeDataString($code))" +
        "&redirect_uri=$([Uri]::EscapeDataString($redirectUri))" +
        "&code_verifier=$verifier&scope=$([Uri]::EscapeDataString($scope))"
    try {
        $token = Invoke-RestMethod "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Method POST -Body $tokenBody -ContentType "application/x-www-form-urlencoded" -ErrorAction Stop
    } catch {
        Write-StepUpdate "Token exchange failed: $_" FAIL
        return $null
    }
    Write-StepUpdate "Graph authenticated (interactive)" OK
    return $token.access_token
}

function Send-AzureStorageBlock {
    param([string]$SasUri, [string]$FilePath)
    $chunkSize = 5 * 1024 * 1024
    $stream    = [System.IO.File]::OpenRead($FilePath)
    $blockIds  = [System.Collections.Generic.List[string]]::new()
    $seq       = 0
    try {
        $buffer = New-Object byte[] $chunkSize
        while (($read = $stream.Read($buffer, 0, $chunkSize)) -gt 0) {
            $seq++
            $blockId = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($seq.ToString("D6")))
            $blockIds.Add($blockId)
            if ($read -eq $chunkSize) {
                $chunk = $buffer
            } else {
                $chunk = New-Object byte[] $read
                [Array]::Copy($buffer, 0, $chunk, 0, $read)
            }
            $putUri = "$SasUri&comp=block&blockid=$([Uri]::EscapeDataString($blockId))"
            $wc     = New-Object System.Net.WebClient
            $wc.Headers.Add("x-ms-blob-type", "BlockBlob")
            $wc.UploadData($putUri, "PUT", $chunk) | Out-Null
            $wc.Dispose()
        }
    } finally { $stream.Close() }
    $blockListBody = '<?xml version="1.0" encoding="utf-8"?><BlockList>' + (($blockIds | ForEach-Object { "<Latest>$_</Latest>" }) -join '') + '</BlockList>'
    $wc2 = New-Object System.Net.WebClient
    $wc2.Headers.Add("Content-Type", "application/xml")
    $wc2.UploadString("$SasUri&comp=blocklist", "PUT", $blockListBody) | Out-Null
    $wc2.Dispose()
}

function Wait-GraphState {
    param([string]$Uri, [hashtable]$Headers, [string]$Property, [string[]]$SuccessValues, [string[]]$FailValues, [int]$TimeoutSec = 300)
    $deadline = (Get-BeamUtcNow).AddSeconds($TimeoutSec)
    while ((Get-BeamUtcNow) -lt $deadline) {
        $r = Invoke-BeamGraph $Uri $Headers
        $v = $r.$Property
        if ($SuccessValues -contains $v) { return $r }
        if ($FailValues    -contains $v) { throw "Graph state failure: $v" }
        Start-Sleep 5
    }
    throw "Timeout waiting for Graph state: $Property"
}

function Test-InstallerHeader {
    # Returns $true if the file's magic bytes match the expected installer type.
    # EXE is a PE ("MZ"); MSI is an OLE compound document (D0 CF 11 E0).
    param([string]$Path, [switch]$IsMsi)
    try {
        $fs  = [IO.File]::OpenRead($Path)
        $hdr = New-Object byte[] 8
        [void]$fs.Read($hdr, 0, 8)
        $fs.Close()
    } catch { return $false }
    if ($IsMsi) {
        return ($hdr.Length -ge 4 -and $hdr[0] -eq 0xD0 -and $hdr[1] -eq 0xCF -and $hdr[2] -eq 0x11 -and $hdr[3] -eq 0xE0)
    }
    if ($hdr.Length -lt 2) { return $false }
    return ([Text.Encoding]::ASCII.GetString($hdr, 0, 2) -eq "MZ")
}

function Resolve-DirectDownload {
    # Some portals answer a download URL with a landing page rather than the binary. Given the
    # saved page, pull a real installer link out of the markup and return its absolute URL, or
    # $null if nothing usable is found. Relative links are resolved against the page URL. Silent:
    # the caller decides whether to retry and what, if anything, to print.
    param([string]$PagePath, [string]$PageUrl)
    try { $html = Get-Content -LiteralPath $PagePath -Raw -ErrorAction Stop } catch { return $null }
    if (-not $html) { return $null }

    $cand = New-Object System.Collections.Generic.List[string]
    # id-based package endpoint some portals use to serve the installer
    foreach ($m in [regex]::Matches($html, '(?i)([a-z0-9_\-./]*mkdefault\.asp\?id=[a-z0-9%]+)')) { $cand.Add($m.Groups[1].Value) }
    # href/src/action attributes pointing at an installer or that same endpoint
    foreach ($m in [regex]::Matches($html, '(?i)(?:href|src|action)\s*=\s*["'']([^"''<>]+?(?:kcssetup\.exe|mkdefault\.asp\?id=[a-z0-9%]+|\.exe|\.msi))(?:["''?#]|$)')) { $cand.Add($m.Groups[1].Value) }
    # bare absolute installer URLs anywhere in the page
    foreach ($m in [regex]::Matches($html, '(?i)(https?://[^\s"''<>]+?\.(?:exe|msi))(?:["''?#]|\b)')) { $cand.Add($m.Groups[1].Value) }

    $list = @($cand | Where-Object { $_ } | Select-Object -Unique)
    if ($list.Count -eq 0) { return $null }

    # Prefer the id-based endpoint, then a named setup, then any exe, then msi, then whatever is left.
    $pick = $list | Where-Object { $_ -match '(?i)mkdefault\.asp\?id=' } | Select-Object -First 1
    if (-not $pick) { $pick = $list | Where-Object { $_ -match '(?i)kcssetup\.exe' } | Select-Object -First 1 }
    if (-not $pick) { $pick = $list | Where-Object { $_ -match '(?i)\.exe($|[?#])' } | Select-Object -First 1 }
    if (-not $pick) { $pick = $list | Where-Object { $_ -match '(?i)\.msi($|[?#])' } | Select-Object -First 1 }
    if (-not $pick) { $pick = $list | Select-Object -First 1 }
    if (-not $pick) { return $null }

    try {
        if ($pick -match '^(?i)https?://') { return $pick }
        return ([Uri]::new([Uri]$PageUrl, $pick)).AbsoluteUri
    } catch { return $pick }
}

function Invoke-BeamGraph {
    # Every Graph call goes through here so a failure is always a terminating error. Without
    # -ErrorAction Stop, Invoke-RestMethod can write a non-terminating error and carry on, and the
    # caller would print OK for a call that failed. String bodies are sent as UTF-8 bytes: PS 5.1
    # sends string bodies as Latin-1, which mangles non-ASCII names.
    param([string]$Uri, [hashtable]$Headers, [string]$Method = 'GET', $Body)
    $p = @{ Uri = $Uri; Headers = $Headers; Method = $Method; ErrorAction = 'Stop' }
    if ($null -ne $Body) {
        if ($Body -is [string]) { $p.Body = [System.Text.Encoding]::UTF8.GetBytes($Body) } else { $p.Body = $Body }
        $p.ContentType = 'application/json; charset=utf-8'
    }
    Invoke-RestMethod @p
}

function Get-BeamAppKey {
    # A stable, file-name-safe identity for one deployed app: a readable slice of the name plus a
    # short hash of the exact display name, so two names that sanitize to the same text still differ.
    param([string]$DisplayName)
    $readable = ($DisplayName -replace '[^A-Za-z0-9._-]+', '_').Trim('_')
    if ($readable.Length -gt 40) { $readable = $readable.Substring(0, 40) }
    if (-not $readable) { $readable = 'app' }
    $sha  = [System.Security.Cryptography.SHA256]::Create()
    $hash = $sha.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($DisplayName.ToLowerInvariant()))
    $sha.Dispose()
    $short = ([BitConverter]::ToString($hash, 0, 4) -replace '-', '').ToLowerInvariant()
    return "$readable-$short"
}

function New-BeamInstallScript {
    # The PowerShell the Intune SYSTEM install command runs (wrapped in -EncodedCommand). It runs the
    # package, reads the installer's real exit code, and stamps the run ID marker only for an accepted
    # success: 0, or a reboot-required 3010 / 1641 (both reported to Intune as 3010, a soft reboot).
    # Any other code is returned unchanged with no marker, so detection fails and Intune sees the failure.
    # Working directory at install time is the extracted content folder, so the file is referenced
    # relatively. Single quotes in the file name / switches are doubled for the PS literal.
    param([string]$SetupFile, [string]$InstallArgs, [bool]$IsMsi, [string]$MarkerDir, [string]$MarkerPath, [string]$RunId)
    $sfEsc = $SetupFile   -replace "'", "''"
    $iaEsc = $InstallArgs -replace "'", "''"
    if ($IsMsi) {
        $al = "/i `"$sfEsc`""
        if ($iaEsc) { $al = "$al $iaEsc" }
        $runLine = "`$p = Start-Process -FilePath 'msiexec.exe' -ArgumentList '$al' -Wait -PassThru"
    } else {
        if ($iaEsc) {
            $runLine = "`$p = Start-Process -FilePath '.\$sfEsc' -ArgumentList '$iaEsc' -Wait -PassThru"
        } else {
            $runLine = "`$p = Start-Process -FilePath '.\$sfEsc' -Wait -PassThru"
        }
    }
    @"
`$ErrorActionPreference = 'Stop'
`$c = 1
try { $runLine; `$c = [int]`$p.ExitCode } catch { `$c = 1 }
if (`$c -eq 0 -or `$c -eq 3010 -or `$c -eq 1641) {
    try { New-Item -ItemType Directory -Path '$MarkerDir' -Force | Out-Null; Set-Content -Path '$MarkerPath' -Value '$RunId' -Encoding ascii -Force } catch { `$c = 1 }
    if (`$c -eq 1641) { `$c = 3010 }
}
exit `$c
"@
}

function New-BeamPublishResult {
    # What a publish returns: whether the app is published and assigned, and each target's outcome
    # ("Intune queued" or "FAIL"). Targets not listed here are failed by the caller.
    param([bool]$Published, $Targets)
    if ($null -eq $Targets) { $Targets = [ordered]@{} }
    [pscustomobject]@{ Published = $Published; Targets = $Targets }
}
function Invoke-BeamPublish {
    param(
        [string[]]$Targets,
        [string]$Url,
        [string]$LocalPath,
        [string]$SetupFile,
        [bool]$IsMsi,
        [string]$InstallArgs,
        [string]$AppName,
        [string]$Publisher,
        [string]$TenantId,
        [string]$RunId,
        [string]$RunRoot
    )
    $runId = $RunId
    $intuneWorkDir = Join-Path (Get-BeamWorkRoot) "IntuneWork"

    Write-HLine dashed
    Write-StepUpdate "Intune Win32 deployment" INFO
    Write-StepUpdate "Targets: $($Targets -join ', ')" INFO
    Write-Host ""

    if (-not $script:BeamToolPath) { Write-StepUpdate "IntuneWinAppUtil unavailable: aborted" FAIL; return (New-BeamPublishResult $false) }

    Write-StepUpdate "Authenticating to Graph (tenant: $TenantId)..." INFO
    $token = Get-GraphToken -TenantId $TenantId
    if (-not $token) { return (New-BeamPublishResult $false) }
    $h = @{ Authorization = "Bearer $token" }

    # Fast capability check: a tenant without Intune returns "not applicable to target tenant"
    # here, so bail before wasting the download/packaging. Only bail on that specific signal.
    try {
        Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$top=1" -Headers $h -ErrorAction Stop | Out-Null
    } catch {
        if ("$_" -match "not applicable to target tenant|not licensed|Tenant is not") {
            Write-StepUpdate "This tenant isn't provisioned/licensed for Intune: Win32 deployment unavailable." FAIL
            Write-Host ""
            return (New-BeamPublishResult $false)
        }
    }

    $safeName = ($AppName -replace '[\\/:*?"<>|]', '_')
    # Every run packages from folders it creates itself, named for the run ID, so nothing left by an
    # earlier or concurrent run (or by another app whose name sanitizes to the same text) can end up
    # in the package. A folder that already exists is not ours: refuse it instead of reusing it.
    if (-not $RunRoot) { $RunRoot = Join-Path $intuneWorkDir ("run-" + ($RunId -replace '[^A-Za-z0-9._-]', '_')) }
    if (Test-Path -LiteralPath $RunRoot) {
        Write-StepUpdate "Staging folder $RunRoot already exists, so it is not this run's: refusing to reuse it." FAIL
        return (New-BeamPublishResult $false)
    }
    $srcDir = Join-Path $RunRoot "src"
    $outDir = Join-Path $RunRoot "out"
    try {
        New-Item -ItemType Directory -Path $srcDir -Force -ErrorAction Stop | Out-Null
        New-Item -ItemType Directory -Path $outDir -Force -ErrorAction Stop | Out-Null
    } catch { Write-StepUpdate "Could not create the staging folder: $_" FAIL; return (New-BeamPublishResult $false) }

    $localExe = Join-Path $srcDir $SetupFile

    # Stage a fresh copy of the installer on every run (no caching): a re-run against a tenant where
    # the app/group already exist must still replace the content, so we re-stage, re-package, and
    # upload a new content version each time.
    if ($Url) {
        Write-StepUpdate "Downloading $SetupFile..." INFO
        try {
            Invoke-WebRequest -Uri ([uri]$Url) -OutFile $localExe -ErrorAction Stop
        } catch { Write-StepUpdate "Download failed: $_" FAIL; return (New-BeamPublishResult $false) }

        # If the URL handed back a page instead of the binary, pull the real download link out of
        # it and fetch that instead. Runs silently, so a resolvable page just downloads normally.
        $sourceUrl = $Url
        $scraped = $false
        if (-not (Test-InstallerHeader -Path $localExe -IsMsi:$IsMsi)) {
            $direct = Resolve-DirectDownload -PagePath $localExe -PageUrl $Url
            if ($direct -and $direct -ne $Url) {
                Write-StepUpdate "The URL returned a page, not an installer. It links to: $direct" WARN
                try { Invoke-WebRequest -Uri ([uri]$direct) -OutFile $localExe -ErrorAction Stop; $sourceUrl = $direct; $scraped = $true }
                catch { Write-StepUpdate "Download of the linked file failed: $_" FAIL; return (New-BeamPublishResult $false) }
            }
        }
        Write-StepUpdate "Downloaded: $localExe ($([math]::Round((Get-Item $localExe).Length/1MB,1)) MB)" OK
        Write-StepUpdate "Identity of the file that will be deployed as SYSTEM:" INFO
        $sigStatus = Show-BeamInstallerIdentity -Path $localExe -SourceUrl $sourceUrl
        if ($scraped) {
            # The admin typed a different URL. A link pulled out of a web page is not something they
            # chose, so it is never deployed without an explicit yes.
            if (-not (Read-BeamConfirm "This file was found in a web page, not typed by you. Deploy it")) {
                Write-StepUpdate "Not confirmed: nothing was packaged or published." FAIL
                return (New-BeamPublishResult $false)
            }
        }
    } else {
        Write-StepUpdate "Staging local file: $LocalPath" INFO
        try {
            Copy-Item -LiteralPath $LocalPath -Destination $localExe -Force -ErrorAction Stop
            Write-StepUpdate "Staged: $localExe ($([math]::Round((Get-Item $localExe).Length/1MB,1)) MB)" OK
            Write-StepUpdate "Identity of the file that will be deployed as SYSTEM:" INFO
            Show-BeamInstallerIdentity -Path $localExe | Out-Null
        } catch { Write-StepUpdate "Staging failed: $_" FAIL; return (New-BeamPublishResult $false) }
    }

    # Verify the staged file header matches its type. EXE is a PE (MZ). MSI is an OLE compound
    # document (D0 CF 11 E0). A download URL that returns HTML fails both, catching a bad link.
    try {
        $fs = [IO.File]::OpenRead($localExe)
        $hdr = New-Object byte[] 8
        [void]$fs.Read($hdr, 0, 8)
        $fs.Close()
    } catch { $hdr = @() }
    if ($IsMsi) {
        $okHdr = ($hdr.Length -ge 4 -and $hdr[0] -eq 0xD0 -and $hdr[1] -eq 0xCF -and $hdr[2] -eq 0x11 -and $hdr[3] -eq 0xE0)
        if (-not $okHdr) {
            Write-StepUpdate "Staged file is not a valid .msi (bad OLE header, $([math]::Round((Get-Item $localExe).Length/1KB,0)) KB)." FAIL
            Write-StepUpdate "If this came from a URL, it likely returned a page. Use a direct-download link." FAIL
            return (New-BeamPublishResult $false)
        }
    } else {
        $mz = ""
        if ($hdr.Length -ge 2) { $mz = [Text.Encoding]::ASCII.GetString($hdr, 0, 2) }
        if ($mz -ne "MZ") {
            Write-StepUpdate "Staged file is not a valid .exe (header '$mz', $([math]::Round((Get-Item $localExe).Length/1KB,0)) KB)." FAIL
            Write-StepUpdate "If this came from a URL, it likely returned a page. Use a direct-download link." FAIL
            return (New-BeamPublishResult $false)
        }
    }

    # The package source must hold exactly the one installer BEAM staged and nothing else.
    if (-not (Test-BeamStagingManifest -Dir $srcDir -SetupFile $SetupFile)) {
        Write-StepUpdate "The staging folder does not contain exactly ${SetupFile}: refusing to package it." FAIL
        return (New-BeamPublishResult $false)
    }

    Write-StepUpdate "Packaging with IntuneWinAppUtil..." INFO
    Invoke-BeamPackager -Tool $script:BeamToolPath -SourceDir $srcDir -SetupFile $SetupFile -OutDir $outDir
    $intunewinFiles = @(Get-ChildItem -LiteralPath $outDir -Filter "*.intunewin")
    if ($intunewinFiles.Count -ne 1) { Write-StepUpdate "Packaging failed: expected one .intunewin output, found $($intunewinFiles.Count)" FAIL; return (New-BeamPublishResult $false) }
    $intunewinFile = $intunewinFiles[0]
    Write-StepUpdate "Packaged: $($intunewinFile.Name) ($([math]::Round($intunewinFile.Length/1MB,1)) MB)" OK

    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip       = [System.IO.Compression.ZipFile]::OpenRead($intunewinFile.FullName)
    $metaEntry = $zip.Entries | Where-Object { $_.Name -eq "Detection.xml" } | Select-Object -First 1
    $ms        = New-Object System.IO.MemoryStream
    $metaEntry.Open().CopyTo($ms)
    $metaXml   = [xml][System.Text.Encoding]::UTF8.GetString($ms.ToArray())
    $contEntry = $zip.Entries | Where-Object { $_.FullName -like "IntuneWinPackage/Contents/*" } | Select-Object -First 1
    $encPath   = Join-Path $outDir $contEntry.Name
    [System.IO.Compression.ZipFileExtensions]::ExtractToFile($contEntry, $encPath, $true)
    $zip.Dispose()

    $unencSize = [long]$metaXml.ApplicationInfo.UnencryptedContentSize
    $encSize   = (Get-Item $encPath).Length
    $enc       = $metaXml.ApplicationInfo.EncryptionInfo
    Write-StepUpdate "Unencrypted: $([math]::Round($unencSize/1MB,1)) MB  Encrypted: $([math]::Round($encSize/1MB,1)) MB" INFO

    # Per-run marker detection. Each BEAM execution has a unique run ID ($runId, generated once per
    # run so every targeted machine in this run shares it). The install runs the package, then stamps
    # the run ID into a marker file. Detection passes ONLY when the marker holds THIS run's ID. Result:
    # within a run each machine installs exactly once (marker written, no loop); on a later run the new
    # run ID won't match the old marker, so it installs again. This is app-agnostic by design, so no
    # per-app product code or file path is needed for any installer.
    $markerDir  = 'C:\ProgramData\BEAM'
    $appDisplayName = "$AppName (BEAM)"
    # One marker file per app, so a second app on the same device never overwrites the first one's run ID.
    $markerPath = "$markerDir\$(Get-BeamAppKey $appDisplayName).marker"
    $detectScript = @"
`$m = '$markerPath'
if ((Test-Path `$m) -and ((Get-Content `$m -Raw -ErrorAction SilentlyContinue).Trim() -eq '$runId')) { Write-Output 'Detected'; exit 0 }
exit 1
"@
    $detectB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($detectScript))
    $detectRules = @(
        @{
            "@odata.type"         = "#microsoft.graph.win32LobAppPowerShellScriptRule"
            ruleType              = "detection"
            enforceSignatureCheck = $false
            runAs32Bit            = $false
            scriptContent         = $detectB64
            operationType         = "notConfigured"
            operator              = "notConfigured"
            comparisonValue       = $null
        }
    )

    # The install command is one powershell.exe -EncodedCommand (see New-BeamInstallScript).
    $installScript = New-BeamInstallScript -SetupFile $SetupFile -InstallArgs $InstallArgs -IsMsi $IsMsi -MarkerDir $markerDir -MarkerPath $markerPath -RunId $runId
    $stampB64      = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($installScript))
    $installCmd    = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand $stampB64"

    # Uninstall command line is required by Intune. MSI uninstalls by file; EXE is a no-op so a manual
    # uninstall never re-runs the installer (detection is marker-based and won't trigger it anyway).
    if ($IsMsi) { $uninstallCmd = "msiexec /x `"$SetupFile`" /qn" } else { $uninstallCmd = "cmd /c exit 0" }

    $app = $null
    $odataName = $appDisplayName -replace "'", "''"
    try {
        $existingApp = Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=displayName eq '$odataName'" $h
        $app = $existingApp.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.win32LobApp' } | Select-Object -First 1
    } catch {
        # Carrying on would create a second app with the same name, so stop instead.
        Write-StepUpdate "Could not look up an existing app named '$appDisplayName': $_" FAIL
        return (New-BeamPublishResult $false)
    }

    if ($app) {
        Write-StepUpdate "Reusing existing app: $($app.displayName) ($($app.id))" OK
        Write-StepUpdate "Adding a new content version to the existing app..." INFO
        # The live install command and detection rule are NOT touched here. They change together with
        # the committed content version in one patch, only after the new content is uploaded and
        # committed. Until then devices keep a coherent command, detection and package.
    }
    if (-not $app) {
    Write-StepUpdate "Creating Win32 app in Intune..." INFO
    $appBody = @{
        "@odata.type"        = "#microsoft.graph.win32LobApp"
        displayName          = $appDisplayName
        description          = "Deployed by BEAM v$_ver"
        publisher            = $Publisher
        fileName             = $SetupFile
        setupFilePath        = $SetupFile
        installCommandLine   = $installCmd
        uninstallCommandLine = $uninstallCmd
        applicableArchitectures = "x64,x86"
        minimumSupportedOperatingSystem = @{
            "@odata.type" = "#microsoft.graph.windowsMinimumOperatingSystem"
            v10_1607 = $true
        }
        installExperience = @{
            "@odata.type"         = "#microsoft.graph.win32LobAppInstallExperience"
            runAsAccount          = "system"
            deviceRestartBehavior = "suppress"
        }
        returnCodes = @(
            @{ returnCode = 0;    type = "success"    }
            @{ returnCode = 3010; type = "softReboot" }
            @{ returnCode = 1618; type = "retry"      }
            @{ returnCode = 1603; type = "failed"     }
        )
        rules = $detectRules
    } | ConvertTo-Json -Depth 10
    try {
        # UTF-8 bytes: the display name / description / publisher can contain non-ASCII, and PS 5.1
        # sends string bodies as Latin-1, which mangles it.
        $app = Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" -Method POST -Headers $h -Body ([System.Text.Encoding]::UTF8.GetBytes($appBody)) -ContentType "application/json; charset=utf-8" -ErrorAction Stop
        Write-StepUpdate "App created: $($app.displayName) ($($app.id))" OK
    } catch {
        if ("$_" -match "not applicable to target tenant|not licensed") { Write-StepUpdate "This tenant isn't provisioned for Intune: Win32 deployment unavailable." FAIL }
        else { Write-StepUpdate "App creation failed: $_" FAIL }
        return (New-BeamPublishResult $false)
    }
    }

    $fileBody = @{
        "@odata.type" = "#microsoft.graph.mobileAppContentFile"
        name          = "$SetupFile.intunewin"
        size          = $unencSize
        sizeEncrypted = $encSize
        manifest      = $null
        isDependency  = $false
    } | ConvertTo-Json
    try {
        $cv = Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions" $h POST "{}"
        Write-StepUpdate "Content version: $($cv.id)" INFO
        $fileEntry = Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions/$($cv.id)/files" $h POST $fileBody
    } catch { Write-StepUpdate "Creating the content version failed: $_" FAIL; return (New-BeamPublishResult $false) }
    $baseFileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions/$($cv.id)/files/$($fileEntry.id)"

    Write-StepUpdate "Waiting for Azure Storage URI..." INFO
    try { $fileReady = Wait-GraphState $baseFileUri $h "uploadState" @("azureStorageUriRequestSuccess") @("azureStorageUriRequestFailed") }
    catch { Write-StepUpdate "Waiting for the Azure Storage URI failed: $_" FAIL; return (New-BeamPublishResult $false) }
    Write-StepUpdate "Azure Storage URI ready" OK

    Write-StepUpdate "Uploading encrypted content to Azure Storage..." INFO
    try {
        Send-AzureStorageBlock -SasUri $fileReady.azureStorageUri -FilePath $encPath
        Write-StepUpdate "Upload complete" OK
    } catch { Write-StepUpdate "Upload failed: $_" FAIL; return (New-BeamPublishResult $false) }

    Write-StepUpdate "Committing file with encryption info..." INFO
    $commitBody = @{
        fileEncryptionInfo = @{
            encryptionKey        = $enc.EncryptionKey
            macKey               = $enc.MacKey
            initializationVector = $enc.InitializationVector
            mac                  = $enc.Mac
            profileIdentifier    = $enc.ProfileIdentifier
            fileDigest           = $enc.FileDigest
            fileDigestAlgorithm  = $enc.FileDigestAlgorithm
        }
    } | ConvertTo-Json -Depth 5
    try { Invoke-BeamGraph "$baseFileUri/commit" $h POST $commitBody | Out-Null }
    catch { Write-StepUpdate "Committing the file failed: $_" FAIL; return (New-BeamPublishResult $false) }
    try { Wait-GraphState $baseFileUri $h "uploadState" @("commitFileSuccess") @("commitFileFailed") | Out-Null }
    catch { Write-StepUpdate "Waiting for the file commit failed: $_" FAIL; return (New-BeamPublishResult $false) }
    Write-StepUpdate "File committed" OK

    # One patch publishes the new install command, uninstall command, detection rule and committed
    # content version together, and only now that the content is uploaded and committed. It is then
    # read back, so a patch that was accepted but did not take effect is not reported as published.
    $patchBody = @{
        "@odata.type"           = "#microsoft.graph.win32LobApp"
        fileName                = $SetupFile
        setupFilePath           = $SetupFile
        installCommandLine      = $installCmd
        uninstallCommandLine    = $uninstallCmd
        rules                   = $detectRules
        committedContentVersion = $cv.id
    } | ConvertTo-Json -Depth 10
    try {
        Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)" $h PATCH $patchBody | Out-Null
        $live = Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)" $h
    } catch { Write-StepUpdate "Committing the app version failed: $_" FAIL; return (New-BeamPublishResult $false) }
    if ("$($live.committedContentVersion)" -ne "$($cv.id)" -or $live.installCommandLine -ne $installCmd) {
        Write-StepUpdate "Committing the app version failed: Intune did not apply the new content version and install command." FAIL
        return (New-BeamPublishResult $false)
    }
    Write-StepUpdate "App version committed" OK

    # Every run gets its own group, named with the run ID. A group reused by name would carry the
    # members of earlier runs, and the assignment would then install the app on machines this run
    # never chose. The mailNickname is ASCII, free of characters Graph rejects, and at most 64 long.
    $groupName = "BEAM-$safeName-$RunId"
    $runTag = ($RunId -replace '[^A-Za-z0-9-]', '')
    if ($runTag.Length -gt 40) { $runTag = $runTag.Substring($runTag.Length - 40) }
    $nickBase = ("BEAM-$safeName" -replace '[^\x21-\x7E]', '-') -replace '[@()\\\[\]";:<>,\s]', ''
    $room = 64 - $runTag.Length - 1
    if ($nickBase.Length -gt $room) { $nickBase = $nickBase.Substring(0, $room) }
    $mailNick = "$nickBase-$runTag"
    Write-StepUpdate "Creating security group: $groupName" INFO
    $groupBody = @{ displayName = $groupName; mailEnabled = $false; mailNickname = $mailNick; securityEnabled = $true } | ConvertTo-Json
    $group = $null
    try { $group = Invoke-BeamGraph "https://graph.microsoft.com/v1.0/groups" $h POST $groupBody }
    catch { Write-StepUpdate "Creating the group failed: $_" FAIL; return (New-BeamPublishResult $false) }
    if (-not $group -or -not $group.id) { Write-StepUpdate "Creating the group failed: Graph returned no group ID" FAIL; return (New-BeamPublishResult $false) }
    Write-StepUpdate "Group created: $($group.id)" OK
    Write-Host ""

    # Resolve each target to exactly one device and add it to the group. A target is only queued
    # after the group read-back lists its device and the assignment succeeds.
    $tres = [ordered]@{}
    foreach ($target in $Targets) { $tres[$target] = 'FAIL' }
    $pending = New-Object System.Collections.ArrayList
    foreach ($target in $Targets) {
        Write-StepUpdate "--- $target ---" INFO
        $esc = $target -replace "'", "''"

        $md = $null; $mdErr = $false; $mdAmbiguous = $false
        try {
            $mdList = @((Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=deviceName eq '$esc'" $h).value | Where-Object { $_.deviceName -ieq $target })
            if ($mdList.Count -gt 1) { $mdAmbiguous = $true }
            elseif ($mdList.Count -eq 1) { $md = $mdList[0] }
        } catch { $mdErr = $true }
        if ($mdAmbiguous) {
            Write-Detail "Intune device ID" "ambiguous: $($mdList.Count) managed devices are named $target" FAIL
            continue
        }
        if ($mdErr) {
            Write-Detail "Intune device ID" "lookup failed: sync skipped" WARN
        } elseif ($md) {
            Write-Detail "Intune device ID" $md.id
            Write-Detail "OS" "$($md.operatingSystem) $($md.osVersion)"
        } else {
            Write-Detail "Intune device ID" "not found: sync skipped" WARN
        }

        # Prefer the Entra device that the managed device points at; fall back to an exact name that
        # matches exactly one Entra device.
        $entra = $null; $entraErr = $null; $entraCount = 0
        try {
            if ($md -and $md.azureADDeviceId) {
                $byId = @((Invoke-BeamGraph "https://graph.microsoft.com/v1.0/devices?`$filter=deviceId eq '$($md.azureADDeviceId)'" $h).value)
                $entraCount = $byId.Count
                if ($byId.Count -eq 1) { $entra = $byId[0] }
            }
            if (-not $entra -and $entraCount -eq 0) {
                $byName = @((Invoke-BeamGraph "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$esc'" $h).value | Where-Object { $_.displayName -ieq $target })
                $entraCount = $byName.Count
                if ($byName.Count -eq 1) { $entra = $byName[0] }
            }
        } catch { $entraErr = $_ }
        if ($entraErr) { Write-Detail "Entra device ID" "lookup failed (needs Device.Read.All consent?): $entraErr" FAIL; continue }
        if ($entraCount -gt 1) { Write-Detail "Entra device ID" "ambiguous: $entraCount Entra devices match $target" FAIL; continue }
        if (-not $entra) { Write-Detail "Entra device ID" "not found: $target cannot be added to the group" FAIL; continue }
        Write-Detail "Entra device ID" $entra.id

        $memberBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($entra.id)" } | ConvertTo-Json
        $added = $true
        try { Invoke-BeamGraph "https://graph.microsoft.com/v1.0/groups/$($group.id)/members/`$ref" $h POST $memberBody | Out-Null }
        catch {
            if ("$_" -notmatch 'already exist') { Write-Detail "Group" "add failed: $_" FAIL; $added = $false }
        }
        if (-not $added) { continue }
        [void]$pending.Add([pscustomobject]@{ Target = $target; EntraId = "$($entra.id)"; Md = $md })
        Write-Host ""
    }

    # Read the group back and keep only the targets that are really in it.
    $verified = New-Object System.Collections.ArrayList
    if ($pending.Count -gt 0) {
        $memberIds = New-Object System.Collections.ArrayList
        $readOk = $true
        try {
            $next = "https://graph.microsoft.com/v1.0/groups/$($group.id)/members?`$select=id&`$top=999"
            $pages = 0
            while ($next -and $pages -lt 100) {
                $page = Invoke-BeamGraph $next $h
                foreach ($m in @($page.value)) { [void]$memberIds.Add("$($m.id)") }
                $next = $page.'@odata.nextLink'
                $pages++
            }
        } catch { $readOk = $false; Write-StepUpdate "Reading the group back failed: $_" FAIL }
        foreach ($pnd in $pending) {
            if ($readOk -and $memberIds -contains $pnd.EntraId) {
                Write-Detail $pnd.Target "verified in $groupName" OK
                [void]$verified.Add($pnd)
            } else {
                Write-Detail $pnd.Target "not listed as a member of $groupName after the add" FAIL
            }
        }
    }
    if ($verified.Count -eq 0) {
        Write-StepUpdate "No target could be verified in the group. The app was published but not assigned to anything." FAIL
        return (New-BeamPublishResult $false $tres)
    }

    # Nudge each verified device to check in. A failed sync only delays the install: the device
    # still picks the app up on its normal check-in.
    foreach ($v in $verified) {
        if ($v.Md) {
            try {
                Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($v.Md.id)/syncDevice" $h POST | Out-Null
                Write-Detail "Sync $($v.Target)" "triggered" OK
            } catch { Write-Detail "Sync $($v.Target)" "trigger failed: $_" WARN }
        }
    }
    Write-Host ""

    Write-StepUpdate "Assigning app to group $groupName..." INFO
    $assignBody = @{
        mobileAppAssignments = @(
            @{
                "@odata.type" = "#microsoft.graph.mobileAppAssignment"
                intent        = "required"
                target        = @{
                    "@odata.type" = "#microsoft.graph.groupAssignmentTarget"
                    groupId       = $group.id
                }
                settings      = @{
                    "@odata.type"                = "#microsoft.graph.win32LobAppAssignmentSettings"
                    notifications                = "showAll"
                    installTimeSettings          = $null
                    restartSettings              = $null
                    deliveryOptimizationPriority = "notConfigured"
                }
            }
        )
    } | ConvertTo-Json -Depth 10
    try { Invoke-BeamGraph "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/assign" $h POST $assignBody | Out-Null }
    catch { Write-StepUpdate "Assigning the app failed: $_" FAIL; return (New-BeamPublishResult $false $tres) }
    foreach ($v in $verified) { $tres[$v.Target] = 'Intune queued' }
    Write-StepUpdate "App assigned. Devices will install on next Intune check-in (15-60 min)." OK
    Write-Host ""
    Write-StepUpdate "App ID: $($app.id)" INFO
    Write-StepUpdate "Group: $groupName ($($group.id))" INFO
    Write-Host ""
    return (New-BeamPublishResult $true $tres)
}
function Test-BeamStagingManifest {
    # True only when the folder holds exactly one file and it is the expected installer.
    param([string]$Dir, [string]$SetupFile)
    $items = @(Get-ChildItem -LiteralPath $Dir -Force -ErrorAction SilentlyContinue)
    return ($items.Count -eq 1 -and -not $items[0].PSIsContainer -and $items[0].Name -ieq $SetupFile)
}

function Invoke-IntuneInstall {
    # Publishes the app and returns an object: Published (bool) and Targets (target -> "Intune queued"
    # or "FAIL"). The run's own staging folder is removed when the
    # run ends, success or not. A folder that already existed is never ours, so it is left alone.
    # Any unexpected error is a failed run with a message, never an unhandled throw.
    param(
        [string[]]$Targets,
        [string]$Url,
        [string]$LocalPath,
        [string]$SetupFile,
        [bool]$IsMsi,
        [string]$InstallArgs,
        [string]$AppName,
        [string]$Publisher,
        [string]$TenantId,
        [string]$RunId
    )
    $runRoot = Join-Path (Join-Path (Get-BeamWorkRoot) "IntuneWork") ("run-" + ($RunId -replace '[^A-Za-z0-9._-]', '_'))
    $existed = Test-Path -LiteralPath $runRoot
    $res = New-BeamPublishResult $false
    try {
        $out = @(Invoke-BeamPublish -Targets $Targets -Url $Url -LocalPath $LocalPath -SetupFile $SetupFile -IsMsi $IsMsi -InstallArgs $InstallArgs `
            -AppName $AppName -Publisher $Publisher -TenantId $TenantId -RunId $RunId -RunRoot $runRoot)
        $got = @($out | Where-Object { $_ -is [pscustomobject] -and $_.PSObject.Properties['Published'] }) | Select-Object -Last 1
        if ($got) { $res = $got }
    } catch {
        Write-StepUpdate "Unexpected error: $($_.Exception.Message)" FAIL
        $res = New-BeamPublishResult $false
    } finally {
        if (-not $existed -and (Test-Path -LiteralPath $runRoot)) {
            try { Remove-Item -LiteralPath $runRoot -Recurse -Force -ErrorAction Stop } catch { }
        }
    }
    return $res
}

function Invoke-BeamMain {
    try { Clear-Host } catch { }
    Show-BeamBanner

    $in = $null; $inErr = $null
    try { $in = Read-BeamInputs } catch { $inErr = $_ }
    if ($inErr) {
        if ($inErr.Exception.Message -eq 'BEAM_NOCONSOLE') {
            Exit-BeamRun 1 "there is no console to read the deployment inputs from. Nothing was changed."
        }
        Exit-BeamRun 1 "reading the inputs failed: $($inErr.Exception.Message)"
    }
    if ($in.Abort) { Exit-BeamRun $in.Code $in.Text }

    $script:BeamToolPath = Initialize-BeamTool
    if (-not $script:BeamToolPath) { Exit-BeamRun 1 "IntuneWinAppUtil is unavailable. Nothing was deployed." }

    $targets = @($in.Targets)
    $results = [ordered]@{}
    $targets | ForEach-Object { $results[$_] = "Intune pending" }

    # One run ID per BEAM execution, shared by every target this run. The install stamps it into a
    # marker and detection checks for it, so each run installs once per machine; re-running BEAM
    # generates a new ID and installs again.
    $runId = (Get-Date -Format "yyyyMMddHHmmss") + "-" + ([guid]::NewGuid().ToString("N").Substring(0,8))

    Write-HLine
    Write-Host ""
    Write-StepUpdate "Deploying '$($in.AppName)' to $($targets.Count) target(s) via Intune..." INFO
    Write-StepUpdate "Run ID: $runId" INFO
    Write-Host ""

    $pub = Invoke-IntuneInstall -Targets $targets -Url $in.InstallerUrl -LocalPath $in.LocalInstaller -SetupFile $in.SetupFile -IsMsi $in.IsMsi -InstallArgs $in.InstallArgs -AppName $in.AppName -Publisher $in.Publisher -TenantId $in.TenantId -RunId $runId
    # Each target is queued only if the publish verified it. Anything not reported is a failure.
    foreach ($tg in $targets) {
        $results[$tg] = "FAIL"
        if ($pub.Published -and $pub.Targets -and $pub.Targets.Contains($tg) -and $pub.Targets[$tg] -eq "Intune queued") { $results[$tg] = "Intune queued" }
    }
    $nFail = @($results.Values | Where-Object { $_ -eq "FAIL" }).Count
    if ($nFail -gt 0) {
        Write-Host ""
        Write-StepUpdate "Intune deployment did not complete for $nFail of $($targets.Count) target(s) (see above)." FAIL
    }

    Write-HLine
    Write-Host ""
    Write-StepUpdate "Deployment Summary" INFO
    Write-Host ""

    $results.GetEnumerator() | Where-Object { $_.Value -eq "Intune queued" } | ForEach-Object { Write-StepUpdate "$($_.Key)  (Intune: pending check-in)" WARN }
    $results.GetEnumerator() | Where-Object { $_.Value -eq "FAIL"          } | ForEach-Object { Write-StepUpdate "$($_.Key)" FAIL }
    Write-Host ""

    $queued = @($results.Values | Where-Object { $_ -eq "Intune queued" }).Count
    $failed = @($results.Values | Where-Object { $_ -eq "FAIL" }).Count
    Write-Host "  Intune queued : " -NoNewline -ForegroundColor $InfoCol; Write-Host $queued -ForegroundColor $AccentCol -NoNewline; Write-Host "    Failed : " -NoNewline -ForegroundColor $InfoCol; Write-Host "$failed    of $($targets.Count) total" -ForegroundColor $ArtCol
    Write-StepUpdate "Devices install on next Intune check-in. Verify in Intune > the app > Device install status." INFO
    Write-Host ""

    if ($failed -eq 0 -and $queued -eq $targets.Count) {
        Show-BeamFooter "  DEPLOYMENT COMPLETE"
        Exit-BeamRun 0 "'$($in.AppName)' is queued in Intune for $queued of $($targets.Count) target(s). Devices install on their next check-in."
    }
    if ($queued -gt 0) {
        Show-BeamFooter "  DEPLOYMENT FINISHED WITH PROBLEMS"
        Exit-BeamRun 0 "'$($in.AppName)' is queued for $queued of $($targets.Count) target(s); $failed failed."
    }
    Show-BeamFooter "  DEPLOYMENT FAILED"
    Exit-BeamRun 1 "'$($in.AppName)' was not deployed to any of the $($targets.Count) target(s)."
}

if (-not $global:KillerScriptsTestMode) { Invoke-BeamMain }
