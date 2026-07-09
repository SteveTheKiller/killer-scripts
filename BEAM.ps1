<#
.SYNOPSIS
    Bulk Endpoint App Mover (BEAM) v1.00
    Created by Steve the Killer | Updated: 2026-07-09
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

    Each run has a unique run ID. The install runs your package, stamps the ID into a
    marker, and detection matches that ID, so each run installs once per machine and does
    NOT loop. Re-running installs again. Detection is marker-based and does not inspect
    the app itself, so it works for any installer.
#>

#region UI Initialization
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding            = [System.Text.Encoding]::UTF8
$ProgressPreference        = 'SilentlyContinue'
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }
Clear-Host

$script:Width = 85
$_ver         = "1.00"
$_fver        = "| v$_ver"

# ESC handling requires a true interactive console. Detect it without touching the input buffer
# (accessing [Console]::KeyAvailable here breaks Read-Host in relayed consoles). In a
# redirected/relayed console ESC can't be read at all, so Ctrl+C is the cancel there.
$script:IsInteractive = [Environment]::UserInteractive -and -not [Console]::IsInputRedirected

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
    param([string]$Prompt)
    if (-not $script:IsInteractive) { return (Read-Host $Prompt) }

    $disp = "$Prompt  [F = load CSV/TXT]"
    Write-Host "${disp}: " -NoNewline
    $buf = ""
    while ($true) {
        $k = [Console]::ReadKey($true)
        if ($k.Key -eq [ConsoleKey]::Enter) { Write-Host ""; return $buf }
        if ($buf -eq "" -and ($k.KeyChar -eq 'f' -or $k.KeyChar -eq 'F')) {
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
        if ($k.KeyChar) { $buf += $k.KeyChar; Write-Host $k.KeyChar -NoNewline }
    }
}

# Header
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
#endregion


#region Input Collection
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
    $srcInput = (Read-Host).Trim().Trim('"')
    Write-Host ""

    if ($srcInput -eq "") {
        if (-not $script:IsInteractive) {
            Write-StepUpdate "The file picker needs an interactive console. Provide a URL instead: aborting" FAIL
            return
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
            $ans   = (Read-Host "  Couldn't read a file type from the URL. Is it an MSI? (y/N)").Trim()
            $isMsi = ($ans -match '^(y|yes)$')
            $setupFile = if ($isMsi) { "installer.msi" } else { "installer.exe" }
        }
        Write-StepUpdate "Source URL accepted ($setupFile)" OK
    }
}
Write-Host ""

# 2. App name shown in Intune. Blank derives it from the installer file name.
$defaultName = [IO.Path]::GetFileNameWithoutExtension($setupFile)
$appName     = (Read-Host "  App name  [blank = '$defaultName']").Trim()
if ($appName -eq "") { $appName = $defaultName }

# 3. Publisher (required by Intune, cosmetic). Blank defaults to BEAM.
$publisher = (Read-Host "  Publisher  [blank = 'BEAM']").Trim()
if ($publisher -eq "") { $publisher = "BEAM" }

# 4. Install switches. Optional. MSI defaults to /qn when blank; EXE runs bare when blank.
$phSwitch    = if ($isMsi) { "MSI: blank uses /qn" } else { "EXE: e.g. /S  /silent  /qn  -s" }
$installArgs = (Read-Host "  Install switches (optional)  [$phSwitch]").Trim()
if ($isMsi -and $installArgs -eq "") { $installArgs = "/qn" }
Write-Host ""

# 5. Targets. Intune matches by hostname, so IP entries are dropped.
$targetInput  = Read-TargetsField "  Targets (comma-separated hostnames, IPs can't match Intune)"
$targets      = $targetInput -split "," | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne "" }
$ipTargets    = @($targets | Where-Object { $_ -match '^\d{1,3}(\.\d{1,3}){3}$' })
if ($ipTargets.Count -gt 0) {
    Write-StepUpdate "Dropping IP target(s), Intune matches by hostname only: $($ipTargets -join ', ')" WARN
    $targets = @($targets | Where-Object { $_ -notmatch '^\d{1,3}(\.\d{1,3}){3}$' })
}
if ($targets.Count -eq 0) { Write-StepUpdate "No hostname targets to deploy: aborting" FAIL; return }
Write-Host ""

# 6. Global Admin for Intune. The domain resolves the tenant ID for the sign-in.
$adminUpn = Read-Host "  Global admin email for Intune (required)"
$tenantId = $null
if (-not $adminUpn) {
    Write-StepUpdate "A Global Admin / Intune Admin sign-in is required for this tool: aborting" FAIL
    return
}
$adminDomain = ($adminUpn -split "@")[-1]
if (-not $adminDomain -or $adminDomain -eq $adminUpn) {
    Write-StepUpdate "Could not parse a domain from '$adminUpn': aborting" FAIL
    return
}
Write-StepUpdate "Resolving tenant for domain: $adminDomain" INFO
try {
    $oidc     = Invoke-RestMethod "https://login.microsoftonline.com/$adminDomain/v2.0/.well-known/openid-configuration" -ErrorAction Stop
    $tenantId = ([regex]::Match($oidc.issuer, '(?i)[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}')).Value
} catch { }
if (-not $tenantId) { Write-StepUpdate "Tenant lookup failed for '$adminDomain': aborting" FAIL; return }
Write-StepUpdate "Tenant ID: $tenantId" OK

Write-Host ""
Write-HLine dashed
Write-Host ""
#endregion

#region IntuneWin Tool Setup
$workRoot          = Join-Path $env:LOCALAPPDATA "BEAM"
$intunewinToolPath = Join-Path $workRoot "IntuneWinAppUtil.exe"
$intuneWorkDir     = Join-Path $workRoot "IntuneWork"
if (-not (Test-Path $workRoot)) { New-Item -ItemType Directory -Path $workRoot -Force | Out-Null }
if (-not (Test-Path $intunewinToolPath)) {
    Write-StepUpdate "IntuneWinAppUtil not found, downloading..." INFO
    try {
        Invoke-WebRequest -Uri "https://github.com/microsoft/Microsoft-Win32-Content-Prep-Tool/raw/master/IntuneWinAppUtil.exe" -OutFile $intunewinToolPath -ErrorAction Stop
        Write-StepUpdate "IntuneWinAppUtil ready at $intunewinToolPath" OK
    } catch {
        Write-StepUpdate "IntuneWinAppUtil download failed: cannot continue" FAIL
        return
    }
} else {
    Write-StepUpdate "IntuneWinAppUtil found at $intunewinToolPath" OK
}
Write-Host ""
#endregion


#region Helper Functions
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

    # Loopback listener on a free port. "localhost" prefix needs no URL ACL / elevation.
    $listener = New-Object System.Net.HttpListener
    $port     = 0
    for ($p = 8400; $p -le 8420; $p++) {
        try { $listener.Prefixes.Clear(); $listener.Prefixes.Add("http://localhost:$p/"); $listener.Start(); $port = $p; break } catch { }
    }
    if ($port -eq 0) { Write-StepUpdate "Could not bind a loopback port (8400-8420) for auth" FAIL; return $null }
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
    try { Start-Process $authUrl } catch { }

    # Wait for the loopback redirect (auto-capture), with a timeout so it never hangs forever.
    Write-StepUpdate "Waiting for sign-in to complete (up to 180s, ESC to cancel)..." INFO
    $query    = $null
    $ctxTask  = $listener.GetContextAsync()
    $deadline = (Get-Date).AddSeconds(180)
    while (-not $ctxTask.IsCompleted -and (Get-Date) -lt $deadline) {
        if (Test-EscPressed) { try { $listener.Stop() } catch { }; Write-StepUpdate "Sign-in cancelled by operator" WARN; return $null }
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
    try { $listener.Stop() } catch { }

    # Last resort (e.g. browser on another machine): read the redirect URL from the clipboard.
    if (-not ($query -and $query -match 'code=')) {
        Write-Host ""
        Write-StepUpdate "No local redirect captured (browser likely on another machine)." WARN
        Write-StepUpdate "Copy the FULL address-bar URL from the browser, then press Enter here." INFO
        Read-Host "  (press Enter after copying the URL)" | Out-Null
        try { $query = (Get-Clipboard -Raw) } catch { try { $query = (Get-Clipboard) -join '' } catch { $query = $null } }
    }
    if (-not $query) { Write-StepUpdate "No authorization response obtained: aborting auth" FAIL; return $null }

    # Extract code (guard UnescapeDataString so a stray % can't throw). [?&] anchor skips session_state.
    $code = $null
    if ($query -match '[?&]code=([^&\s]+)') {
        try { $code = [Uri]::UnescapeDataString($matches[1]) } catch { $code = $matches[1] }
        if ($query -match '[?&]state=([^&\s]+)' -and $matches[1] -ne $state) {
            Write-StepUpdate "OAuth state mismatch: aborting auth" FAIL; return $null
        }
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
    $deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSec)
    while ([DateTime]::UtcNow -lt $deadline) {
        $r = Invoke-RestMethod $Uri -Headers $Headers
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

function Invoke-IntuneInstall {
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
    $runId = $RunId

    Write-HLine dashed
    Write-StepUpdate "Intune Win32 deployment" INFO
    Write-StepUpdate "Targets: $($Targets -join ', ')" INFO
    Write-Host ""

    if (-not $intunewinToolPath) { Write-StepUpdate "IntuneWinAppUtil unavailable: aborted" FAIL; return $false }

    Write-StepUpdate "Authenticating to Graph (tenant: $TenantId)..." INFO
    $token = Get-GraphToken -TenantId $TenantId
    if (-not $token) { return $false }
    $h = @{ Authorization = "Bearer $token" }

    # Fast capability check: a tenant without Intune returns "not applicable to target tenant"
    # here, so bail before wasting the download/packaging. Only bail on that specific signal.
    try {
        Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$top=1" -Headers $h -ErrorAction Stop | Out-Null
    } catch {
        if ("$_" -match "not applicable to target tenant|not licensed|Tenant is not") {
            Write-StepUpdate "This tenant isn't provisioned/licensed for Intune: Win32 deployment unavailable." FAIL
            Write-Host ""
            return $false
        }
    }

    $safeName = ($AppName -replace '[\\/:*?"<>|]', '_')
    $srcDir = Join-Path $intuneWorkDir "src\$safeName"
    $outDir = Join-Path $intuneWorkDir "out\$safeName"
    @($intuneWorkDir, $srcDir, $outDir) | ForEach-Object { if (-not (Test-Path $_)) { New-Item -ItemType Directory -Path $_ -Force | Out-Null } }

    $localExe = Join-Path $srcDir $SetupFile

    # Stage a fresh copy of the installer on every run (no caching): a re-run against a tenant where
    # the app/group already exist must still replace the content, so we re-stage, re-package, and
    # upload a new content version each time. Remove any stale copy first.
    if (Test-Path $localExe) { Remove-Item $localExe -Force -ErrorAction SilentlyContinue }
    if ($Url) {
        Write-StepUpdate "Downloading $SetupFile..." INFO
        try {
            Invoke-WebRequest -Uri ([uri]$Url) -OutFile $localExe -ErrorAction Stop
        } catch { Write-StepUpdate "Download failed: $_" FAIL; return $false }

        # If the URL handed back a page instead of the binary, pull the real download link out of
        # it and fetch that instead. Runs silently, so a resolvable page just downloads normally.
        if (-not (Test-InstallerHeader -Path $localExe -IsMsi:$IsMsi)) {
            $direct = Resolve-DirectDownload -PagePath $localExe -PageUrl $Url
            if ($direct -and $direct -ne $Url) {
                try { Invoke-WebRequest -Uri ([uri]$direct) -OutFile $localExe -ErrorAction Stop } catch { }
            }
        }
        Write-StepUpdate "Downloaded: $localExe ($([math]::Round((Get-Item $localExe).Length/1MB,1)) MB)" OK
    } else {
        Write-StepUpdate "Staging local file: $LocalPath" INFO
        try {
            Copy-Item -LiteralPath $LocalPath -Destination $localExe -Force -ErrorAction Stop
            Write-StepUpdate "Staged: $localExe ($([math]::Round((Get-Item $localExe).Length/1MB,1)) MB)" OK
        } catch { Write-StepUpdate "Staging failed: $_" FAIL; return $false }
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
            return $false
        }
    } else {
        $mz = ""
        if ($hdr.Length -ge 2) { $mz = [Text.Encoding]::ASCII.GetString($hdr, 0, 2) }
        if ($mz -ne "MZ") {
            Write-StepUpdate "Staged file is not a valid .exe (header '$mz', $([math]::Round((Get-Item $localExe).Length/1KB,0)) KB)." FAIL
            Write-StepUpdate "If this came from a URL, it likely returned a page. Use a direct-download link." FAIL
            return $false
        }
    }

    Write-StepUpdate "Packaging with IntuneWinAppUtil..." INFO
    Get-ChildItem $outDir | Remove-Item -Force
    & $intunewinToolPath -c $srcDir -s $SetupFile -o $outDir -q 2>&1 | Out-Null
    $intunewinFile = Get-ChildItem $outDir -Filter "*.intunewin" | Select-Object -First 1
    if (-not $intunewinFile) { Write-StepUpdate "Packaging failed: no .intunewin output" FAIL; return $false }
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
    $markerPath = 'C:\ProgramData\BEAM\BEAM.marker'
    $detectScript = @"
`$m = '$markerPath'
if ((Test-Path `$m) -and ((Get-Content `$m -Raw -ErrorAction SilentlyContinue).Trim() -eq '$runId')) { Write-Output 'Detected'; exit 0 }
exit 1
"@
    $detectB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($detectScript))

    # Build the install command per file type. Everything is wrapped in a single powershell.exe
    # -EncodedCommand so it runs the package and then stamps the marker regardless of the installer's
    # exit code. Working directory at install time is the extracted content folder, so the file is
    # referenced relatively. Single quotes in the file name / switches are doubled for the PS literal.
    $sfEsc = $SetupFile   -replace "'", "''"
    $iaEsc = $InstallArgs -replace "'", "''"
    if ($IsMsi) {
        $al = "/i `"$sfEsc`""
        if ($iaEsc) { $al = "$al $iaEsc" }
        $runLine = "Start-Process -FilePath 'msiexec.exe' -ArgumentList '$al' -Wait"
    } else {
        if ($iaEsc) {
            $runLine = "Start-Process -FilePath '.\$sfEsc' -ArgumentList '$iaEsc' -Wait"
        } else {
            $runLine = "Start-Process -FilePath '.\$sfEsc' -Wait"
        }
    }
    $stampInner = "$runLine; New-Item -ItemType Directory -Path '$markerDir' -Force | Out-Null; Set-Content -Path '$markerPath' -Value '$runId' -Encoding ascii -Force"
    $stampB64   = [Convert]::ToBase64String([System.Text.Encoding]::Unicode.GetBytes($stampInner))
    $installCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -EncodedCommand $stampB64"

    # Uninstall command line is required by Intune. MSI uninstalls by file; EXE is a no-op so a manual
    # uninstall never re-runs the installer (detection is marker-based and won't trigger it anyway).
    if ($IsMsi) { $uninstallCmd = "msiexec /x `"$SetupFile`" /qn" } else { $uninstallCmd = "cmd /c exit 0" }

    $appDisplayName = "$AppName (BEAM)"
    $app = $null
    try {
        $existingApp = Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps?`$filter=displayName eq '$appDisplayName'" -Headers $h -ErrorAction Stop
        $app = $existingApp.value | Where-Object { $_.'@odata.type' -eq '#microsoft.graph.win32LobApp' } | Select-Object -First 1
    } catch { }

    if ($app) {
        Write-StepUpdate "Reusing existing app: $($app.displayName) ($($app.id))" OK
        Write-StepUpdate "Adding a new content version to the existing app..." INFO
        # Update BOTH the install command and the detection rule to this run's ID. Patching only the
        # command would leave the old run ID in detection, so the prior marker would still satisfy it
        # and the re-run would never fire. Both must carry the new run ID.
        try {
            $fixBody = @{
                "@odata.type"      = "#microsoft.graph.win32LobApp"
                installCommandLine = $installCmd
                rules = @(
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
            } | ConvertTo-Json -Depth 10
            Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)" -Method PATCH -Headers $h -Body ([System.Text.Encoding]::UTF8.GetBytes($fixBody)) -ContentType "application/json; charset=utf-8" | Out-Null
        } catch { }
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
        rules = @(
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
    } | ConvertTo-Json -Depth 10
    try {
        # UTF-8 bytes: the display name / description / publisher can contain non-ASCII, and PS 5.1
        # sends string bodies as Latin-1, which mangles it.
        $app = Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps" -Method POST -Headers $h -Body ([System.Text.Encoding]::UTF8.GetBytes($appBody)) -ContentType "application/json; charset=utf-8" -ErrorAction Stop
        Write-StepUpdate "App created: $($app.displayName) ($($app.id))" OK
    } catch {
        if ("$_" -match "not applicable to target tenant|not licensed") { Write-StepUpdate "This tenant isn't provisioned for Intune: Win32 deployment unavailable." FAIL }
        else { Write-StepUpdate "App creation failed: $_" FAIL }
        return $false
    }
    }

    $cv = Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions" -Method POST -Headers $h -Body "{}" -ContentType "application/json"
    Write-StepUpdate "Content version: $($cv.id)" INFO

    $fileBody = @{
        "@odata.type" = "#microsoft.graph.mobileAppContentFile"
        name          = "$SetupFile.intunewin"
        size          = $unencSize
        sizeEncrypted = $encSize
        manifest      = $null
        isDependency  = $false
    } | ConvertTo-Json
    $fileEntry   = Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions/$($cv.id)/files" -Method POST -Headers $h -Body $fileBody -ContentType "application/json"
    $baseFileUri = "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/microsoft.graph.win32LobApp/contentVersions/$($cv.id)/files/$($fileEntry.id)"

    Write-StepUpdate "Waiting for Azure Storage URI..." INFO
    $fileReady = Wait-GraphState $baseFileUri $h "uploadState" @("azureStorageUriRequestSuccess") @("azureStorageUriRequestFailed")
    Write-StepUpdate "Azure Storage URI ready" OK

    Write-StepUpdate "Uploading encrypted content to Azure Storage..." INFO
    try {
        Send-AzureStorageBlock -SasUri $fileReady.azureStorageUri -FilePath $encPath
        Write-StepUpdate "Upload complete" OK
    } catch { Write-StepUpdate "Upload failed: $_" FAIL; return $false }

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
    Invoke-RestMethod "$baseFileUri/commit" -Method POST -Headers $h -Body $commitBody -ContentType "application/json" | Out-Null
    Wait-GraphState $baseFileUri $h "uploadState" @("commitFileSuccess") @("commitFileFailed") | Out-Null
    Write-StepUpdate "File committed" OK

    $patchBody = @{ "@odata.type" = "#microsoft.graph.win32LobApp"; committedContentVersion = $cv.id } | ConvertTo-Json
    Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)" -Method PATCH -Headers $h -Body $patchBody -ContentType "application/json" | Out-Null
    Write-StepUpdate "App version committed" OK

    $groupName = "BEAM-$safeName"
    $group = $null
    try {
        $existingGroup = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups?`$filter=displayName eq '$groupName'" -Headers $h -ErrorAction Stop
        $group = $existingGroup.value | Select-Object -First 1
    } catch { }

    if ($group) {
        Write-StepUpdate "Reusing existing group: $groupName ($($group.id))" OK
    }
    if (-not $group) {
        Write-StepUpdate "Creating security group: $groupName" INFO
        # mailNickname must be ASCII with no special chars: map non-ASCII to hyphens and strip
        # anything else Graph rejects. displayName keeps the real name.
        $mailNick = ($groupName -replace '[^\x00-\x7F]', '-') -replace '[@()\\\[\]";:<>,\s]', ''
        if ($mailNick.Length -gt 64) { $mailNick = $mailNick.Substring(0, 64) }
        $groupBody = @{ displayName = $groupName; mailEnabled = $false; mailNickname = $mailNick; securityEnabled = $true } | ConvertTo-Json
        $group = Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups" -Method POST -Headers $h -Body ([System.Text.Encoding]::UTF8.GetBytes($groupBody)) -ContentType "application/json; charset=utf-8"
        Write-StepUpdate "Group created: $($group.id)" OK
    }
    Write-Host ""

    foreach ($target in $Targets) {
        Write-StepUpdate "--- $target ---" INFO
        $entraDevice = $null; $entraErr = $false
        try {
            $entraResp   = Invoke-RestMethod "https://graph.microsoft.com/v1.0/devices?`$filter=displayName eq '$target'" -Headers $h -ErrorAction Stop
            $entraDevice = $entraResp.value | Select-Object -First 1
        } catch { $entraErr = $true }
        if ($entraErr) {
            Write-Detail "Entra device ID" "lookup denied: needs Device.Read.All consent" WARN
        } elseif ($entraDevice) {
            Write-Detail "Entra device ID" $entraDevice.id
            $memberBody = @{ "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$($entraDevice.id)" } | ConvertTo-Json
            try {
                Invoke-RestMethod "https://graph.microsoft.com/v1.0/groups/$($group.id)/members/`$ref" -Method POST -Headers $h -Body $memberBody -ContentType "application/json" | Out-Null
                Write-Detail "Group" "added to $groupName" OK
            } catch {
                if ("$_" -match "already exist") { Write-Detail "Group" "already a member of $groupName" }
                else { Write-Detail "Group" "add failed: $_" WARN }
            }
        } else {
            Write-Detail "Entra device ID" "not found by displayName: skipping group add" WARN
        }
        $md = $null; $mdErr = $false
        try {
            $mdResp = Invoke-RestMethod "https://graph.microsoft.com/beta/deviceManagement/managedDevices?`$filter=deviceName eq '$target'" -Headers $h -ErrorAction Stop
            $md     = $mdResp.value | Select-Object -First 1
        } catch { $mdErr = $true }
        if ($mdErr) {
            Write-Detail "Intune device ID" "lookup failed: sync skipped" WARN
        } elseif ($md) {
            Write-Detail "Intune device ID" $md.id
            Write-Detail "OS" "$($md.operatingSystem) $($md.osVersion)"
            try {
                Invoke-RestMethod "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($md.id)/syncDevice" -Method POST -Headers $h | Out-Null
                Write-Detail "Sync" "triggered" OK
            } catch { Write-Detail "Sync" "trigger failed: $_" WARN }
        } else {
            Write-Detail "Intune device ID" "not found: sync skipped" WARN
        }
        Write-Host ""
    }

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
    Invoke-RestMethod "https://graph.microsoft.com/beta/deviceAppManagement/mobileApps/$($app.id)/assign" -Method POST -Headers $h -Body $assignBody -ContentType "application/json" | Out-Null
    Write-StepUpdate "App assigned. Devices will install on next Intune check-in (15-60 min)." OK
    Write-Host ""
    Write-StepUpdate "App ID: $($app.id)" INFO
    Write-StepUpdate "Group: $groupName ($($group.id))" INFO
    Write-Host ""
    return $true
}
#endregion


#region Main
$results = [ordered]@{}
$targets | ForEach-Object { $results[$_] = "Intune pending" }

# One run ID per BEAM execution, shared by every target this run. The install stamps it into a
# marker and detection checks for it, so each run installs once per machine; re-running BEAM
# generates a new ID and installs again.
$runId = (Get-Date -Format "yyyyMMddHHmmss") + "-" + ([guid]::NewGuid().ToString("N").Substring(0,8))

Write-HLine
Write-Host ""
Write-StepUpdate "Deploying '$appName' to $($targets.Count) target(s) via Intune..." INFO
Write-StepUpdate "Run ID: $runId" INFO
Write-Host ""

$intuneOk = Invoke-IntuneInstall -Targets $targets -Url $installerUrl -LocalPath $localInstaller -SetupFile $setupFile -IsMsi $isMsi -InstallArgs $installArgs -AppName $appName -Publisher $publisher -TenantId $tenantId -RunId $runId
if ($intuneOk) {
    $targets | ForEach-Object { $results[$_] = "Intune queued" }
} else {
    $targets | ForEach-Object { $results[$_] = "FAIL" }
    Write-Host ""
    Write-StepUpdate "Intune deployment did not complete (see error above): $($targets.Count) target(s) marked FAILED." FAIL
}
#endregion


#region Summary & Footer
Write-HLine
Write-Host ""
Write-StepUpdate "Deployment Summary" INFO
Write-Host ""

$results.GetEnumerator() | Where-Object { $_.Value -eq "Intune queued" } | ForEach-Object { Write-StepUpdate "$($_.Key)  (Intune: pending check-in)" WARN }
$results.GetEnumerator() | Where-Object { $_.Value -eq "FAIL"          } | ForEach-Object { Write-StepUpdate "$($_.Key)" FAIL }
Write-Host ""

$queued = ($results.Values | Where-Object { $_ -eq "Intune queued" }).Count
$failed = ($results.Values | Where-Object { $_ -eq "FAIL" }).Count
Write-Host "  Intune queued : " -NoNewline -ForegroundColor $InfoCol; Write-Host $queued -ForegroundColor $AccentCol -NoNewline; Write-Host "    Failed : " -NoNewline -ForegroundColor $InfoCol; Write-Host "$failed    of $($targets.Count) total" -ForegroundColor $ArtCol
Write-StepUpdate "Devices install on next Intune check-in. Verify in Intune > the app > Device install status." INFO
Write-Host ""

# Footer
$_sfx    = "█"
$_ffillW = $script:Width - $_artW - 1 - $_sfx.Length
$_footer = "  DEPLOYMENT COMPLETE"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))

Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
#endregion