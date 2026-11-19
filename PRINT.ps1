<#
.SYNOPSIS
    PRINTER RESPONSE & INTERFACE NETWORK TOOL (PRINT) v2.0
    Developed by SteveTheKiller | Updated: 2026-11-19
.DESCRIPTION
    Printer management utility that inventories, adds, and removes devices. Features a
    multi-threaded network scanner that discovers printers on the local subnet via port
    scan, identifies specific models through HTTP/Web server scraping, and automates
    driver matching from the local store. Supports manual IP entry, UNC shared printer
    paths, and includes an intelligent fallback to the Microsoft IPP Class Driver for
    universal compatibility.
.NOTES
    Parameters:
      -Silent     Non-interactive mode, lists printers only.
    Exit Codes:
      0  - Success
      1  - Failure
      2  - Not elevated, or a precondition failed (for example no console and no -Silent)
      3  - Partial success (some changes succeeded, some failed)
#>
param(
    [switch]$Silent
)
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

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-PrintElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Test-PrintInteractive {
    try { [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected } catch { $false }
}
function Read-PrintKey {
    # Returns $null when there is no console, so callers treat it as ESC and never spin.
    try { [Console]::ReadKey($true) } catch { $null }
}
function Test-PrintIppSupported {
    # Add-Printer -IppURL exists on newer Windows 10/11 and Server builds only.
    try { (Get-Command Add-Printer -ErrorAction Stop).Parameters.ContainsKey('IppURL') } catch { $false }
}
function Get-PrintCursorTop { try { [Console]::CursorTop } catch { -1 } }
function Set-PrintCursor {
    param([int]$Left, [int]$Top)
    try { [Console]::SetCursorPosition($Left, $Top) } catch { }
}
#region [0] - HELPERS & THEME
# ============================================================================
$script:Width   = 85
$script:Version = "v2.0"

$LineCol   = "Green"
$MainCol   = "Yellow"
$ArtCol    = "DarkYellow"
$BorderCol = "DarkGreen"
$DimCol    = "DarkGray"
$WarnCol   = "DarkYellow"
$OkCol     = "Green"
$AccentCol = "Yellow"

function Write-HLine {
    param([string]$Style = "dashed", [int]$Width = $script:Width)
    $line = if ($Style -eq "dashed") { ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width) } else { "-" * $Width }
    $colors = @([ConsoleColor]$LineCol, [ConsoleColor]$ArtCol, [ConsoleColor]$DimCol)
    $useConsole = $true
    try { $saved = [Console]::ForegroundColor } catch { $useConsole = $false }
    $i = 0
    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') { $fg = [ConsoleColor]$DimCol }
        else { $fg = $colors[$i % $colors.Count]; $i++ }
        if ($useConsole) { [Console]::ForegroundColor = $fg; [Console]::Write($char) }
        else { Write-Host $char -NoNewline -ForegroundColor $fg }
    }
    if ($useConsole) { [Console]::ForegroundColor = $saved; [Console]::WriteLine() }
    else { Write-Host "" }
}

function Read-Input {
    # Like Read-Host but returns $null on ESC (caller treats $null as "back")
    $buf = ""
    while ($true) {
        $k = Read-PrintKey
        if ($null -eq $k)                        { return $null }
        if ($k.Key -eq [ConsoleKey]::Escape)     { Write-Host ""; return $null }
        if ($k.Key -eq [ConsoleKey]::Enter)      { Write-Host ""; return $buf  }
        if ($k.Key -eq [ConsoleKey]::Backspace) {
            if ($buf.Length -gt 0) {
                $buf = $buf.Substring(0, $buf.Length - 1)
                Write-Host "`b `b" -NoNewline
            }
        } else {
            $buf += $k.KeyChar
            Write-Host $k.KeyChar -NoNewline
        }
    }
}

function Write-Banner {
    try { Clear-Host } catch { }
    $_pfx  = "█  "
    $_art1 = "╔═╗ ╦═╗ ╦ ╔╗╔ ╔╦╗ "
    $_art2 = "╠═╝ ╠╦╝ ║ ║║║  ║  "
    $_art3 = "╩   ╩╚═ ╩ ╝╚╝  ╩  "
    $script:ArtW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1 = $_art1.PadRight($script:ArtW)
    $_art2 = $_art2.PadRight($script:ArtW)
    $_art3 = $_art3.PadRight($script:ArtW)
    $script:Art1 = $_art1; $script:Art2 = $_art2; $script:Art3 = $_art3
    $_fillW = $script:Width - $_pfx.Length - $script:ArtW
    $_title = "PRINTER RESPONSE & INTERFACE NETWORK TOOL"
    $script:Ver = "| $($script:Version)"
    $_pad   = " " * [Math]::Max(0, ($_fillW - $_title.Length - $script:Ver.Length))
    Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_pad$($script:Ver)" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host "[>] Device Name  : " -NoNewline -ForegroundColor $LineCol
    Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $AccentCol
    Write-HLine
}

function Show-PrinterList {
    param([ref]$PrinterRef)
    $printers = Get-Printer -ErrorAction SilentlyContinue | Sort-Object Name
    $PrinterRef.Value = $printers
    if (-not $printers -or $printers.Count -eq 0) {
        Write-Host "  No printers installed." -ForegroundColor $DimCol
        return
    }
    Write-Host ("{0,-4}{1,-44}{2}" -f "ID", "Printer Name", "Port") -ForegroundColor Cyan
    Write-HLine
    $i = 1
    foreach ($p in $printers) {
        $nameDisplay = if ($p.Name.Length     -gt 42) { $p.Name.Substring(0, 41)     + [char]0x2026 } else { $p.Name }
        $portDisplay = if ($p.PortName.Length -gt 34) { $p.PortName.Substring(0, 33) + [char]0x2026 } else { $p.PortName }
        $color = if ($p.Default) { $OkCol } else { $MainCol }
        Write-Host ("{0,-4}" -f $i) -ForegroundColor $DimCol -NoNewline
        Write-Host ("{0,-44}" -f $nameDisplay) -ForegroundColor $color -NoNewline
        Write-Host $portDisplay -ForegroundColor $DimCol
        $i++
    }
}
#endregion

#region [1] - NETWORK SCAN & DETECTION
# ============================================================================
function ConvertTo-PrintUInt32 {
    param([string]$IP)
    $o = $IP.Split('.')
    [long]([long]$o[0] * 16777216 + [long]$o[1] * 65536 + [long]$o[2] * 256 + [long]$o[3])
}
function ConvertFrom-PrintUInt32 {
    param([long]$N)
    "{0}.{1}.{2}.{3}" -f (($N -shr 24) -band 255), (($N -shr 16) -band 255), (($N -shr 8) -band 255), ($N -band 255)
}

function Get-PrintScanRange {
    # Host addresses of the network that IPAddress/PrefixLength belongs to, excluding the network and
    # broadcast addresses. Networks larger than a /22 are bounded to the /24 around the host.
    # Returns $null when the input is invalid or the network has no scannable hosts (/31, /32).
    param([string]$IPAddress, [int]$PrefixLength)
    $parsed = $null
    if (-not [System.Net.IPAddress]::TryParse($IPAddress, [ref]$parsed)) { return $null }
    if ($parsed.AddressFamily -ne [System.Net.Sockets.AddressFamily]::InterNetwork) { return $null }
    if ($PrefixLength -lt 0 -or $PrefixLength -gt 30) { return $null }
    $note = $null
    $prefix = $PrefixLength
    if ($prefix -lt 22) {
        $note = "/$PrefixLength is too large to scan, limiting to the /24 around $IPAddress"
        $prefix = 24
    }
    $ip      = ConvertTo-PrintUInt32 $IPAddress
    $mask    = (4294967295 -shl (32 - $prefix)) -band 4294967295
    $network = $ip -band $mask
    $bcast   = $network + (4294967295 -bxor $mask)
    $addrs   = New-Object System.Collections.Generic.List[string]
    for ($n = $network + 1; $n -lt $bcast; $n++) { $addrs.Add((ConvertFrom-PrintUInt32 $n)) }
    [pscustomobject]@{
        Label     = "$(ConvertFrom-PrintUInt32 $network)/$prefix"
        Addresses = $addrs.ToArray()
        Note      = $note
    }
}

function Get-PrintScanTarget {
    # Scan range for the interface that owns the best IPv4 default route (lowest route + interface metric).
    # Returns $null when there is no default route or no usable address on it.
    $routes = @(Get-NetRoute -AddressFamily IPv4 -DestinationPrefix '0.0.0.0/0' -ErrorAction SilentlyContinue)
    if ($routes.Count -eq 0) { return $null }
    $ifMetric = @{}
    foreach ($i in @(Get-NetIPInterface -AddressFamily IPv4 -ErrorAction SilentlyContinue)) { $ifMetric[[int]$i.InterfaceIndex] = [int]$i.InterfaceMetric }
    $ordered = $routes | Sort-Object { [int]$_.RouteMetric + [int]$ifMetric[[int]$_.InterfaceIndex] }
    $all = @(Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue)
    foreach ($r in $ordered) {
        $addr = $all | Where-Object { $_.InterfaceIndex -eq $r.InterfaceIndex -and
                                      $_.IPAddress -notlike '169.254.*' -and
                                      $_.IPAddress -notlike '127.*' } | Select-Object -First 1
        if (-not $addr) { continue }
        $range = Get-PrintScanRange -IPAddress $addr.IPAddress -PrefixLength ([int]$addr.PrefixLength)
        if ($range) { return $range }
    }
    return $null
}
function Get-PrinterModel {
    param([string]$IP)
    try {
        $resp = Invoke-WebRequest -Uri "http://$IP/" -TimeoutSec 2 -UseBasicParsing -ErrorAction Stop
        if ($resp.Content -match '<title[^>]*>\s*([^<]{4,60}?)\s*</title>') {
            $t = $Matches[1].Trim() -replace '\s+', ' '
            if ($t -notmatch '(?i)index|home|welcome|web server|untitled') { return $t }
        }
        if ($resp.Content -match 'MDL:([^;\"<]{3,40})') { return $Matches[1].Trim() }
        if ($resp.Content -match '"model"\s*:\s*"([^"]{3,40})"') { return $Matches[1].Trim() }
    } catch {}
    return $null
}

function Invoke-PrinterScan {
    param([string[]]$Addresses, [string]$Label)
    $e       = [char]27
    $spinner = @('|','/','-','\')
    $tick    = 0
    $ports   = @(9100, 631)
    $timeout = 600

    Write-Host -NoNewline "      > Scanning $Label..." -ForegroundColor $DimCol

    $found = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
    $pool  = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 60)
    $pool.Open()

    $jobs = $Addresses | ForEach-Object {
        $ip = $_
        $ps = [System.Management.Automation.PowerShell]::Create()
        $ps.RunspacePool = $pool
        $ps.AddScript({
            param($ip, $ports, $timeout, $found)
            foreach ($port in $ports) {
                try {
                    $tcp = [System.Net.Sockets.TcpClient]::new()
                    if ($tcp.ConnectAsync($ip, $port).Wait($timeout) -and $tcp.Connected) {
                        $found.Add($ip); $tcp.Close(); return
                    }
                    $tcp.Close()
                } catch {}
            }
        }).AddArgument($ip).AddArgument($ports).AddArgument($timeout).AddArgument($found) | Out-Null
        @{ PS = $ps; Handle = $ps.BeginInvoke() }
    }

    while ($jobs | Where-Object { -not $_.Handle.IsCompleted }) {
        Write-Host -NoNewline "${e}[2K`r      > Scanning $Label... $($spinner[$tick % 4])  ($($found.Count) found)" -ForegroundColor $DimCol
        $tick++; Start-Sleep -Milliseconds 150
    }
    foreach ($j in $jobs) { $j.PS.Dispose() }
    $pool.Close(); $pool.Dispose()

    Write-Host "${e}[2K`r      > Scan complete - $($found.Count) device(s) responding on printer ports." -ForegroundColor $DimCol
    return $found | Sort-Object { [System.Version]$_ }
}

function Find-PrinterDriver {
    param([string]$ModelName)
    $drivers = Get-PrinterDriver -ErrorAction SilentlyContinue
    # Exact substring match
    $match = $drivers | Where-Object { $_.Name -like "*$ModelName*" } | Select-Object -First 1
    if ($match) { return $match.Name }
    # Multi-keyword match on significant words
    $words = $ModelName -split '[\s\-]+' | Where-Object { $_.Length -gt 3 } | Select-Object -First 3
    foreach ($word in $words) {
        $match = $drivers | Where-Object { $_.Name -like "*$word*" } | Select-Object -First 1
        if ($match) { return $match.Name }
    }
    # IPP universal fallback (always present on Win10/11)
    $ipp = $drivers | Where-Object { $_.Name -like "*IPP*" } | Select-Object -First 1
    if ($ipp) { return $ipp.Name }
    return "Microsoft IPP Class Driver"
}
#endregion

function Resolve-PrintDriverChoice {
    # Maps a typed 1-based driver ID to a driver. Empty, 0, non-numeric or out-of-range input returns $null
    # so the caller keeps the suggested driver instead of indexing -1 (the last driver).
    param([object[]]$Drivers, [string]$Raw)
    $n = 0
    if (-not [int]::TryParse(("$Raw").Trim(), [ref]$n)) { return $null }
    if ($n -lt 1 -or $n -gt @($Drivers).Count) { return $null }
    return @($Drivers)[$n - 1]
}

#region [2] - ADD WIZARD
# ============================================================================
function Invoke-AddPrinterWizard {
    $e = [char]27

    # Draws banner + persistent "Add Printer" header with optional breadcrumb
    function Write-WizardBanner {
        param([string]$Crumb = "")
        Write-Banner
        Write-Host "[>] Add Printer" -NoNewline -ForegroundColor Cyan
        if ($Crumb) {
            $available = $script:Width - 15 - 5  # "[>] Add Printer" + "  >  "
            $display   = if ($Crumb.Length -gt $available) { [char]0x2026 + $Crumb.Substring($Crumb.Length - $available + 1) } else { $Crumb }
            Write-Host "  >  " -NoNewline -ForegroundColor $DimCol
            Write-Host $display -ForegroundColor $MainCol
        } else {
            Write-Host ""
        }
        Write-HLine
    }

    :addmenu while ($true) {
        Write-WizardBanner
        Write-Host "  " -NoNewline; Write-Host "[1]" -NoNewline -ForegroundColor Gray; Write-Host " Scan network for printers" -ForegroundColor $MainCol
        Write-Host "  " -NoNewline; Write-Host "[2]" -NoNewline -ForegroundColor Gray; Write-Host " Enter IP address manually" -ForegroundColor $MainCol
        Write-Host "  " -NoNewline; Write-Host "[3]" -NoNewline -ForegroundColor Gray; Write-Host " Connect to shared printer  (\\server\name)" -ForegroundColor $MainCol
        Write-Host "  " -NoNewline; Write-Host "[ESC]" -NoNewline -ForegroundColor Gray; Write-Host " Back" -ForegroundColor $DimCol
        Write-Host ""
        Write-Host "  Choice: " -NoNewline -ForegroundColor $LineCol
        $key = Read-PrintKey
        if ($null -eq $key -or $key.Key -eq [ConsoleKey]::Escape) { return }
        $addChoice = $key.KeyChar.ToString()
        Write-Host $addChoice

        $targetIP    = $null
        $targetModel = $null

        switch ($addChoice) {
            '1' {
                $target = Get-PrintScanTarget
                if (-not $target) {
                    Write-Host "`n  [!] Could not determine the local network from the default route. Enter the IP manually." -ForegroundColor Red
                    Start-Sleep -Seconds 2; continue addmenu
                }
                Write-WizardBanner "Network Scan"
                if ($target.Note) { Write-Host "      > $($target.Note)" -ForegroundColor $WarnCol }
                $ips = @(Invoke-PrinterScan -Addresses $target.Addresses -Label $target.Label)
                if ($ips.Count -eq 0) {
                    Write-Host "  [!] No printers found. Try entering the IP manually." -ForegroundColor $WarnCol
                    Start-Sleep -Seconds 2; continue addmenu
                }
                $candidates = @()
                $n = 1
                foreach ($ip in $ips) {
                    Write-Host -NoNewline "${e}[2K`r      > Identifying $ip..." -ForegroundColor $DimCol
                    $model = Get-PrinterModel -IP $ip
                    $candidates += [PSCustomObject]@{ Index = $n; IP = $ip; Model = if ($model) { $model } else { "Unknown Printer" } }
                    $n++
                }
                Write-WizardBanner "Network Scan"
                Write-Host ("{0,-4} {1,-18} {2}" -f "ID", "IP Address", "Detected Model") -ForegroundColor Cyan
                Write-HLine
                foreach ($c in $candidates) {
                    Write-Host ("{0,-4}" -f $c.Index) -ForegroundColor $DimCol -NoNewline
                    Write-Host ("{0,-18}" -f $c.IP)   -ForegroundColor $AccentCol -NoNewline
                    Write-Host $c.Model                -ForegroundColor $MainCol
                }
                Write-Host ""
                Write-Host "  Select ID: " -NoNewline -ForegroundColor $LineCol
                $raw = Read-Input
                if ($null -eq $raw) { continue addmenu }
                $sel = $candidates | Where-Object { $_.Index -eq [int]$raw }
                if (-not $sel) { Write-Host "  Invalid selection." -ForegroundColor Red; Start-Sleep -Seconds 1; continue addmenu }
                $targetIP    = $sel.IP
                $targetModel = $sel.Model
            }

            '2' {
                Write-WizardBanner "Manual IP"
                Write-Host "  IP address: " -NoNewline -ForegroundColor $LineCol
                $targetIP = Read-Input
                if ($null -eq $targetIP) { continue addmenu }
                $targetIP = $targetIP.Trim()
                if (-not $targetIP) { continue addmenu }
                Write-Host -NoNewline "      > Checking $targetIP..." -ForegroundColor $DimCol
                $model = Get-PrinterModel -IP $targetIP
                $targetModel = if ($model) { $model } else { "Printer @ $targetIP" }
                Write-Host "${e}[2K`r      > Detected: $targetModel" -ForegroundColor $DimCol
            }

            '3' {
                Write-WizardBanner "Shared Printer"
                Write-Host "  UNC path (e.g. \\\\server\\printer): " -NoNewline -ForegroundColor $LineCol
                $unc = Read-Input
                if ($null -eq $unc) { continue addmenu }
                $unc = $unc.Trim()
                if (-not $unc) { continue addmenu }
                Write-Host "  Friendly name: " -NoNewline -ForegroundColor $LineCol
                $fname = Read-Input
                if ($null -eq $fname) { continue addmenu }
                $fname = $fname.Trim()
                if (-not $fname) { $fname = ($unc -split '\\')[-1] }
                Write-Host ""
                Write-Host "      > Connecting to $unc..." -ForegroundColor $DimCol
                try {
                    Add-Printer -ConnectionName $unc -ErrorAction Stop
                    Write-Host "      > Connected. Renaming to '$fname'..." -ForegroundColor $DimCol
                    $added = Get-Printer | Where-Object { $_.PortName -like "*$(($unc -split '\\')[-1])*" } | Select-Object -First 1
                    if ($added -and $added.Name -ne $fname) {
                        Rename-Printer -InputObject $added -NewName $fname -ErrorAction SilentlyContinue
                    }
                    Write-Host "      > '$fname' added successfully." -ForegroundColor $OkCol
                    $script:PrintSuccesses++
                } catch {
                    $script:PrintFailures++
                    Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""; Write-Host "  Press any key..." -ForegroundColor $DimCol
                Read-PrintKey | Out-Null; return
            }

            default { continue addmenu }
        }

        if (-not $targetIP) { continue addmenu }

        # Crumb helper: truncate a segment for display in the breadcrumb trail
        function Get-Seg { param([string]$s, [int]$max)
            if ($s.Length -gt $max) { $s.Substring(0, $max - 1) + [char]0x2026 } else { $s }
        }

        $driverName = Find-PrinterDriver -ModelName $targetModel
        $fname      = $null
        $wizStep    = 0

        :wizloop while ($true) {
            # Build the breadcrumb for the current step
            $segIP    = Get-Seg $targetIP    15
            $segModel = Get-Seg $targetModel 20
            $segDrv   = Get-Seg $driverName  20
            $segName  = if ($fname) { Get-Seg $fname 18 } else { $null }
            $crumb0   = "$segIP - $segModel"
            $crumb1   = "$crumb0  >  $segDrv"
            $crumb2   = if ($segName) { "$crumb1  >  $segName" } else { $crumb1 }

            switch ($wizStep) {
                0 {
                    # --- Driver selection ---
                    $driverSource = if ($driverName -eq "Microsoft IPP Class Driver") { "(IPP universal fallback)" } else { "(matched from driver store)" }
                    Write-WizardBanner $crumb0
                    Write-Host "      > Suggested: $driverName  $driverSource" -ForegroundColor $DimCol
                    Write-Host ""
                    Write-Host "  " -NoNewline; Write-Host "[A]" -NoNewline -ForegroundColor Gray; Write-Host " Accept suggested driver" -ForegroundColor $MainCol
                    Write-Host "  " -NoNewline; Write-Host "[C]" -NoNewline -ForegroundColor Gray; Write-Host " Choose from installed drivers" -ForegroundColor $MainCol
                    Write-Host "  " -NoNewline; Write-Host "[P]" -NoNewline -ForegroundColor Gray; Write-Host " Provide INF file path" -ForegroundColor $MainCol
                    Write-Host "  Choice [A]: " -NoNewline -ForegroundColor $LineCol
                    $drvKey    = Read-PrintKey
                    if ($null -eq $drvKey -or $drvKey.Key -eq [ConsoleKey]::Escape) { break wizloop }
                    $drvChoice = $drvKey.KeyChar.ToString().ToUpper()
                    Write-Host $(if ($drvChoice -in 'C','P') { $drvChoice } else { "A" })

                    if ($drvChoice -eq 'C') {
                        $allDrivers = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Sort-Object Name)
                        if ($allDrivers.Count -eq 0) {
                            Write-Host "  [!] No drivers found in store." -ForegroundColor $WarnCol
                        } else {
                            Write-Host ""
                            $d = 1
                            foreach ($drv in $allDrivers) {
                                Write-Host ("{0,-4}" -f $d) -ForegroundColor $DimCol -NoNewline
                                Write-Host $drv.Name -ForegroundColor $MainCol
                                $d++
                            }
                            Write-Host ""
                            Write-Host "  Select driver ID: " -NoNewline -ForegroundColor $LineCol
                            $drvRaw = Read-Input
                            if ($null -eq $drvRaw) { continue wizloop }
                            $drvSel = Resolve-PrintDriverChoice -Drivers $allDrivers -Raw $drvRaw
                            if ($drvSel) { $driverName = $drvSel.Name } else { Write-Host "  Invalid - using suggested." -ForegroundColor $WarnCol }
                        }
                    } elseif ($drvChoice -eq 'P') {
                        Write-Host "  INF file path: " -NoNewline -ForegroundColor $LineCol
                        $infPath = Read-Input
                        if ($null -eq $infPath) { continue wizloop }
                        $infPath = $infPath.Trim()
                        if ($infPath -and (Test-Path $infPath)) {
                            Write-Host "      > Importing driver from $infPath..." -ForegroundColor $DimCol
                            try {
                                $beforeNames = @(Get-PrinterDriver -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name)
                                pnputil /add-driver $infPath /install 2>&1 | Out-Null
                                $newDriver = Get-PrinterDriver -ErrorAction SilentlyContinue |
                                    Where-Object { $_.Name -notin $beforeNames } |
                                    Select-Object -First 1
                                if ($newDriver) {
                                    $driverName = $newDriver.Name
                                    Write-Host "      > Imported: $driverName" -ForegroundColor $OkCol
                                } else {
                                    Write-Host "      [!] Could not confirm import - using suggested." -ForegroundColor $WarnCol
                                }
                            } catch {
                                Write-Host "      [!] Import failed: $($_.Exception.Message)" -ForegroundColor Red
                            }
                        } else {
                            Write-Host "  [!] File not found - using suggested." -ForegroundColor $WarnCol
                        }
                    }
                    $wizStep = 1
                }

                1 {
                    # --- Friendly name ---
                    Write-WizardBanner $crumb1
                    Write-Host ""
                    Write-Host "  Friendly name " -NoNewline -ForegroundColor $MainCol
                    Write-Host "(Enter to use: $targetModel): " -NoNewline -ForegroundColor $DimCol
                    $fname = Read-Input
                    if ($null -eq $fname) { $wizStep = 0; continue wizloop }
                    if (-not $fname.Trim()) { $fname = $targetModel } else { $fname = $fname.Trim() }
                    $wizStep = 2
                }

                2 {
                    # --- Protocol ---
                    Write-WizardBanner $crumb2
                    Write-Host "      > Driver: $driverName" -ForegroundColor $DimCol
                    Write-Host "      > Name:   $fname" -ForegroundColor $DimCol
                    Write-Host ""
                    Write-Host "  Port protocol:" -ForegroundColor $MainCol
                    Write-Host "  " -NoNewline; Write-Host "[1]" -NoNewline -ForegroundColor Gray; Write-Host " RAW / JetDirect  (port 9100)  - recommended" -ForegroundColor $MainCol
                    Write-Host "  " -NoNewline; Write-Host "[2]" -NoNewline -ForegroundColor Gray; Write-Host " IPP              (port 631)   - fallback" -ForegroundColor $MainCol
                    Write-Host "  Choice [1]: " -NoNewline -ForegroundColor $LineCol
                    $protoKey = Read-PrintKey
                    if ($null -eq $protoKey) { return }
                    if ($protoKey.Key -eq [ConsoleKey]::Escape) { $wizStep = 1; continue wizloop }
                    $proto   = $protoKey.KeyChar.ToString()
                    $useIPP  = ($proto -eq '2')
                    $portNum = if ($useIPP) { 631 } else { 9100 }
                    Write-Host $(if ($useIPP) { "2 (IPP)" } else { "1 (RAW)" })

                    # --- Install ---
                    Write-Host ""
                    $portName = $null
                    try {
                        if ($useIPP) {
                            # IPP is HTTP, so a raw TCP/IP port on 631 cannot print. Use the native IPP queue.
                            if (-not (Test-PrintIppSupported)) {
                                throw "This Windows build cannot add IPP printers natively (Add-Printer has no -IppURL). Choose RAW (port 9100) or update Windows."
                            }
                            $ippUrl = "http://${targetIP}:631/ipp/print"
                            Write-Host "      > Adding IPP printer '$fname' ($ippUrl, Microsoft IPP Class Driver)..." -ForegroundColor $DimCol
                            Add-Printer -Name $fname -IppURL $ippUrl -ErrorAction Stop
                            Write-Host "      > '$fname' installed successfully." -ForegroundColor $OkCol
                            $script:PrintSuccesses++
                            Write-Host ""; Write-Host "  Press any key to continue..." -ForegroundColor $DimCol
                            Read-PrintKey | Out-Null
                            return
                        }
                        # RAW: reuse a same-IP port only when it is a TCP/IP port to this host on 9100.
                        $portName   = "IP_$targetIP"
                        $portExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
                        if ($portExists -and -not ($portExists.PortNumber -eq $portNum -and $portExists.PrinterHostAddress -eq $targetIP)) {
                            Write-Host "      > Port $portName exists with different settings, using a separate RAW port." -ForegroundColor $DimCol
                            $portName   = "IP_${targetIP}_$portNum"
                            $portExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
                            if ($portExists -and -not ($portExists.PortNumber -eq $portNum -and $portExists.PrinterHostAddress -eq $targetIP)) {
                                throw "Port $portName already exists with different settings. Remove it in Print Server Properties and try again."
                            }
                        }
                        if (-not $portExists) {
                            Write-Host "      > Creating port $portName (port $portNum)..." -ForegroundColor $DimCol
                            Add-PrinterPort -Name $portName -PrinterHostAddress $targetIP -PortNumber $portNum -ErrorAction Stop
                        } else {
                            Write-Host "      > Reusing existing port $portName." -ForegroundColor $DimCol
                        }
                        $driverInstalled = Get-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
                        if (-not $driverInstalled) {
                            Write-Host "      > Installing driver '$driverName'..." -ForegroundColor $DimCol
                            try {
                                Add-PrinterDriver -Name $driverName -ErrorAction Stop
                            } catch {
                                Write-Host "      [!] Driver install failed - falling back to Microsoft IPP Class Driver." -ForegroundColor $WarnCol
                                $driverName = "Microsoft IPP Class Driver"
                                Add-PrinterDriver -Name $driverName -ErrorAction SilentlyContinue
                            }
                        }
                        Write-Host "      > Adding printer '$fname'..." -ForegroundColor $DimCol
                        Add-Printer -Name $fname -DriverName $driverName -PortName $portName -ErrorAction Stop
                        Write-Host "      > '$fname' installed successfully." -ForegroundColor $OkCol
                        $script:PrintSuccesses++
                    } catch {
                        $script:PrintFailures++
                        Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }

                    Write-Host ""; Write-Host "  Press any key to continue..." -ForegroundColor $DimCol
                    Read-PrintKey | Out-Null
                    return
                }
            }
        }
        # ESC at driver step - fall through to next addmenu iteration
    }
}
#endregion

#region [3] - DELETE PRINTER
# ============================================================================
function Invoke-DeletePrinter {
    param([object[]]$Printers)
    Write-Banner
    Show-PrinterList -PrinterRef ([ref]$Printers)
    if (-not $Printers -or $Printers.Count -eq 0) { Start-Sleep -Seconds 1; return }
    Write-Host ""
    Write-Host "  Enter printer ID to remove: " -NoNewline -ForegroundColor $WarnCol
    $raw = Read-Input
    if ($null -eq $raw) { return }
    $idx = [int]$raw - 1
    if ($idx -lt 0 -or $idx -ge $Printers.Count) {
        Write-Host "  [!] Invalid ID." -ForegroundColor Red; Start-Sleep -Seconds 1; return
    }
    $target = $Printers[$idx]
    Write-Host ""
    Write-Host "  Remove '$($target.Name)'? " -NoNewline -ForegroundColor $WarnCol
    Write-Host "[Y/N]: " -NoNewline -ForegroundColor $DimCol
    $confirmKey = Read-PrintKey
    if ($null -eq $confirmKey -or $confirmKey.Key -eq [ConsoleKey]::Escape) { return }
    $confirm = $confirmKey.KeyChar.ToString().ToUpper()
    Write-Host $confirm

    if ($confirm -eq 'Y') {
        try {
            Remove-Printer -Name $target.Name -ErrorAction Stop
            Write-Host "  '$($target.Name)' removed." -ForegroundColor $OkCol
            $script:PrintSuccesses++
        } catch {
            $script:PrintFailures++
            Write-Host "  [!] $($_.Exception.Message)" -ForegroundColor Red
        }
        Start-Sleep -Seconds 1
    }
}
#endregion

#region [4] - MAIN
# ============================================================================
function Write-PrintFooter {
    $e = [char]27
    $top = Get-PrintCursorTop
    if ($top -ge 0 -and $script:MenuRow -ge 0) {
        Set-PrintCursor -Left 0 -Top $script:MenuRow
        Write-Host "${e}[0J" -NoNewline  # clear from cursor to end of screen
    }
    $_sfx    = "█"
    $_ftrW   = $script:ArtW + 1
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    $_footer = "  PRINT SEQUENCE COMPLETE"
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $script:Ver.Length))
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $($script:Art1)" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
    Write-Host "$_footer$_fpad$($script:Ver)" -ForegroundColor $MainCol  -NoNewline; Write-Host " $($script:Art2)" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $($script:Art3)" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
}

function Invoke-PrintMain {
    param([switch]$Silent)
    $script:PrintFailures  = 0
    $script:PrintSuccesses = 0
    $script:MenuRow        = -1
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

    if (-not (Test-PrintElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Exit-KillerScript 3
        return
    }

    if ($Silent) {
        Write-Banner
        try {
            $list = @(Get-Printer -ErrorAction Stop | Sort-Object Name)
        } catch {
            Exit-KillerScript 1
            return
        }
        $list | Format-Table Name, DriverName, PortName -AutoSize | Out-String -Width 200 | Write-Host
        Exit-KillerScript 0
        return
    }

    if (-not (Test-PrintInteractive)) {
        Exit-KillerScript 1
        return
    }

    $printers  = @()
    $running   = $true
    while ($running) {
        Write-Banner
        Show-PrinterList -PrinterRef ([ref]$printers)
        Write-Host ""
        $script:MenuRow = Get-PrintCursorTop
        Write-HLine
        Write-Host "  " -NoNewline
        Write-Host "[A]" -NoNewline -ForegroundColor Gray;  Write-Host " Add   "  -NoNewline -ForegroundColor $MainCol
        Write-Host "[D]" -NoNewline -ForegroundColor Gray;  Write-Host " Delete   " -NoNewline -ForegroundColor $MainCol
        Write-Host "[Q]" -NoNewline -ForegroundColor Gray;  Write-Host " Quit"    -ForegroundColor $MainCol
        Write-Host "  Choice: " -NoNewline -ForegroundColor $LineCol
        $key = Read-PrintKey
        if ($null -eq $key) {
            Exit-KillerScript 1
            return
        }
        $choice = $key.KeyChar.ToString().ToUpper()
        Write-Host $choice

        if ($key.Key -eq [ConsoleKey]::Escape) { $running = $false; continue }

        switch ($choice) {
            'A' { Invoke-AddPrinterWizard }
            'D' { Invoke-DeletePrinter -Printers $printers }
            'Q' { $running = $false }
        }
    }

    Write-PrintFooter

    if ($script:PrintFailures -gt 0 -and $script:PrintSuccesses -eq 0) {
        Exit-KillerScript 1
        return
    }
    Exit-KillerScript 0
}
#endregion

if (-not $global:KillerScriptsTestMode) { Invoke-PrintMain -Silent:$Silent }