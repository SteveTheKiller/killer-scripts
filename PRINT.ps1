<#
.SYNOPSIS
    PRINTER RESPONSE & INTERFACE NETWORK TOOL (PRINT) v1.0
    Developed by SteveTheKiller | Updated: 2026-03-14
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
      0  - Completed successfully
      3  - Insufficient privileges
#>
param(
    [switch]$Silent
)

$EXIT_SUCCESS = 0
$EXIT_DENIED  = 3

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit $EXIT_DENIED
}

#region [0] - HELPERS & THEME
# ============================================================================
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

$script:Width   = 85
$script:Version = "v1.0"

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
        $k = [Console]::ReadKey($true)
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
    Clear-Host
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
function Get-LocalSubnet {
    $adapter = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction SilentlyContinue |
        Where-Object { $_.PrefixOrigin -ne 'WellKnown' -and
                       $_.IPAddress -notlike '169.*'   -and
                       $_.IPAddress -ne '127.0.0.1' } |
        Sort-Object PrefixLength -Descending |
        Select-Object -First 1
    if (-not $adapter) { return $null }
    $parts = $adapter.IPAddress -split '\.'
    return "$($parts[0]).$($parts[1]).$($parts[2])"
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
    param([string]$Subnet)
    $e       = [char]27
    $spinner = @('|','/','-','\')
    $tick    = 0
    $ports   = @(9100, 631)
    $timeout = 600

    Write-Host -NoNewline "      > Scanning $Subnet.0/24..." -ForegroundColor $DimCol

    $found = [System.Collections.Concurrent.ConcurrentBag[string]]::new()
    $pool  = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, 60)
    $pool.Open()

    $jobs = 1..254 | ForEach-Object {
        $ip = "$Subnet.$_"
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
        Write-Host -NoNewline "${e}[2K`r      > Scanning $Subnet.0/24... $($spinner[$tick % 4])  ($($found.Count) found)" -ForegroundColor $DimCol
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
        $key = [Console]::ReadKey($true)
        if ($key.Key -eq [ConsoleKey]::Escape) { return }
        $addChoice = $key.KeyChar.ToString()
        Write-Host $addChoice

        $targetIP    = $null
        $targetModel = $null

        switch ($addChoice) {
            '1' {
                $subnet = Get-LocalSubnet
                if (-not $subnet) {
                    Write-Host "`n  [!] Could not determine local subnet." -ForegroundColor Red
                    Start-Sleep -Seconds 2; continue addmenu
                }
                Write-WizardBanner "Network Scan"
                $ips = @(Invoke-PrinterScan -Subnet $subnet)
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
                } catch {
                    Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
                }
                Write-Host ""; Write-Host "  Press any key..." -ForegroundColor $DimCol
                [Console]::ReadKey($true) | Out-Null; return
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
                    $drvKey    = [Console]::ReadKey($true)
                    if ($drvKey.Key -eq [ConsoleKey]::Escape) { break wizloop }
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
                            $drvSel = $allDrivers[[int]$drvRaw - 1]
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
                    $protoKey = [Console]::ReadKey($true)
                    if ($protoKey.Key -eq [ConsoleKey]::Escape) { $wizStep = 1; continue wizloop }
                    $proto   = $protoKey.KeyChar.ToString()
                    $useIPP  = ($proto -eq '2')
                    $portNum = if ($useIPP) { 631 } else { 9100 }
                    Write-Host $(if ($useIPP) { "2 (IPP)" } else { "1 (RAW)" })

                    # --- Install ---
                    Write-Host ""
                    $portName   = "IP_$targetIP"
                    $portExists = Get-PrinterPort -Name $portName -ErrorAction SilentlyContinue
                    try {
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
                    } catch {
                        Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
                    }

                    Write-Host ""; Write-Host "  Press any key to continue..." -ForegroundColor $DimCol
                    [Console]::ReadKey($true) | Out-Null
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
    $confirmKey = [Console]::ReadKey($true)
    if ($confirmKey.Key -eq [ConsoleKey]::Escape) { return }
    $confirm = $confirmKey.KeyChar.ToString().ToUpper()
    Write-Host $confirm

    if ($confirm -eq 'Y') {
        try {
            Remove-Printer -Name $target.Name -ErrorAction Stop
            Write-Host "  '$($target.Name)' removed." -ForegroundColor $OkCol
        } catch {
            Write-Host "  [!] $($_.Exception.Message)" -ForegroundColor Red
        }
        Start-Sleep -Seconds 1
    }
}
#endregion

#region [4] - MAIN LOOP
if ($Silent) {
    Write-Banner
    Get-Printer | Sort-Object Name | Format-Table Name, DriverName, PortName -AutoSize
    exit $EXIT_SUCCESS
}

$printers  = @()
$running   = $true
$e         = [char]27
while ($running) {
    Write-Banner
    Show-PrinterList -PrinterRef ([ref]$printers)
    Write-Host ""
    $script:MenuRow = [Console]::CursorTop
    Write-HLine
    Write-Host "  " -NoNewline
    Write-Host "[A]" -NoNewline -ForegroundColor Gray;  Write-Host " Add   "  -NoNewline -ForegroundColor $MainCol
    Write-Host "[D]" -NoNewline -ForegroundColor Gray;  Write-Host " Delete   " -NoNewline -ForegroundColor $MainCol
    Write-Host "[Q]" -NoNewline -ForegroundColor Gray;  Write-Host " Quit"    -ForegroundColor $MainCol
    Write-Host "  Choice: " -NoNewline -ForegroundColor $LineCol
    $key    = [Console]::ReadKey($true)
    $choice = $key.KeyChar.ToString().ToUpper()
    Write-Host $choice

    if ($key.Key -eq [ConsoleKey]::Escape) { $running = $false; continue }

    switch ($choice) {
        'A' { Invoke-AddPrinterWizard }
        'D' { Invoke-DeletePrinter -Printers $printers }
        'Q' { $running = $false }
    }
}
#endregion

#region [5] - FOOTER
# ============================================================================
[Console]::SetCursorPosition(0, $script:MenuRow)
Write-Host "${e}[0J" -NoNewline  # clear from cursor to end of screen
$_sfx    = "█"
$_ftrW   = $script:ArtW + 1
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  PRINT SEQUENCE COMPLETE"
$_fpad   = " " * ($_ffillW - $_footer.Length - $script:Ver.Length)
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $($script:Art1)" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
Write-Host "$_footer$_fpad$($script:Ver)" -ForegroundColor $MainCol  -NoNewline; Write-Host " $($script:Art2)" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $($script:Art3)" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
#endregion

exit $EXIT_SUCCESS
