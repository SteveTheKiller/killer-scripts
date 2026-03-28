<#
.SYNOPSIS
    Output Device Diagnostic (ODD) v1.9
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Inventories all physical, USB, and Bluetooth audio devices with
    health status and driver versions across input and output categories.
    Highlights the default device in green and shows Audio Service state,
    sample rate, bit depth, and exclusive mode from the Windows registry.
#>

#region [0] - HELPER FUNCTIONS
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Error "O.D.D. requires Elevated Privileges. Please run as Administrator."
    exit 1
}
function Get-StrippedWrapper([string]$Name) {
    if ($Name -match '\((.+)\)\s*$') { return $Matches[1].Trim() }
    return $Name.Trim()
}
function Get-Tokens([string]$Name) {
    ($Name -replace '[()®™+]', '') -split '\s+' | Where-Object { $_.Length -ge 3 }
}
function Test-IsDefault([string]$PnpName, [string[]]$Tokens) {
    if ($null -eq $Tokens -or $Tokens.Count -eq 0) { return $false }
    $matchCount = 0
    foreach ($t in $Tokens) { if ($PnpName -match [regex]::Escape($t)) { $matchCount++ } }
    $threshold = [Math]::Max(1, [Math]::Ceiling($Tokens.Count / 2))
    return $matchCount -ge $threshold
}

$script:Width    = 85
$script:Version  = "v1.9"
$LineCol   = "Magenta"
$MainCol   = "White"
$ArtCol    = "Yellow"
$DimCol    = "DarkGray"
$AccentCol = "Yellow"
function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [Math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "-" * $Width
    }
    $colors = @(
        [ConsoleColor]$LineCol,
        [ConsoleColor]$ArtCol,
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

function Get-DefaultDevice([string]$Branch) {
    $result = [PSCustomObject]@{ Guid = ""; Name = "" }
    $basePath = "SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Branch"
    try { $baseKey = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($basePath) } catch { return $result }
    if ($null -eq $baseKey) { return $result }

    $bestLevel = -1
    foreach ($subName in $baseKey.GetSubKeyNames()) {
        try {
            $devKey = $baseKey.OpenSubKey($subName)
            if ($null -eq $devKey) { continue }
            $state = [int]$devKey.GetValue("DeviceState")
            $level = $devKey.GetValue("Level:0")
            if ($null -ne $level) { $level = [int]$level } else { $level = -1 }
            $devKey.Close()
            if ($state -ne 1 -or $level -le $bestLevel) { continue }

            $propKey = $baseKey.OpenSubKey("$subName\Properties")
            if ($null -eq $propKey) { continue }
            $nameVal = $propKey.GetValue("{b3f8fa53-0004-438e-9003-51a46e139bfc},6")
            $propKey.Close()

            if ($nameVal -is [string] -and $nameVal.Length -gt 0) {
                $bestLevel    = $level
                $result.Guid  = $subName
                $result.Name  = $nameVal
            }
        } catch { continue }
    }
    $baseKey.Close()
    return $result
}

function Get-AudioQualityInfo([string]$Branch, [string]$Guid) {
    $result = [PSCustomObject]@{ SampleRate = "N/A"; BitDepth = "N/A"; ExclusiveMode = "N/A" }
    if (-not $Guid) { return $result }
    try {
        $propPath = "SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Branch\$Guid\Properties"
        $propKey  = [Microsoft.Win32.Registry]::LocalMachine.OpenSubKey($propPath)
        if ($null -eq $propKey) { return $result }

        $fmtRaw = $propKey.GetValue("{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0")
        if ($null -ne $fmtRaw) {
            try {
                $bytes = [byte[]]$fmtRaw
                if ($bytes.Count -ge 24) {
                    $sampleRate = [BitConverter]::ToInt32($bytes, 12)
                    $bitDepth   = [BitConverter]::ToInt16($bytes, 22)
                    if ($sampleRate -gt 8000 -and $sampleRate -lt 400000) {
                        $result.SampleRate = "$([Math]::Round($sampleRate / 1000, 1)) kHz"
                    }
                    if ($bitDepth -gt 0 -and $bitDepth -le 64) {
                        $result.BitDepth = "${bitDepth}-bit"
                    }
                }
            } catch {}
        }

        $exclRaw = $propKey.GetValue("{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},7")
        if ($null -ne $exclRaw) {
            $result.ExclusiveMode = if ([int]$exclRaw -eq 1) { "Allowed" } else { "Disabled" }
        }
        $propKey.Close()
    } catch {}
    return $result
}

$DriverTable   = @{}
$DriverTableID = @{}
try {
    Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue |
        Where-Object { $_.DeviceName -and $_.DeviceName.Trim() -ne "" } |
        ForEach-Object {
            $ver  = if ($_.DriverVersion) { $_.DriverVersion } else { "" }
            $date = if ($_.DriverDate)    { $_.DriverDate.ToString("yyyy-MM-dd") } else { "" }
            $parts = @($ver, $date) | Where-Object { $_ -ne "" }
            $entry = if ($parts.Count -gt 0) { $parts -join "  " } else { "N/A" }
            $DriverTable[$_.DeviceName.Trim()] = $entry
            if ($_.DeviceID) { $DriverTableID[$_.DeviceID.ToUpper()] = $entry }
        }
} catch {}

$PnpDriverMap = @{}
try {
    $allPnp = Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue | Where-Object { $_.DeviceID }
    foreach ($pnp in $allPnp) {
        $uid = $pnp.DeviceID.ToUpper()
        if ($DriverTableID.ContainsKey($uid)) { $PnpDriverMap[$pnp.Name.Trim()] = $DriverTableID[$uid] }
    }
} catch {}

function Get-DriverString([string]$DevName) {
    $DevName = $DevName.Trim()
    if ($DriverTable.ContainsKey($DevName))    { return $DriverTable[$DevName] }
    if ($PnpDriverMap.ContainsKey($DevName))   { return $PnpDriverMap[$DevName] }
    $best = $DriverTable.Keys | Where-Object { ($DevName -like "*$_*") -or ($_ -like "*$DevName*") } | Select-Object -First 1
    if ($best) { return $DriverTable[$best] }
    return "N/A"
}

function Format-DriverString([string]$DevName, [int]$ColWidth = 25) {
    $driverString = Get-DriverString $DevName
    if ($driverString -eq "N/A") { return "N/A".PadLeft($ColWidth) }
    $parts   = $driverString -split "  ", 2
    $version = $parts[0].Trim()
    $date    = if ($parts.Count -gt 1) { $parts[1].Trim() } else { "" }
    if ($date) {
        $gap = [Math]::Max(1, $ColWidth - $version.Length - $date.Length)
        return $version + (" " * $gap) + $date
    } else {
        return $version.PadRight($ColWidth)
    }
}

$OutDevice = Get-DefaultDevice "Render"
$InDevice  = Get-DefaultDevice "Capture"
$DefaultOutputTokens = @(Get-Tokens (Get-StrippedWrapper $OutDevice.Name))
$DefaultInputTokens  = @(Get-Tokens (Get-StrippedWrapper $InDevice.Name))
$OutQuality          = Get-AudioQualityInfo "Render"  $OutDevice.Guid
$InQuality           = Get-AudioQualityInfo "Capture" $InDevice.Guid
#endregion

#region [1] - HEADER
Clear-Host
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$AudioSrv       = Get-Service -Name "AudioSrv"
$AudioSrvReport = if ($AudioSrv.Status -eq "Running") { "RUNNING" } else { "STOPPED" }
$AudioSrvColor  = if ($AudioSrv.Status -eq "Running") { "Green"   } else { "Red"     }
$OutName  = if ($OutDevice.Name) { $OutDevice.Name } else { "[not detected]" }
$OutColor = if ($OutDevice.Name) { "Green" } else { "Yellow" }
$InName   = if ($InDevice.Name)  { $InDevice.Name  } else { "[not detected]" }
$InColor  = if ($InDevice.Name)  { "Green" } else { "Yellow" }

# Header
$_pfx  = "█  "
$_art1 = "╔═╗ ╔╦╗ ╔╦╗ "
$_art2 = "║ ║  ║║  ║║ "
$_art3 = "╚═╝ ╩╩╝ ╩╩╝ "
$_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1  = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "OUTPUT DEVICE DIAGNOSTIC"
$_ver   = "| $($script:Version)"
$_pad   = " " * [Math]::Max(0, ($_fillW - $_title.Length - $_ver.Length))
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_pad$_ver" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
# Titles White, Answers Yellow
Write-Host "  Audio Service   : " -NoNewline -ForegroundColor White; Write-Host $AudioSrvReport -ForegroundColor $AudioSrvColor
Write-Host "  Default Output  : " -NoNewline -ForegroundColor White; Write-Host $OutName -ForegroundColor $OutColor
Write-Host "                    Format: " -NoNewline -ForegroundColor White; Write-Host "$($OutQuality.SampleRate) / $($OutQuality.BitDepth)" -ForegroundColor Yellow
Write-Host "                    Exclusive: " -NoNewline -ForegroundColor White; Write-Host $OutQuality.ExclusiveMode -ForegroundColor Yellow
Write-Host "  Default Input   : " -NoNewline -ForegroundColor White; Write-Host $InName  -ForegroundColor $InColor
Write-Host "                    Format: " -NoNewline -ForegroundColor White; Write-Host "$($InQuality.SampleRate) / $($InQuality.BitDepth)" -ForegroundColor Yellow
Write-Host "                    Exclusive: " -NoNewline -ForegroundColor White; Write-Host $InQuality.ExclusiveMode -ForegroundColor Yellow
#Write-Host ""
#endregion

#region [2] - DEVICES (WIRED AND BLUETOOTH)
Write-Host "$("Device Name".PadRight(50)) $("Status".PadRight(10)) $("Driver")" -ForegroundColor Cyan
$_ul = ("- " * 24).Substring(0, 48) + " " + ("- " * 5).Substring(0, 10) + " " + ("- " * 13).Substring(0, 25)
$_ui = 0; $_uc = @([ConsoleColor]$LineCol, [ConsoleColor]$AccentCol, [ConsoleColor]$DimCol, [ConsoleColor]$MainCol)
foreach ($_ch in $_ul.ToCharArray()) { Write-Host $_ch -NoNewline -ForegroundColor $_uc[$_ui % $_uc.Count]; if ($_ch -ne ' ') { $_ui++ } }
Write-Host ""
Write-Host "[>] Physical & USB: Output" -ForegroundColor Magenta
$PhysOut = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.ClassGuid -ceq "{4d36e96c-e325-11ce-bfc1-08002be10318}" -and
    $_.Service -notlike "Bth*" -and $_.Service -ne "MSKSSRV" -and
    $_.Name -notlike "*Streaming*" -and $_.Name -notlike "*Bus*" -and $_.Name -notlike "*Microphone*"
} | Sort-Object Name

foreach ($Dev in $PhysOut) {
    $Diag  = switch ($Dev.ConfigManagerErrorCode) { 0 { "Healthy" } 22 { "Disabled" } default { "Error $($Dev.ConfigManagerErrorCode)" } }
    $Drv   = Format-DriverString $Dev.Name
    $Color = if     ($Dev.ConfigManagerErrorCode -eq 22)            { "Red"   }
             elseif (Test-IsDefault $Dev.Name $DefaultOutputTokens) { "Green" }
             else                                                  { "White" }
    Write-Host "$($Dev.Name.PadRight(48).Substring(0,48)) $($Diag.PadRight(10)) $Drv" -ForegroundColor $Color
}

Write-Host "`n[>] Physical & USB: Input" -ForegroundColor Magenta
$PhysIn = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.ClassGuid -eq "{4d36e96c-e325-11ce-bfc1-08002be10318}" -and
    $_.Service -notlike "Bth*" -and $_.Service -ne "MSKSSRV" -and
    $_.Name -notlike "*Streaming*" -and
    ($_.Name -like "*Microphone*" -or $_.Name -like "*Line In*" -or $_.Name -like "*Capture*")
} | Sort-Object Name

foreach ($Dev in $PhysIn) {
    $Diag  = switch ($Dev.ConfigManagerErrorCode) { 0 { "Healthy" } 22 { "Disabled" } default { "Error $($Dev.ConfigManagerErrorCode)" } }
    $Drv   = Format-DriverString $Dev.Name
    $Color = if     ($Dev.ConfigManagerErrorCode -eq 22)           { "Red"   }
             elseif (Test-IsDefault $Dev.Name $DefaultInputTokens) { "Green" }
             else                                                  { "White" }
    Write-Host "$($Dev.Name.PadRight(48).Substring(0,48)) $($Diag.PadRight(10)) $Drv" -ForegroundColor $Color
}

Write-Host "`n[>] Bluetooth: Output" -ForegroundColor Magenta
$BTOut = Get-CimInstance Win32_PnPEntity | Where-Object { $_.Service -eq "BthA2dp" } | Sort-Object Name

foreach ($Dev in $BTOut) {
    $Diag  = switch ($Dev.ConfigManagerErrorCode) { 0 { "Healthy" } 22 { "Disabled" } default { "Error $($Dev.ConfigManagerErrorCode)" } }
    $Drv   = Format-DriverString $Dev.Name
    $Color = if     ($Dev.ConfigManagerErrorCode -eq 22)            { "Red"   }
             elseif (Test-IsDefault $Dev.Name $DefaultOutputTokens) { "Green" }
             else                                                  { "White" }
    Write-Host "$($Dev.Name.PadRight(48).Substring(0,48)) $($Diag.PadRight(10)) $Drv" -ForegroundColor $Color
}

Write-Host "`n[>] Bluetooth: Input" -ForegroundColor Magenta
$BTIn = Get-CimInstance Win32_PnPEntity | Where-Object {
    $_.Service -eq "BthHFAud" -or $_.Service -eq "BthHFEnum"
} | Sort-Object Name

foreach ($Dev in $BTIn) {
    $Diag  = switch ($Dev.ConfigManagerErrorCode) { 0 { "Healthy" } 22 { "Disabled" } default { "Error $($Dev.ConfigManagerErrorCode)" } }
    $Drv   = Format-DriverString $Dev.Name
    $Color = if     ($Dev.ConfigManagerErrorCode -eq 22)           { "Red"   }
             elseif (Test-IsDefault $Dev.Name $DefaultInputTokens) { "Green" }
             else                                                  { "White" }
    Write-Host "$($Dev.Name.PadRight(48).Substring(0,48)) $($Diag.PadRight(10)) $Drv" -ForegroundColor $Color
}
#endregion

Write-HLine -Style dashed