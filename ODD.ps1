<#
.SYNOPSIS
    Output Device Diagnostic (ODD) v2.0
    Developed by SteveTheKiller | Updated: 2026-11-19
.DESCRIPTION
    Inventories all physical, USB, and Bluetooth audio devices with
    health status and driver versions across input and output categories.
    Highlights the default device in green and shows Audio Service state,
    sample rate, bit depth, and exclusive mode from the Windows registry.
    The default is read from the Windows Core Audio default endpoint and
    matched to a device by its Plug and Play identity, never by name. If
    Core Audio cannot be queried the registry guess is used and labelled
    "(inferred)".
    Exit codes: 0 report produced and no problem found, 2 not elevated,
    3 report produced but the audio service is stopped or a device is in
    error.
#>

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

#region [0] - HELPER FUNCTIONS
# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-OddElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}
function Initialize-OddConsole {
    try { Clear-Host -ErrorAction Stop } catch { }
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
}
function Get-OddAudioService { Get-Service -Name "AudioSrv" -ErrorAction SilentlyContinue }
function Get-OddPnpEntities { Get-CimInstance Win32_PnPEntity -ErrorAction SilentlyContinue }
function Get-OddDriverRows { Get-CimInstance Win32_PnPSignedDriver -ErrorAction SilentlyContinue }

# Asks Windows Core Audio for the default endpoint (console role). Returns Ok = $false when the
# API could not be used, otherwise Ok = $true with the endpoint GUID, or Id = $null when Windows
# reports no default endpoint for that direction.
function Get-OddDefaultEndpoint {
    param([ValidateSet('Render', 'Capture')][string]$Flow)
    try {
        if (-not ('KillerOdd.DefaultEndpoint' -as [type])) {
            Add-Type -TypeDefinition $script:OddCoreAudioSource -ErrorAction Stop
        }
        $id = $null
        $hr = [KillerOdd.DefaultEndpoint]::Get($(if ($Flow -eq 'Render') { 0 } else { 1 }), 0, [ref]$id)
        if ($hr -eq 0 -and $id -match '\{([0-9a-fA-F-]{36})\}\s*$') {
            return [pscustomobject]@{ Ok = $true; Id = $Matches[1].ToLowerInvariant() }
        }
        # 0x80070490 is E_NOTFOUND: Windows has no default endpoint for this direction.
        if ($hr -eq -2147023728) { return [pscustomobject]@{ Ok = $true; Id = $null } }
    } catch { }
    [pscustomobject]@{ Ok = $false; Id = $null }
}

# Reads every endpoint under MMDevices through the 64-bit registry view, so a 32-bit host
# sees the same list as a 64-bit one.
function Get-OddEndpoints {
    param([ValidateSet('Render', 'Capture')][string]$Branch)
    $list = New-Object System.Collections.ArrayList
    $hklm = $null
    $baseKey = $null
    try {
        $hklm = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::LocalMachine, [Microsoft.Win32.RegistryView]::Registry64)
        $baseKey = $hklm.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\MMDevices\Audio\$Branch")
        if ($null -eq $baseKey) { return @() }
        foreach ($subName in $baseKey.GetSubKeyNames()) {
            try {
                $devKey = $baseKey.OpenSubKey($subName)
                if ($null -eq $devKey) { continue }
                $state = [int64]$devKey.GetValue("DeviceState")
                $level = $devKey.GetValue("Level:0")
                $level = if ($null -ne $level) { [int]$level } else { -1 }
                $devKey.Close()
                $propKey = $baseKey.OpenSubKey("$subName\Properties")
                if ($null -eq $propKey) { continue }
                $name   = $propKey.GetValue("{b3f8fa53-0004-438e-9003-51a46e139bfc},6")
                # The parent Plug and Play device, stored as "{1}.<DeviceID>".
                $parent = $propKey.GetValue("{b3f8fa53-0004-438e-9003-51a46e139bfc},2")
                $fmt    = $propKey.GetValue("{f19f064d-082c-4e27-bc73-6882a1bb8e4c},0")
                $excl   = $propKey.GetValue("{1da5d803-d492-4edd-8c23-e0c0ffee7f0e},7")
                $propKey.Close()
                [void]$list.Add([pscustomobject]@{
                    Guid = $subName.Trim('{', '}').ToLowerInvariant(); State = $state; Level = $level
                    Name = $(if ($name -is [string]) { $name } else { "" })
                    ParentId = $(if ($parent -is [string]) { ($parent -replace '^\{\d+\}\.', '').Trim() } else { "" })
                    Format = $fmt; Exclusive = $excl
                })
            } catch { continue }
        }
    } catch { return @() } finally {
        if ($baseKey) { $baseKey.Close() }
        if ($hklm) { $hklm.Close() }
    }
    return @($list)
}

# Picks the default endpoint. Core Audio is authoritative. Only when it cannot be queried is the
# highest active endpoint used, and the result says so with Source = Inferred.
function Get-DefaultDevice {
    param([ValidateSet('Render', 'Capture')][string]$Branch, $Endpoints)
    $result = [pscustomobject]@{ Guid = ""; Name = ""; ParentId = ""; Source = "None"; Endpoint = $null }
    $core = Get-OddDefaultEndpoint -Flow $Branch
    if ($core.Ok) {
        if (-not $core.Id) { return $result }
        $result.Guid = $core.Id
        $result.Source = "CoreAudio"
        $ep = @($Endpoints | Where-Object { $_.Guid -eq $core.Id }) | Select-Object -First 1
        if ($ep) { $result.Name = $ep.Name; $result.ParentId = $ep.ParentId; $result.Endpoint = $ep }
        return $result
    }
    # DeviceState is a bit field, only the low nibble is the state (1 = active).
    $best = @($Endpoints | Where-Object { ($_.State -band 0xF) -eq 1 -and $_.Name } | Sort-Object Level -Descending) | Select-Object -First 1
    if ($best) {
        $result.Guid = $best.Guid; $result.Name = $best.Name; $result.ParentId = $best.ParentId
        $result.Source = "Inferred"; $result.Endpoint = $best
    }
    return $result
}

# A listed device is the default only when its Plug and Play identity is the default endpoint's parent.
function Test-IsDefault {
    param([string]$DeviceId, [string]$DefaultParentId)
    if ([string]::IsNullOrWhiteSpace($DeviceId) -or [string]::IsNullOrWhiteSpace($DefaultParentId)) { return $false }
    return ($DeviceId.Trim().ToUpperInvariant() -eq $DefaultParentId.Trim().ToUpperInvariant())
}

function Get-AudioQualityInfo {
    param($Endpoint)
    $result = [PSCustomObject]@{ SampleRate = "N/A"; BitDepth = "N/A"; ExclusiveMode = "N/A" }
    if ($null -eq $Endpoint) { return $result }
    if ($null -ne $Endpoint.Format) {
        try {
            $bytes = [byte[]]$Endpoint.Format
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
    if ($null -ne $Endpoint.Exclusive) {
        $result.ExclusiveMode = if ([int]$Endpoint.Exclusive -eq 1) { "Allowed" } else { "Disabled" }
    }
    return $result
}

function Get-OddDriverTables {
    $byName = @{}
    $byId = @{}
    try {
        Get-OddDriverRows |
            Where-Object { $_.DeviceName -and $_.DeviceName.Trim() -ne "" } |
            ForEach-Object {
                $ver  = if ($_.DriverVersion) { $_.DriverVersion } else { "" }
                $date = if ($_.DriverDate)    { $_.DriverDate.ToString("yyyy-MM-dd") } else { "" }
                $parts = @($ver, $date) | Where-Object { $_ -ne "" }
                $entry = if ($parts.Count -gt 0) { $parts -join "  " } else { "N/A" }
                $byName[$_.DeviceName.Trim()] = $entry
                if ($_.DeviceID) { $byId[$_.DeviceID.ToUpper()] = $entry }
            }
    } catch {}
    [pscustomobject]@{ ByName = $byName; ById = $byId }
}

function Get-DriverString([string]$DevName, [string]$DevId) {
    $DevName = $DevName.Trim()
    if ($DevId -and $script:DriverTableID.ContainsKey($DevId.ToUpper())) { return $script:DriverTableID[$DevId.ToUpper()] }
    if ($script:DriverTable.ContainsKey($DevName)) { return $script:DriverTable[$DevName] }
    $best = $script:DriverTable.Keys | Where-Object { ($DevName -like "*$_*") -or ($_ -like "*$DevName*") } | Select-Object -First 1
    if ($best) { return $script:DriverTable[$best] }
    return "N/A"
}

function Format-DriverString([string]$DevName, [string]$DevId, [int]$ColWidth = 26) {
    $driverString = Get-DriverString $DevName $DevId
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

$script:Width    = 85
$script:Version  = "v2.0"
$LineCol   = "Magenta"
$MainCol   = "White"
$ArtCol    = "Yellow"
$DimCol    = "DarkGray"
$AccentCol = "Yellow"
$script:DriverTable   = @{}
$script:DriverTableID = @{}
$script:OddCoreAudioSource = @'
using System;
using System.Runtime.InteropServices;
namespace KillerOdd {
    [ComImport, Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDeviceEnumerator {
        [PreserveSig] int EnumAudioEndpoints(int dataFlow, int stateMask, out IntPtr devices);
        [PreserveSig] int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice endpoint);
    }
    [ComImport, Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    interface IMMDevice {
        [PreserveSig] int Activate(ref Guid iid, int clsCtx, IntPtr activationParams, [MarshalAs(UnmanagedType.IUnknown)] out object iface);
        [PreserveSig] int OpenPropertyStore(int access, out IntPtr properties);
        [PreserveSig] int GetId([MarshalAs(UnmanagedType.LPWStr)] out string id);
    }
    [ComImport, Guid("BCDE0395-E52F-467C-8E3D-C4579291692E")]
    class MMDeviceEnumeratorCom { }
    public static class DefaultEndpoint {
        // Returns the HRESULT. id is the endpoint id when the call succeeded, otherwise null.
        public static int Get(int dataFlow, int role, out string id) {
            id = null;
            IMMDeviceEnumerator enumerator = null;
            IMMDevice device = null;
            try {
                enumerator = (IMMDeviceEnumerator)new MMDeviceEnumeratorCom();
                int hr = enumerator.GetDefaultAudioEndpoint(dataFlow, role, out device);
                if (hr != 0 || device == null) return hr;
                return device.GetId(out id);
            } finally {
                if (device != null) Marshal.ReleaseComObject(device);
                if (enumerator != null) Marshal.ReleaseComObject(enumerator);
            }
        }
    }
}
'@
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
#endregion

# Prints one device section and returns how many listed devices were in error.
function Write-OddSection {
    param([string]$Title, $Devices, [string]$DefaultParentId, [string]$Lead = "")
    Write-Host "$Lead[>] $Title" -ForegroundColor Magenta
    $bad = 0
    foreach ($Dev in @($Devices | Where-Object { $null -ne $_ })) {
        $code  = [int]$Dev.ConfigManagerErrorCode
        $Diag  = switch ($code) { 0 { "Healthy" } 22 { "Disabled" } default { "Error $code" } }
        if ($code -ne 0 -and $code -ne 22) { $bad++ }
        $Drv   = Format-DriverString $Dev.Name $Dev.DeviceID
        $Color = if     ($code -eq 22)                                    { "Red"   }
                 elseif (Test-IsDefault $Dev.DeviceID $DefaultParentId)    { "Green" }
                 else                                                     { "White" }
        Write-Host "$($Dev.Name.PadRight(47).Substring(0,47)) $($Diag.PadRight(10)) $Drv" -ForegroundColor $Color
    }
    return $bad
}

function Invoke-OddMain {
    if (-not (Test-OddElevated)) {
        Write-Error "O.D.D. requires Elevated Privileges. Please run as Administrator."
        Write-Output "[FAIL] ODD not elevated, nothing was read (exit 2)"
        Exit-KillerScript 2
    }
    $endpointsOut = @(Get-OddEndpoints -Branch Render)
    $endpointsIn  = @(Get-OddEndpoints -Branch Capture)
    $OutDevice = Get-DefaultDevice -Branch Render  -Endpoints $endpointsOut
    $InDevice  = Get-DefaultDevice -Branch Capture -Endpoints $endpointsIn
    $OutQuality = Get-AudioQualityInfo $OutDevice.Endpoint
    $InQuality  = Get-AudioQualityInfo $InDevice.Endpoint
    $tables = Get-OddDriverTables
    $script:DriverTable   = $tables.ByName
    $script:DriverTableID = $tables.ById
    $allPnp = @(Get-OddPnpEntities | Where-Object { $_.DeviceID })

    #region [1] - HEADER
    Initialize-OddConsole
    $AudioSrv       = Get-OddAudioService
    $AudioRunning   = ($null -ne $AudioSrv -and "$($AudioSrv.Status)" -eq "Running")
    $AudioSrvReport = if ($AudioRunning) { "RUNNING" } else { "STOPPED" }
    $AudioSrvColor  = if ($AudioRunning) { "Green"   } else { "Red"     }
    $suffixOut = if ($OutDevice.Source -eq "Inferred") { " (inferred)" } else { "" }
    $suffixIn  = if ($InDevice.Source  -eq "Inferred") { " (inferred)" } else { "" }
    $OutName  = if ($OutDevice.Name) { "$($OutDevice.Name)$suffixOut" } else { "[not detected]" }
    $OutColor = if ($OutDevice.Name) { "Green" } else { "Yellow" }
    $InName   = if ($InDevice.Name)  { "$($InDevice.Name)$suffixIn" } else { "[not detected]" }
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
    #endregion

    #region [2] - DEVICES (WIRED AND BLUETOOTH)
    Write-Host "$("Device Name".PadRight(47)) $("Status".PadRight(10)) $("Driver")" -ForegroundColor Cyan
    $_ul = ("- " * 24).Substring(0, 47) + " " + ("- " * 5).Substring(0, 10) + " " + ("- " * 13).Substring(0, 26)
    $_ui = 0; $_uc = @([ConsoleColor]$LineCol, [ConsoleColor]$AccentCol, [ConsoleColor]$DimCol, [ConsoleColor]$MainCol)
    foreach ($_ch in $_ul.ToCharArray()) { Write-Host $_ch -NoNewline -ForegroundColor $_uc[$_ui % $_uc.Count]; if ($_ch -ne ' ') { $_ui++ } }
    Write-Host ""
    $PhysOut = $allPnp | Where-Object {
        $_.ClassGuid -ceq "{4d36e96c-e325-11ce-bfc1-08002be10318}" -and
        $_.Service -notlike "Bth*" -and $_.Service -ne "MSKSSRV" -and
        $_.Name -notlike "*Streaming*" -and $_.Name -notlike "*Bus*" -and $_.Name -notlike "*Microphone*"
    } | Sort-Object Name
    $PhysIn = $allPnp | Where-Object {
        $_.ClassGuid -eq "{4d36e96c-e325-11ce-bfc1-08002be10318}" -and
        $_.Service -notlike "Bth*" -and $_.Service -ne "MSKSSRV" -and
        $_.Name -notlike "*Streaming*" -and
        ($_.Name -like "*Microphone*" -or $_.Name -like "*Line In*" -or $_.Name -like "*Capture*")
    } | Sort-Object Name
    $BTOut = $allPnp | Where-Object { $_.Service -eq "BthA2dp" } | Sort-Object Name
    $BTIn  = $allPnp | Where-Object { $_.Service -eq "BthHFAud" -or $_.Service -eq "BthHFEnum" } | Sort-Object Name

    $errors = 0
    $errors += [int](Write-OddSection "Physical & USB: Output" $PhysOut $OutDevice.ParentId)
    $errors += [int](Write-OddSection "Physical & USB: Input"  $PhysIn  $InDevice.ParentId "`n")
    $errors += [int](Write-OddSection "Bluetooth: Output"     $BTOut    $OutDevice.ParentId "`n")
    $errors += [int](Write-OddSection "Bluetooth: Input"       $BTIn     $InDevice.ParentId "`n")
    #endregion

    Write-HLine -Style dashed
    $notes = @()
    if (-not $AudioRunning) { $notes += "the Windows Audio service is not running" }
    if ($errors -gt 0) { $notes += "$errors listed device(s) report an error" }
    if ($notes.Count -gt 0) {
        Write-Output "[WARN] ODD report produced, but $($notes -join ' and ') (exit 3)"
        Exit-KillerScript 3
    }
    Write-Output "[PASS] ODD report produced, audio service running and no device errors (exit 0)"
    Exit-KillerScript 0
}

if (-not $global:KillerScriptsTestMode) { Invoke-OddMain }
