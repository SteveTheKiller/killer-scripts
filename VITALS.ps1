<#
.SYNOPSIS
    Visual Interface for Technical Asset & Logistics Summary (V.I.T.A.L.S.) v2.2
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Collects and displays a full hardware and network snapshot:
    make/model/serial, CPU/RAM/GPU specs, disk health, BitLocker status
    and recovery keys, network configuration, battery wear, TPM/Secure
    Boot status, domain membership (with DC detection), and local admin
    members.
#>

#region Pre-Flight Checks
# ============================================================================
if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Exit
}
function Get-WMI {
    param([string]$Class)
    if ($PSVersionTable.PSEdition -eq "Core") {
        return Get-CimInstance -ClassName $Class -ErrorAction SilentlyContinue
    } else {
        return Get-WmiObject -Class $Class -ErrorAction SilentlyContinue
    }
}
Clear-Host
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OS  = Get-WMI Win32_OperatingSystem
$Sys = Get-WMI Win32_ComputerSystem
$VMType = $null
if     ($Sys.Manufacturer -match "VMware")                                             { $VMType = "VMware" }
elseif ($Sys.Manufacturer -match "innotek|Oracle" -or $Sys.Model -match "VirtualBox") { $VMType = "VirtualBox" }
elseif ($Sys.Manufacturer -match "Microsoft" -and $Sys.Model -match "Virtual")        { $VMType = "Hyper-V" }
elseif ($Sys.Manufacturer -match "QEMU" -or $Sys.Model -match "QEMU")                 { $VMType = "QEMU/KVM" }
elseif ($Sys.Manufacturer -match "Xen")                                                { $VMType = "Xen" }
elseif ($Sys.Manufacturer -match "Parallels" -or $Sys.Model -match "Parallels")        { $VMType = "Parallels" }
$LastBootRaw = $OS.LastBootUpTime
$LastBoot = if ($LastBootRaw -is [DateTime]) { $LastBootRaw } else { [Management.ManagementDateTimeConverter]::ToDateTime($LastBootRaw) }
$Uptime = (Get-Date) - $LastBoot
$RebootPending = $false
$RegChecks = @(
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
    "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
)
foreach ($Key in $RegChecks) { if (Test-Path $Key) { $RebootPending = $true } }
if ($Uptime.Days -gt 7) { $UptimeColor = "Red" }
elseif ($Uptime.Days -gt 3) { $UptimeColor = "Yellow" }
else { $UptimeColor = "Green" }
try { $PublicIP = (Invoke-WebRequest -Uri "http://ifconfig.me/ip" -UseBasicParsing -TimeoutSec 2).Content.Trim() } catch { $PublicIP = $null }
$TPMInfo     = Get-Tpm -ErrorAction SilentlyContinue
$TPMVerRaw   = if ($TPMInfo -and $TPMInfo.TpmPresent) {
    if ($PSVersionTable.PSEdition -eq "Core") { (Get-CimInstance -Namespace "root\CIMv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue).SpecVersion }
    else { (Get-WmiObject -Namespace "root\CIMv2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue).SpecVersion }
} else { $null }
$TPMVerClean = if ($TPMVerRaw) { ($TPMVerRaw -split ',')[0].Trim() } else { $null }
$SecureBoot  = try { Confirm-SecureBootUEFI -ErrorAction Stop } catch { $null }

$script:Width    = 85
$script:Version  = "v2.2"
$script:PrefixFG = "Cyan"
$script:LabelFG  = "Gray"
$script:ValueFG  = "Yellow"
$script:WarnFG   = "DarkYellow"
$LineCol  = "Green"
$MainCol  = "White"
$ArtCol   = "Yellow"
$DimCol   = "DarkGray"
function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [Math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "━" * $Width
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

# Header
$_pfx  = "█  "
$_art1 = "╦  ╦ ╦ ╔╦╗ ╔═╗ ╦   ╔═╗ "
$_art2 = "╚╗╔╝ ║  ║  ╠═╣ ║   ╚═╗ "
$_art3 = " ╚╝  ╩  ╩  ╩ ╩ ╩═╝ ╚═╝ "
$_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1  = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "VISUAL INTERFACE FOR TECHNICAL ASSET & LOGISTICS SUMMARY"
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

# System Info
Write-Host "   Device Name  : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $script:ValueFG
Write-Host "   Uptime       : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($Uptime.Days)d $($Uptime.Hours)h $($Uptime.Minutes)m" -ForegroundColor $UptimeColor
Write-Host "   Last Boot    : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($LastBoot.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor $script:ValueFG
Write-Host "   Public IP    : " -NoNewline -ForegroundColor $script:LabelFG; if ($PublicIP) { Write-Host $PublicIP -ForegroundColor $script:ValueFG } else { Write-Host "Offline" -ForegroundColor Red }
$IsDomain    = $Sys.DomainRole -in @(1,3,4,5)
$IsEntra     = $null -ne (Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo" -ErrorAction SilentlyContinue)
Write-Host "   Domain Joined: " -NoNewline -ForegroundColor $script:LabelFG
if ($IsDomain) {
    Write-Host $Sys.Domain -NoNewline -ForegroundColor $script:ValueFG
    if ($Sys.DomainRole -ge 4) { Write-Host "  (DC)" -NoNewline -ForegroundColor DarkGreen }
    Write-Host ""
} else { Write-Host "No" -ForegroundColor $script:WarnFG }
Write-Host "   Entra Joined : " -NoNewline -ForegroundColor $script:LabelFG; if ($IsEntra) { Write-Host "Yes" -ForegroundColor Green } else { Write-Host "No" -ForegroundColor $script:WarnFG }
if ($VMType) { Write-Host "   Environment  : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "Virtual Machine " -NoNewline -ForegroundColor $script:ValueFG; Write-Host "($VMType)" -ForegroundColor $script:WarnFG }
$TPMStatus  = if ($TPMInfo -and $TPMInfo.TpmPresent) { if ($TPMInfo.TpmReady) { "Ready" } else { "Not Ready" } } else { "Not Detected" }
$TPMColor   = if ($TPMInfo -and $TPMInfo.TpmPresent) { if ($TPMInfo.TpmReady) { "Green" } else { "DarkYellow" } } else { "DarkYellow" }
$TPMDisplay = if ($TPMVerClean) { "TPM $TPMVerClean $TPMStatus" } else { "TPM $TPMStatus" }
$SBText     = if ($SecureBoot -eq $true) { "Enabled" } elseif ($SecureBoot -eq $false) { "Disabled" } else { "Unavailable" }
$SBColor    = if ($SecureBoot -eq $true) { "Green" } elseif ($SecureBoot -eq $false) { "Red" } else { "DarkYellow" }
Write-Host "   Security     : " -NoNewline -ForegroundColor $script:LabelFG
Write-Host $TPMDisplay -NoNewline -ForegroundColor $TPMColor
Write-Host "  |  " -NoNewline -ForegroundColor DarkGray
Write-Host "Secure Boot $SBText" -ForegroundColor $SBColor
Write-HLine -Style dashed
#endregion

#region 1. System Identity
# ============================================================================
Write-Host "Identifying Hardware..." -ForegroundColor $script:PrefixFG

$Bios = Get-WMI Win32_Bios
$OS   = Get-WMI Win32_OperatingSystem
$Base = Get-WMI Win32_BaseBoard
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion

# Intelligent OEM filtering - Fallback to Motherboard if generic
$Make  = if ($Sys.Manufacturer -match "To Be Filled" -or [string]::IsNullOrWhiteSpace($Sys.Manufacturer)) { $Base.Manufacturer } else { $Sys.Manufacturer }
$Model = if ($Sys.Model -match "To Be Filled" -or [string]::IsNullOrWhiteSpace($Sys.Model)) { $Base.Product } else { $Sys.Model }
$SNRaw = if ($Bios.SerialNumber -match "To Be Filled" -or [string]::IsNullOrWhiteSpace($Bios.SerialNumber)) { $null } else { $Bios.SerialNumber }

Write-Host "   Manufacturer : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $Make -ForegroundColor $script:ValueFG
Write-Host "   Model        : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $Model -ForegroundColor $script:ValueFG
Write-Host "   SerialNumber : " -NoNewline -ForegroundColor $script:LabelFG; if ($SNRaw) { Write-Host $SNRaw -ForegroundColor $script:ValueFG } else { Write-Host "Serial Unavailable" -ForegroundColor $script:WarnFG }
Write-Host "   OS Version   : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $OS.Caption -ForegroundColor $script:ValueFG
Write-Host "   Build/Ver    : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($OS.BuildNumber) ($WinVer)" -ForegroundColor $script:ValueFG
#endregion

#region 2. CPU & GPU Performance
# ============================================================================
Write-Host "Processing Power..." -ForegroundColor $script:PrefixFG
$CPU = Get-WMI Win32_Processor
$GPU = Get-WMI Win32_VideoController
$RAMSticks = Get-WMI Win32_PhysicalMemory
$TotalRAM = [Math]::Round($Sys.TotalPhysicalMemory / 1GB, 2)
$MaxSpeed = ($RAMSticks | Measure-Object -Property Speed -Maximum).Maximum
Write-Host "   CPU          : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $CPU.Name -ForegroundColor $script:ValueFG
Write-Host "   Cores/Logcl  : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($CPU.NumberOfCores) Cores | $($CPU.NumberOfLogicalProcessors) Threads" -ForegroundColor $script:ValueFG
Write-Host "   Clock Speed  : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($CPU.CurrentClockSpeed)MHz" -ForegroundColor $script:ValueFG
Write-Host "   Memory       : " -NoNewline -ForegroundColor $script:LabelFG
if ($MaxSpeed) {
    Write-Host "$TotalRAM GB Installed @ $($MaxSpeed)MHz" -ForegroundColor $script:ValueFG
} else {
    Write-Host "$TotalRAM GB Installed " -NoNewline -ForegroundColor $script:ValueFG; Write-Host "(speed unavailable)" -ForegroundColor $script:WarnFG
}
foreach ($G in $GPU) {
    $VRAM = [Math]::Round($G.AdapterRAM / 1GB, 2)
    Write-Host "   GPU          : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($G.Name) " -NoNewline -ForegroundColor $script:ValueFG
    if ($VRAM -gt 0) { Write-Host "($VRAM GB VRAM)" -ForegroundColor $script:ValueFG } else { Write-Host "(VRAM unavailable)" -ForegroundColor $script:WarnFG }
}
#endregion

#region 3. Battery Status
# ============================================================================
$Battery = Get-WMI Win32_Battery
if ($Battery) {
    Write-Host "Battery Status..." -ForegroundColor $script:PrefixFG
    $Charge      = $Battery.EstimatedChargeRemaining
    $ChargeColor = if ($Charge -ge 50) { "Green" } elseif ($Charge -ge 20) { "Yellow" } else { "Red" }
    $BattStatus  = switch ($Battery.BatteryStatus) {
        1 { "Discharging" }  2 { "On AC Power" }   3 { "Fully Charged" }
        4 { "Low" }          5 { "Critical" }       6 { "Charging" }
        7 { "Charging/High" } 8 { "Charging/Low" } 9 { "Charging/Critical" }
        default { "Unknown" }
    }
    Write-Host "   Charge       : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$Charge%" -ForegroundColor $ChargeColor
    Write-Host "   Status       : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $BattStatus -ForegroundColor $script:ValueFG
    if ($Battery.DesignCapacity -and $Battery.FullChargeCapacity -and $Battery.DesignCapacity -gt 0) {
        $WearPct   = [Math]::Round((1 - ($Battery.FullChargeCapacity / $Battery.DesignCapacity)) * 100, 1)
        $WearColor = if ($WearPct -lt 20) { "Green" } elseif ($WearPct -lt 40) { "Yellow" } else { "Red" }
        Write-Host "   Wear Level   : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$WearPct% degraded" -ForegroundColor $WearColor
    }
    $Runtime = $Battery.EstimatedRunTime
    if ($Runtime -and $Runtime -ne 71582788) {
        $BHours = [Math]::Floor($Runtime / 60); $BMins = $Runtime % 60
        Write-Host "   Est. Runtime : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "${BHours}h ${BMins}m remaining" -ForegroundColor $script:ValueFG
    }
}
#endregion

#region 4. Resources & Storage Health
# ============================================================================
$BLRecoveryKeys = @()
$DiskReport = Get-PhysicalDisk | ForEach-Object {
    $PhysDisk = $_
    $BLStatus = "Off"
    $DriveLetters = @()
    $Disk = Get-Disk -Number $PhysDisk.DeviceID -ErrorAction SilentlyContinue
    if ($Disk) {
        $Partitions = $Disk | Get-Partition -ErrorAction SilentlyContinue
        foreach ($P in $Partitions) {
            if ($P.DriveLetter) {
                $DriveLetters += "$($P.DriveLetter):"
                try {
                    $Vol = Get-BitLockerVolume -MountPoint "$($P.DriveLetter):" -ErrorAction Stop
                    if ($Vol.ProtectionStatus -eq 'On') {
                        $BLStatus = "On"
                        $RecoveryProtector = $Vol.KeyProtector | Where-Object { $_.KeyProtectorType -eq 'RecoveryPassword' } | Select-Object -First 1
                        if ($RecoveryProtector) {
                            $BLRecoveryKeys += [PSCustomObject]@{
                                Drive    = "$($P.DriveLetter):"
                                ID       = $RecoveryProtector.KeyProtectorId
                                Key      = $RecoveryProtector.RecoveryPassword
                            }
                        }
                    }
                } catch { $BLStatus = "N/A" }
            }
        }
    }
    [PSCustomObject]@{
        "Drive"     = if ($DriveLetters) { $DriveLetters -join ", " } else { "None" }
        FriendlyName = $PhysDisk.FriendlyName
        "Size_GB"   = [Math]::Round($PhysDisk.Size / 1GB, 2)
        MediaType    = $PhysDisk.MediaType
        "BitLocker"  = $BLStatus
        "Status"    = $PhysDisk.HealthStatus
    }
}
Write-Host "`nDrive   FriendlyName         Size_GB  MediaType    BitLocker  Status" -ForegroundColor $script:PrefixFG
foreach ($D in ($DiskReport | Sort-Object Drive)) {
    $FormatStr = "{0,-7} {1,-20} {2,-8} {3,-12} {4,-10} {5}"
    Write-Host ($FormatStr -f $D.Drive, ($D.FriendlyName.PadRight(20).Substring(0,20)), $D.Size_GB, $D.MediaType, $D.BitLocker, $D.Status) -ForegroundColor $script:ValueFG
}
if ($BLRecoveryKeys.Count -gt 0) {
    Write-Host ""
    Write-Host "BitLocker Recovery Keys" -ForegroundColor $script:PrefixFG
    foreach ($K in $BLRecoveryKeys) {
        Write-Host $K.Drive -NoNewline -ForegroundColor Yellow; Write-Host "  ID  : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $K.ID -ForegroundColor $script:ValueFG
        Write-Host "    Key : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $K.Key -ForegroundColor Yellow
    }
}
#endregion

#region 5. Network Configuration
# ============================================================================
function Get-SubnetMask ([int]$Prefix) {
    $mask = @(0,0,0,0)
    for ($i=0; $i -lt $Prefix; $i++) { $mask[[Math]::Floor($i/8)] += [Math]::Pow(2, (7 - ($i % 8))) }
    return $mask -join "."
}
$NetworkReport = Get-NetAdapter | Where-Object Status -eq "Up" | ForEach-Object {
    $Interface = $_
    $IPs = Get-NetIPAddress -InterfaceIndex $Interface.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue
    $DNS = (Get-DnsClientServerAddress -InterfaceIndex $Interface.InterfaceIndex -AddressFamily IPv4 -ErrorAction SilentlyContinue).ServerAddresses
    foreach ($IP in $IPs) {
        if ($IP.IPAddress -notlike "169.254*") {
            $Gateway = (Get-NetRoute -InterfaceIndex $Interface.InterfaceIndex -DestinationPrefix "0.0.0.0/0" -ErrorAction SilentlyContinue).NextHop | Select-Object -First 1
            [PSCustomObject]@{
                Interface = $Interface.InterfaceAlias
                IPAddress = $IP.IPAddress
                Subnet    = Get-SubnetMask $IP.PrefixLength
                Gateway   = if ($Gateway) { $Gateway } else { "None" }
                DNS       = if ($DNS) { [string[]]$DNS } else { [string[]]@("None") }
            }
        }
    }
}
Write-Host "`nInterface            IPAddress       Subnet          Gateway         DNS Servers" -ForegroundColor $script:PrefixFG
foreach ($N in $NetworkReport) {
    $FormatStr = "{0,-20} {1,-15} {2,-15} {3,-15} {4}"
    Write-Host ($FormatStr -f ($N.Interface.PadRight(20).Substring(0,20)), $N.IPAddress, $N.Subnet, $N.Gateway, $N.DNS[0]) -ForegroundColor $script:ValueFG
    if ($N.DNS.Count -gt 1) {
        $Indent = " " * 69
        foreach ($D in $N.DNS | Select-Object -Skip 1) {
            Write-Host "$Indent$D" -ForegroundColor $script:ValueFG
        }
    }
}
#endregion

#region 6. Local Administrators
# ============================================================================
Write-HLine -Style dashed
Write-Host "Local Administrators..." -ForegroundColor $script:PrefixFG
try {
    $AdminMembers = ([ADSI]"WinNT://./Administrators,group").psbase.Invoke('Members') | ForEach-Object {
        $MemberName = $_.GetType().InvokeMember('Name',    'GetProperty', $null, $_, $null)
        $AdsPath    = $_.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $_, $null)
        $Domain     = if ($AdsPath -match "WinNT://([^/]+)/[^/]+$") { $Matches[1] } else { $env:COMPUTERNAME }
        "$Domain\$MemberName"
    }
    foreach ($Admin in $AdminMembers) {
        Write-Host "" -NoNewline; Write-Host $Admin -ForegroundColor $script:ValueFG
    }
} catch {
    Write-Host "" -NoNewline; Write-Host "Unable to enumerate local administrators." -ForegroundColor $script:WarnFG
}
#endregion

if ($RebootPending) { Write-Host "[!] ALERT: Pending reboot flag detected." -ForegroundColor Red }

# Footer (art reuses header $_art1/$_art2/$_art3, prefixed with a space)
$_sfx    = "█"
$_ftr1   = " " + $_art1; $_ftr2 = " " + $_art2; $_ftr3 = " " + $_art3
$_ftrW   = $_artW + 1
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  TECHNICAL ASSET DATA COLLECTION COMPLETE"
$_fver   = "| $($script:Version)"
$_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_fver.Length))
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr1 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline; Write-Host $_ftr2 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host $_ftr3 -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $LineCol

exit 0
