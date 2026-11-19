<#
.SYNOPSIS
    Visual Interface for Technical Asset & Logistics Summary (V.I.T.A.L.S.) v2.2
    Developed by SteveTheKiller | Updated: 2026-03-13
.DESCRIPTION
    Collects and displays a full hardware and network snapshot:
    make/model/serial, CPU/RAM/GPU specs, disk health, BitLocker status
    and every BitLocker key protector with its ID (and the recovery
    password of each recovery-password protector), network configuration,
    battery wear, TPM/Secure Boot status, domain membership (with DC
    detection), and local admin members.
    Exit codes: 0 report produced, 2 not elevated, 3 report produced but
    the disk or BitLocker information could not be read.
#>

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-VitalsElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
function Get-WMI {
    param([string]$Class)
    if ($PSVersionTable.PSEdition -eq "Core") {
        return Get-CimInstance -ClassName $Class -ErrorAction SilentlyContinue
    } else {
        return Get-WmiObject -Class $Class -ErrorAction SilentlyContinue
    }
}
function Initialize-VitalsConsole {
    try { Clear-Host -ErrorAction Stop } catch { }
    try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
}
function Test-VitalsRebootPending {
    $RegChecks = @(
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired",
        "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
    )
    foreach ($Key in $RegChecks) { if (Test-Path $Key) { return $true } }
    return $false
}
function Get-VitalsPublicIp {
    try { (Invoke-WebRequest -Uri "http://ifconfig.me/ip" -UseBasicParsing -TimeoutSec 2).Content.Trim() } catch { $null }
}
function Get-VitalsTpm {
    $info = Get-Tpm -ErrorAction SilentlyContinue
    $ver = $null
    if ($info -and $info.TpmPresent) {
        $raw = if ($PSVersionTable.PSEdition -eq "Core") { (Get-CimInstance -Namespace "root\CIMv2\Security\MicrosoftTpm" -ClassName Win32_Tpm -ErrorAction SilentlyContinue).SpecVersion }
               else { (Get-WmiObject -Namespace "root\CIMv2\Security\MicrosoftTpm" -Class Win32_Tpm -ErrorAction SilentlyContinue).SpecVersion }
        if ($raw) { $ver = ($raw -split ',')[0].Trim() }
    }
    [pscustomobject]@{ Info = $info; Version = $ver }
}
function Get-VitalsSecureBoot { try { Confirm-SecureBootUEFI -ErrorAction Stop } catch { $null } }
function Get-VitalsIsEntra { $null -ne (Get-ChildItem "HKLM:\SYSTEM\CurrentControlSet\Control\CloudDomainJoin\JoinInfo" -ErrorAction SilentlyContinue) }
function Get-VitalsWinVer { (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue).DisplayVersion }
# Physical disks with the drive letters of their partitions. Ok is $false when the disks could not be listed.
function Get-VitalsDiskInfo {
    try { $phys = @(Get-PhysicalDisk -ErrorAction Stop) } catch { return [pscustomobject]@{ Ok = $false; Disks = @() } }
    $list = New-Object System.Collections.ArrayList
    foreach ($PhysDisk in $phys) {
        $letters = @()
        $Disk = Get-Disk -Number $PhysDisk.DeviceID -ErrorAction SilentlyContinue
        if ($Disk) {
            foreach ($P in @($Disk | Get-Partition -ErrorAction SilentlyContinue)) {
                if ($P.DriveLetter) { $letters += "$($P.DriveLetter):" }
            }
        }
        [void]$list.Add([pscustomobject]@{
            FriendlyName = $PhysDisk.FriendlyName; Size = $PhysDisk.Size; MediaType = $PhysDisk.MediaType
            HealthStatus = $PhysDisk.HealthStatus; DriveLetters = $letters
        })
    }
    [pscustomobject]@{ Ok = $true; Disks = @($list) }
}
# Every BitLocker volume on the machine, with or without a drive letter.
function Get-VitalsBitLocker {
    try { [pscustomobject]@{ Ok = $true; Volumes = @(Get-BitLockerVolume -ErrorAction Stop) } }
    catch { [pscustomobject]@{ Ok = $false; Volumes = @() } }
}
function Get-VitalsNetworkReport {
    Get-NetAdapter | Where-Object Status -eq "Up" | ForEach-Object {
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
}
function Get-VitalsLocalAdmins {
    ([ADSI]"WinNT://./Administrators,group").psbase.Invoke('Members') | ForEach-Object {
        $MemberName = $_.GetType().InvokeMember('Name',    'GetProperty', $null, $_, $null)
        $AdsPath    = $_.GetType().InvokeMember('AdsPath', 'GetProperty', $null, $_, $null)
        $Domain     = if ($AdsPath -match "WinNT://([^/]+)/[^/]+$") { $Matches[1] } else { $env:COMPUTERNAME }
        "$Domain\$MemberName"
    }
}

function Get-SubnetMask ([int]$Prefix) {
    $mask = @(0,0,0,0)
    for ($i=0; $i -lt $Prefix; $i++) { $mask[[Math]::Floor($i/8)] += [Math]::Pow(2, (7 - ($i % 8))) }
    return $mask -join "."
}
function Get-VitalsVmType {
    param($Sys)
    if     ($Sys.Manufacturer -match "VMware")                                             { return "VMware" }
    elseif ($Sys.Manufacturer -match "innotek|Oracle" -or $Sys.Model -match "VirtualBox") { return "VirtualBox" }
    elseif ($Sys.Manufacturer -match "Microsoft" -and $Sys.Model -match "Virtual")        { return "Hyper-V" }
    elseif ($Sys.Manufacturer -match "QEMU" -or $Sys.Model -match "QEMU")                 { return "QEMU/KVM" }
    elseif ($Sys.Manufacturer -match "Xen")                                                { return "Xen" }
    elseif ($Sys.Manufacturer -match "Parallels" -or $Sys.Model -match "Parallels")        { return "Parallels" }
    return $null
}

# One entry per key protector on every volume, so a technician sees all of them. Only a
# recovery-password protector has a password to show.
function Get-VitalsProtectorRows {
    param($Volumes)
    $rows = New-Object System.Collections.ArrayList
    foreach ($Vol in @($Volumes | Where-Object { $null -ne $_ })) {
        $mp = if ($Vol.MountPoint) { "$($Vol.MountPoint)" } else { "(no drive letter)" }
        foreach ($Prot in @($Vol.KeyProtector | Where-Object { $null -ne $_ })) {
            [void]$rows.Add([pscustomobject]@{
                Drive      = $mp
                Protection = "$($Vol.ProtectionStatus)"
                Type       = "$($Prot.KeyProtectorType)"
                ID         = "$($Prot.KeyProtectorId)"
                Key        = $(if ("$($Prot.KeyProtectorType)" -eq 'RecoveryPassword' -and $Prot.RecoveryPassword) { "$($Prot.RecoveryPassword)" } else { $null })
            })
        }
    }
    return @($rows)
}

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

# Header art
$_pfx  = "█  "
$_art1 = "╦  ╦ ╦ ╔╦╗ ╔═╗ ╦   ╔═╗ "
$_art2 = "╚╗╔╝ ║  ║  ╠═╣ ║   ╚═╗ "
$_art3 = " ╚╝  ╩  ╩  ╩ ╩ ╩═╝ ╚═╝ "
$_artW  = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1  = $_art1.PadRight($_artW)
$_art2  = $_art2.PadRight($_artW)
$_art3  = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "VISUAL INTERFACE FOR TECHNICAL ASSET & LOGISTICS SUMMARY"

function Invoke-VitalsMain {
    #region Pre-Flight Checks
    # ============================================================================
    if (-not (Test-VitalsElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        Write-Output "[FAIL] VITALS not elevated, nothing was collected (exit 2)"
        Exit-KillerScript 2
    }
    $Problems = New-Object System.Collections.ArrayList
    Initialize-VitalsConsole
    $OS  = Get-WMI Win32_OperatingSystem
    $Sys = Get-WMI Win32_ComputerSystem
    $VMType = Get-VitalsVmType $Sys
    $LastBootRaw = $OS.LastBootUpTime
    $LastBoot = if ($LastBootRaw -is [DateTime]) { $LastBootRaw } else { [Management.ManagementDateTimeConverter]::ToDateTime($LastBootRaw) }
    $Uptime = (Get-Date) - $LastBoot
    $RebootPending = Test-VitalsRebootPending
    $UptimeColor = if ($Uptime.Days -gt 7) { "Red" } elseif ($Uptime.Days -gt 3) { "Yellow" } else { "Green" }
    $PublicIP    = Get-VitalsPublicIp
    $Tpm         = Get-VitalsTpm
    $TPMInfo     = $Tpm.Info
    $TPMVerClean = $Tpm.Version
    $SecureBoot  = Get-VitalsSecureBoot

    # Header
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

    # System Info
    Write-Host "   Device Name  : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $script:ValueFG
    Write-Host "   Uptime       : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($Uptime.Days)d $($Uptime.Hours)h $($Uptime.Minutes)m" -ForegroundColor $UptimeColor
    Write-Host "   Last Boot    : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host "$($LastBoot.ToString('yyyy-MM-dd HH:mm'))" -ForegroundColor $script:ValueFG
    Write-Host "   Public IP    : " -NoNewline -ForegroundColor $script:LabelFG; if ($PublicIP) { Write-Host $PublicIP -ForegroundColor $script:ValueFG } else { Write-Host "Offline" -ForegroundColor Red }
    $IsDomain    = $Sys.DomainRole -in @(1,3,4,5)
    $IsEntra     = Get-VitalsIsEntra
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
    $WinVer = Get-VitalsWinVer

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
    $BL = Get-VitalsBitLocker
    $DiskInfo = Get-VitalsDiskInfo
    if (-not $DiskInfo.Ok) { [void]$Problems.Add("the physical disks could not be listed") }
    if (-not $BL.Ok) { [void]$Problems.Add("BitLocker volumes could not be read") }
    $DiskReport = foreach ($PhysDisk in @($DiskInfo.Disks)) {
        $BLStatus = "Off"
        if ($PhysDisk.DriveLetters.Count -gt 0 -and -not $BL.Ok) { $BLStatus = "N/A" }
        elseif ($BL.Ok) {
            foreach ($Letter in $PhysDisk.DriveLetters) {
                $Vol = @($BL.Volumes | Where-Object { "$($_.MountPoint)" -eq $Letter }) | Select-Object -First 1
                if ($Vol -and "$($Vol.ProtectionStatus)" -eq 'On') { $BLStatus = "On" }
            }
        }
        [PSCustomObject]@{
            "Drive"     = if ($PhysDisk.DriveLetters.Count -gt 0) { $PhysDisk.DriveLetters -join ", " } else { "None" }
            FriendlyName = $PhysDisk.FriendlyName
            "Size_GB"   = [Math]::Round($PhysDisk.Size / 1GB, 2)
            MediaType    = $PhysDisk.MediaType
            "BitLocker"  = $BLStatus
            "Status"    = $PhysDisk.HealthStatus
        }
    }
    Write-Host "`nDrive   FriendlyName         Size_GB  MediaType    BitLocker  Status" -ForegroundColor $script:PrefixFG
    foreach ($D in (@($DiskReport) | Sort-Object Drive)) {
        $FormatStr = "{0,-7} {1,-20} {2,-8} {3,-12} {4,-10} {5}"
        Write-Host ($FormatStr -f $D.Drive, ("$($D.FriendlyName)".PadRight(20).Substring(0,20)), $D.Size_GB, $D.MediaType, $D.BitLocker, $D.Status) -ForegroundColor $script:ValueFG
    }
    # Every protector on every volume, grouped by volume, with its ID.
    $ProtectorRows = @(Get-VitalsProtectorRows $BL.Volumes)
    if ($ProtectorRows.Count -gt 0) {
        Write-Host ""
        Write-Host "BitLocker Protectors" -ForegroundColor $script:PrefixFG
        $lastDrive = $null
        foreach ($K in $ProtectorRows) {
            if ($K.Drive -ne $lastDrive) {
                Write-Host $K.Drive -NoNewline -ForegroundColor Yellow; Write-Host "  (Protection $($K.Protection))" -ForegroundColor $script:LabelFG
                $lastDrive = $K.Drive
            }
            Write-Host "  Type: " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $K.Type -NoNewline -ForegroundColor $script:ValueFG
            Write-Host "  ID: " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $K.ID -ForegroundColor $script:ValueFG
            if ($K.Key) { Write-Host "    Key : " -NoNewline -ForegroundColor $script:LabelFG; Write-Host $K.Key -ForegroundColor Yellow }
        }
    }
    #endregion

    #region 5. Network Configuration
    # ============================================================================
    $NetworkReport = @(Get-VitalsNetworkReport)
    Write-Host "`nInterface            IPAddress       Subnet          Gateway         DNS Servers" -ForegroundColor $script:PrefixFG
    foreach ($N in $NetworkReport) {
        $FormatStr = "{0,-20} {1,-15} {2,-15} {3,-15} {4}"
        Write-Host ($FormatStr -f ($N.Interface.PadRight(20).Substring(0,20)), $N.IPAddress, $N.Subnet, $N.Gateway, @($N.DNS)[0]) -ForegroundColor $script:ValueFG
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
        $AdminMembers = @(Get-VitalsLocalAdmins)
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

    if ($Problems.Count -gt 0) {
        Write-Output "[WARN] VITALS report produced, but $($Problems -join ' and ') (exit 3)"
        Exit-KillerScript 3
    }
    Write-Output "[PASS] VITALS report produced (exit 0)"
    Exit-KillerScript 0
}

if (-not $global:KillerScriptsTestMode) { Invoke-VitalsMain }
