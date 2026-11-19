<#
.SYNOPSIS
    Windows Update, Repair, & System Alignment (W.U.R.S.A.) v3.0
    Developed by Steve the Killer | Updated: 2026-10-16
.DESCRIPTION
    Enforces OS patches, OEM driver updates, and Chocolatey third-party app upgrades
    (skips in-use apps). Performs unattended feature upgrades to 26H2, dispatched to a
    detached SYSTEM task so it survives RMM / LiveConnect disconnects: 24H2 and 25H2 take
    the KB5121794 eKB, Windows 10 and older Windows 11 take the ISO. Reboot is deferred to
    the caller.
.NOTES
    Params: -InplaceUpgrade (auto-confirm, unattended/RMM), -No3rdParty, -NoUpgrade.
    Feature-upgrade result is written to C:\Windows\Temp\25H2IPU\ipu_status.txt; poll it.
    Status: REBOOT_REQUIRED (staged, reboot to finish), DISPATCHED/RUNNING/DOWNLOADING/
    REPAIRING/VERIFY_PENDING/WU_HANDOFF (upgrade active), VERIFIED (committed), BLOCK_*/FAILED_*/
    UNEXPECTED (not complete).
    Only REBOOT_REQUIRED should trigger a reboot, always left to the caller.
#>
param(
    [switch]$InplaceUpgrade,   # Auto-confirm the feature upgrade prompt (implies -Update3rdParty)
    [switch]$Update3rdParty,   # Auto-confirm the third-party update prompt (unattended)
    [switch]$No3rdParty,       # Skip Chocolatey / third-party app updates entirely
    [switch]$NoUpgrade         # Skip the feature upgrade check entirely (region 5)
)

$_ver    = "| v3.0"

# Define the latest known Windows release
$LatestVersion = "26H2"
# Build of the target release. CurrentBuild at or above this is current or newer
# (24H2 = 26100, 25H2 = 26200, 26H2 = 26300), so the upgrade is never offered as a downgrade.
$LatestBuild   = 26300

$script:Width  = 95
$LineCol   = "DarkCyan"
$MainCol   = "Cyan"
$WarnCol   = "DarkYellow"
$BorderCol = "Red"
$ArtCol    = "White"
$AccentCol = "Yellow"
$DimCol    = "DarkGray"
$OkCol     = "Green"
$InfoCol   = "Cyan"

#region Exit and system boundaries
# Every exit goes through Exit-KillerScript. In test mode it throws instead of exiting.
# Exit codes: 0 success, 1 failure, 2 not elevated or precondition failed, 3 partial success,
# 3010 success with a reboot required. The final status line always agrees with the code.
function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# Each function below wraps one system boundary so tests can replace it.
function Test-WursaElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}

function Invoke-WursaNative {
    # Runs a native tool with a real argument array and returns its exit code and output lines.
    param([string]$FilePath, [string[]]$Arguments)
    $out = & $FilePath @Arguments 2>&1
    [pscustomobject]@{ ExitCode = $LASTEXITCODE; Output = @($out | ForEach-Object { [string]$_ }) }
}

function Invoke-WursaChoco {
    param([string[]]$Arguments)
    Invoke-WursaNative -FilePath 'choco' -Arguments $Arguments
}

function New-WursaUpdateSession { New-Object -ComObject Microsoft.Update.Session }
function New-WursaUpdateColl    { New-Object -ComObject Microsoft.Update.UpdateColl }

function Register-WursaProductService {
    $sm = New-Object -ComObject Microsoft.Update.ServiceManager
    $sm.AddService2("7971f918-a847-4430-9279-4a52d1efe18d", 7, "") | Out-Null
}

function Get-WursaPowerLineStatus {
    # 'Online', 'Offline' or 'Unknown' from the OS power line state, or $null when it cannot be read.
    try {
        Add-Type -AssemblyName System.Windows.Forms -ErrorAction Stop
        return [string][System.Windows.Forms.SystemInformation]::PowerStatus.PowerLineStatus
    } catch { return $null }
}

function Save-WursaFile {
    param([string]$Uri, [string]$OutFile)
    (New-Object System.Net.WebClient).DownloadFile($Uri, $OutFile)
}

function Expand-WursaZip {
    param([string]$ZipPath, [string]$Destination)
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ZipPath, $Destination)
}

function Invoke-WursaScriptFile {
    param([string]$Path)
    & $Path *>&1 | Out-Null
}

function Invoke-WursaRestorePoint {
    # Creates a restore point in a timeboxed child job. Returns @{ Result = OK|ERR|TIMEOUT; Message }.
    param([int]$TimeoutSeconds = 90)
    # Bypass the 24-hour cooldown Windows enforces between restore points
    Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SystemRestore" -Name "SystemRestorePointCreationFrequency" -Value 0 -Type DWord -Force -ErrorAction SilentlyContinue
    # The checkpoint calls into VSS/SrClient, which can stall or hard-fault the host
    # process under RMM and take the whole run down before try/catch can fire. Run it
    # in a timeboxed child job so a VSS hang can't kill WURSA; treat it as non-fatal.
    $desc = "WURSA Pre-Update $($env:COMPUTERNAME) $(Get-Date -Format 'yyyy-MM-dd HH:mm')"
    $job = Start-Job -ScriptBlock {
        param($Desc)
        try {
            $ProgressPreference = 'SilentlyContinue'
            # Ensure System Protection is enabled on C: (often disabled on managed endpoints)
            Enable-ComputerRestore -Drive "C:\" -ErrorAction SilentlyContinue *>$null
            # Call the SystemRestore WMI method directly. Checkpoint-Computer wraps the
            # same call but also emits a native progress bar that ignores $ProgressPreference
            # (it comes from SrClient/VSS, not Write-Progress) and paints past the shell
            # width. The WMI call is silent. 12 = MODIFY_SETTINGS, 100 = BEGIN_SYSTEM_CHANGE.
            $sr = [wmiclass]"\\.\root\default:SystemRestore"
            $r  = $sr.CreateRestorePoint($Desc, 12, 100)
            if ($r.ReturnValue -eq 0) { "OK" } else { "ERR: CreateRestorePoint returned $($r.ReturnValue)" }
        } catch {
            "ERR: $($_.Exception.Message)"
        }
    } -ArgumentList $desc
    $result = $null
    if (Wait-Job -Job $job -Timeout $TimeoutSeconds) {
        $text = [string](Receive-Job -Job $job)
        if ($text -match '^OK') { $result = @{ Result = 'OK'; Message = '' } }
        else { $result = @{ Result = 'ERR'; Message = ($text -replace '^ERR: ', '') } }
    } else {
        Stop-Job -Job $job -ErrorAction SilentlyContinue
        $result = @{ Result = 'TIMEOUT'; Message = "timed out after ${TimeoutSeconds}s (VSS stalled)" }
    }
    Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
    $result
}
#endregion

#region UI helpers
$script:StepRow  = -1
$script:StepMsg  = ""
$script:StepOpen = $false

# RMM output panes have no cursor. Every cursor call degrades to plain sequential output.
function Get-WursaCursorTop {
    try { return [int][Console]::CursorTop } catch { return -1 }
}
function Set-WursaCursor {
    param([int]$Left, [int]$Top)
    try { [Console]::SetCursorPosition($Left, $Top); return $true } catch { return $false }
}

function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo)
    if ($Success) {
        if ($script:StepOpen) {
            $combined = if ($CustomInfo) { "$($script:StepMsg) $CustomInfo" } else { $script:StepMsg }
            $_pad = " " * [math]::Max(1, $script:Width - $combined.Length - "[SUCCESS]".Length)
            $savedTop = Get-WursaCursorTop
            if ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop -and (Set-WursaCursor 0 $script:StepRow)) {
                Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
                if ($CustomInfo) { Write-Host " $CustomInfo" -ForegroundColor $WarnCol -NoNewline }
                Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
                [void](Set-WursaCursor 0 $savedTop)
            } elseif ($script:StepRow -lt 0) {
                Write-Host "$combined$_pad[SUCCESS]" -ForegroundColor $OkCol
            }
        }
        $script:StepRow  = -1
        $script:StepOpen = $false
    } else {
        $script:StepRow  = Get-WursaCursorTop
        $script:StepMsg  = $Message
        $script:StepOpen = $true
        Write-Host $Message -ForegroundColor $InfoCol
    }
}

function Write-SubResult {
    param([string]$Tag, [string]$Color)
    $left = -1
    try { $left = [int][Console]::CursorLeft } catch {}
    $_pad = if ($left -ge 0) { " " * [math]::Max(1, $script:Width - $left - $Tag.Length) } else { "  " }
    Write-Host "$_pad$Tag" -ForegroundColor $Color
}

function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "-" * $Width
    }
    $colors = @(
        [ConsoleColor]$BorderCol,
        [ConsoleColor]$ArtCol,
        [ConsoleColor]$MainCol,
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

#region Windows Update pass (discovery and installation)
function Get-WursaRebootPending {
    # A pending reboot from prior updates means Windows won't commit any further
    # updates until the boot happens.
    foreach ($rp in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired',
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\PostRebootReporting'
    )) { if (Test-Path $rp) { return $true } }
    $pfro = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
    if ($pfro -and $pfro.PendingFileRenameOperations) { return $true }
    return $false
}

function Install-WursaUpdate {
    # Downloads and installs one update. Success is decided by ResultCode (2=Succeeded,
    # 3=SucceededWithErrors) and HResult (0=clean; 0x80242014 = WU_E_UH_POSTREBOOTREQUIRED).
    param($Session, $Update)
    $coll = New-WursaUpdateColl
    $coll.Add($Update) | Out-Null
    $Downloader = $Session.CreateUpdateDownloader()
    $Downloader.Updates = $coll
    $null = $Downloader.Download()
    $Installer = $Session.CreateUpdateInstaller()
    $Installer.Updates = $coll
    $Installer.AllowSourcePrompts = $false
    $r  = $Installer.Install()
    $rc = $r.ResultCode
    $hr = $r.HResult
    $res = [pscustomobject]@{ Tag = ''; Color = $OkCol; Outcome = 'Success'; RebootRequired = $false }
    if ($rc -eq 2 -and $hr -eq 0) {
        $res.Tag = "[SUCCESS]"
        if ($r.RebootRequired) { $res.Tag = "[REBOOT REQ]"; $res.Color = $WarnCol; $res.Outcome = 'Reboot'; $res.RebootRequired = $true }
    } elseif ($hr -eq -2145116140 -or $hr -eq 0x80242014) {   # WU_E_UH_POSTREBOOTREQUIRED
        $res.Tag = "[REBOOT REQ]"; $res.Color = $WarnCol; $res.Outcome = 'Reboot'; $res.RebootRequired = $true
    } elseif ($rc -eq 3) {
        $res.Tag = "[SUCCESS*]"; $res.Color = $WarnCol; $res.Outcome = 'Warn'   # succeeded with errors
    } else {
        $res.Tag = "[FAILED 0x$('{0:X8}' -f $hr)]"; $res.Color = "Red"; $res.Outcome = 'Failed'
    }
    $res
}

function Write-WursaUpdateLines {
    # Appends the result tag to the last wrapped title line. Without cursor control the
    # tag is printed on its own line instead.
    param([string[]]$Lines, [int]$Row, [string]$Tag, [string]$Color)
    $savedTop = Get-WursaCursorTop
    if ($Row -ge 0 -and $savedTop -ge 0 -and (Set-WursaCursor 0 $Row)) {
        for ($li = 0; $li -lt $Lines.Count; $li++) {
            if ($li -eq $Lines.Count - 1) {
                $pad = " " * [math]::Max(1, $script:Width - $Lines[$li].Length - $Tag.Length)
                Write-Host "$($Lines[$li])$pad" -NoNewline -ForegroundColor $DimCol
                Write-Host $Tag -ForegroundColor $Color
            } else {
                Write-Host $Lines[$li] -ForegroundColor $DimCol
            }
        }
        [void](Set-WursaCursor 0 $savedTop)
    } else {
        Write-Host "        $Tag" -ForegroundColor $Color
    }
}

function Invoke-WursaWindowsUpdatePass {
    # Regions 2 and 3: discovery, then one-at-a-time installation. Returns a result object.
    $res = [pscustomobject]@{
        Status = 'UpToDate'; Found = 0; Succeeded = 0; Failed = 0
        RebootRequired = $false; Error = $null; Searcher = $null; FeatureOffered = $false
    }
    Write-StepUpdate "[3/5] Scanning for Drivers & OS Patches..."
    # Detect a pending reboot, skip cleanly to the rest of the script, and let the caller
    # decide when to reboot. Feature-upgrade check still runs.
    if (Get-WursaRebootPending) {
        Write-StepUpdate -Success -CustomInfo "(Deferred - Reboot Pending)"
        Write-Host "[4/5] Windows Update Installation..." -NoNewline -ForegroundColor $DimCol
        Write-SubResult "[DEFERRED]" $WarnCol
        $res.Status = 'Deferred'
        return $res
    }
    # Resume Windows Update if it is paused. While paused, WU returns nothing, so an
    # update run would silently find zero updates and the feature upgrade is never
    # offered. Clears the local "Pause updates" values (UX\Settings); leaves any
    # policy-driven pause alone.
    $_uxk = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
    if (Test-Path $_uxk) {
        $_wasPaused = $false
        foreach ($_pn in 'PauseUpdatesExpiryTime','PauseUpdatesStartTime','PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime','PauseQualityUpdatesStartTime','PauseQualityUpdatesEndTime') {
            if ($null -ne (Get-ItemProperty -Path $_uxk -Name $_pn -ErrorAction SilentlyContinue)) {
                Remove-ItemProperty -Path $_uxk -Name $_pn -ErrorAction SilentlyContinue
                $_wasPaused = $true
            }
        }
        if ($_wasPaused) { Write-Host "      [i] Windows Update was paused - resumed." -ForegroundColor $WarnCol }
    }
    # Direct API search to bypass the 'remoteIpNoProxy' crash
    $UpdateSession = New-WursaUpdateSession
    $UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
    $res.Searcher = $UpdateSearcher
    $UpdateList = $null
    try {
        # IsPresent=0 excludes updates that are already staged on disk and only pending
        # a reboot to commit, so the same KBs are not re-staged on every run.
        $s1 = $UpdateSearcher.Search("IsInstalled=0 and IsPresent=0 and Type='Software' and IsHidden=0").Updates
        $s2 = $UpdateSearcher.Search("IsInstalled=0 and IsPresent=0 and Type='Driver' and IsHidden=0").Updates
        $UpdateList = New-WursaUpdateColl
        foreach ($u in $s1) {
            # Do not install a feature update through the generic WU loop. A download or
            # pre-download is not a completed upgrade. Region 5 handles the feature upgrade.
            if ($u.Title -match [regex]::Escape($LatestVersion) -or $u.Title -match 'Feature update to Windows') {
                $res.FeatureOffered = $true
                continue
            }
            $UpdateList.Add($u) | Out-Null
        }
        foreach ($u in $s2) { $UpdateList.Add($u) | Out-Null }
        if ($res.FeatureOffered) {
            Write-Host "      [i] $LatestVersion detected; reserved for the upgrade path." -ForegroundColor $DimCol
        }
    } catch {
        Write-StepUpdate -Success -CustomInfo "(Scan Error)"
        Write-Host "      [!] WU search failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "[4/5] Windows Update Installation..." -NoNewline -ForegroundColor $DimCol
        Write-SubResult "[FAILED]" "Red"
        $res.Status = 'ScanFailed'
        $res.Error = $_.Exception.Message
        $res.Failed++
        return $res
    }
    if (-not $UpdateList -or $UpdateList.Count -eq 0) {
        Write-StepUpdate -Success -CustomInfo "(System Up to Date)"
        Write-Host "[4/5] Windows Update Installation..." -NoNewline -ForegroundColor $DimCol
        Write-SubResult "[SKIPPED]" $WarnCol
        return $res
    }
    $res.Found = $UpdateList.Count
    Write-StepUpdate -Success -CustomInfo "($($UpdateList.Count) Found)"
    Write-StepUpdate "[4/5] Installing Windows Updates..."
    foreach ($Update in $UpdateList) {
        # Word-wrap the title at word boundaries so long names don't blow the
        # cursor-reposition math. All lines use the same indent; the tag is appended
        # to the last line only.
        $_uPrefix = "      > Deploying: "
        $_uIndent = " " * $_uPrefix.Length
        $_uAvail  = $script:Width - "[SUCCESS]".Length - $_uPrefix.Length
        $_uLines  = @()
        $_rem     = $Update.Title
        $_first   = $true
        while ($_rem.Length -gt 0) {
            $_pfx = if ($_first) { $_uPrefix } else { $_uIndent }
            if ($_rem.Length -le $_uAvail) {
                $_uLines += "$_pfx$_rem"
                break
            }
            $_chunk = $_rem.Substring(0, $_uAvail)
            $_split = $_chunk.LastIndexOf(' ')
            if ($_split -le 0) { $_split = $_uAvail }   # no space found - hard break
            $_uLines += "$_pfx$($_rem.Substring(0, $_split).TrimEnd())"
            $_rem    = $_rem.Substring($_split).TrimStart()
            $_first  = $false
        }
        $_uRow = Get-WursaCursorTop
        foreach ($_line in $_uLines) { Write-Host $_line -ForegroundColor $InfoCol }
        try {
            $r = Install-WursaUpdate -Session $UpdateSession -Update $Update
            if ($r.RebootRequired) { $res.RebootRequired = $true }
            if ($r.Outcome -eq 'Failed') { $res.Failed++ } else { $res.Succeeded++ }
            Write-WursaUpdateLines -Lines $_uLines -Row $_uRow -Tag $r.Tag -Color $r.Color
        } catch {
            $res.Failed++
            Write-WursaUpdateLines -Lines $_uLines -Row $_uRow -Tag "[FAILED]" -Color "Red"
        }
    }
    Write-Progress -Activity "W.U.R.S.A.: Deploying Updates" -Completed
    Write-StepUpdate "[4/5] Windows Updates" -Success
    $res.Status = 'Installed'
    $res
}
#endregion

#region Third-party software updates
$script:WursaChocoPublisher = '^CN=Chocolatey Software(,|$)'
function Get-WursaThirdPartyCatalog {
    # Install detection uses Chocolatey's own local package DB rather than probing
    # filesystem paths. Id is the bare package id. Args are extra choco arguments, one
    # array element each, so nothing is ever split or re-parsed from a packed string.
    @(
        @{ Name = "Google Chrome"; Id = "googlechrome"; Args = @("--install-arguments='--system-level'", "--ignore-checksums"); Process = "chrome" },
        @{ Name = "Microsoft Edge"; Id = "microsoft-edge"; Args = @(); Process = "msedge" },
        @{ Name = "Mozilla Firefox"; Id = "firefox"; Args = @(); Process = "firefox" },
        @{ Name = "Brave"; Id = "brave"; Args = @(); Process = "brave" },
        @{ Name = "Adobe Acrobat"; Id = "adobereader"; Args = @(); Process = @("Acrobat","AcroRd32") },
        @{ Name = "Zoom"; Id = "zoom"; Args = @("--install-arguments='/quiet /norestart'"); Process = "Zoom" },
        @{ Name = "Microsoft Teams"; Id = "microsoft-teams"; Args = @(); Process = @("Teams","ms-teams") },
        @{ Name = "Webex"; Id = "webex"; Args = @(); Process = "CiscoCollabHost" },
        @{ Name = "Slack"; Id = "slack"; Args = @(); Process = "slack" },
        @{ Name = "RingCentral"; Id = "ringcentral"; Args = @(); Process = "RingCentral" },
        @{ Name = "Notepad++"; Id = "notepadplusplus"; Args = @(); Process = "notepad++" },
        @{ Name = "VLC"; Id = "vlc"; Args = @(); Process = "vlc" },
        @{ Name = "7-Zip"; Id = "7zip"; Args = @(); Process = "7zFM" },
        @{ Name = "KillerPDF"; Id = "killerpdf"; Args = @(); Process = "KillerPDF" },
        @{ Name = "KillerNotes"; Id = "killernotes"; Args = @(); Process = "KillerNotes" },
        @{ Name = "KillerScan"; Id = "killerscan"; Args = @(); Process = "KillerScan" },
        @{ Name = "KillerShell"; Id = "killershell"; Args = @(); Process = "KillerShell" },
        @{ Name = "Killendar"; Id = "killendar"; Args = @(); Process = "Killendar" },
        @{ Name = "KillerMCP"; Id = "killermcp"; Args = @(); Process = @("KillerMCP", "KillerMCP.Configure", "KillerMCP.Notify") },
        @{ Name = "KillerShot"; Id = "killershot"; Args = @(); Process = @("KillerShot", "KillerShot.PrintScreen") },
        @{ Name = "KillerBench"; Id = "killerbench"; Args = @(); Process = "KillerBench" }
    )
}

function New-WursaSecureTempDir {
    # A fresh, unpredictably named directory with a protected admin-only ACL for downloads that will be
    # run as an administrator. Returns the path, or $null when a trusted directory cannot be had.
    param([string]$Base = (Join-Path $env:ProgramData 'WURSA'))
    $path = Join-Path $Base ([guid]::NewGuid().ToString('N'))
    $r = Initialize-WursaWorkDir -Path $path
    if ($r -and $r.Ready) { return $path }
    return $null
}

function Install-WursaChocolatey {
    # Returns $true when choco is available afterwards. The package is fetched into a private directory and
    # chocolateyInstall.ps1 runs only when the extracted choco.exe carries a valid Chocolatey signature.
    $work = $null
    try {
        Set-ExecutionPolicy Bypass -Scope Process -Force
        [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072
        $work = New-WursaSecureTempDir
        if (-not $work) {
            Write-Host "      [!] Could not create a protected download folder - skipping third-party updates." -ForegroundColor Red
            return $false
        }
        $ChocoInstallDir = "$env:ProgramData\chocolatey"
        $ZipPath = Join-Path $work 'chocolatey.zip'
        $ExtractPath = Join-Path $work 'extracted'
        Save-WursaFile -Uri "https://community.chocolatey.org/api/v2/package/chocolatey" -OutFile $ZipPath
        # Extract using pure .NET - bypasses Microsoft.PowerShell.Archive entirely
        Expand-WursaZip -ZipPath $ZipPath -Destination $ExtractPath
        $exe = Get-ChildItem -LiteralPath $ExtractPath -Recurse -Filter 'choco.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
        if (-not $exe -or -not (Test-WursaSignedBinary -Path $exe.FullName -PublisherPattern $script:WursaChocoPublisher)) {
            Write-Host "      [!] The downloaded Chocolatey package is not signed by Chocolatey Software - not running its installer." -ForegroundColor Red
            return $false
        }
        $ChocoInstallScript = Get-ChildItem -LiteralPath $ExtractPath -Recurse -Filter 'chocolateyInstall.ps1' -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($ChocoInstallScript) {
            $env:ChocolateyInstall = $ChocoInstallDir
            Invoke-WursaScriptFile -Path $ChocoInstallScript.FullName
        }
        # Refresh PATH
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        if (Get-Command choco -ErrorAction SilentlyContinue) {
            Write-Host "      Chocolatey installed successfully." -ForegroundColor Green
            return $true
        }
        Write-Host "      [!] Chocolatey install failed - skipping third-party updates." -ForegroundColor Red
    } catch {
        Write-Host "      [!] Chocolatey install failed: $($_.Exception.Message)" -ForegroundColor Red
    } finally {
        if ($work) { Remove-Item -LiteralPath $work -Recurse -Force -ErrorAction SilentlyContinue }
    }
    return $false
}

function Get-WursaChocoMajor {
    # Major version of the installed choco, or $null when it cannot be read.
    try {
        $r = Invoke-WursaChoco -Arguments @('--version')
        if ($r.ExitCode -ne 0) { return $null }
        foreach ($line in $r.Output) {
            if ($line -match '^\s*(\d+)\.\d+') { return [int]$Matches[1] }
        }
    } catch {}
    return $null
}

function Get-WursaChocoLocalInventory {
    # Local package ids installed through Chocolatey, as a hashtable of lowercase ids, or $null when the
    # inventory cannot be trusted. Chocolatey 1.x lists the remote feed unless --local-only is given;
    # 2.x and later list local packages only and no longer take that switch.
    $major = Get-WursaChocoMajor
    if ($null -eq $major) { return $null }
    $listArgs = if ($major -lt 2) { @('list', '--local-only', '--limit-output') } else { @('list', '--limit-output') }
    $inv = @{}
    try {
        $r = Invoke-WursaChoco -Arguments $listArgs
        if ($r.ExitCode -ne 0) { return $null }
        foreach ($line in $r.Output) {
            if ($line -match '^([^|]+)\|') { $inv[$Matches[1].ToLower()] = $true }
        }
    } catch { return $null }
    $inv
}

function Get-WursaChocoOutcome {
    # Classifies a choco exit code: Success, RebootRequired (1641 and 3010 are success plus a restart), or Failed.
    param([int]$ExitCode)
    if ($ExitCode -eq 0) { return 'Success' }
    if ($ExitCode -eq 1641 -or $ExitCode -eq 3010) { return 'RebootRequired' }
    'Failed'
}

function Invoke-WursaThirdPartyUpdate {
    # Region 4. Gated behind a Y/N prompt. -No3rdParty skips outright, -Update3rdParty and
    # -InplaceUpgrade auto-confirm (unattended/RMM), and a non-interactive console defaults to skip.
    param([bool]$Skip, [bool]$AutoConfirm, [string]$AutoFlag)
    $res = [pscustomobject]@{ Status = 'Skipped'; Succeeded = 0; Failed = 0; RebootRequired = $false }
    $ThirdParty = Get-WursaThirdPartyCatalog
    $run3rdParty = $false
    if ($Skip) {
        Write-Host "[i] Third-party updates skipped (-No3rdParty)." -ForegroundColor $DimCol
    } elseif ($AutoConfirm) {
        Write-Host "      [$AutoFlag] Auto-confirming third-party updates." -ForegroundColor $DimCol
        $run3rdParty = $true
    } else {
        try { while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null } } catch {}
        try {
            Write-Host "Would you like to update installed third-party software via Chocolatey? (Y/N): " -NoNewline -ForegroundColor $WarnCol
            $key    = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $choice = $key.Character
            if ($choice -ne "`n" -and $choice -ne "`r") { Write-Host $choice -NoNewline }
            Write-Host ""
            if ($choice -in @('Y','y')) { $run3rdParty = $true }
        } catch {
            Write-Host "      [i] Non-interactive - defaulting to NO. Use -Update3rdParty for unattended runs." -ForegroundColor $DimCol
        }
    }
    if (-not $run3rdParty) {
        if (-not $Skip) { $res.Status = 'Declined' }
        return $res
    }

    # Chocolatey install-if-missing runs only after the user (or a switch) has said
    # yes, so a decline is never paid for with a Chocolatey install.
    $ChocoAvailable = Get-Command choco -ErrorAction SilentlyContinue
    if (-not $ChocoAvailable) {
        Write-Host "`n[!] Chocolatey not found - Installing..." -ForegroundColor Yellow
        $null = Install-WursaChocolatey
        $ChocoAvailable = Get-Command choco -ErrorAction SilentlyContinue
    }
    if (-not $ChocoAvailable) { $res.Status = 'ChocoUnavailable'; return $res }

    Write-StepUpdate "[5/5] Updating Installed Third-Party Software..."
    # Update Chocolatey itself first
    Write-Host "      > " -NoNewline -ForegroundColor $DimCol
    Write-Host "Chocolatey: " -NoNewline -ForegroundColor $DimCol
    $self = Invoke-WursaChoco -Arguments @('upgrade', 'chocolatey', '-y', '--no-progress')
    $selfOutcome = Get-WursaChocoOutcome -ExitCode $self.ExitCode
    if ($selfOutcome -eq 'Failed') {
        Write-SubResult "[FAILED]" Red
        $res.Failed++
    } else {
        if ($selfOutcome -eq 'RebootRequired') { $res.RebootRequired = $true }
        $chocoSelfMatch = $self.Output | Select-String -Pattern 'upgraded (\d+)/'
        if ($chocoSelfMatch -and [int]$chocoSelfMatch.Matches[0].Groups[1].Value -gt 0) {
            Write-SubResult "[UPDATED]" Green
            $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        } else {
            Write-SubResult "[ALREADY UPDATED]" Cyan
        }
    }
    # Cache Chocolatey's local package list once. One shell-out per run beats one per app. If it cannot be
    # read, nothing can be told apart from "not installed", so no app is touched and the run counts as partial.
    $ChocoLocal = Get-WursaChocoLocalInventory
    if ($null -eq $ChocoLocal) {
        Write-Host "      [!] Could not read Chocolatey's local package inventory - skipping third-party updates." -ForegroundColor Yellow
        $res.Failed++
        $res.Status = 'InventoryUnverified'
        return $res
    }
    # Cache Add/Remove Programs display names once. If an app is installed outside
    # Chocolatey (MSI, the app's own installer) we still want to upgrade it - the
    # per-app installer's own /silent handler is idempotent, so choco upgrade will
    # refresh the existing install in place rather than orphan it.
    $ArpNames = New-Object System.Collections.Generic.List[string]
    foreach ($_arpRoot in @(
        'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*',
        'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*'
    )) {
        try {
            Get-ItemProperty $_arpRoot -ErrorAction SilentlyContinue |
                Where-Object { $_.DisplayName } |
                ForEach-Object { $ArpNames.Add($_.DisplayName) }
        } catch {}
    }
    foreach ($App in $ThirdParty) {
        $_pkgId = $App.Id.ToLower()
        $_procs = @($App.Process)

        Write-Host "      > " -NoNewline -ForegroundColor Gray
        Write-Host "$($App.Name): " -NoNewline -ForegroundColor White

        # Two ways an app counts as installed on this box: Chocolatey knows about
        # it OR Add/Remove Programs shows it. Either one is enough to run choco
        # upgrade. Nothing in either list means the app is not on this machine and we
        # do not install it fresh.
        $_inChoco = $ChocoLocal.ContainsKey($_pkgId)
        $_inArp   = ($ArpNames | Where-Object { $_ -like "*$($App.Name)*" } | Select-Object -First 1) -ne $null
        if (-not $_inChoco -and -not $_inArp) {
            Write-SubResult "[NOT INSTALLED]" DarkGray
            continue
        }
        $IsRunning = $_procs | ForEach-Object { Get-Process -Name $_ -ErrorAction SilentlyContinue } | Select-Object -First 1
        if ($IsRunning) {
            Write-SubResult "[IN USE - SKIPPED]" Yellow
            continue
        }
        $up = Invoke-WursaChoco -Arguments (@('upgrade', $App.Id) + @($App.Args) + @('-y', '--no-progress'))
        $outcome = Get-WursaChocoOutcome -ExitCode $up.ExitCode
        if ($App.Name -eq "Google Chrome") {
            $GUpdate = "C:\Program Files (x86)\Google\Update\GoogleUpdate.exe"
            if (Test-Path $GUpdate) {
                $null = Invoke-WursaNative -FilePath $GUpdate -Arguments @('/ua', '/installsource', 'scheduler')
            }
        }
        if ($outcome -eq 'Failed') {
            Write-SubResult "[FAILED]" Red
            $res.Failed++
            continue
        }
        if ($outcome -eq 'RebootRequired') { $res.RebootRequired = $true }
        $upgradeMatch = $up.Output | Select-String -Pattern 'upgraded (\d+)/'
        $upgradeCount = if ($upgradeMatch) { [int]$upgradeMatch.Matches[0].Groups[1].Value } else { 0 }
        if ($upgradeCount -gt 0) {
            Write-SubResult "[UPDATED]" Green
            $res.Succeeded++
        } elseif (-not $_inChoco -and $_inArp) {
            Write-SubResult "[ADOPTED]" Green
            $res.Succeeded++
        } else {
            Write-SubResult "[ALREADY UPDATED]" Cyan
        }
    }
    Write-StepUpdate "[5/5] Third-Party Updates" -Success
    $res.Status = 'Done'
    $res
}
#endregion

#region Feature upgrade helpers
# Windows Update policy values. These wrappers are the only code that touches the policy hive, and they are
# copied verbatim into the post-reboot verifier script, so keep them self-contained and plain ASCII.
function Initialize-WursaPolicyKey {
    param([string]$Key)
    if (-not (Test-Path $Key)) { New-Item -Path $Key -Force | Out-Null }
}

function Get-WursaPolicyValue {
    param([string]$Key, [string]$Name)
    $p = Get-ItemProperty -Path $Key -Name $Name -ErrorAction SilentlyContinue
    if ($p) { return $p.$Name }
    return $null
}

function Set-WursaPolicyValue {
    param([string]$Key, [string]$Name, $Value, [string]$Type = 'DWord')
    Initialize-WursaPolicyKey -Key $Key
    Set-ItemProperty -Path $Key -Name $Name -Value $Value -Type $Type -Force
}

function Remove-WursaPolicyValue {
    param([string]$Key, [string]$Name)
    Remove-ItemProperty -Path $Key -Name $Name -ErrorAction SilentlyContinue
}

function Restore-WursaPolicyRecord {
    # Puts every value in the restore record back exactly (or removes it if it did not exist before),
    # reads each one back, and deletes the record only when every value verified. Returns $true when the
    # policy is back as it was, $false when anything did not stick (the record is kept so it can retry).
    param([string]$RecordFile)
    if (-not [System.IO.File]::Exists($RecordFile)) { return $true }
    $polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auk  = $polk + '\AU'
    $map  = @{}
    try {
        foreach ($line in [System.IO.File]::ReadAllLines($RecordFile)) {
            if ($line -match '^([^=]+)=(.+)$') { $map[$Matches[1]] = $Matches[2] }
        }
    } catch { return $false }
    $targets = @{ NoAutoUpdate = $auk; AUOptions = $auk; SetActiveHours = $polk; ActiveHoursStart = $polk; ActiveHoursEnd = $polk; ProductVersion = $polk; TargetReleaseVersion = $polk; TargetReleaseVersionInfo = $polk }
    $allOk = $true
    foreach ($name in @($targets.Keys)) {
        $key = $targets[$name]
        $orig = $map[$name]
        # The release pin is restored only when it was recorded (it is set only on a Windows Update handoff).
        if ($null -eq $orig -and $name -match '^(ProductVersion|TargetReleaseVersion)') { continue }
        $isString = ($name -eq 'ProductVersion' -or $name -eq 'TargetReleaseVersionInfo')
        $absent = ($null -eq $orig -or $orig -eq 'ABSENT')
        try {
            if ($absent) {
                Remove-WursaPolicyValue -Key $key -Name $name
            } elseif ($isString) {
                Set-WursaPolicyValue -Key $key -Name $name -Value $orig -Type 'String'
            } else {
                Set-WursaPolicyValue -Key $key -Name $name -Value ([int]$orig) -Type 'DWord'
            }
            $now = Get-WursaPolicyValue -Key $key -Name $name
            if ($absent) { $good = ($null -eq $now) } else { $good = ($null -ne $now -and ([string]$now -eq [string]$orig)) }
            if (-not $good) { $allOk = $false }
        } catch {
            $allOk = $false
        }
    }
    if ($allOk) {
        try { [System.IO.File]::Delete($RecordFile) } catch { $allOk = $false }
    }
    return $allOk
}

function Restore-WUUpgradeOverrides {
    # Runs in the MAIN script (not the detached runner), when the box is found already on the latest
    # build. Undoes the temporary Windows Update policy changes through the verified restore.
    if (-not (Test-Path $IPU_AURestoreFile)) { return }
    if (Restore-WursaPolicyRecord -RecordFile $IPU_AURestoreFile) {
        Write-Host "      Restored original Windows Update policy values (auto-update, active hours and release pin)." -ForegroundColor $DimCol
    } else {
        Write-Host "      [!] Could not fully restore the Windows Update policy values; the record is kept and the next run will retry." -ForegroundColor $WarnCol
    }
}

function Test-WUFeatureUpdateStaged {
    # Only a feature update that Windows reports as reboot-required is considered staged.
    # Offered, downloading, and pre-downloaded updates do not suppress the deterministic
    # Region 5 upgrade path. This prevents WURSA from declaring a handoff while the box
    # remains on the old feature release.
    param($Searcher)
    try {
        $_res = $Searcher.Search("RebootRequired=1").Updates
        foreach ($_u in $_res) {
            if ($_u.Title -match [regex]::Escape($LatestVersion) -or $_u.Title -match 'Feature update to Windows') { return $true }
        }
    } catch {}
    return $false
}

function Write-WursaStatusFile {
    param([string]$Path, [string]$Text)
    try { [System.IO.File]::WriteAllText($Path, ($Text + "`r`n"), [System.Text.Encoding]::ASCII) } catch {}
}

function Invoke-WursaVerifyRun {
    # The post-reboot check, as a function so the shipped code is the tested code. Returns the exit code.
    # A committed build restores the temporary policy changes and removes the task, but only if the
    # restore verified; otherwise the task and the record stay so the next boot retries.
    param([string]$StatusFile, [string]$RestoreFile, [string]$VerifyTaskName, [int]$TargetBuild, [int]$SettleSeconds = 180)
    if ($SettleSeconds -gt 0) { Start-Sleep -Seconds $SettleSeconds }
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $ver = $cv.DisplayVersion
    $build = $cv.CurrentBuild
    $ubr = $cv.UBR
    $stamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    if ([int]$build -ge $TargetBuild) {
        if (Restore-WursaPolicyRecord -RecordFile $RestoreFile) {
            Write-WursaStatusFile -Path $StatusFile -Text "VERIFIED $ver $build.$ubr $stamp"
            $null = Invoke-WursaNative -FilePath 'schtasks.exe' -Arguments @('/Delete', '/TN', $VerifyTaskName, '/F')
            return 0
        }
        Write-WursaStatusFile -Path $StatusFile -Text "VERIFIED_RESTORE_PENDING $ver $build.$ubr $stamp"
        return 1
    }
    if (Test-Path 'HKLM:\SYSTEM\Setup\MoSetup\Volatile') {
        Write-WursaStatusFile -Path $StatusFile -Text "VERIFY_PENDING VERSION_$ver BUILD_$build.$ubr $stamp"
        return 0
    }
    Write-WursaStatusFile -Path $StatusFile -Text "FAILED_POSTREBOOT VERSION_$ver BUILD_$build.$ubr $stamp"
    $null = Invoke-WursaNative -FilePath 'schtasks.exe' -Arguments @('/Delete', '/TN', $VerifyTaskName, '/F')
    return 1
}

function Get-WursaVerifierScript {
    # Builds a one-shot post-reboot verifier. It runs as SYSTEM at startup and carries the tested
    # functions it needs as source text, so what ships is what the tests ran.
    param([string]$StatusFile, [string]$RestoreFile, [string]$VerifyTaskName, [int]$TargetBuild)
    $names = 'Invoke-WursaNative', 'Initialize-WursaPolicyKey', 'Get-WursaPolicyValue', 'Set-WursaPolicyValue',
             'Remove-WursaPolicyValue', 'Restore-WursaPolicyRecord', 'Write-WursaStatusFile', 'Invoke-WursaVerifyRun'
    $src = ($names | ForEach-Object { Get-WursaFunctionSource -Name $_ }) -join "`r`n"
    "`$ErrorActionPreference = 'SilentlyContinue'`r`n" + $src + "`r`n" +
        "exit (Invoke-WursaVerifyRun -StatusFile '$StatusFile' -RestoreFile '$RestoreFile' -VerifyTaskName '$VerifyTaskName' -TargetBuild $TargetBuild)`r`n"
}

function Get-WursaPowerState {
    # 'AC', 'Battery' or 'Unknown'. The OS power line is trusted first. When it cannot say, the battery
    # status decides: no battery means a desktop on AC; any AC-ish battery means AC (BatteryStatus 2 and 3
    # are AC, 6 to 9 are charging, 11 is partially charged); only batteries that all read discharging
    # (1, 4, 5) mean Battery; anything else is Unknown and the caller must not guess.
    $line = $null
    try { $line = Get-WursaPowerLineStatus } catch {}
    if ($line -eq 'Online')  { return 'AC' }
    if ($line -eq 'Offline') { return 'Battery' }
    try {
        $batteries = @(Get-CimInstance Win32_Battery -ErrorAction Stop)
    } catch {
        return 'Unknown'
    }
    if ($batteries.Count -eq 0) { return 'AC' }
    $ac = 0; $discharging = 0
    foreach ($b in $batteries) {
        switch ([int]$b.BatteryStatus) {
            { $_ -in 2, 3, 6, 7, 8, 9, 11 } { $ac++ }
            { $_ -in 1, 4, 5 } { $discharging++ }
        }
    }
    if ($ac -gt 0) { return 'AC' }
    if ($discharging -gt 0) { return 'Battery' }
    'Unknown'
}
#endregion

#region Staging directory
# The SYSTEM upgrade task and the post-reboot verifier run scripts and media from the staging
# directory, so a standard user must never own or write to it. C:\Windows\Temp lets any user
# create a subfolder, and a folder a user creates there is theirs to rewrite.
$script:WursaTrustedSids = @(
    'S-1-5-18',        # SYSTEM
    'S-1-5-32-544',    # Administrators
    'S-1-5-80-956008885-3418522649-1831038044-1853292631-2271478464'   # TrustedInstaller
)

function Get-WursaOwnerSid {
    param([string]$Path)
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    $acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
}

function Get-WursaDirectoryProblem {
    # Returns $null when the directory is safe to run code from, otherwise the reason it is not.
    param([string]$Path)
    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        if (-not $item.PSIsContainer) { return 'not a directory' }
        if ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { return 'reparse point' }
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        $owner = $acl.GetOwner([System.Security.Principal.SecurityIdentifier]).Value
        if ($script:WursaTrustedSids -notcontains $owner) { return "owner $owner" }
        if (-not $acl.AreAccessRulesProtected) { return 'inherited permissions' }
        foreach ($rule in $acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier])) {
            if ($rule.AccessControlType -eq 'Allow' -and $script:WursaTrustedSids -notcontains $rule.IdentityReference.Value) {
                return "access granted to $($rule.IdentityReference.Value)"
            }
        }
    } catch {
        return "unreadable ($($_.Exception.Message))"
    }
    return $null
}

function Set-WursaProtectedAcl {
    # Removes inherited permissions and leaves SYSTEM and Administrators with full control.
    param([string]$Path)
    $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
    $acl.SetAccessRuleProtection($true, $false)
    # Drop any explicit grant to anyone else as well, in case the directory carried one.
    foreach ($rule in @($acl.GetAccessRules($true, $false, [System.Security.Principal.SecurityIdentifier]))) {
        if ($script:WursaTrustedSids -notcontains $rule.IdentityReference.Value) { $acl.PurgeAccessRules($rule.IdentityReference) }
    }
    foreach ($sid in 'S-1-5-18', 'S-1-5-32-544') {
        $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            ([System.Security.Principal.SecurityIdentifier]$sid), 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($rule)
    }
    Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
}

function Initialize-WursaWorkDir {
    # Returns @{ Ready; Reason; MovedAside }. A usable directory is one we created with a protected
    # admin-only ACL, or an existing one that already meets that bar. Anything else is renamed aside
    # (never run from) and replaced. If that is not possible the caller must refuse to continue.
    param([string]$Path)
    $res = [pscustomobject]@{ Ready = $false; Reason = ''; MovedAside = $null }
    $asideTrusted = $false
    if (Test-Path -LiteralPath $Path) {
        $problem = Get-WursaDirectoryProblem -Path $Path
        if (-not $problem) { $res.Ready = $true; return $res }
        $asideTrusted = ($problem -notmatch '^(owner|reparse|not a directory|unreadable)')
        $aside = "$Path.untrusted-$(Get-Date -Format 'yyyyMMddHHmmss')"
        try {
            Move-Item -LiteralPath $Path -Destination $aside -ErrorAction Stop
        } catch {
            $res.Reason = "existing directory is not trusted ($problem) and could not be moved aside: $($_.Exception.Message)"
            return $res
        }
        $res.MovedAside = $aside
    }
    try {
        # No -Force: creation must fail if something else got there first.
        New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
        Set-WursaProtectedAcl -Path $Path
    } catch {
        $res.Reason = "could not create a protected directory: $($_.Exception.Message)"
        return $res
    }
    $problem = Get-WursaDirectoryProblem -Path $Path
    if ($problem) { $res.Reason = "new directory failed verification ($problem)"; return $res }
    if (@(Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue).Count -ne 0) {
        $res.Reason = 'new directory was not empty'
        return $res
    }
    # Carry the policy restore record across, but only if an admin or SYSTEM wrote it.
    if ($res.MovedAside -and $asideTrusted) {
        $old = Join-Path $res.MovedAside 'au_restore.txt'
        try {
            if ((Test-Path -LiteralPath $old) -and ($script:WursaTrustedSids -contains (Get-WursaOwnerSid -Path $old))) {
                Copy-Item -LiteralPath $old -Destination (Join-Path $Path 'au_restore.txt') -ErrorAction Stop
            }
        } catch {}
    }
    $res.Ready = $true
    $res
}

function Confirm-WursaWorkDir {
    # Once per run: make sure the staging directory is safe, tell the technician what happened.
    param([string]$Path)
    if ($script:WorkDirChecked -and $script:WorkDirChecked.Path -eq $Path) { return $script:WorkDirChecked.Ready }
    $r = Initialize-WursaWorkDir -Path $Path
    if ($r.MovedAside) {
        Write-Host "      [!] $Path was not trusted and was moved to $($r.MovedAside). Review and delete it." -ForegroundColor $WarnCol
    }
    if (-not $r.Ready) {
        Write-Host "      [!] Staging directory refused: $($r.Reason)" -ForegroundColor Red
    }
    $script:WorkDirChecked = [pscustomobject]@{ Path = $Path; Ready = [bool]$r.Ready }
    [bool]$r.Ready
}
#endregion

#region Trust checks and the generated runner
function Test-WursaSignedBinary {
    # True only for an existing file with a valid Authenticode signature from the expected publisher.
    # A pinned SHA-256 is optional, but when given it must match as well. This function is embedded
    # in the generated SYSTEM runner, so keep it self-contained and ASCII.
    param([string]$Path, [string]$PublisherPattern, [string]$Sha256)
    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $sig = Get-AuthenticodeSignature -LiteralPath $Path -ErrorAction Stop
        if ([string]$sig.Status -ne 'Valid') { return $false }
        if (-not $sig.SignerCertificate) { return $false }
        if ($sig.SignerCertificate.Subject -notmatch $PublisherPattern) { return $false }
        if ($Sha256) {
            if ((Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash -ne $Sha256) { return $false }
        }
        return $true
    } catch {
        return $false
    }
}

function Get-WursaFunctionSource {
    # Source text of a function defined in this script, ready to paste into a generated script.
    param([string]$Name)
    $sb = (Get-Command -Name $Name -CommandType Function).ScriptBlock
    "function $Name {" + [Environment]::NewLine + $sb.ToString() + [Environment]::NewLine + "}"
}

function New-WursaRunnerScript {
    # Config preamble (values injected by the caller), the shared functions, then the runner body.
    param([string]$Config)
    $shared = (Get-WursaFunctionSource -Name 'Test-WursaSignedBinary')
    $Config + "`r`n" + $shared + "`r`n" + (Get-WursaRunnerBody)
}

function Get-WursaRunnerBody {
    # Runner body (single-quoted here-string: written verbatim, runs later).
    @'
$ErrorActionPreference = "Continue"
function Set-Status { param([string]$Text) try { [System.IO.File]::WriteAllText($IPU_StatusFile, ($Text + "`r`n"), [System.Text.Encoding]::ASCII) } catch {} }
$IPU_IAExe  = Join-Path $IPU_WorkDir "Windows11InstallationAssistant.exe"
$IPU_IAArgs = "/QuietInstall /SkipEULA /auto upgrade /NoReboot"
function Register-IPUPostRebootVerifier {
    $verifyCmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$IPU_VerifierPath`""
    $verifyOut = schtasks.exe /Create /TN $IPU_VerifyTaskName /TR $verifyCmd /SC ONSTART /RU SYSTEM /RL HIGHEST /F 2>&1
    if ($LASTEXITCODE -eq 0) {
        Write-Output "Post-reboot verifier registered as $IPU_VerifyTaskName."
        return $true
    }
    Write-Output "WARN: post-reboot verifier could not be registered: $verifyOut"
    return $false
}
function Set-RebootRequired {
    # A reboot is needed. Record whether the post-reboot verifier (which restores the temporary policy
    # changes) is armed. If it could not be registered the state says so, and the final report tells the
    # technician to run WURSA again after the reboot.
    param([string]$Detail)
    if (Register-IPUPostRebootVerifier) { Set-Status "REBOOT_REQUIRED $Detail" }
    else { Set-Status "REBOOT_REQUIRED_UNARMED $Detail" }
}
function Invoke-IPUInstallationAssistant {
    param([string]$Reason)
    Write-Output "Routing to Windows 11 Installation Assistant ($Reason)."
    $iaStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    # Suspend BitLocker for the reboots the Assistant will drive
    try {
        $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($bl -and $bl.ProtectionStatus -eq 'On') { Suspend-BitLocker -MountPoint "C:" -RebootCount 4 -ErrorAction Stop; Write-Output "BitLocker suspended." }
    } catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }
    if (Get-Process -Name 'Windows11InstallationAssistant' -ErrorAction SilentlyContinue) {
        Write-Output "Installation Assistant already running; leaving it to finish."
        Set-Status "RUNNING IA_INPROGRESS $iaStamp"
        return
    }
    if (-not (Test-Path $IPU_IAExe)) {
        try {
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $IPU_IAUrl -OutFile $IPU_IAExe -UseBasicParsing
        } catch {
            Write-Output "Installation Assistant download failed: $_"
            Invoke-IPUWindowsUpdate -Reason "Assistant download failed"
            return
        }
    }
    # Verify every copy, downloaded now or already present, before it runs as SYSTEM.
    if (-not (Test-WursaSignedBinary -Path $IPU_IAExe -PublisherPattern $IPU_IAPublisher -Sha256 $IPU_IASha256)) {
        Write-Output "Installation Assistant failed verification (valid Microsoft signature and any pinned digest required); not launching it."
        Remove-Item $IPU_IAExe -Force -ErrorAction SilentlyContinue
        Set-Status "FAILED IA_SIGNATURE $iaStamp"
        return
    }
    Write-Output "Installation Assistant signature verified."
    # Run to completion and wait. /NoReboot stages the upgrade without restarting,
    # so the reboot is deferred to the caller exactly like the ISO path. We never
    # auto-reboot here. ( /NoRestartUI is deliberately NOT used: as SYSTEM it forces
    # an immediate silent restart, which is the opposite of what we want. )
    try {
        $iaProc = Start-Process -FilePath $IPU_IAExe -ArgumentList $IPU_IAArgs -Wait -PassThru
        $iaCode = $iaProc.ExitCode
        Write-Output "Installation Assistant exited with code $iaCode."
        if ($iaCode -eq 0 -or $iaCode -eq 3010) {
            Set-RebootRequired -Detail "IA $iaStamp"
        } else {
            Write-Output "Installation Assistant could not complete (exit $iaCode); falling back to Windows Update."
            Invoke-IPUWindowsUpdate -Reason "Assistant exit $iaCode"
        }
    } catch {
        Write-Output "Installation Assistant launch failed: $_"
        Invoke-IPUWindowsUpdate -Reason "Assistant launch failed"
    }
}
function Set-WUUpgradeOverrides {
    # Bypass with reboot protection. Some boxes can only upgrade through Windows Update
    # (Enterprise composition, ARM64), but a NoAutoUpdate=1 policy stops WU from ever
    # delivering the feature update. Clear it and pin active hours so the WU-driven
    # reboot lands out of hours. Originals are recorded and restored once the box is on
    # the latest build; an enforced policy also self-reverts, and each run re-opens the
    # window until the upgrade lands.
    $polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $auk  = "$polk\AU"
    if (-not (Test-Path $polk)) { New-Item -Path $polk -Force | Out-Null }
    if (-not (Test-Path $auk))  { New-Item -Path $auk  -Force | Out-Null }
    if (-not (Test-Path $IPU_AURestoreFile)) {
        $noAuto = (Get-ItemProperty -Path $auk  -Name NoAutoUpdate     -ErrorAction SilentlyContinue).NoAutoUpdate
        $auOpt  = (Get-ItemProperty -Path $auk  -Name AUOptions        -ErrorAction SilentlyContinue).AUOptions
        $setAH  = (Get-ItemProperty -Path $polk -Name SetActiveHours   -ErrorAction SilentlyContinue).SetActiveHours
        $ahS    = (Get-ItemProperty -Path $polk -Name ActiveHoursStart -ErrorAction SilentlyContinue).ActiveHoursStart
        $ahE    = (Get-ItemProperty -Path $polk -Name ActiveHoursEnd   -ErrorAction SilentlyContinue).ActiveHoursEnd
        @(
            "NoAutoUpdate=$(if ($null -eq $noAuto) {'ABSENT'} else {$noAuto})"
            "AUOptions=$(if ($null -eq $auOpt) {'ABSENT'} else {$auOpt})"
            "SetActiveHours=$(if ($null -eq $setAH) {'ABSENT'} else {$setAH})"
            "ActiveHoursStart=$(if ($null -eq $ahS) {'ABSENT'} else {$ahS})"
            "ActiveHoursEnd=$(if ($null -eq $ahE) {'ABSENT'} else {$ahE})"
        ) | Set-Content -Path $IPU_AURestoreFile -Encoding ASCII -Force
    }
    Set-ItemProperty -Path $auk  -Name NoAutoUpdate     -Value 0 -Type DWord -Force
    # NoAutoUpdate=0 alone does not release an already-pending feature update on the
    # re-keyed Enterprise-composition boxes; AUOptions=4 (auto download + schedule
    # install) is what actually moves it. Set both.
    Set-ItemProperty -Path $auk  -Name AUOptions        -Value 4 -Type DWord -Force
    Set-ItemProperty -Path $polk -Name SetActiveHours   -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $polk -Name ActiveHoursStart -Value $IPU_ActiveHoursStart -Type DWord -Force
    Set-ItemProperty -Path $polk -Name ActiveHoursEnd   -Value $IPU_ActiveHoursEnd -Type DWord -Force
    # Pin Windows Update to the target release so it offers that feature update and not an
    # older one. Originals go into the same restore file; a file written by an older run
    # gets the missing entries appended before anything is changed.
    $trvRec = @(Get-Content -Path $IPU_AURestoreFile -ErrorAction SilentlyContinue)
    foreach ($trvName in 'ProductVersion','TargetReleaseVersion','TargetReleaseVersionInfo') {
        if (-not ($trvRec -match "^$trvName=")) {
            $trvVal = (Get-ItemProperty -Path $polk -Name $trvName -ErrorAction SilentlyContinue).$trvName
            Add-Content -Path $IPU_AURestoreFile -Value "$trvName=$(if ($null -eq $trvVal) {'ABSENT'} else {$trvVal})" -Encoding ASCII
        }
    }
    Set-ItemProperty -Path $polk -Name ProductVersion           -Value 'Windows 11'   -Type String -Force
    Set-ItemProperty -Path $polk -Name TargetReleaseVersion     -Value 1              -Type DWord  -Force
    Set-ItemProperty -Path $polk -Name TargetReleaseVersionInfo -Value $LatestVersion -Type String -Force
    Write-Output "Cleared NoAutoUpdate, set AUOptions=4, pinned active hours $IPU_ActiveHoursStart-$IPU_ActiveHoursEnd and pinned the target release to $LatestVersion so WU can deliver and reboot out of hours (originals recorded)."
}

function Invoke-IPUStaleCleanup {
    # Clear anything a prior aborted attempt left behind: orphaned setup processes,
    # mounted media (which pile up FsDepends/WIMMount filter instances), and setup
    # scratch folders. Any of these can lock files during the next downlevel finalize
    # and produce a 0xC1900101 sharing-violation failure, or leave a stale ~BT that
    # misleads Windows Update. Scoped to our work dir so an unrelated admin-mounted
    # image is never touched. Runs before routing, so every path (WU, Assistant, ISO)
    # starts from a clean state.
    Write-Output "Clearing stale upgrade state..."
    foreach ($pname in 'setup','setuphost','setupprep','setupplatform') {
        Get-Process -Name $pname -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Output "  Stopping leftover process: $($_.Name) (PID $($_.Id))"
            Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue
        }
    }
    Get-CimInstance -Namespace 'root\Microsoft\Windows\Storage' -ClassName 'MSFT_DiskImage' -ErrorAction SilentlyContinue |
        Where-Object { $_.Attached -and $_.ImagePath -like "$IPU_WorkDir\*" } |
        ForEach-Object {
            try {
                Dismount-DiskImage -ImagePath $_.ImagePath -ErrorAction Stop | Out-Null
                Write-Output "  Dismounted stale image: $($_.ImagePath)"
            } catch {
                Write-Output "  WARN: could not dismount $($_.ImagePath): $_"
            }
        }
    foreach ($d in "$env:SystemDrive\`$WINDOWS.~BT","$env:SystemDrive\`$WINDOWS.~WS","$env:SystemDrive\`$WINDOWS.~Q") {
        if (Test-Path -LiteralPath $d) {
            Write-Output "  Removing setup scratch: $d"
            Remove-Item -LiteralPath $d -Recurse -Force -ErrorAction SilentlyContinue
            if (Test-Path -LiteralPath $d) { cmd /c "rd /s /q `"$d`"" 2>$null }
        }
    }
    Write-Output "Stale state cleanup complete."
}

function Invoke-IPUWindowsUpdate {
    # Universal last-resort path. Windows Update is edition- and language-aware and is
    # the only mechanism that upgrades Enterprise composition and ARM64. It does not
    # stage locally and WU controls the timing, so clear any pause, nudge a scan, and
    # flag it. No reboot flag is set because nothing is staged.
    param([string]$Reason)
    $wuStamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Output "Handing off to Windows Update ($Reason)."
    $uxk = 'HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings'
    if (Test-Path $uxk) {
        foreach ($pn in 'PauseUpdatesExpiryTime','PauseUpdatesStartTime','PauseFeatureUpdatesStartTime','PauseFeatureUpdatesEndTime','PauseQualityUpdatesStartTime','PauseQualityUpdatesEndTime') {
            Remove-ItemProperty -Path $uxk -Name $pn -ErrorAction SilentlyContinue
        }
    }
    Set-WUUpgradeOverrides
    Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Start-Process -FilePath "$env:SystemRoot\System32\UsoClient.exe" -ArgumentList "StartInteractiveScan" -WindowStyle Hidden -ErrorAction SilentlyContinue
    Set-Status "WU_HANDOFF $wuStamp"
}
function Invoke-IPUComponentRepair {
    # Repair component store (StartComponentCleanup, RestoreHealth, SFC). Fixes 0xC1900204 / eKB apply failures. Needs internet.
    Set-Status "REPAIRING $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    Write-Output "Repairing component store (DISM StartComponentCleanup, RestoreHealth, SFC)..."
    try { & dism.exe /Online /Cleanup-Image /StartComponentCleanup | Out-Null } catch {}
    try { & dism.exe /Online /Cleanup-Image /RestoreHealth        | Out-Null } catch {}
    try { & sfc.exe /scannow                                       | Out-Null } catch {}
    Write-Output "Component store repair complete."
}
function Invoke-IPUEnablementPackage {
    # 24H2 (26100) or 25H2 (26200) -> 26H2 via KB5121794 eKB (x64 or ARM64 package). Verify servicing baseline, apply, repair+retry once, else WU fallback.
    $ekCV = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $ekBuild = [int]$ekCV.CurrentBuild
    $ekUBR = [int]$ekCV.UBR
    $ekFrom = if ($ekBuild -eq 26100) { '24H2' } else { '25H2' }
    if ($IPU_IsArm) { $IPU_EkbUrl = $IPU_EkbArmUrl; $IPU_EkbSha256 = $IPU_EkbArmSha256; $IPU_EkbPath = $IPU_EkbArmPath }
    if ($ekUBR -lt $IPU_EkbMinUBR) {
        Write-Output "$ekFrom servicing baseline $ekBuild.$ekUBR is below the $LatestVersion eKB prerequisite $ekBuild.$IPU_EkbMinUBR. Handing off to Windows Update."
        Invoke-IPUWindowsUpdate -Reason "$ekFrom servicing baseline $ekBuild.$ekUBR below $ekBuild.$IPU_EkbMinUBR"
        return
    }

    Write-Output "$ekFrom detected (build $ekBuild.$ekUBR). Applying the $LatestVersion enablement package (KB5121794)."
    try {
        $bl = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
        if ($bl -and $bl.ProtectionStatus -eq 'On') { Suspend-BitLocker -MountPoint "C:" -RebootCount 2 -ErrorAction Stop; Write-Output "BitLocker suspension for the eKB reboot." }
    } catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }

    Set-Status "DOWNLOADING $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    if (Test-Path $IPU_EkbPath) { Remove-Item $IPU_EkbPath -Force -ErrorAction SilentlyContinue }
    $ekOK = $true
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $IPU_EkbUrl -OutFile $IPU_EkbPath -UseBasicParsing -TimeoutSec 180
    } catch { Write-Output "eKB download failed from $IPU_EkbUrl : $_"; $ekOK = $false }
    if ($ekOK -and $IPU_EkbSha256) {
        $h = (Get-FileHash $IPU_EkbPath -Algorithm SHA256).Hash
        if ($h -ne $IPU_EkbSha256) { Write-Output "eKB hash mismatch. Expected $IPU_EkbSha256 got $h."; Remove-Item $IPU_EkbPath -Force -ErrorAction SilentlyContinue; $ekOK = $false }
        else { Write-Output "eKB hash verified." }
    }
    if ($ekOK) {
        $codes = @()
        for ($t = 1; $t -le 2; $t++) {
            Write-Output "Applying eKB via wusa (attempt $t of 2)..."
            $wp = Start-Process -FilePath "$env:SystemRoot\System32\wusa.exe" -ArgumentList "`"$IPU_EkbPath`" /quiet /norestart" -Wait -PassThru
            $ec = $wp.ExitCode
            $codes += ("0x{0:X}" -f $ec)
            "wusa exit code: $ec (0x$('{0:X}' -f $ec))" | Out-File -FilePath $IPU_SetupLog -Encoding ASCII
            # 0 = applied; 3010 = applied, reboot required; 0x240006 (2359302) = already installed
            if ($ec -eq 0 -or $ec -eq 3010) {
                Set-RebootRequired -Detail "3010 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                Write-Output "eKB applied (exit $($codes[-1])). Reboot to finish $LatestVersion."
                return
            }
            if ($ec -eq 2359302) {
                $ekPending = (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending') -or (Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired') -or (Test-Path "$env:SystemRoot\WinSxS\pending.xml")
                if ($ekPending) {
                    Set-RebootRequired -Detail "3010 $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
                    Write-Output "eKB is already installed and servicing reports a pending reboot. Reboot to finish $LatestVersion."
                    return
                }
                Write-Output "eKB is already installed, but no pending reboot is present and the OS is still $ekFrom. Handing off to Windows Update."
                Invoke-IPUWindowsUpdate -Reason "$ekFrom eKB already installed without pending reboot"
                return
            }
            if ($t -eq 1) { Write-Output "eKB apply returned $($codes[-1]). Repairing component store and retrying once."; Invoke-IPUComponentRepair }
        }
        Write-Output "eKB did not apply after repair and retry (codes: $($codes -join ', ')). Handing off to Windows Update."
    } else {
        Write-Output "eKB not available locally. Handing off to Windows Update (it delivers prerequisite CUs and the $LatestVersion feature update)."
    }
    Invoke-IPUWindowsUpdate -Reason "$ekFrom eKB unavailable or failed"
}
$stamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
if (-not (Test-Path $IPU_WorkDir)) { New-Item -ItemType Directory -Path $IPU_WorkDir -Force | Out-Null }
if (-not (Test-Path $IPU_LogDir))  { New-Item -ItemType Directory -Path $IPU_LogDir  -Force | Out-Null }
# Media left by older runs (earlier ISO name, 25H2 ISOs and eKB) is no longer used.
foreach ($IPU_LegacyName in 'Win11_25H2_x64.iso','Win11_25H2_x64_en-US.iso','Win11_25H2_x64_en-GB.iso','Win11_25H2_eKB_KB5054156_x64.msu') {
    $IPU_LegacyIso = Join-Path $IPU_WorkDir $IPU_LegacyName
    if (Test-Path $IPU_LegacyIso) { Remove-Item $IPU_LegacyIso -Force -ErrorAction SilentlyContinue }
}
Set-Status "RUNNING $stamp"
Start-Transcript -Path (Join-Path $IPU_LogDir "ipu_runner.log") -Append -Force | Out-Null

# Hygiene first: clear stale setup scratch / processes / mounts a prior aborted run
# left behind, so it cannot lock files at finalize or leave a stale ~BT that misleads
# Windows Update. Runs ahead of routing, so it benefits the WU, Assistant and ISO paths.
Invoke-IPUStaleCleanup

$cancel = $false

# Universal routing: pick a mechanism that can actually serve THIS box. WURSA runs
# worldwide, so the device may be any architecture, edition, UI language, or build.
# The local ISO only covers x64 + retail editions + the languages we host (en-US/en-GB).
#   LTSC / LTSB            -> skip and report (not feature-upgraded this way)
#   24H2 / 25H2 (26100 / 26200, x64 or ARM64) -> KB5121794 enablement package (light, one reboot)
#   ARM64                  -> Windows Update (x64 ISO and x64-only Assistant cannot run)
#   Enterprise             -> Windows Update (no retail media carries Enterprise)
#   non-hosted UI language -> Installation Assistant (MS serves matching language+edition)
#   hosted language, retail-> the local ISO below (covers Win10 and older Win11 x64)
# The Assistant itself falls back to Windows Update if it cannot serve the box.
$IPU_CV      = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
$IPU_CompEd  = $IPU_CV.CompositionEditionID
$IPU_EdId    = $IPU_CV.EditionID
$IPU_Prod    = $IPU_CV.ProductName
$IPU_Build   = [int]$IPU_CV.CurrentBuild
$IPU_RunLang = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language' -ErrorAction SilentlyContinue).InstallLanguage
# Detect native hardware/OS architecture. Environment variables can describe the
# emulated PowerShell process instead of the Snapdragon/ARM64 machine underneath it.
$IPU_ProcArch = $null
$IPU_OSArch   = $null
try { $IPU_ProcArch = [int](Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1 -ExpandProperty Architecture) } catch {}
try { $IPU_OSArch   = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).OSArchitecture } catch {}
$IPU_IsArm = ($IPU_ProcArch -in @(5,12)) -or ($IPU_OSArch -match 'ARM')
if (-not $IPU_IsArm -and $null -eq $IPU_ProcArch -and -not $IPU_OSArch) {
    $IPU_IsArm = ($env:PROCESSOR_ARCHITECTURE -match 'ARM') -or ($env:PROCESSOR_ARCHITEW6432 -match 'ARM')
}
$IPU_IsLtsc  = ($IPU_EdId -match 'EnterpriseS|IoTEnterpriseS') -or ($IPU_Prod -match 'LTSC|LTSB')
$IPU_Hosted  = ($IPU_RunLang -eq '0409' -or $IPU_RunLang -eq '0809')

if ($IPU_IsLtsc) {
    Write-Output "Edition '$IPU_EdId' ($IPU_Prod) is LTSC/LTSB; not feature-upgraded this way. Skipping for manual handling."
    Set-Status "BLOCK_LTSC $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $cancel = $true
}
elseif ($IPU_Build -in @(26100, 26200)) {
    # 24H2 or 25H2 (x64 or ARM64): take the enablement package, not the ISO.
    Invoke-IPUEnablementPackage
    $cancel = $true
}
elseif ($IPU_IsArm) {
    Write-Output "ARM64 device; the x64 ISO and Installation Assistant cannot run here."
    Invoke-IPUWindowsUpdate -Reason "ARM64"
    $cancel = $true
}
elseif ($IPU_EdId -match 'Enterprise') {
    Write-Output "Edition '$IPU_EdId'; no retail media (ISO or Assistant) carries Enterprise."
    Invoke-IPUWindowsUpdate -Reason "edition $IPU_EdId"
    $cancel = $true
}
elseif (-not $IPU_Hosted) {
    Write-Output "UI language '$IPU_RunLang' is not one we host an ISO for; using the Installation Assistant."
    Invoke-IPUInstallationAssistant -Reason "non-hosted language $IPU_RunLang"
    $cancel = $true
}
# else: hosted language + retail edition -> fall through to the ISO path below

# ISO download with partial-file detection
$NeedDownload = $true
if (Test-Path $IPU_IsoPath) {
    $ExistingGB = [math]::Round((Get-Item $IPU_IsoPath).Length / 1GB, 2)
    if ($ExistingGB -ge $IPU_IsoSizeGB) { $NeedDownload = $false }
    else { Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue }
}

if (-not $cancel -and $NeedDownload) {
    try {
        $HeadResponse = Invoke-WebRequest -Uri $IPU_IsoUrl -Method Head -UseBasicParsing -TimeoutSec 30
        Write-Output "URL OK (HTTP $($HeadResponse.StatusCode))."
    } catch {
        Write-Output "ISO URL not accessible: $_"
        Set-Status "FAILED URL_UNREACHABLE $stamp"
        $cancel = $true
    }
}

if (-not $cancel -and $NeedDownload) {
    try {
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -Uri $IPU_IsoUrl -OutFile $IPU_IsoPath -UseBasicParsing
    } catch {
        Write-Output "ISO download failed: $_"
        if (Test-Path $IPU_IsoPath) { Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue }
        Set-Status "FAILED DOWNLOAD $stamp"
        $cancel = $true
    }
    if (-not $cancel) {
        if (-not (Test-Path $IPU_IsoPath)) { Set-Status "FAILED NO_ISO $stamp"; $cancel = $true }
        else {
            $DownloadedGB = [math]::Round((Get-Item $IPU_IsoPath).Length / 1GB, 2)
            if ($DownloadedGB -lt $IPU_IsoSizeGB) {
                Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue
                Set-Status "FAILED PARTIAL_DOWNLOAD $stamp"
                $cancel = $true
            }
        }
    }
}

if (-not $cancel) {
    $actualHash = (Get-FileHash $IPU_IsoPath -Algorithm SHA256).Hash
    if ($actualHash -ne $IPU_IsoSha256) {
        Write-Output "ISO hash mismatch. Expected $IPU_IsoSha256 got $actualHash"
        Remove-Item $IPU_IsoPath -Force -ErrorAction SilentlyContinue
        Set-Status "FAILED HASH_MISMATCH $stamp"
        $cancel = $true
    } else { Write-Output "ISO hash verified." }
}

if (-not $cancel) {
    $MountedDrive = $null
    try {
        $MountResult  = Mount-DiskImage -ImagePath $IPU_IsoPath -PassThru
        $MountedDrive = ($MountResult | Get-Volume).DriveLetter
    } catch { Write-Output "Could not mount ISO: $_"; Set-Status "FAILED MOUNT $stamp"; $cancel = $true }

    if (-not $cancel) {
        $SetupExe = "${MountedDrive}:\setup.exe"
        if (-not (Test-Path $SetupExe)) {
            Dismount-DiskImage -ImagePath $IPU_IsoPath -ErrorAction SilentlyContinue
            Set-Status "FAILED NO_SETUP $stamp"
            $cancel = $true
        }
    }

    if (-not $cancel) {
        # BitLocker suspension - 4 reboots covers SafeOS + first/second boot + buffer.
        # Done here, after all gates pass, so an aborted run never leaves protection off.
        try {
            $BLStatus = Get-BitLockerVolume -MountPoint "C:" -ErrorAction SilentlyContinue
            if ($BLStatus -and $BLStatus.ProtectionStatus -eq 'On') {
                Suspend-BitLocker -MountPoint "C:" -RebootCount 4 -ErrorAction Stop
                Write-Output "BitLocker suspended."
            }
        } catch { Write-Output "BitLocker suspension failed: $($_.Exception.Message)" }

        $SetupArgs = "/auto upgrade /quiet /compat ignorewarning /DynamicUpdate disable /showoobe None /Telemetry Disable /EULA Accept /noreboot /Copylogs `"$IPU_LogDir`""
        $IPU_ExitCode = -1
        $storeRepaired = $false
        for ($setupTry = 1; $setupTry -le 2; $setupTry++) {
            Write-Output "Launching Windows Setup (downlevel phase, attempt $setupTry of 2)..."
            $IPU_ExitCode = -1
            try {
                $proc = Start-Process -FilePath $SetupExe -ArgumentList $SetupArgs -Wait -PassThru
                $IPU_ExitCode = $proc.ExitCode
            } catch { Write-Output "setup.exe failed to launch: $_" }
            "setup.exe exit code: $IPU_ExitCode" | Out-File -FilePath $IPU_SetupLog -Encoding ASCII
            if ($IPU_ExitCode -eq 0 -or $IPU_ExitCode -eq 3010) { break }
            # 0xC1900204 with matched edition/language = corrupt store, not a real block. Repair once and retry.
            if ($IPU_ExitCode -eq -1047526908 -and -not $storeRepaired -and $setupTry -lt 2) {
                Write-Output "0xC1900204 with matched edition/language: repairing component store and retrying."
                Invoke-IPUComponentRepair
                $storeRepaired = $true
                continue
            }
            break
        }
        Dismount-DiskImage -ImagePath $IPU_IsoPath -ErrorAction SilentlyContinue

        $done = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        switch ($IPU_ExitCode) {
            0           { Set-RebootRequired -Detail "3010 $done" }
            3010        { Set-RebootRequired -Detail "3010 $done" }
            -1047526908 { Invoke-IPUInstallationAssistant -Reason "0xC1900204 persists after store repair" }
            -1047527167 { Set-Status "BLOCK_HARD_COMPAT 0xC1900101 $done" }
            -1047526904 { Set-Status "BLOCK_APP_DRIVER 0xC1900208 $done" }
            -1047526912 { Set-Status "BLOCK_MIN_REQ 0xC1900200 $done" }
            default     { Set-Status "UNEXPECTED $IPU_ExitCode $done" }
        }
    }
}

Stop-Transcript | Out-Null
# Remove the one-shot task so it does not linger
schtasks.exe /Delete /TN "$IPU_TaskName" /F | Out-Null
'@
}
#endregion

#region Final status
function Resolve-WursaFinalStatus {
    # Maps the run state to the exit code and the status lines. The text always agrees with the code.
    param(
        [bool]$NoUpgrade, [int]$FinalBuild, [string]$FinalVersion, [string]$FeatureUpgradeState,
        [string]$IpuState, [bool]$RebootPending, [int]$FailedCount = 0, [int]$SucceededCount = 0
    )
    # Precedence: a failed feature upgrade (1) beats everything, then a needed reboot (3010), then
    # update failures (3 when something else succeeded, 1 when nothing did), then 0. While a detached
    # upgrade is running, a reboot request from the update pass must not interrupt it, so it is not reported.
    $lines = New-Object System.Collections.ArrayList
    $upgradeActive = $false
    $code = 0
    $add = { param($t, $c) [void]$lines.Add(@{ Text = $t; Color = $c }) }
    if (-not $NoUpgrade -and $FinalBuild -ge $LatestBuild) {
        & $add "STATUS: SYSTEM CURRENT ($FinalVersion)" $OkCol
        $code = 0
    } elseif ($NoUpgrade) {
        if ($RebootPending) {
            & $add "[!] STATUS: REBOOT REQUIRED" "Red"
            $code = 3010
        } else {
            & $add "STATUS: UPDATE PASS COMPLETE" $OkCol
            $code = 0
        }
    } elseif ($FeatureUpgradeState -eq "REBOOT_REQUIRED") {
        & $add "[!] STATUS: REBOOT REQUIRED TO FINISH $LatestVersion" "Red"
        $code = 3010
    } elseif ($FeatureUpgradeState -match '^(FAILED|BLOCKED|NOT_STARTED)') {
        & $add "[!] STATUS: $LatestVersion FEATURE UPGRADE NOT COMPLETE" "Red"
        & $add "      Upgrade state: $FeatureUpgradeState" $WarnCol
        $code = 1
    } elseif ($FeatureUpgradeState -eq "DISPATCHED") {
        if ($IpuState -match '^REBOOT_REQUIRED') {
            & $add "[!] STATUS: REBOOT REQUIRED TO FINISH $LatestVersion" "Red"
            & $add "      Upgrade state: $IpuState" $DimCol
            $code = 3010
        } elseif ($IpuState -match '^(FAILED|BLOCK_|UNEXPECTED)') {
            & $add "[!] STATUS: $LatestVersion FEATURE UPGRADE NOT COMPLETE" "Red"
            & $add "      Upgrade state: $IpuState" $WarnCol
            $code = 1
        } else {
            & $add "[>] STATUS: $LatestVersion FEATURE UPGRADE IN PROGRESS" $WarnCol
            if ($IpuState) { & $add "      Upgrade state: $IpuState" $DimCol }
            $code = 0
            $upgradeActive = $true
        }
    } elseif ($IpuState -match '^REBOOT_REQUIRED') {
        & $add "[!] STATUS: REBOOT REQUIRED TO FINISH $LatestVersion" "Red"
        & $add "      Upgrade state: $IpuState" $DimCol
        $code = 3010
    } elseif ($IpuState -match '^(DISPATCHED|RUNNING|DOWNLOADING|REPAIRING|VERIFY_PENDING|WU_HANDOFF)') {
        & $add "[>] STATUS: $LatestVersion FEATURE UPGRADE IN PROGRESS" $WarnCol
        & $add "      Upgrade state: $IpuState" $DimCol
        $code = 0
        $upgradeActive = $true
    } elseif ($IpuState -match '^(FAILED|BLOCK_|UNEXPECTED)') {
        & $add "[!] STATUS: $LatestVersion FEATURE UPGRADE NOT COMPLETE" "Red"
        & $add "      Upgrade state: $IpuState" $WarnCol
        $code = 1
    } else {
        & $add "[!] STATUS: FEATURE UPGRADE NOT COMPLETE (still $FinalVersion)" "Red"
        $code = 1
    }
    if ($IpuState -match '^REBOOT_REQUIRED_UNARMED') {
        & $add "      The post-reboot verifier could not be registered: run WURSA again after the reboot to restore the temporary policy changes." $WarnCol
    }
    # Fold in the update pass: reboot first, then failures. Only a clean 0 is ever changed here.
    if ($code -eq 0 -and $RebootPending -and -not $upgradeActive) {
        $lines.Clear()
        & $add "[!] STATUS: REBOOT REQUIRED" "Red"
        $code = 3010
    }
    if ($FailedCount -gt 0) {
        if ($code -eq 0) {
            $code = if ($SucceededCount -gt 0) { 3 } else { 1 }
            $lines.Clear()
            & $add "[!] STATUS: UPDATES INCOMPLETE ($FailedCount failed, $SucceededCount succeeded)" "Red"
        } else {
            & $add "      Also: $FailedCount update(s) failed ($SucceededCount succeeded)" $WarnCol
        }
    }
    [pscustomobject]@{ Code = $code; Lines = @($lines) }
}
#endregion

function Invoke-WursaMain {
# The body below is deliberately not indented: it keeps the diff small and keeps the
# here-string terminators in the generated scripts at column 0.
trap {
    if ($_.Exception.Message -like 'KILLERSCRIPT_EXIT:*') { throw $_ }
    Write-Host "[!] Fatal: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "[!] STATUS: FAILED" -ForegroundColor Red
    Exit-KillerScript 1
}
$script:ExitCode = 0
$script:StepRow  = -1
$script:StepMsg  = ""
$script:StepOpen = $false
# Feature-upgrade state is tracked separately from ordinary Windows Update.
# A surfaced/downloaded feature update is NOT considered staged. Region 5 owns
# the deterministic feature upgrade path unless Windows reports a real reboot-ready
# feature update.
$script:FeatureUpdateStaged = $false
$script:FeatureUpgradeState = "UNKNOWN"
$script:WorkDirChecked = $null
$ProgressPreference = 'SilentlyContinue'   # suppress console-wide progress overlays (VSS/BITS/Invoke-WebRequest)

#region 0 - Pre-Flight
if (-not (Test-WursaElevated)) {
    Write-Warning "Elevation Required: Please run as Administrator."
    Write-Host "[!] STATUS: NOT ELEVATED (run as Administrator)" -ForegroundColor Red
    Exit-KillerScript 2
}
# 1. Force session bypass to clear environment hurdles
try { Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process -Force -ErrorAction Stop } catch {}
# 2. Re-register core paths to bypass the LiveConnect temp-path bug
$env:PSModulePath = [Environment]::GetEnvironmentVariable("PSModulePath","Machine") + ";" + [Environment]::GetEnvironmentVariable("PSModulePath","User")

# Baseline Data for UI
$OS = Get-CimInstance Win32_OperatingSystem
try { Clear-Host } catch {}

# Header
$_pfx  = "█  "
$_art1 = "╦ ╦ ╦ ╦ ╦═╗ ╔═╗ ╔═╗ "
$_art2 = "║║║ ║ ║ ╠╦╝ ╚═╗ ╠═╣ "
$_art3 = "╚╩╝ ╚═╝ ╩╚═ ╚═╝ ╩ ╩ "
$_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
$_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
$_fillW = $script:Width - $_pfx.Length - $_artW
$_title = "WINDOWS UPDATE, REPAIR, & SYSTEM ALIGNMENT"
$_tpad  = " " * ($_fillW - $_title.Length)
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
Write-Host $_pfx -ForegroundColor $BorderCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

Write-Host "[>] Device Name      : " -NoNewline -ForegroundColor $LineCol
Write-Host "$($env:COMPUTERNAME)" -ForegroundColor $AccentCol
Write-Host "[>] Operating System : " -NoNewline -ForegroundColor $LineCol
$WinVer = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
$WinBuild = [int](Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
Write-Host "$($OS.Caption) $WinVer (Build $($OS.BuildNumber))" -ForegroundColor $AccentCol
Write-HLine -Style dashed

#region 0.5 - System Restore Point
# ============================================================================
Write-StepUpdate "[1/5] Creating System Restore Point..."
$_rp = Invoke-WursaRestorePoint -TimeoutSeconds 90
if ($_rp.Result -eq 'OK') {
    Write-StepUpdate -Success
} elseif ($_rp.Result -eq 'TIMEOUT') {
    Write-Host "      [!] Restore point $($_rp.Message); continuing." -ForegroundColor $WarnCol
} else {
    Write-Host "      [!] Restore point failed: $($_rp.Message)" -ForegroundColor $WarnCol
}
#endregion

#region 1 - Service Registration
# ============================================================================
Write-StepUpdate "[2/5] Opting-in to Microsoft Product Updates..."
try {
    Register-WursaProductService
    Write-StepUpdate -Success
    Start-Sleep -Seconds 3
} catch {
    Write-Host "      [!] Failed: $($_.Exception.Message)" -ForegroundColor Red
}
#endregion

#region 1.4 - Open Windows Update policy for the update pass
# ============================================================================
# Some managed endpoints carry NoAutoUpdate=1, which can prevent the direct Windows
# Update API from servicing ordinary updates. Open the policy here, but DO NOT kick
# a USO feature-update scan. Region 5 owns the feature upgrade path. Starting USO here
# can pre-download the feature update in parallel and race the deterministic ISO/eKB path. Enterprise
# and ARM64 fallback paths explicitly start USO later only when Windows Update is the
# selected upgrade mechanism. Original policy values are recorded for restoration.
$IPU_WorkDir          = "C:\Windows\Temp\25H2IPU"
$IPU_AURestoreFile    = Join-Path $IPU_WorkDir "au_restore.txt"
$IPU_ActiveHoursStart = 7     # WU will not auto-reboot the feature update between these hours (0-23, max 18h span)
$IPU_ActiveHoursEnd   = 19
if (-not $NoUpgrade -and $WinBuild -lt $LatestBuild -and -not (Confirm-WursaWorkDir -Path $IPU_WorkDir)) {
    Write-Host "      [!] Windows Update policy left untouched; the upgrade will not be dispatched." -ForegroundColor $WarnCol
} elseif (-not $NoUpgrade -and $WinBuild -lt $LatestBuild) {
    Write-Host "[>] Opening Windows Update policy for the update pass..." -ForegroundColor $LineCol
    $_polk = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate'
    $_auk  = "$_polk\AU"
    if (-not (Test-Path $_polk)) { New-Item -Path $_polk -Force | Out-Null }
    if (-not (Test-Path $_auk))  { New-Item -Path $_auk  -Force | Out-Null }
    # Record originals once, in the same format Restore-WUUpgradeOverrides expects.
    if (-not (Test-Path $IPU_AURestoreFile)) {
        $_noAuto = (Get-ItemProperty -Path $_auk  -Name NoAutoUpdate     -ErrorAction SilentlyContinue).NoAutoUpdate
        $_auOpt  = (Get-ItemProperty -Path $_auk  -Name AUOptions        -ErrorAction SilentlyContinue).AUOptions
        $_setAH  = (Get-ItemProperty -Path $_polk -Name SetActiveHours   -ErrorAction SilentlyContinue).SetActiveHours
        $_ahS    = (Get-ItemProperty -Path $_polk -Name ActiveHoursStart -ErrorAction SilentlyContinue).ActiveHoursStart
        $_ahE    = (Get-ItemProperty -Path $_polk -Name ActiveHoursEnd   -ErrorAction SilentlyContinue).ActiveHoursEnd
        @(
            "NoAutoUpdate=$(if ($null -eq $_noAuto) {'ABSENT'} else {$_noAuto})"
            "AUOptions=$(if ($null -eq $_auOpt) {'ABSENT'} else {$_auOpt})"
            "SetActiveHours=$(if ($null -eq $_setAH) {'ABSENT'} else {$_setAH})"
            "ActiveHoursStart=$(if ($null -eq $_ahS) {'ABSENT'} else {$_ahS})"
            "ActiveHoursEnd=$(if ($null -eq $_ahE) {'ABSENT'} else {$_ahE})"
        ) | Set-Content -Path $IPU_AURestoreFile -Encoding ASCII -Force
    }
    # NoAutoUpdate=0 alone does not release a pending feature update on the re-keyed
    # Enterprise-composition boxes; AUOptions=4 (auto download + schedule install) is
    # what actually moves it. Pin active hours so any WU-driven reboot lands out of hours.
    Set-ItemProperty -Path $_auk  -Name NoAutoUpdate     -Value 0 -Type DWord -Force
    Set-ItemProperty -Path $_auk  -Name AUOptions        -Value 4 -Type DWord -Force
    Set-ItemProperty -Path $_polk -Name SetActiveHours   -Value 1 -Type DWord -Force
    Set-ItemProperty -Path $_polk -Name ActiveHoursStart -Value $IPU_ActiveHoursStart -Type DWord -Force
    Set-ItemProperty -Path $_polk -Name ActiveHoursEnd   -Value $IPU_ActiveHoursEnd   -Type DWord -Force
    Restart-Service wuauserv -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 3
    Write-Host "      WU opened (NoAutoUpdate=0, AUOptions=4); direct API scan follows." -ForegroundColor $DimCol
}
#endregion

#region 2 and 3 - Windows Update discovery and installation
# ============================================================================
$passRes = Invoke-WursaWindowsUpdatePass
$UpdateSearcher = $passRes.Searcher
#endregion

#region 4 - Third-Party Software Updates
# ============================================================================
$_tpFlag = if ($Update3rdParty) { "-Update3rdParty" } else { "-InplaceUpgrade" }
$tpRes = Invoke-WursaThirdPartyUpdate -Skip ([bool]$No3rdParty) -AutoConfirm ([bool]($Update3rdParty -or $InplaceUpgrade)) -AutoFlag $_tpFlag
#endregion

#region 5 - Post-Update Version Check & Optional In-Place Upgrade
# ============================================================================
# ISO config - hosted on Cloudflare R2 behind a stable custom domain
# (iso.killertools.net) so the link does not expire. To target a different
# feature build, swap the object in the killer-isos bucket and update the URL.
#
# Language-aware: the in-place upgrade requires the ISO language to match the
# running OS UI language, or setup.exe returns 0xC1900204 and the keep files
# path is silently blocked. Read the base install language from the NLS registry
# (reliable under SYSTEM) and pick the matching image: en-GB (0809) gets the
# International ISO, everything else defaults to en-US for the US fleet.
$IPU_InstallLang = (Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Nls\Language' -ErrorAction SilentlyContinue).InstallLanguage
switch ($IPU_InstallLang) {
    '0809'  { $IPU_IsoName = "Win11_26H2_x64_en-GB.iso"; $IPU_IsoSha256 = "7E3F373BD3C2321B5D5125DFCF718DFBF1A0ABC50A267118669951890DF98A5D" }
    default { $IPU_IsoName = "Win11_26H2_x64_en-US.iso"; $IPU_IsoSha256 = "BD4307DF32BC8AF33B39CCECB1174AEB345386630F89A2B86C7A4E36B55EA650" }
}
$IPU_IsoUrl     = "https://iso.killertools.net/$IPU_IsoName"
$IPU_IsoSizeGB  = 6.0
# $IPU_WorkDir, $IPU_AURestoreFile, and $IPU_ActiveHoursStart/End are defined earlier
# (region 1.4 WU ungag, which runs before the scan); reused here.
$IPU_IsoPath    = Join-Path $IPU_WorkDir $IPU_IsoName
$IPU_LogDir     = Join-Path $IPU_WorkDir "SetupLogs"
$IPU_SetupLog   = Join-Path $IPU_WorkDir "setup_exit.log"
$IPU_StatusFile = Join-Path $IPU_WorkDir "ipu_status.txt"
$IPU_RunnerPath = Join-Path $IPU_WorkDir "Invoke-IPU.ps1"
$IPU_TaskName   = "WURSA-25H2-IPU"
$IPU_VerifierPath  = Join-Path $IPU_WorkDir "Verify-25H2.ps1"
$IPU_VerifyTaskName = "WURSA-25H2-Verify"

# 24H2 (26100) or 25H2 (26200) -> 26H2 enablement package (KB5121794), x64 and ARM64.
# Host both MSUs in the killer-isos bucket. Prerequisite: servicing baseline
# 26100.9278 / 26200.9278 (KB5120998) or later; below that, Windows Update takes over.
$IPU_EkbName    = "Win11_26H2_eKB_KB5121794_x64.msu"
$IPU_EkbUrl     = "https://iso.killertools.net/$IPU_EkbName"
$IPU_EkbSha256  = "BF43EB13B7EBDE920C2201473BFD5C531B45E0F46ACD8E0843F27F9744FB4B43"
$IPU_EkbPath    = Join-Path $IPU_WorkDir $IPU_EkbName
$IPU_EkbArmName   = "Win11_26H2_eKB_KB5121794_arm64.msu"
$IPU_EkbArmUrl    = "https://iso.killertools.net/$IPU_EkbArmName"
$IPU_EkbArmSha256 = "367E77739B3B09DC8C48553F9853BF53042F052DDAADC07ED025B83BB1D53A9D"
$IPU_EkbArmPath   = Join-Path $IPU_WorkDir $IPU_EkbArmName
$IPU_EkbMinUBR    = 9278

# Installation Assistant fallback (online, edition-correct). The local ISO is a
# consumer image with no Enterprise SKU, so an Enterprise-composition box fails the
# ISO match with 0xC1900204. The Assistant pulls the matching edition straight from
# Microsoft. Host the exe alongside the ISO in the killer-isos bucket; the hash pin
# is optional (leave empty to skip).
$IPU_IAUrl      = "https://iso.killertools.net/Windows11InstallationAssistant.exe"
$IPU_IASha256   = ""
# Whatever the source, the Assistant only runs as SYSTEM when it carries a valid Microsoft signature.
$IPU_IAPublisher = '^CN=Microsoft Corporation(,|$)'


Write-HLine -Style dashed
Write-Host "[>] Checking Windows Feature Update Level..." -ForegroundColor $LineCol
$InstalledVersion = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").DisplayVersion
$InstalledBuild   = [int](Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentBuild
# Detect the native machine architecture from CIM instead of the PowerShell process
# environment. Snapdragon/ARM64 devices can run x64-emulated management processes,
# where PROCESSOR_ARCHITECTURE may report AMD64 and incorrectly route to x64 media.
$IPU_ProcArch = $null
$IPU_OSArch   = $null
try { $IPU_ProcArch = [int](Get-CimInstance Win32_Processor -ErrorAction Stop | Select-Object -First 1 -ExpandProperty Architecture) } catch {}
try { $IPU_OSArch   = (Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).OSArchitecture } catch {}
$IPU_IsArm64 = ($IPU_ProcArch -in @(5,12)) -or ($IPU_OSArch -match 'ARM')
if (-not $IPU_IsArm64 -and $null -eq $IPU_ProcArch -and -not $IPU_OSArch) {
    # Last-resort fallback only when CIM is unavailable.
    $IPU_IsArm64 = ($env:PROCESSOR_ARCHITECTURE -match 'ARM') -or ($env:PROCESSOR_ARCHITEW6432 -match 'ARM')
}
Write-Host "      Installed Version : $InstalledVersion" -ForegroundColor Yellow
Write-Host "      Latest Version    : $LatestVersion"    -ForegroundColor Yellow
if ($IPU_IsArm64) { Write-Host "      Architecture      : ARM64" -ForegroundColor $AccentCol }
# ARM64 24H2/25H2 takes the ARM64 eKB; older ARM64 builds can only upgrade through Windows Update.
$IPU_ArmWUOnly = $IPU_IsArm64 -and $InstalledBuild -notin @(26100, 26200)
if ($NoUpgrade) {
    $script:FeatureUpgradeState = "SKIPPED"
    Write-Host "      [-NoUpgrade] Feature upgrade check skipped." -ForegroundColor $DimCol
} elseif ($InstalledBuild -lt $LatestBuild -and (Test-WUFeatureUpdateStaged -Searcher $UpdateSearcher)) {
    # This is the only Windows Update state that suppresses the dedicated upgrade path.
    # Windows has actually staged the feature update and is asking for a reboot.
    $script:FeatureUpdateStaged = $true
    $script:FeatureUpgradeState = "REBOOT_REQUIRED"
    Write-Host "      > $LatestVersion is staged by Windows Update and requires a reboot." -ForegroundColor $DimCol
    if (Confirm-WursaWorkDir -Path $IPU_WorkDir) {
        ('REBOOT_REQUIRED WU ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
    }
} elseif ($InstalledBuild -lt $LatestBuild) {
    Write-Host "      > Feature update available." -ForegroundColor $WarnCol
    # Battery safety
    $_power = Get-WursaPowerState
    $OnBattery = ($_power -ne 'AC')

    $proceed = $false
    if ($_power -eq 'Battery') {
        $script:FeatureUpgradeState = "BLOCKED_BATTERY"
        Write-Host "      [!] Upgrade skipped: device is running on battery power." -ForegroundColor Yellow
    } elseif ($_power -eq 'Unknown') {
        $script:FeatureUpgradeState = "BLOCKED_POWER_UNKNOWN"
        Write-Host "      [!] Upgrade skipped: could not determine the power source (BLOCKED_POWER_UNKNOWN). Plug in AC and run again." -ForegroundColor Yellow
    } elseif ($InplaceUpgrade) {
        if ($IPU_ArmWUOnly) {
            Write-Host "      [-InplaceUpgrade] ARM64 detected; auto-dispatching Windows Update feature-upgrade handoff." -ForegroundColor $DimCol
        } else {
            Write-Host "      [-InplaceUpgrade] Auto-dispatching detached feature-upgrade workflow." -ForegroundColor $DimCol
        }
        $proceed = $true
    } else {
        # Flush buffered keystrokes; guarded so redirected (RMM) consoles do not throw
        try { while ([Console]::KeyAvailable) { [Console]::ReadKey($true) | Out-Null } } catch {}
        try {
            if ($IPU_ArmWUOnly) {
                Write-Host "Would you like to hand off the $($LatestVersion) upgrade to Windows Update? (Y/N): " -NoNewline -ForegroundColor $WarnCol
            } else {
                Write-Host "Would you like to perform an in-place upgrade to $($LatestVersion)? (Y/N): " -NoNewline -ForegroundColor $WarnCol
            }
            $key    = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $choice = $key.Character
            if ($choice -ne "`n" -and $choice -ne "`r") { Write-Host $choice -NoNewline }
            Write-Host ""
            if ($choice -in @('Y','y')) { $proceed = $true }
        } catch {
            Write-Host "      [i] Non-interactive - defaulting to NO. Use -InplaceUpgrade for unattended runs." -ForegroundColor $DimCol
        }
    }

    if (-not $OnBattery -and -not $proceed) {
        $script:FeatureUpgradeState = "NOT_STARTED"
    }

    if (-not $OnBattery -and $proceed) {
        if ($IPU_ArmWUOnly) {
            Write-Host "[>] Dispatching detached ARM64 Windows Update handoff (survives session disconnect)..." -ForegroundColor $LineCol
        } else {
            Write-Host "[>] Dispatching detached feature-upgrade workflow (survives session disconnect)..." -ForegroundColor $LineCol
        }

        if (-not (Confirm-WursaWorkDir -Path $IPU_WorkDir)) {
            $script:FeatureUpgradeState = "BLOCKED_WORKDIR"
            Write-Host "      [!] Upgrade not dispatched: $IPU_WorkDir is not a protected directory." -ForegroundColor Red
        } else {

        # One-shot post-reboot verifier, built by Get-WursaVerifierScript.
        $VerifierBody = Get-WursaVerifierScript -StatusFile $IPU_StatusFile -RestoreFile $IPU_AURestoreFile -VerifyTaskName $IPU_VerifyTaskName -TargetBuild $LatestBuild
        $VerifierBody | Out-File -FilePath $IPU_VerifierPath -Encoding ASCII -Force

        # --- Build the standalone runner the scheduled task executes as SYSTEM ---
        # Config preamble (double-quoted here-string: values are injected now).
        $RunnerConfig = @"
`$IPU_IsoUrl     = '$IPU_IsoUrl'
`$IPU_IsoSizeGB  = $IPU_IsoSizeGB
`$IPU_WorkDir    = '$IPU_WorkDir'
`$IPU_IsoPath    = '$IPU_IsoPath'
`$IPU_LogDir     = '$IPU_LogDir'
`$IPU_SetupLog   = '$IPU_SetupLog'
`$IPU_StatusFile = '$IPU_StatusFile'
`$IPU_TaskName   = '$IPU_TaskName'
`$IPU_VerifierPath  = '$IPU_VerifierPath'
`$IPU_VerifyTaskName = '$IPU_VerifyTaskName'
`$IPU_IsoSha256  = '$IPU_IsoSha256'
`$IPU_IAUrl     = '$IPU_IAUrl'
`$IPU_IASha256  = '$IPU_IASha256'
`$IPU_IAPublisher = '$IPU_IAPublisher'
`$IPU_AURestoreFile    = '$IPU_AURestoreFile'
`$IPU_ActiveHoursStart = $IPU_ActiveHoursStart
`$IPU_ActiveHoursEnd   = $IPU_ActiveHoursEnd
`$IPU_EkbUrl     = '$IPU_EkbUrl'
`$IPU_EkbSha256  = '$IPU_EkbSha256'
`$IPU_EkbPath    = '$IPU_EkbPath'
`$IPU_EkbArmUrl    = '$IPU_EkbArmUrl'
`$IPU_EkbArmSha256 = '$IPU_EkbArmSha256'
`$IPU_EkbArmPath   = '$IPU_EkbArmPath'
`$IPU_EkbMinUBR    = $IPU_EkbMinUBR
`$LatestVersion    = '$LatestVersion'
"@


        New-WursaRunnerScript -Config $RunnerConfig | Out-File -FilePath $IPU_RunnerPath -Encoding ASCII -Force

        # --- Register + start the detached SYSTEM task ---
        $null = ('DISPATCHED ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
        $TR = "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File $IPU_RunnerPath"
        $createOut = schtasks.exe /Create /TN $IPU_TaskName /TR $TR /SC ONCE /ST 00:00 /RU SYSTEM /RL HIGHEST /F 2>&1
        if ($LASTEXITCODE -eq 0) {
            $runOut = schtasks.exe /Run /TN $IPU_TaskName 2>&1
            if ($LASTEXITCODE -eq 0) {
                $script:FeatureUpgradeState = "DISPATCHED"
                Write-Host "      > Upgrade dispatched as SYSTEM. It continues after you disconnect." -ForegroundColor Green
                Write-Host "      > Status file : $IPU_StatusFile" -ForegroundColor $DimCol
                Write-Host "      > Setup logs  : $IPU_LogDir" -ForegroundColor $DimCol
                Write-Host "      > A reboot will be required once the detached upgrade completes." -ForegroundColor $WarnCol
            } else {
                $script:FeatureUpgradeState = "FAILED"
                ('FAILED TASK_START ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
                Write-Host "      [!] Task created but failed to start: $runOut" -ForegroundColor Red
            }
        } else {
            $script:FeatureUpgradeState = "FAILED"
            ('FAILED TASK_CREATE ' + (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')) | Out-File -FilePath $IPU_StatusFile -Encoding ASCII -Force
            Write-Host "      [!] Failed to register scheduled task: $createOut" -ForegroundColor Red
        }
        }   # staging directory trusted
    }
} else {
    $script:FeatureUpgradeState = "CURRENT"
    if ($InstalledBuild -gt $LatestBuild) {
        Write-Host "      System is on $InstalledVersion, newer than $LatestVersion. No feature upgrade needed." -ForegroundColor Green
    } else {
        Write-Host "      System is already on the latest feature update." -ForegroundColor Green
    }
    Restore-WUUpgradeOverrides
}
#endregion

#region 6 - Finalization
# ============================================================================
# Checking reboot status via direct API to bypass Get-WURebootStatus crash.
$RebootPending = [bool]$passRes.RebootRequired -or [bool]$tpRes.RebootRequired
if (-not $RebootPending) {
    $RebootPending = $null -ne (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction SilentlyContinue)
}

# Never report SYSTEM CURRENT solely because there is no reboot flag. A feature
# upgrade is complete only when the installed build is actually the target release or newer.
$_finalCV      = Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -ErrorAction SilentlyContinue
$_finalVersion = $_finalCV.DisplayVersion
$_ipuState     = ""
# Only read the status file from a directory that is still protected; anything else could be forged.
if (-not (Get-WursaDirectoryProblem -Path $IPU_WorkDir) -and (Test-Path $IPU_StatusFile)) {
    try { $_ipuState = (Get-Content $IPU_StatusFile -ErrorAction Stop | Select-Object -First 1).Trim() } catch {}
}

Write-HLine -Style dashed
$_final = Resolve-WursaFinalStatus -NoUpgrade ([bool]$NoUpgrade) -FinalBuild ([int]$_finalCV.CurrentBuild) -FinalVersion $_finalVersion `
    -FeatureUpgradeState $script:FeatureUpgradeState -IpuState $_ipuState -RebootPending $RebootPending `
    -FailedCount ([int]$passRes.Failed + [int]$tpRes.Failed) -SucceededCount ([int]$passRes.Succeeded + [int]$tpRes.Succeeded)
foreach ($_l in $_final.Lines) { Write-Host $_l.Text -ForegroundColor $_l.Color }
$script:ExitCode = $_final.Code
# Footer
$_sfx    = "█"
$_ftrW   = $_artW + 1
$_ffillW = $script:Width - $_ftrW - $_sfx.Length
$_footer = "  WINDOWS UPDATE SEQUENCE COMPLETE"
$_fpad   = " " * ($_ffillW - $_footer.Length - $_ver.Length)
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host $_sfx -ForegroundColor $BorderCol
#endregion

Exit-KillerScript $script:ExitCode
}

if (-not $global:KillerScriptsTestMode) { Invoke-WursaMain }
