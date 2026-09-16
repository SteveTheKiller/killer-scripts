<#
.SYNOPSIS
    Patch & Event Examination Kit (P.E.E.K.) v1.0
    Created by Steve the Killer | Updated: 2026-08-27   
.DESCRIPTION
    Read-only live Windows Update and Windows Setup servicing monitor for
    remote support. Correlates WindowsUpdateClient events, servicing processes,
    Panther setup activity, compatibility blocks, reboot indicators, services,
    disk space, power state, and recent update history in one continuously
    refreshing console view.
.NOTES
    Compatible with Windows PowerShell 5.1 and PowerShell 7.
    Safe for Kaseya LiveConnect and SYSTEM execution.
    Makes no registry changes, starts no updates, opens no UI, and never reboots.
    Press Q or ESC to exit cleanly. Ctrl+C also stops the monitor.
#>

param(
    [ValidateRange(1,60)]
    [int]$RefreshSeconds = 3,

    [switch]$Once
)

#region [0] - PRE-FLIGHT & CONFIGURATION
# ============================================================================

$EXIT_SUCCESS = 0
$EXIT_DENIED  = 3

if (-not ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "P.E.E.K. requires elevated privileges. Please run as Administrator or SYSTEM."
    exit $EXIT_DENIED
}

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding            = [System.Text.Encoding]::UTF8

$script:Acronym    = "PEEK"
$script:AcronymFmt = "P.E.E.K."
$script:ScriptName = "Patch & Event Examination Kit"
$script:Version    = "v1.0"

$LineCol   = "Blue"
$MainCol   = "Magenta"
$ArtCol    = "Yellow"
$AccentCol = "Cyan"
$DimCol    = "Gray"

$script:LabelFG    = "Cyan"
$script:ValueFG    = "White"
$script:MutedFG    = "DarkGray"
$script:DoneFG     = "Green"
$script:WarnFG     = "DarkYellow"
$script:FailFG     = "Red"
$script:ActiveFG   = "Yellow"
$script:LineColors = @(
    [ConsoleColor]::Yellow,
    [ConsoleColor]::Blue,
    [ConsoleColor]::Magenta,
    [ConsoleColor]::Cyan,
    [ConsoleColor]::Gray
)

$script:Width              = 85
$script:FrameIndex         = 0
$script:PreviousCpu        = @{}
$script:PreviousSampleTime = $null
$script:CompatCacheKey       = $null
$script:CompatCacheValue     = $null
$script:BitLockerState       = "Unknown"
$script:SetupCommandLinePath = ""
$script:SetupCommandLine     = ""
$script:SetupDeepScanTime    = [datetime]::MinValue

#endregion

#region [1] - DISPLAY HELPERS
# ============================================================================

function Update-ConsoleWidth {
    $width = 85
    try {
        $candidate = [int]$Host.UI.RawUI.WindowSize.Width
        if ($candidate -gt 0) { $width = $candidate - 1 }
    } catch {
        try {
            $candidate = [int][Console]::WindowWidth
            if ($candidate -gt 0) { $width = $candidate - 1 }
        } catch {}
    }

    if ($width -lt 76)  { $width = 76 }
    if ($width -gt 120) { $width = 120 }
    $script:Width = $width
}

function Get-TruncatedText {
    param(
        [AllowNull()][string]$Text,
        [int]$MaxLength
    )

    if ($null -eq $Text) { return "" }
    if ($MaxLength -le 0) { return "" }
    if ($Text.Length -le $MaxLength) { return $Text }
    if ($MaxLength -le 3) { return $Text.Substring(0, $MaxLength) }
    return $Text.Substring(0, $MaxLength - 3) + "..."
}

function Format-Age {
    param([datetime]$Time)

    if ($Time -eq [datetime]::MinValue) { return "N/A" }
    $span = (Get-Date) - $Time
    if ($span.TotalSeconds -lt 0) { return "0s" }
    if ($span.TotalSeconds -lt 60) { return "{0}s" -f [math]::Floor($span.TotalSeconds) }
    if ($span.TotalMinutes -lt 60) { return "{0}m {1}s" -f [math]::Floor($span.TotalMinutes), $span.Seconds }
    if ($span.TotalHours -lt 24)   { return "{0}h {1}m" -f [math]::Floor($span.TotalHours), $span.Minutes }
    return "{0}d {1}h" -f [math]::Floor($span.TotalDays), $span.Hours
}

function Format-Uptime {
    param([timespan]$Uptime)
    return "{0}d {1}h {2}m" -f $Uptime.Days, $Uptime.Hours, $Uptime.Minutes
}

function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )

    if ($Width -lt 1) { return }

    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "━" * $Width
    }

    $colors = $script:LineColors
    $useConsole = $true
    try { $saved = [Console]::ForegroundColor } catch { $useConsole = $false }
    $i = 0

    foreach ($char in $line.ToCharArray()) {
        if ($char -eq ' ') {
            $fg = [ConsoleColor]$script:MutedFG
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
    $_pfx  = "█  "
    $_art1 = "╔═╗ ╔═╗ ╔═╗ ╦╔═ "
    $_art2 = "╠═╝ ║╣  ║╣  ╠╩╗ "
    $_art3 = "╩   ╚═╝ ╚═╝ ╩ ╩ "
    $_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1 = $_art1.PadRight($_artW)
    $_art2 = $_art2.PadRight($_artW)
    $_art3 = $_art3.PadRight($_artW)

    $_fillW = $script:Width - $_pfx.Length - $_artW
    if ($_fillW -lt 12) { $_fillW = 12 }

    $_ver      = "| $($script:Version)"
    $_titleMax = [Math]::Max(1, $_fillW - $_ver.Length - 1)
    $_title    = Get-TruncatedText -Text $script:ScriptName.ToUpper() -MaxLength $_titleMax
    $_pad      = " " * [Math]::Max(1, ($_fillW - $_title.Length - $_ver.Length))

    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline
    Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline
    Write-Host ("-" * $_fillW) -ForegroundColor $LineCol

    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline
    Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline
    Write-Host "$_title$_pad$_ver" -ForegroundColor $MainCol

    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline
    Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline
    Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
}

function Write-Footer {
    param([string]$Text = "READ-ONLY LIVE SERVICING MONITOR")

    $_sfx   = "█"
    $_ftr1  = " ╔═╗ ╔═╗ ╔═╗ ╦╔═ "
    $_ftr2  = " ╠═╝ ║╣  ║╣  ╠╩╗ "
    $_ftr3  = " ╩   ╚═╝ ╚═╝ ╩ ╩ "
    $_ftrW  = [Math]::Max($_ftr1.Length, [Math]::Max($_ftr2.Length, $_ftr3.Length))
    $_ftr1  = $_ftr1.PadRight($_ftrW)
    $_ftr2  = $_ftr2.PadRight($_ftrW)
    $_ftr3  = $_ftr3.PadRight($_ftrW)
    $_ffillW = $script:Width - $_ftrW - $_sfx.Length
    if ($_ffillW -lt 12) { $_ffillW = 12 }

    $_fver    = "| $($script:Version)"
    $_textMax = [Math]::Max(1, $_ffillW - $_fver.Length - 2)
    $_footer  = "  " + (Get-TruncatedText -Text $Text -MaxLength $_textMax)
    $_fpad    = " " * [Math]::Max(1, ($_ffillW - $_footer.Length - $_fver.Length))

    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline
    Write-Host $_ftr1 -ForegroundColor $ArtCol -NoNewline
    Write-Host $_sfx -ForegroundColor $LineCol

    Write-Host "$_footer$_fpad$_fver" -ForegroundColor $MainCol -NoNewline
    Write-Host $_ftr2 -ForegroundColor $ArtCol -NoNewline
    Write-Host $_sfx -ForegroundColor $LineCol

    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline
    Write-Host $_ftr3 -ForegroundColor $ArtCol -NoNewline
    Write-Host $_sfx -ForegroundColor $LineCol
}

function Write-Detail {
    param(
        [string]$Label,
        [AllowNull()][string]$Value,
        [string]$ValueColor = $script:ValueFG,
        [int]$LabelWidth = 18
    )

    $prefix = "  " + $Label.PadRight($LabelWidth) + ": "
    $maxValue = [Math]::Max(1, $script:Width - $prefix.Length)
    $display = Get-TruncatedText -Text $Value -MaxLength $maxValue

    Write-Host $prefix -NoNewline -ForegroundColor $script:LabelFG
    Write-Host $display -ForegroundColor $ValueColor
}

function Write-StepUpdate {
    param(
        [ValidateSet("OK","FAIL","INFO","WARN","SKIP","TRY")]
        [string]$Status,
        [string]$Message,
        [string]$Detail = ""
    )

    $tagColor = switch ($Status) {
        "OK"   { $script:DoneFG }
        "FAIL" { $script:FailFG }
        "WARN" { $script:WarnFG }
        "TRY"  { $script:ActiveFG }
        "SKIP" { $script:MutedFG }
        default { $AccentCol }
    }

    $tag = "[$Status]".PadRight(7)
    $detailText = if ($Detail) { " $Detail" } else { "" }
    $maxMessage = [Math]::Max(1, $script:Width - 11 - $detailText.Length)
    $msg = Get-TruncatedText -Text $Message -MaxLength $maxMessage

    Write-Host "  $tag" -ForegroundColor $tagColor -NoNewline
    Write-Host $msg -ForegroundColor $script:ValueFG -NoNewline
    if ($detailText) { Write-Host $detailText -ForegroundColor $script:MutedFG } else { Write-Host "" }
}

function Write-SectionTitle {
    param([string]$Title)
    Write-Host "[>] $Title" -ForegroundColor $AccentCol
}

#endregion

#region [2] - WINDOWS UPDATE & SERVICING DATA
# ============================================================================

function Get-UpdateTitleFromMessage {
    param([AllowNull()][string]$Message)

    if (-not $Message) { return "" }
    $clean = ($Message -replace "`r|`n", " " -replace '\s+', ' ').Trim()

    if ($clean -match 'following update(?: with error [^:]+)?:\s*(.+)$') {
        return $Matches[1].Trim()
    }
    if ($clean -match 'following update:\s*(.+)$') {
        return $Matches[1].Trim()
    }
    if ($clean -match 'update:\s*(.+)$') {
        return $Matches[1].Trim()
    }

    return $clean
}

function Convert-UpdateErrorCode {
    param($Value)

    if ($null -eq $Value) { return "" }
    try {
        if ($Value -is [int] -or $Value -is [long] -or $Value -is [uint32] -or $Value -is [uint64]) {
            return "0x{0:X8}" -f ([uint32]$Value)
        }
    } catch {}

    $text = [string]$Value
    if ($text -match '0x[0-9A-Fa-f]{8}') { return $Matches[0].ToUpper() }
    return $text
}

function Get-PEEKUpdateEvents {
    $result = @()
    $filter = @{
        LogName      = 'System'
        ProviderName = 'Microsoft-Windows-WindowsUpdateClient'
        Id           = @(17,19,20,43,44,212,214)
        StartTime    = (Get-Date).AddDays(-2)
    }

    try {
        $rawEvents = @(Get-WinEvent -FilterHashtable $filter -MaxEvents 100 -ErrorAction SilentlyContinue)
    } catch {
        $rawEvents = @()
    }

    foreach ($wuEvent in $rawEvents) {
        $title = ""
        $status = "INFO"
        $label = "Event"
        $errorCode = ""

        try {
            switch ([int]$wuEvent.Id) {
                17 {
                    $label  = "Ready"
                    $status = "INFO"
                    if ($wuEvent.Properties.Count -gt 0) { $title = [string]$wuEvent.Properties[0].Value }
                }
                19 {
                    $label  = "Installed"
                    $status = "OK"
                    if ($wuEvent.Properties.Count -gt 0) { $title = [string]$wuEvent.Properties[0].Value }
                }
                20 {
                    $label  = "Failed"
                    $status = "FAIL"
                    if ($wuEvent.Properties.Count -gt 0) { $errorCode = Convert-UpdateErrorCode $wuEvent.Properties[0].Value }
                    if ($wuEvent.Properties.Count -gt 1) { $title = [string]$wuEvent.Properties[1].Value }
                }
                43 {
                    $label  = "Installing"
                    $status = "TRY"
                    if ($wuEvent.Properties.Count -gt 0) { $title = [string]$wuEvent.Properties[0].Value }
                }
                44 {
                    $label  = "Downloading"
                    $status = "INFO"
                    $title  = "Windows Update download"
                }
                212 {
                    $label  = "Reverted"
                    $status = "WARN"
                }
                214 {
                    $label  = "Rollback"
                    $status = "WARN"
                }
            }
        } catch {}

        if (-not $title) { $title = Get-UpdateTitleFromMessage -Message $wuEvent.Message }
        $title = ($title -replace "`r|`n", " " -replace '\s+', ' ').Trim()
        if (-not $title) { $title = "Windows Update event" }

        $result += [PSCustomObject]@{
            TimeCreated = [datetime]$wuEvent.TimeCreated
            Id          = [int]$wuEvent.Id
            Status      = $status
            Label       = $label
            Title       = $title
            ErrorCode   = $errorCode
        }
    }

    return @($result | Sort-Object TimeCreated -Descending)
}

function Get-PEEKUpdateContext {
    param([object[]]$Events)

    $eventsArray = @($Events)
    $activeStart = $null

    foreach ($start in @($eventsArray | Where-Object { $_.Id -eq 43 } | Sort-Object TimeCreated -Descending)) {
        $terminal = @($eventsArray | Where-Object {
            ($_.Id -eq 19 -or $_.Id -eq 20) -and
            $_.TimeCreated -gt $start.TimeCreated -and
            $_.Title -eq $start.Title
        } | Sort-Object TimeCreated | Select-Object -First 1)

        if (-not $terminal) {
            $activeStart = $start
            break
        }
    }

    $current = $null
    if ($activeStart) {
        $current = $activeStart
    } elseif ($eventsArray.Count -gt 0) {
        $current = $eventsArray[0]
    }

    $featureTarget = ""
    $isFeature = $false
    if ($current -and $current.Title -match '(?i)(Windows 11, version|Feature update to Windows).*?([0-9]{2}H[0-9])') {
        $isFeature = $true
        $featureTarget = $Matches[2].ToUpper()
    }

    return [PSCustomObject]@{
        Current       = $current
        ActiveStart   = $activeStart
        IsFeature     = $isFeature
        FeatureTarget = $featureTarget
    }
}

function Get-PEEKProcessSnapshot {
    $names = @(
        'Windows11InstallationAssistant',
        'setup',
        'setuphost',
        'setupprep',
        'wuaulcore',
        'TiWorker',
        'TrustedInstaller',
        'MoUsoCoreWorker',
        'UsoClient'
    )

    $priority = @{
        'Windows11InstallationAssistant' = 1
        'setup'                          = 2
        'setuphost'                      = 3
        'setupprep'                      = 4
        'wuaulcore'                      = 5
        'TiWorker'                       = 6
        'TrustedInstaller'               = 7
        'MoUsoCoreWorker'                = 8
        'UsoClient'                      = 9
    }

    $now = Get-Date
    $elapsed = 0.0
    if ($script:PreviousSampleTime) {
        $elapsed = ($now - $script:PreviousSampleTime).TotalSeconds
    }

    try { $processes = @(Get-Process -Name $names -ErrorAction SilentlyContinue) } catch { $processes = @() }
    $nextCpu = @{}
    $rows = @()
    $cpuCount = [Math]::Max(1, [Environment]::ProcessorCount)

    foreach ($process in $processes) {
        $cpuTotal = $null
        try { $cpuTotal = [double]$process.TotalProcessorTime.TotalSeconds } catch {}

        $cpuPct = $null
        if ($null -ne $cpuTotal -and $elapsed -gt 0.25 -and $script:PreviousCpu.ContainsKey($process.Id)) {
            $delta = $cpuTotal - [double]$script:PreviousCpu[$process.Id]
            if ($delta -lt 0) { $delta = 0 }
            $cpuPct = [Math]::Round(($delta / $elapsed / $cpuCount) * 100, 1)
        }

        if ($null -ne $cpuTotal) { $nextCpu[$process.Id] = $cpuTotal }

        $start = $null
        try { $start = [datetime]$process.StartTime } catch {}

        $rank = 99
        if ($priority.ContainsKey($process.ProcessName)) { $rank = [int]$priority[$process.ProcessName] }

        $rows += [PSCustomObject]@{
            Name     = $process.ProcessName
            Id       = $process.Id
            CPU      = $cpuPct
            RAMMB    = [Math]::Round($process.WorkingSet64 / 1MB, 1)
            Start    = $start
            Priority = $rank
        }
    }

    $script:PreviousCpu        = $nextCpu
    $script:PreviousSampleTime = $now

    return @($rows | Sort-Object Priority, Name, Id)
}

function Get-PEEKSetupState {
    $actPaths = @(
        "$env:SystemDrive\`$WINDOWS.~BT\Sources\Panther\setupact.log",
        "$env:SystemDrive\`$WINDOWS.~BT\Sources\Rollback\setupact.log",
        "$env:SystemRoot\Panther\setupact.log"
    )

    $errPaths = @(
        "$env:SystemDrive\`$WINDOWS.~BT\Sources\Panther\setuperr.log",
        "$env:SystemDrive\`$WINDOWS.~BT\Sources\Rollback\setuperr.log",
        "$env:SystemRoot\Panther\setuperr.log"
    )

    $actFiles = @()
    foreach ($path in $actPaths) {
        if (Test-Path -LiteralPath $path) {
            try { $actFiles += Get-Item -LiteralPath $path -ErrorAction Stop } catch {}
        }
    }

    $errFiles = @()
    foreach ($path in $errPaths) {
        if (Test-Path -LiteralPath $path) {
            try { $errFiles += Get-Item -LiteralPath $path -ErrorAction Stop } catch {}
        }
    }

    $act = $actFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $err = $errFiles | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    $mode = "None"
    $commandLine = ""
    $featureCode = ""
    $featureText = ""

    if ($act) {
        try {
            $tail = @(Get-Content -LiteralPath $act.FullName -Tail 350 -ErrorAction SilentlyContinue)
            for ($i = $tail.Count - 1; $i -ge 0; $i--) {
                if (-not $featureCode -and $tail[$i] -match '(0xC190[0-9A-Fa-f]{4})') {
                    $featureCode = $Matches[1].ToUpper()
                    break
                }
            }
        } catch {}

        $deepScanDue = ($script:SetupCommandLinePath -ne $act.FullName) -or
                       (-not $script:SetupCommandLine) -or
                       (((Get-Date) - $script:SetupDeepScanTime).TotalSeconds -ge 15)

        if ($deepScanDue) {
            try {
                $deepTail = @(Get-Content -LiteralPath $act.FullName -Tail 50000 -ErrorAction SilentlyContinue)
                $foundCommandLine = ""
                for ($i = $deepTail.Count - 1; $i -ge 0; $i--) {
                    if ($deepTail[$i] -match 'SetupHost::Initialize:\s+CmdLine\s+=\s+\[(.+)\]') {
                        $foundCommandLine = $Matches[1].Trim()
                        break
                    }
                }
                if ($foundCommandLine) {
                    $script:SetupCommandLine = $foundCommandLine
                    $script:SetupCommandLinePath = $act.FullName
                }
                $script:SetupDeepScanTime = Get-Date
            } catch {}
        }

        if ($script:SetupCommandLinePath -eq $act.FullName) {
            $commandLine = $script:SetupCommandLine
        }
    }

    $useErrorLog = $false
    if ($err) {
        if (-not $act) {
            $useErrorLog = $true
        } else {
            $logDeltaMinutes = [Math]::Abs(($act.LastWriteTime - $err.LastWriteTime).TotalMinutes)
            if ($logDeltaMinutes -le 5) { $useErrorLog = $true }
        }
    }

    if ($useErrorLog -and -not $featureCode) {
        try {
            $tailErr = @(Get-Content -LiteralPath $err.FullName -Tail 250 -ErrorAction SilentlyContinue)
            for ($i = $tailErr.Count - 1; $i -ge 0; $i--) {
                if ($tailErr[$i] -match '(0xC190[0-9A-Fa-f]{4})') {
                    $featureCode = $Matches[1].ToUpper()
                    break
                }
            }
        } catch {}
    }

    if ($commandLine) {
        if ($commandLine -match '(?i)/Compat\s+ScanOnly') {
            $mode = "Compatibility Scan"
        } elseif ($commandLine -match '(?i)/PreDownload') {
            $mode = "PreDownload / Package"
        } elseif ($commandLine -match '(?i)/auto\s+upgrade') {
            $mode = "In-Place Upgrade"
        } elseif ($commandLine -match '(?i)/Finalize') {
            $mode = "Finalize"
        } else {
            $mode = "Windows Setup"
        }
    }

    if ($featureCode) {
        $featureText = switch ($featureCode) {
            '0XC1900208' { 'App or driver compatibility block' }
            '0XC1900200' { 'Minimum requirements block' }
            '0XC1900204' { 'Migration choice / edition-language block' }
            '0XC1900101' { 'Driver or rollback failure' }
            default      { 'Windows Setup feature-upgrade error' }
        }
    }

    $active = $false
    $actAgeSeconds = $null
    if ($act) {
        $actAgeSeconds = ((Get-Date) - $act.LastWriteTime).TotalSeconds
        if ($actAgeSeconds -le 15) { $active = $true }
    }

    return [PSCustomObject]@{
        ActFile       = $act
        ErrFile       = $err
        Active        = $active
        ActAgeSeconds = $actAgeSeconds
        Mode          = $mode
        CommandLine   = $commandLine
        FeatureCode   = $featureCode
        FeatureText   = $featureText
    }
}

function Get-PEEKCompatState {
    $patterns = @(
        "$env:SystemDrive\`$WINDOWS.~BT\Sources\Panther\CompatData*.xml",
        "$env:SystemRoot\Panther\CompatData*.xml"
    )

    $files = @()
    foreach ($pattern in $patterns) {
        try { $files += @(Get-ChildItem -Path $pattern -ErrorAction SilentlyContinue) } catch {}
    }

    $file = $files | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if (-not $file) {
        return [PSCustomObject]@{
            File          = $null
            LastWriteTime = [datetime]::MinValue
            BlockCount    = 0
            Summary       = "No compatibility report"
            HasReport     = $false
        }
    }

    $cacheKey = "$($file.FullName)|$($file.LastWriteTime.Ticks)|$($file.Length)"
    if ($script:CompatCacheKey -eq $cacheKey -and $script:CompatCacheValue) {
        return $script:CompatCacheValue
    }

    $blocks = @()
    try {
        [xml]$xml = Get-Content -LiteralPath $file.FullName -Raw -ErrorAction Stop

        $hardNodes = @($xml.SelectNodes("//*[@BlockingType='Hard']"))
        foreach ($node in $hardNodes) {
            $title = ""
            try { $title = [string]$node.GetAttribute('Title') } catch {}
            if (-not $title -and $node.ParentNode) {
                try { $title = [string]$node.ParentNode.GetAttribute('Name') } catch {}
            }
            if (-not $title) { $title = "Hard compatibility block" }
            $blocks += "HARD: $title"
        }

        $driverNodes = @($xml.SelectNodes("//DriverPackage[@BlockMigration='True']"))
        foreach ($node in $driverNodes) {
            $inf = ""
            try { $inf = [string]$node.GetAttribute('Inf') } catch {}
            if (-not $inf) { $inf = "Unknown INF" }
            $blocks += "DRIVER: $inf"
        }
    } catch {
        $blocks += "Compatibility report could not be parsed"
    }

    $blocks = @($blocks | Select-Object -Unique)
    $summary = if ($blocks.Count -gt 0) { $blocks -join '; ' } else { "None detected" }

    $value = [PSCustomObject]@{
        File          = $file
        LastWriteTime = [datetime]$file.LastWriteTime
        BlockCount    = $blocks.Count
        Summary       = $summary
        HasReport     = $true
    }

    $script:CompatCacheKey   = $cacheKey
    $script:CompatCacheValue = $value
    return $value
}

function Get-PEEKRebootState {
    $cbs = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending'
    $wu  = Test-Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired'
    $mo  = Test-Path 'HKLM:\SYSTEM\Setup\MoSetup\Volatile'
    $pendingXml = Test-Path "$env:SystemRoot\WinSxS\pending.xml"

    $fileRename = $false
    try {
        $sessionManager = Get-ItemProperty 'HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager' -Name PendingFileRenameOperations -ErrorAction SilentlyContinue
        if ($null -ne $sessionManager.PendingFileRenameOperations) { $fileRename = $true }
    } catch {}

    $updateExe = $false
    try {
        $updateVolatile = (Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Updates' -Name UpdateExeVolatile -ErrorAction SilentlyContinue).UpdateExeVolatile
        if ($null -ne $updateVolatile -and [int]$updateVolatile -ne 0) { $updateExe = $true }
    } catch {}

    $sources = @()
    if ($wu)         { $sources += 'WU' }
    if ($cbs)        { $sources += 'CBS' }
    if ($mo)         { $sources += 'MoSetup' }
    if ($fileRename) { $sources += 'FileRename' }
    if ($updateExe)  { $sources += 'UpdateExe' }
    if ($pendingXml) { $sources += 'pending.xml' }

    return [PSCustomObject]@{
        WindowsUpdate = $wu
        CBS           = $cbs
        MoSetup       = $mo
        FileRename    = $fileRename
        UpdateExe     = $updateExe
        PendingXml    = $pendingXml
        Any           = ($sources.Count -gt 0)
        Sources       = $sources
    }
}

function Get-PEEKServices {
    $definitions = @(
        @{ Name = 'wuauserv';         Label = 'WU'   },
        @{ Name = 'BITS';             Label = 'BITS' },
        @{ Name = 'UsoSvc';           Label = 'USO'  },
        @{ Name = 'TrustedInstaller'; Label = 'TI'   },
        @{ Name = 'DoSvc';            Label = 'DO'   }
    )

    $rows = @()
    foreach ($definition in $definitions) {
        $status = 'Missing'
        try {
            $service = Get-Service -Name $definition.Name -ErrorAction Stop
            $status = [string]$service.Status
        } catch {}

        $rows += [PSCustomObject]@{
            Name   = $definition.Name
            Label  = $definition.Label
            Status = $status
        }
    }

    return $rows
}

function Get-PEEKReadiness {
    $diskText = "Unavailable"
    $diskColor = $script:MutedFG
    try {
        $disk = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$env:SystemDrive'" -ErrorAction Stop
        if ($disk -and $disk.Size -gt 0) {
            $freeGB = [Math]::Round($disk.FreeSpace / 1GB, 1)
            $freePct = [Math]::Round(($disk.FreeSpace / $disk.Size) * 100, 0)
            $diskText = "$freeGB GB free ($freePct%)"
            if ($freeGB -lt 20) { $diskColor = $script:FailFG }
            elseif ($freeGB -lt 30) { $diskColor = $script:WarnFG }
            else { $diskColor = $script:DoneFG }
        }
    } catch {}

    $powerText = "AC / no battery detected"
    $powerColor = $script:DoneFG
    try {
        $battery = Get-CimInstance Win32_Battery -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($battery) {
            $charge = [int]$battery.EstimatedChargeRemaining
            $isAC = $battery.BatteryStatus -in @(2,3,6,7,8,9,11)
            if ($isAC) {
                $powerText = "AC power, battery $charge%"
                $powerColor = $script:DoneFG
            } else {
                $powerText = "BATTERY, $charge% remaining"
                $powerColor = if ($charge -lt 30) { $script:FailFG } else { $script:WarnFG }
            }
        }
    } catch {}

    return [PSCustomObject]@{
        DiskText   = $diskText
        DiskColor  = $diskColor
        PowerText  = $powerText
        PowerColor = $powerColor
    }
}

function Get-PEEKBitLockerState {
    if (-not (Get-Command Get-BitLockerVolume -ErrorAction SilentlyContinue)) { return "Unavailable" }
    try {
        $volume = Get-BitLockerVolume -MountPoint $env:SystemDrive -ErrorAction Stop
        if (-not $volume) { return "Unavailable" }
        return "Protection $($volume.ProtectionStatus), Volume $($volume.VolumeStatus)"
    } catch {
        return "Unavailable"
    }
}

#endregion

#region [3] - STATE CORRELATION
# ============================================================================

function Get-PEEKSystemState {
    param(
        [object]$OS,
        [object]$UpdateContext,
        [object[]]$Processes,
        [object]$Setup,
        [object]$Compat,
        [object]$Reboot
    )

    $now = Get-Date
    $processNames = @($Processes | ForEach-Object { $_.Name })
    $strongInstallProcess = @($processNames | Where-Object {
        $_ -match '^(Windows11InstallationAssistant|setup|setuphost|setupprep|wuaulcore|TiWorker)$'
    }).Count -gt 0
    $trustedInstallerProcess = @($processNames | Where-Object { $_ -eq 'TrustedInstaller' }).Count -gt 0
    $orchestratorProcess = @($processNames | Where-Object { $_ -match '^(MoUsoCoreWorker|UsoClient)$' }).Count -gt 0

    $current = $UpdateContext.Current
    $activeStart = $UpdateContext.ActiveStart
    $installProcess = $strongInstallProcess -or ($trustedInstallerProcess -and ($activeStart -or $Setup.Active))
    $targetInstalled = $false
    if ($UpdateContext.FeatureTarget -and $OS.DisplayVersion) {
        $targetInstalled = ($OS.DisplayVersion.ToUpper() -eq $UpdateContext.FeatureTarget.ToUpper())
    }

    $currentAgeHours = 9999
    if ($current) { $currentAgeHours = ($now - $current.TimeCreated).TotalHours }

    $compatFresh = $false
    if ($Compat.HasReport -and $Compat.BlockCount -gt 0) {
        if ($activeStart) {
            $compatFresh = ($Compat.LastWriteTime -ge $activeStart.TimeCreated.AddMinutes(-5))
        } else {
            $compatFresh = (($now - $Compat.LastWriteTime).TotalHours -le 12)
        }
    }

    $state = "IDLE"
    $stateColor = $DimCol
    $rebootAction = "NO REBOOT REQUIRED"
    $rebootColor = $script:DoneFG

    if ($targetInstalled) {
        $state = "FEATURE UPDATE COMPLETE - $($UpdateContext.FeatureTarget) INSTALLED"
        $stateColor = $script:DoneFG
    }
    elseif ($current -and $current.Id -eq 214 -and $currentAgeHours -le 6) {
        $state = "ROLLBACK ACTIVE"
        $stateColor = $script:FailFG
    }
    elseif ($current -and $current.Id -eq 212 -and $currentAgeHours -le 24) {
        $state = "UPDATE ROLLED BACK"
        $stateColor = $script:FailFG
    }
    elseif ($current -and $current.Id -eq 20 -and $currentAgeHours -le 12 -and -not $activeStart) {
        $state = "UPDATE FAILED"
        if ($current.ErrorCode) { $state += " - $($current.ErrorCode)" }
        $stateColor = $script:FailFG
    }
    elseif ($activeStart -and ($installProcess -or $Setup.Active)) {
        if ($UpdateContext.IsFeature) { $state = "FEATURE UPGRADE ACTIVE" }
        else { $state = "UPDATE INSTALLING" }
        $stateColor = $script:ActiveFG
    }
    elseif ($installProcess -or $Setup.Active) {
        $state = "WINDOWS SERVICING ACTIVE"
        $stateColor = $script:ActiveFG
    }
    elseif ($activeStart) {
        $state = "INSTALL STARTED - NO LIVE ACTIVITY DETECTED"
        $stateColor = $script:WarnFG
    }
    elseif ($compatFresh) {
        $state = "COMPATIBILITY BLOCK DETECTED"
        $stateColor = $script:FailFG
    }
    elseif ($Reboot.Any -and $current -and $current.Id -eq 19) {
        $state = "READY TO REBOOT"
        $stateColor = $script:WarnFG
    }
    elseif ($Reboot.Any) {
        $state = "REBOOT PENDING"
        $stateColor = $script:WarnFG
    }
    elseif ($current -and $current.Id -eq 19 -and $currentAgeHours -le 12) {
        $state = "UPDATE INSTALLED"
        $stateColor = $script:DoneFG
    }
    elseif ($orchestratorProcess) {
        $state = "WINDOWS UPDATE ACTIVE"
        $stateColor = $AccentCol
    }

    if ($activeStart -or $installProcess -or $Setup.Active) {
        $rebootAction = "DO NOT REBOOT - SERVICING ACTIVE OR INCOMPLETE"
        $rebootColor = $script:FailFG
    }
    elseif ($state -eq "READY TO REBOOT") {
        $rebootAction = "READY TO REBOOT - NO ACTIVE SERVICING DETECTED"
        $rebootColor = $script:WarnFG
    }
    elseif ($Reboot.Any) {
        $rebootAction = "REBOOT PENDING - VERIFY CHANGE WINDOW"
        $rebootColor = $script:WarnFG
    }

    return [PSCustomObject]@{
        State             = $state
        StateColor        = $stateColor
        RebootAction      = $rebootAction
        RebootColor       = $rebootColor
        InstallProcess    = $installProcess
        Orchestrator      = $orchestratorProcess
        CompatibilityFresh = $compatFresh
    }
}

#endregion

#region [4] - LIVE VIEW
# ============================================================================

function Write-ProcessTable {
    param([object[]]$Processes)

    $rows = @($Processes)
    if ($rows.Count -eq 0) {
        Write-Host "  No Windows servicing processes detected." -ForegroundColor $script:MutedFG
        return
    }

    $nameWidth = [Math]::Min(32, [Math]::Max(22, $script:Width - 39))

    Write-Host ("  {0,-$nameWidth} {1,7} {2,8} {3,10}" -f "Process", "PID", "CPU", "RAM MB") -ForegroundColor $AccentCol
    foreach ($row in $rows) {
        $name = Get-TruncatedText -Text $row.Name -MaxLength $nameWidth
        $cpuText = if ($null -eq $row.CPU) { "--" } else { "$($row.CPU)%" }
        Write-Host ("  {0,-$nameWidth} {1,7} {2,8} {3,10:N1}" -f $name, $row.Id, $cpuText, $row.RAMMB) -ForegroundColor $script:ValueFG
    }
}

function Write-ServiceLine {
    param([object[]]$Services)

    Write-Host "  Services          : " -ForegroundColor $script:LabelFG -NoNewline
    $items = @($Services)
    for ($i = 0; $i -lt $items.Count; $i++) {
        $service = $items[$i]
        Write-Host "$($service.Label)=" -ForegroundColor $DimCol -NoNewline
        $color = if ($service.Status -eq 'Running') { $script:DoneFG } elseif ($service.Status -eq 'Stopped') { $script:MutedFG } else { $script:WarnFG }
        Write-Host $service.Status -ForegroundColor $color -NoNewline
        if ($i -lt ($items.Count - 1)) { Write-Host "  |  " -ForegroundColor $script:MutedFG -NoNewline }
    }
    Write-Host ""
}

function Write-RebootFlags {
    param([object]$Reboot)

    $flags = @(
        @{ Label = 'WU';         Value = $Reboot.WindowsUpdate },
        @{ Label = 'CBS';        Value = $Reboot.CBS },
        @{ Label = 'MoSetup';    Value = $Reboot.MoSetup },
        @{ Label = 'FileRename'; Value = $Reboot.FileRename },
        @{ Label = 'pending.xml';Value = $Reboot.PendingXml }
    )

    Write-Host "  Reboot Flags      : " -ForegroundColor $script:LabelFG -NoNewline
    for ($i = 0; $i -lt $flags.Count; $i++) {
        $flag = $flags[$i]
        Write-Host "$($flag.Label)=" -ForegroundColor $DimCol -NoNewline
        if ($flag.Value) { Write-Host "YES" -ForegroundColor $script:WarnFG -NoNewline }
        else { Write-Host "no" -ForegroundColor $script:MutedFG -NoNewline }
        if ($i -lt ($flags.Count - 1)) { Write-Host "  |  " -ForegroundColor $script:MutedFG -NoNewline }
    }
    Write-Host ""
}

function Show-PEEKSnapshot {
    Update-ConsoleWidth

    $now = Get-Date
    $cv = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction SilentlyContinue
    $osCim = Get-CimInstance Win32_OperatingSystem -ErrorAction SilentlyContinue

    $displayVersion = if ($cv.DisplayVersion) { [string]$cv.DisplayVersion } elseif ($cv.ReleaseId) { [string]$cv.ReleaseId } else { "Unknown" }
    $build = if ($cv.CurrentBuild) { [string]$cv.CurrentBuild } else { [string]$osCim.BuildNumber }
    $ubr = if ($null -ne $cv.UBR) { [string]$cv.UBR } else { "0" }
    $caption = if ($osCim.Caption) { [string]$osCim.Caption } else { "Windows" }
    $lastBoot = if ($osCim.LastBootUpTime) { [datetime]$osCim.LastBootUpTime } else { [datetime]::MinValue }
    $uptime = if ($lastBoot -ne [datetime]::MinValue) { $now - $lastBoot } else { [timespan]::Zero }

    $os = [PSCustomObject]@{
        DisplayVersion = $displayVersion
        Build          = $build
        UBR            = $ubr
        Caption        = $caption
        LastBoot       = $lastBoot
        Uptime         = $uptime
    }

    $events        = @(Get-PEEKUpdateEvents)
    $updateContext = Get-PEEKUpdateContext -Events $events
    $processes     = @(Get-PEEKProcessSnapshot)
    $setup         = Get-PEEKSetupState
    $compat        = Get-PEEKCompatState
    $reboot        = Get-PEEKRebootState
    $services      = @(Get-PEEKServices)
    $readiness     = Get-PEEKReadiness
    $state         = Get-PEEKSystemState -OS $os -UpdateContext $updateContext -Processes $processes -Setup $setup -Compat $compat -Reboot $reboot

    Clear-Host
    Write-Banner

    Write-Detail -Label "Device Name" -Value $env:COMPUTERNAME -ValueColor $script:ValueFG
    Write-Detail -Label "Operating System" -Value "$caption $displayVersion (Build $build.$ubr)" -ValueColor $script:ValueFG
    Write-Detail -Label "Uptime" -Value (Format-Uptime -Uptime $uptime) -ValueColor $DimCol
    Write-Detail -Label "Overall Status" -Value $state.State -ValueColor $state.StateColor

    if ($updateContext.Current) {
        Write-Detail -Label "Current Update" -Value $updateContext.Current.Title -ValueColor $(if ($updateContext.ActiveStart) { $script:ActiveFG } else { $script:ValueFG })
        $latestEventText = "ID $($updateContext.Current.Id) $($updateContext.Current.Label) at $($updateContext.Current.TimeCreated.ToString('yyyy-MM-dd HH:mm:ss'))"
        if ($updateContext.Current.ErrorCode) { $latestEventText += " $($updateContext.Current.ErrorCode)" }
        Write-Detail -Label "Latest WU Event" -Value $latestEventText -ValueColor $(if ($updateContext.Current.Id -eq 20) { $script:FailFG } elseif ($updateContext.Current.Id -eq 19) { $script:DoneFG } else { $AccentCol })
        if ($updateContext.ActiveStart) {
            Write-Detail -Label "Install Elapsed" -Value (Format-Age -Time $updateContext.ActiveStart.TimeCreated) -ValueColor $script:ActiveFG
        }
    } else {
        Write-Detail -Label "Current Update" -Value "No Windows Update event detected in the last 48 hours" -ValueColor $script:MutedFG
    }

    Write-HLine -Style dashed
    Write-SectionTitle "SERVICING ACTIVITY"
    Write-ProcessTable -Processes $processes

    if ($setup.ActFile) {
        $setupAge = Format-Age -Time $setup.ActFile.LastWriteTime
        $setupColor = if ($setup.Active) { $script:DoneFG } else { $script:MutedFG }
        Write-Detail -Label "Panther setupact" -Value "$(Split-Path $setup.ActFile.FullName -Leaf) updated $setupAge ago" -ValueColor $setupColor
    } else {
        Write-Detail -Label "Panther setupact" -Value "Not present" -ValueColor $script:MutedFG
    }

    Write-Detail -Label "Setup Mode" -Value $setup.Mode -ValueColor $(if ($setup.Active) { $AccentCol } else { $DimCol })

    if ($setup.FeatureCode) {
        $codeValue = "$($setup.FeatureCode) - $($setup.FeatureText)"
        Write-Detail -Label "Setup Code" -Value $codeValue -ValueColor $script:FailFG
    } else {
        Write-Detail -Label "Setup Code" -Value "No 0xC190 feature-upgrade code detected" -ValueColor $script:MutedFG
    }

    if ($compat.HasReport) {
        $compatAge = Format-Age -Time $compat.LastWriteTime
        $compatValue = "$($compat.Summary)  [report age $compatAge]"
        $compatColor = if ($compat.BlockCount -gt 0) { $script:FailFG } else { $script:DoneFG }
        Write-Detail -Label "Compatibility" -Value $compatValue -ValueColor $compatColor
    } else {
        Write-Detail -Label "Compatibility" -Value $compat.Summary -ValueColor $script:MutedFG
    }

    Write-ServiceLine -Services $services

    Write-HLine -Style dashed
    Write-SectionTitle "REBOOT & READINESS"
    Write-RebootFlags -Reboot $reboot
    Write-Detail -Label "Reboot Action" -Value $state.RebootAction -ValueColor $state.RebootColor
    Write-Detail -Label "$env:SystemDrive Free Space" -Value $readiness.DiskText -ValueColor $readiness.DiskColor
    Write-Detail -Label "Power" -Value $readiness.PowerText -ValueColor $readiness.PowerColor
    Write-Detail -Label "BitLocker" -Value $script:BitLockerState -ValueColor $DimCol

    Write-HLine -Style dashed
    Write-SectionTitle "RECENT WINDOWS UPDATE EVENTS"

    $timeline = @($events | Select-Object -First 5)
    if ($timeline.Count -eq 0) {
        Write-StepUpdate -Status "INFO" -Message "No WindowsUpdateClient events found in the last 48 hours."
    } else {
        foreach ($wuEvent in $timeline) {
            $message = "$($wuEvent.TimeCreated.ToString('HH:mm:ss'))  $($wuEvent.Label): $($wuEvent.Title)"
            $detail = "ID $($wuEvent.Id)"
            if ($wuEvent.ErrorCode) { $detail += " $($wuEvent.ErrorCode)" }
            Write-StepUpdate -Status $wuEvent.Status -Message $message -Detail $detail
        }
    }

    Write-HLine -Style dashed
    $frames = @('|','/','-','\')
    $frame = $frames[$script:FrameIndex % $frames.Count]
    $script:FrameIndex++
    $liveText = "LIVE $frame  Updated $($now.ToString('HH:mm:ss'))  |  Refresh ${RefreshSeconds}s  |  Q / ESC quit  |  READ-ONLY"
    Write-Host ("  " + (Get-TruncatedText -Text $liveText -MaxLength ($script:Width - 2))) -ForegroundColor $AccentCol
}

$script:BitLockerState = Get-PEEKBitLockerState
$quit = $false

while (-not $quit) {
    Show-PEEKSnapshot

    if ($Once) { break }

    $deadline = (Get-Date).AddSeconds($RefreshSeconds)
    while ((Get-Date) -lt $deadline -and -not $quit) {
        try {
            if ([Console]::KeyAvailable) {
                $key = [Console]::ReadKey($true)
                if ($key.Key -eq [ConsoleKey]::Q -or $key.Key -eq [ConsoleKey]::Escape) {
                    $quit = $true
                    break
                }
            }
        } catch {}
        Start-Sleep -Milliseconds 100
    }
}

#endregion

#region [5] - EXIT
# ============================================================================

try {
    Write-Footer -Text "READ-ONLY LIVE SERVICING MONITOR"
} catch {}

exit $EXIT_SUCCESS
#endregion