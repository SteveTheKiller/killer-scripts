<#
.SYNOPSIS
    Microsoft Application Cleanse & Eradication (MACE) v1.1
    Developed by SteveTheKiller | Updated: 2026-03-06
.DESCRIPTION
    Completely removes OneDrive, New Outlook, Office/M365, Microsoft
    Project, and Microsoft Teams from a Windows endpoint. Clears all
    associated registry keys, cached credentials, profile data, and
    temp folders. Repairs OneDrive shell folder redirects and path
    pollution. Supports an interactive per-user or system-wide scope
    selection with ESC-to-abort.
#>

#region 0: HELPERS
# ============================================================================
function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

# --- System boundaries (thin wrappers so tests can replace them) ---
function Test-MaceElevated {
    ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")
}
function Test-MaceInteractive {
    try { [Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected } catch { $false }
}
function Test-MaceConsole {
    try { -not [Console]::IsOutputRedirected -and ([Console]::CursorTop -ge 0) } catch { $false }
}
function Get-MaceCursorTop  { try { [Console]::CursorTop } catch { -1 } }
function Get-MaceCursorLeft { try { [Console]::CursorLeft } catch { -1 } }
function Set-MaceCursor {
    param([int]$Left, [int]$Top)
    try { [Console]::SetCursorPosition($Left, $Top) } catch { }
}
function Read-MaceKey {
    try { $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") } catch { $null }
}
function Get-MaceProfileDirs {
    Get-ChildItem "C:\Users" -Directory -ErrorAction SilentlyContinue
}
function Get-MaceProfileList {
    @(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\*" -ErrorAction SilentlyContinue |
        Where-Object { $_.PSChildName -match "S-1-5-21" })
}
function Resolve-MaceAccount {
    param([string]$Sid)
    try { (New-Object Security.Principal.SecurityIdentifier($Sid)).Translate([Security.Principal.NTAccount]).Value } catch { $null }
}
function Get-MaceProcess {
    param([string]$Name)
    @(Get-Process -Name $Name -IncludeUserName -ErrorAction SilentlyContinue)
}
function Get-MaceUninstallEntries {
    @(Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*", "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*" -ErrorAction SilentlyContinue)
}
function Invoke-MaceWinget {
    param([string[]]$Arguments)
    & winget @Arguments 2>$null | Out-Null
}
function Write-MaceFile {
    param([string]$Path, [string]$Text)
    Set-Content -LiteralPath $Path -Value $Text -Encoding UTF8
}
function Test-MaceWinget {
    $null -ne (Get-Command winget -ErrorAction SilentlyContinue)
}
function Get-MaceAppx {
    param([string]$Name, [string]$Sid)
    if ($Sid) { @(Get-AppxPackage -User $Sid -Name $Name -ErrorAction SilentlyContinue) }
    else      { @(Get-AppxPackage -AllUsers -Name $Name -ErrorAction SilentlyContinue) }
}
function Remove-MaceAppx {
    param([string]$Name, [string]$Sid)
    if ($Sid) {
        foreach ($pkg in @(Get-AppxPackage -User $Sid -Name $Name -ErrorAction SilentlyContinue)) {
            Remove-AppxPackage -Package $pkg.PackageFullName -User $Sid -ErrorAction SilentlyContinue
        }
    } else {
        Get-AppxPackage -AllUsers -Name $Name -ErrorAction SilentlyContinue | Remove-AppxPackage -AllUsers -ErrorAction SilentlyContinue
    }
}
function Remove-MaceProvisioned {
    param([string]$Like)
    Get-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue | Where-Object { $_.DisplayName -like $Like } | Remove-AppxProvisionedPackage -Online -ErrorAction SilentlyContinue
}

# --- Output ---
$script:StepRow = -1
$script:StepMsg = ""
function Write-StepUpdate {
    param([string]$Message, [switch]$Success, [string]$CustomInfo, [switch]$Pending)
    if ($Success) {
        $savedTop = Get-MaceCursorTop
        if ($savedTop -lt 0) {
            Write-Host "      [SUCCESS]" -ForegroundColor $OkCol
        } elseif ($script:StepRow -ge 0 -and $script:StepRow -lt $savedTop) {
            Set-MaceCursor 0 $script:StepRow
            $_pad = " " * [math]::Max(1, $script:Width - $script:StepMsg.Length - "[SUCCESS]".Length)
            Write-Host $script:StepMsg -ForegroundColor $DimCol -NoNewline
            Write-Host "$_pad[SUCCESS]" -ForegroundColor $OkCol
            Set-MaceCursor 0 $savedTop
        }
        $script:StepRow = -1
    } else {
        $script:StepRow = Get-MaceCursorTop
        $script:StepMsg = $Message
        Write-Host $Message -ForegroundColor $InfoCol
    }
}

function Write-SubResult {
    param([string]$Tag, [string]$Color)
    $left = Get-MaceCursorLeft
    if ($left -lt 0) {
        Write-Host " $Tag" -ForegroundColor $Color
    } else {
        $_pad = " " * [math]::Max(1, $script:Width - $left - $Tag.Length)
        Write-Host "$_pad$Tag" -ForegroundColor $Color
    }
}

function Write-HLine {
    param(
        [string]$Style = "dashed",
        [int]$Width    = $script:Width
    )
    if ($Style -eq "dashed") {
        $line = ("- " * [math]::Ceiling($Width / 2)).Substring(0, $Width)
    } else {
        $line = "━" * $Width
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
function Write-Banner {
    param([string]$Text, [switch]$NoNewline)
    $esc = [char]27
    if ($script:BannerBG -eq "Black") {
        Write-Host "${esc}[40m$Text${esc}[0m" -NoNewline:$NoNewline -ForegroundColor $script:BannerFG
    } else {
        Write-Host $Text -NoNewline:$NoNewline -ForegroundColor $script:BannerFG -BackgroundColor $script:BannerBG
    }
}

function Invoke-PerUserRegistry {
    param([string]$HiveRoot, [string]$ProfilePath)
    foreach ($key in $UserRegKeys) {
        $full = "$HiveRoot\$key"
        if (Test-Path $full) { Remove-Item $full -Recurse -Force -ErrorAction SilentlyContinue }
    }
    $runPath = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Run"
    foreach ($val in $UserRunValues) {
        if (Get-ItemProperty -Path $runPath -Name $val -ErrorAction SilentlyContinue) {
            Remove-ItemProperty -Path $runPath -Name $val -ErrorAction SilentlyContinue
        }
    }
    $USFPath = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
    $SFPath  = "$HiveRoot\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
    foreach ($fix in $ShellFixes) {
        $current = (Get-ItemProperty -Path $USFPath -Name $fix.Name -ErrorAction SilentlyContinue).$($fix.Name)
        if ($current -like "*OneDrive*") {
            Set-ItemProperty -Path $USFPath -Name $fix.Name -Value $fix.Default -ErrorAction SilentlyContinue
            if ($ProfilePath) {
                $absPath = Join-Path $ProfilePath $fix.Dir
                Set-ItemProperty -Path $SFPath -Name $fix.Name -Value $absPath -ErrorAction SilentlyContinue
                New-Item -ItemType Directory -Path $absPath -Force -ErrorAction SilentlyContinue | Out-Null
            }
        }
    }
    $navKey = "$HiveRoot\Software\Classes\CLSID\{018D5C66-4533-4307-9B53-224DE2ED1FE6}"
    if (Test-Path $navKey) {
        Set-ItemProperty -Path $navKey -Name "System.IsPinnedToNameSpaceTree" -Value 0 -ErrorAction SilentlyContinue
    }
}
#endregion

function Invoke-MaceRun {
    #region 1: INITIALIZATION
    # ========================================================================
    if (-not (Test-MaceElevated)) {
        Write-Warning "Elevation Required: Please run as Administrator."
        $script:MaceExit = 2
        $script:MaceNote = "not elevated; nothing was changed"
        return
    }

    if (Test-MaceConsole) { try { Clear-Host } catch { } }
    $script:Width = 85

    $LineCol   = "Red"
    $MainCol   = "Gray"
    $WarnCol   = "DarkYellow"
    $BorderCol = "Red"
    $ArtCol    = "White"
    $AccentCol = "Red"
    $DimCol    = "DarkGray"
    $OkCol     = "Green"
    $InfoCol   = "Cyan"
    $_ver    = "| v1.2"

    # Header Art & Logic
    $_pfx  = "█  "
    $_art1 = "╔╦╗ ╔═╗ ╔═╗ ╔═╗ "
    $_art2 = "║║║ ╠═╣ ║   ║╣  "
    $_art3 = "╩ ╩ ╩ ╩ ╚═╝ ╚═╝ "
    $_artW = [Math]::Max($_art1.Length, [Math]::Max($_art2.Length, $_art3.Length))
    $_art1 = $_art1.PadRight($_artW); $_art2 = $_art2.PadRight($_artW); $_art3 = $_art3.PadRight($_artW)
    $_fillW = $script:Width - $_pfx.Length - $_artW
    $_title = "MICROSOFT APPLICATION CLEANSE & ERADICATION"

    # Safety padding to prevent negative multiplication error
    $_tpad  = " " * [Math]::Max(0, ($_fillW - $_title.Length))

    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art1 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art2 -ForegroundColor $ArtCol -NoNewline; Write-Host "$_title$_tpad" -ForegroundColor $MainCol
    Write-Host $_pfx -ForegroundColor $LineCol -NoNewline; Write-Host $_art3 -ForegroundColor $ArtCol -NoNewline; Write-Host ("-" * $_fillW) -ForegroundColor $LineCol
    #endregion

    #region 2: SCOPE SELECTION
    # ========================================================================
    $IsInteractive = Test-MaceInteractive
    $ExcludedProfiles = @("Public", "Default", "Default User", "All Users")
    $DetectedProfiles = @(Get-MaceProfileDirs | Where-Object { $_.Name -notin $ExcludedProfiles })

    Write-Host "[>] Device  : " -NoNewline -ForegroundColor $LineCol
    Write-Host $env:COMPUTERNAME -ForegroundColor $MainCol

    # Scope is never assumed. It comes from $MaceScope ('AllUsers' or a profile folder name, set
    # before running the script) or from an explicit menu choice. Anything else stops the script
    # before it changes anything.
    $PresetScope = Get-Variable -Name MaceScope -ValueOnly -ErrorAction SilentlyContinue
    $Choice = $null      # 'All' | 'Profile' | 'Abort' | 'Invalid'
    $SelectedProfile = $null
    $menuRow = Get-MaceCursorTop
    $menuShown = $false

    if ($PresetScope) {
        if ([string]$PresetScope -eq 'AllUsers') {
            $Choice = 'All'
        } else {
            $SelectedProfile = $DetectedProfiles | Where-Object { $_.Name -eq [string]$PresetScope } | Select-Object -First 1
            if ($SelectedProfile) { $Choice = 'Profile' } else { $Choice = 'Invalid' }
        }
    } elseif ($IsInteractive) {
        $menuShown = $true
        Write-Host "[?] Select target user profile:" -ForegroundColor Yellow
        Write-Host ""
        for ($i = 0; $i -lt $DetectedProfiles.Count; $i++) {
            Write-Host "    " -NoNewline
            Write-Host "[$i]" -NoNewline -ForegroundColor $AccentCol
            Write-Host " $($DetectedProfiles[$i].Name)" -ForegroundColor White
        }
        Write-Host "    " -NoNewline
        Write-Host '[A]' -NoNewline -ForegroundColor $AccentCol
        Write-Host " All Users (system-wide, removes Office and Teams for every account)" -ForegroundColor White
        Write-Host "    " -NoNewline
        Write-Host '[ESC]' -NoNewline -ForegroundColor $DimCol
        Write-Host " Exit" -ForegroundColor $DimCol
        Write-Host ""
        Write-Host "Enter choice: " -NoNewline -ForegroundColor $AccentCol
        $key = Read-MaceKey
        if ($null -eq $key) {
            $Choice = 'Invalid'
        } elseif ($key.VirtualKeyCode -eq 27) {
            $Choice = 'Abort'
        } else {
            $ScopeChoice = $key.Character.ToString()
            if ($ScopeChoice -match '^\d+$' -and [int]$ScopeChoice -lt $DetectedProfiles.Count) {
                $SelectedProfile = $DetectedProfiles[[int]$ScopeChoice]
                $Choice = 'Profile'
            } elseif ($ScopeChoice -eq 'A' -or $ScopeChoice -eq 'a') {
                $Choice = 'All'
            } else {
                $Choice = 'Invalid'
            }
        }
    }

    if ($menuShown -and $menuRow -ge 0) {
        $endRow = Get-MaceCursorTop
        for ($r = $menuRow; $r -le $endRow; $r++) {
            Set-MaceCursor 0 $r
            try { [Console]::Write(" " * $script:Width) } catch { }
        }
        Set-MaceCursor 0 $menuRow
    }

    if ($Choice -eq 'Abort') {
        Write-Host ""
        Write-Host "[!] Aborted." -ForegroundColor $DimCol
        $script:MaceExit = 0
        $script:MaceNote = "aborted by user; nothing was changed"
        return
    }
    if ($null -eq $Choice -or $Choice -eq 'Invalid') {
        if ($PresetScope) {
            Write-Host "[!] `$MaceScope '$PresetScope' does not match AllUsers or a profile folder under C:\Users. Nothing was changed." -ForegroundColor Yellow
            $script:MaceNote = "scope '$PresetScope' not recognized; nothing was changed"
        } elseif ($IsInteractive) {
            Write-Host "[!] No valid scope was chosen. Nothing was changed." -ForegroundColor Yellow
            $script:MaceNote = "no valid scope chosen; nothing was changed"
        } else {
            Write-Host "[!] Non-interactive session and no scope set. Set `$MaceScope = 'AllUsers' (or a profile folder name) before running. Nothing was changed." -ForegroundColor Yellow
            $script:MaceNote = "no scope set in a non-interactive session; nothing was changed"
        }
        $script:MaceExit = 2
        return
    }

    $ProfileList = Get-MaceProfileList
    if ($Choice -eq 'Profile') {
        $TargetProfiles = @($SelectedProfile.FullName)
        $AllUsers       = $false
        $ScopeLabel     = "$($SelectedProfile.Name)  ($($SelectedProfile.FullName))"
    } else {
        $TargetProfiles = @($DetectedProfiles | Select-Object -ExpandProperty FullName)
        $AllUsers       = $true
        $ScopeLabel     = "All Users (System-Wide)"
    }

    # Identity of the single selected profile, so per-user work can be bound to it
    $TargetSid     = $null
    $TargetAccount = $null
    if (-not $AllUsers) {
        $entry = $ProfileList | Where-Object { $_.ProfileImagePath -eq $TargetProfiles[0] } | Select-Object -First 1
        if ($entry) {
            $TargetSid     = $entry.PSChildName
            $TargetAccount = Resolve-MaceAccount $TargetSid
        }
    }

    Write-Host "[>] Scope   : " -NoNewline -ForegroundColor Red
    Write-Host $ScopeLabel -ForegroundColor Yellow
    Write-HLine -Style dashed
    $Partial = New-Object System.Collections.ArrayList
    #endregion

    #region 3: PROCESS TERMINATION
    # ========================================================================
    Write-StepUpdate "[1/5] Terminating Active Processes..." -Pending
    $Processes = @(
        "OneDrive", "OneDriveSetup",
        "olk", "Outlook",
        "Teams", "ms-teams", "Update",
        "WINWORD", "EXCEL", "POWERPNT", "ONENOTE", "MSACCESS", "MSPUB", "WINPROJ",
        "OfficeClickToRun", "OfficeC2RClient",
        "msedgewebview2", "WebViewHost"
    )
    foreach ($proc in $Processes) {
        foreach ($p in (Get-MaceProcess -Name $proc)) {
            Stop-Process -Id $p.Id -Force -ErrorAction SilentlyContinue
        }
    }
    Start-Sleep -Seconds 3
    Write-StepUpdate -Success
    #endregion

    #region 4: APPLICATION UNINSTALL
    # ========================================================================
    Write-StepUpdate "[2/5] Uninstalling Applications..."
    $winget = Test-MaceWinget

    # Determine if the single target profile (if any) is the currently logged-in user
    $TargetIsSelf = (-not $AllUsers) -and ($TargetProfiles[0] -eq $env:USERPROFILE)

    # --- OneDrive ---
    Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " OneDrive..." -NoNewline -ForegroundColor Gray
    $ODSetups = @(
        "$env:SYSTEMROOT\System32\OneDriveSetup.exe",
        "$env:SYSTEMROOT\SysWOW64\OneDriveSetup.exe",
        "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe"
    )
    # System32/SysWOW64 OneDriveSetup.exe ships with Windows itself - not an install indicator
    $ODUserFound = $null -ne ($TargetProfiles | Where-Object { Test-Path "$_\AppData\Local\Microsoft\OneDrive" } | Select-Object -First 1)
    $ODFound = (Test-Path "$env:LOCALAPPDATA\Microsoft\OneDrive\OneDriveSetup.exe") -or
               (Test-Path "HKLM:\SOFTWARE\Microsoft\OneDrive") -or $ODUserFound
    if ($ODFound) {
        $ODArgs = if ($AllUsers) { "/uninstall /allusers" } else { "/uninstall" }
        foreach ($setup in $ODSetups) {
            if (Test-Path $setup) {
                Start-Process $setup -ArgumentList $ODArgs -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue
                break
            }
        }
        if ($winget) { Invoke-MaceWinget @("uninstall", "--id", "Microsoft.OneDrive", "--silent", "--accept-source-agreements") }
        Write-SubResult "[DONE]" $OkCol
    } else {
        Write-SubResult "[NOT-INSTALLED]" $InfoCol
    }

    # --- New Outlook ---
    Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " New Outlook (olk)..." -NoNewline -ForegroundColor Gray
    $OlkFound = (Get-MaceAppx -Name "Microsoft.OutlookForWindows").Count -gt 0
    if ($OlkFound) {
        if ($AllUsers) {
            Remove-MaceAppx -Name "Microsoft.OutlookForWindows"
            Remove-MaceProvisioned -Like "*OutlookForWindows*"
        } elseif ($TargetIsSelf) {
            Remove-MaceAppx -Name "Microsoft.OutlookForWindows" -Sid $TargetSid
        } else {
            Write-Host " (AppX skipped - target user not active session)" -NoNewline -ForegroundColor DarkGray
        }
        Write-SubResult "[DONE]" $OkCol
    } else {
        Write-SubResult "[NOT-INSTALLED]" $InfoCol
    }

    # --- Office / M365 ---
    Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " Office / M365 / Project..." -NoNewline -ForegroundColor Gray
    if ($true) {
        $C2RSetup = @(
            "C:\Program Files\Common Files\Microsoft Shared\ClickToRun\setup64.exe",
            "C:\Program Files (x86)\Common Files\Microsoft Shared\ClickToRun\setup64.exe"
        ) | Where-Object { Test-Path $_ } | Select-Object -First 1
        $OfficeMSIPkgs = @(Get-MaceUninstallEntries |
            Where-Object { $_.DisplayName -match "Microsoft (Office|365|Project)" -and $_.UninstallString -match "msiexec" })
        $OfficeFound = ($null -ne $C2RSetup) -or ($OfficeMSIPkgs.Count -gt 0)

        if ($OfficeFound) {
            if ($C2RSetup) {
                Write-Host " (ClickToRun found...)" -NoNewline -ForegroundColor DarkGray
                $XMLPath = "$env:TEMP\MACE_OfficeRemove.xml"
                $xml = @"
<Configuration>
  <Remove All="TRUE"/>
  <Display Level="None" AcceptEULA="TRUE"/>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
</Configuration>
"@
                Write-MaceFile -Path $XMLPath -Text $xml
                $proc = Start-Process $C2RSetup -ArgumentList "/configure `"$XMLPath`"" -PassThru -WindowStyle Hidden
                $proc | Wait-Process -Timeout 900 -ErrorAction SilentlyContinue
                if (-not $proc.HasExited) { $proc | Stop-Process -Force -ErrorAction SilentlyContinue }
                Remove-Item $XMLPath -Force -ErrorAction SilentlyContinue
            } else {
                Write-Host " (no local C2R, using MSI fallback)" -NoNewline -ForegroundColor DarkGray
            }
            # MSI fallback for any remaining installs
            $OfficePkgs = @(Get-MaceUninstallEntries |
                Where-Object { $_.DisplayName -match "Microsoft (Office|365|Project)" -and $_.UninstallString -match "msiexec" })
            foreach ($pkg in $OfficePkgs) {
                $guid = [regex]::Match($pkg.UninstallString, '\{[A-Z0-9-]+\}').Value
                if ($guid) { Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue }
            }
            Write-SubResult "[DONE]" $OkCol
        } else {
            Write-SubResult "[NOT-INSTALLED]" $InfoCol
        }
    }

    # --- Teams ---
    Write-Host "      " -NoNewline; Write-Host ">" -NoNewline -ForegroundColor $DimCol; Write-Host " Microsoft Teams..." -NoNewline -ForegroundColor Gray
    $TeamsPkgs = @(Get-MaceUninstallEntries |
        Where-Object { $_.DisplayName -like "*Teams Machine-Wide*" -or $_.DisplayName -eq "Microsoft Teams" })
    if ($true) {
        $TeamsAppx = @(Get-MaceAppx -Name "MSTeams") + @(Get-MaceAppx -Name "MicrosoftTeams")
        $TeamsFound = ($TeamsPkgs.Count -gt 0) -or ($TeamsAppx.Count -gt 0)
        if ($TeamsFound) {
            foreach ($pkg in $TeamsPkgs) {
                if ($pkg.UninstallString -match "msiexec") {
                    $guid = [regex]::Match($pkg.UninstallString, '\{[A-Z0-9-]+\}').Value
                    if ($guid) { Start-Process "msiexec.exe" -ArgumentList "/x $guid /qn /norestart" -Wait -WindowStyle Hidden -ErrorAction SilentlyContinue }
                }
            }
            if ($AllUsers) {
                Remove-MaceAppx -Name "MSTeams"
                Remove-MaceAppx -Name "MicrosoftTeams"
                Remove-MaceProvisioned -Like "*Teams*"
            } elseif ($TargetIsSelf) {
                Remove-MaceAppx -Name "MSTeams" -Sid $TargetSid
                Remove-MaceAppx -Name "MicrosoftTeams" -Sid $TargetSid
            } else {
                Write-Host " (AppX skipped - target user not active session)" -NoNewline -ForegroundColor DarkGray
            }
            if ($winget) {
                Invoke-MaceWinget @("uninstall", "--id", "Microsoft.Teams", "--silent", "--accept-source-agreements")
                Invoke-MaceWinget @("uninstall", "--name", "Microsoft Teams", "--silent", "--accept-source-agreements")
            }
            Write-SubResult "[DONE]" $OkCol
        } else {
            Write-SubResult "[NOT-INSTALLED]" $InfoCol
        }
    }
    Write-StepUpdate -Success
    #endregion

    #region 5: FILESYSTEM PURGE
    # ========================================================================
    Write-StepUpdate "[3/5] Purging Residual Files & Folders..." -Pending
    $SystemFolders = @(
        "$env:PROGRAMDATA\Microsoft OneDrive",
        "C:\Program Files\Microsoft OneDrive",
        "C:\Program Files\Microsoft Office",
        "C:\Program Files (x86)\Microsoft Office",
        "C:\Program Files\Common Files\Microsoft Shared\ClickToRun",
        "C:\ProgramData\Microsoft\ClickToRun",
        "C:\Program Files\Microsoft\Teams",
        "C:\Program Files (x86)\Microsoft\Teams",
        "C:\Program Files (x86)\Microsoft\TeamsMachineInstaller"
    )
    # Client and cache folders only. The OneDrive sync root (<profile>\OneDrive and any
    # "OneDrive - <org>" folder) holds user files that may exist nowhere else, so it is never removed.
    $RelUserPaths = @(
        "AppData\Local\Microsoft\OneDrive",
        "AppData\Local\Microsoft\Olk",
        "AppData\Local\Packages\Microsoft.OutlookForWindows_8wekyb3d8bbwe",
        "AppData\Local\Microsoft\Outlook\HubAppCache",
        "AppData\Local\Microsoft\Office",
        "AppData\Roaming\Microsoft\Office",
        "AppData\Roaming\Microsoft\Outlook",
        "AppData\Local\Microsoft\Teams",
        "AppData\Roaming\Microsoft\Teams",
        "AppData\Local\Microsoft\IdentityCache",
        "AppData\Local\Microsoft\OneAuth",
        "AppData\Local\Microsoft\TokenBroker\Cache",
        "AppData\Local\Microsoft\Edge\User Data\Default\EBWebView",
        "AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy"
    )
    foreach ($folder in $SystemFolders) {
        if (Test-Path $folder) { Remove-Item $folder -Recurse -Force -ErrorAction SilentlyContinue }
    }
    foreach ($profilePath in $TargetProfiles) {
        foreach ($rel in $RelUserPaths) {
            $full = Join-Path $profilePath $rel
            if (Test-Path $full) { Remove-Item -LiteralPath $full -Recurse -Force -ErrorAction SilentlyContinue }
        }
        foreach ($keep in @(Get-ChildItem -LiteralPath $profilePath -Directory -Filter "OneDrive*" -ErrorAction SilentlyContinue)) {
            Write-Host "      Left in place (may hold unsynced user files): $($keep.FullName)" -ForegroundColor $DimCol
        }
    }
    Write-StepUpdate -Success
    #endregion

    #region 6: REGISTRY, PATHS & SHELL FOLDER REPAIR
    # ========================================================================
    Write-StepUpdate "[4/5] Registry, Paths & Shell Folder Repair..." -Pending

    # Machine-level keys
    $MachineKeys = @(
        "HKLM:\SOFTWARE\Microsoft\OneDrive",
        "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun",
        "HKLM:\SOFTWARE\Microsoft\Office\16.0",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Office\16.0",
        "HKLM:\SOFTWARE\Microsoft\AppVISV",
        "HKLM:\SOFTWARE\Microsoft\Teams",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Teams",
        "HKLM:\SOFTWARE\WOW6432Node\Microsoft\MicrosoftTeams"
    )
    foreach ($key in $MachineKeys) {
        if (Test-Path $key) { Remove-Item $key -Recurse -Force -ErrorAction SilentlyContinue }
    }

    $UserRegKeys   = @(
        "Software\Microsoft\OneDrive",
        "Software\Microsoft\Office",
        "Software\Microsoft\Teams",
        "Software\Microsoft\Olk",
        "Software\Microsoft\Office\16.0\Outlook"
    )
    $UserRunValues = @("OneDrive", "Teams", "com.squirrel.Teams.Teams")
    $ShellFixes    = @(
        @{ Name = "Personal";    Default = "%USERPROFILE%\Documents"; Dir = "Documents" },
        @{ Name = "Desktop";     Default = "%USERPROFILE%\Desktop";   Dir = "Desktop"   },
        @{ Name = "My Pictures"; Default = "%USERPROFILE%\Pictures";  Dir = "Pictures"  },
        @{ Name = "My Music";    Default = "%USERPROFILE%\Music";     Dir = "Music"     },
        @{ Name = "My Video";    Default = "%USERPROFILE%\Videos";    Dir = "Videos"    }
    )

    # Match selected profiles to ProfileList entries for hive loading
    $HiveProfiles = $ProfileList | Where-Object {
        (Test-Path "$($_.ProfileImagePath)\NTUSER.DAT") -and
        ($TargetProfiles -contains $_.ProfileImagePath)
    }

    foreach ($prof in $HiveProfiles) {
        $HiveName   = "MACE_$($prof.PSChildName)"
        reg load "HKU\$HiveName" "$($prof.ProfileImagePath)\NTUSER.DAT" 2>&1 | Out-Null
        if ($LASTEXITCODE -eq 0) {
            # Hive loaded cleanly (user not logged in)
            Invoke-PerUserRegistry -HiveRoot "Registry::HKU\$HiveName" -ProfilePath $prof.ProfileImagePath
            [GC]::Collect()
            Start-Sleep -Milliseconds 500
            reg unload "HKU\$HiveName" 2>$null | Out-Null
        } elseif (Test-Path "Registry::HKU\$($prof.PSChildName)") {
            # User is currently logged in - their hive is already mounted under their SID
            Invoke-PerUserRegistry -HiveRoot "Registry::HKU\$($prof.PSChildName)" -ProfilePath $prof.ProfileImagePath
        } else {
            [void]$Partial.Add("registry for $($prof.ProfileImagePath) was locked and skipped")
        }
    }

    # Repair PSModulePath - strip dead/OneDrive-redirected entries
    $env:PSModulePath = ($env:PSModulePath -split ';' | Where-Object { $_ -and (Test-Path $_) }) -join ';'

    Write-StepUpdate -Success
    #endregion

    #region 7: CREDENTIAL FLUSH
    # ========================================================================
    Write-StepUpdate "[5/5] Flushing Cached Credentials..." -Pending
    # cmdkey only sees the vault of the account running this script. It is only valid for the
    # selected user when that user is the caller, so another profile's vault is left alone.
    cmdkey /list 2>$null | ForEach-Object {
        if ($_ -match "Target:\s*(.+)") {
            $target = $Matches[1].Trim()
            if ($target -match "Office|Outlook|Teams|OneDrive|MicrosoftOffice|MicrosoftTeams|msteams|lync") {
                cmdkey /delete:$target 2>$null | Out-Null
            }
        }
    }
    Write-StepUpdate -Success
    #endregion

    # Footer (reuses header art - art on right, fill on left)
    $_ffillW = $script:Width - $_artW - 2  # 1 leading space + 1 sfx char
    $_footer = "  MACE SEQUENCE COMPLETE"
    $_fpad   = " " * [Math]::Max(0, ($_ffillW - $_footer.Length - $_ver.Length))

    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art1" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    Write-Host "$_footer$_fpad$_ver" -ForegroundColor $MainCol -NoNewline; Write-Host " $_art2" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol
    Write-Host ("-" * $_ffillW) -ForegroundColor $LineCol -NoNewline; Write-Host " $_art3" -ForegroundColor $ArtCol -NoNewline; Write-Host "█" -ForegroundColor $LineCol

    if ($Partial.Count -gt 0) {
        $script:MaceExit = 3
        $script:MaceNote = ($Partial -join "; ")
    } else {
        $script:MaceExit = 0
        $script:MaceNote = "scope: $ScopeLabel"
    }
}

function Invoke-MaceMain {
    $script:MaceExit = 1
    $script:MaceNote = "MACE did not finish"
    $script:StepRow  = -1
    try {
        Invoke-MaceRun | Out-Null
    } catch {
        $script:MaceExit = 1
        $script:MaceNote = "failed: $($_.Exception.Message)"
    }
    $label = switch ($script:MaceExit) {
        0       { "SUCCESS" }
        2       { "NOT RUN" }
        3       { "PARTIAL" }
        default { "FAILED" }
    }
    # Goes to the output stream so an RMM console captures it
    Write-Output ("MACE RESULT: {0} (exit {1}) - {2}" -f $label, $script:MaceExit, $script:MaceNote)
    Exit-KillerScript $script:MaceExit
}

if (-not $global:KillerScriptsTestMode) { Invoke-MaceMain }
