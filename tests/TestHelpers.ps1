# Shared test helpers for killer-scripts. Test-only: shipped scripts never load this file.
# Dot-source inside BeforeAll:  . (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$script:KillerRepoRoot = Split-Path -Parent $PSScriptRoot

function Get-KillerScriptPath {
    # Full path of a shipped script by name, for example Get-KillerScriptPath WURSA.
    param([Parameter(Mandatory)][string]$Name)
    $p = Join-Path $script:KillerRepoRoot ("{0}.ps1" -f $Name)
    if (-not (Test-Path -LiteralPath $p)) { throw "No such script: $p" }
    $p
}

function Enable-KillerTestMode {
    # Set before dot-sourcing a script so its Main does not run and Exit-KillerScript throws.
    $global:KillerScriptsTestMode = $true
}

function Disable-KillerTestMode {
    Remove-Variable -Name KillerScriptsTestMode -Scope Global -ErrorAction SilentlyContinue
}

function Test-KillerSafePath {
    # True when every path argument points inside Pester's TestDrive or TestRegistry.
    # Pester 5 filters expose parameters as variables, not $PSBoundParameters, so callers pass them in.
    param([object[]]$Paths)
    $paths = @($Paths | Where-Object { $null -ne $_ -and [string]$_ -ne '' })
    if ($paths.Count -eq 0) { return $false }
    $drive = $null
    try { $drive = (Get-PSDrive -Name TestDrive -ErrorAction Stop).Root } catch { }
    foreach ($p in $paths) {
        $s = [string]$p
        if ($s -like 'TestDrive:*' -or $s -like 'TestRegistry:*') { continue }
        if ($drive -and $s.StartsWith($drive, [StringComparison]::OrdinalIgnoreCase)) { continue }
        return $false
    }
    $true
}

function Initialize-KillerSafetyNet {
    # Every destructive command throws unless the test mocks it. File and registry cmdlets are
    # allowed only inside TestDrive or TestRegistry. Call from BeforeAll, before any script code runs.
    $block = { throw "UNMOCKED DESTRUCTIVE CALL: $($MyInvocation.MyCommand.Name) was not mocked by this test." }

    $pathScoped = @(
        'Remove-Item', 'Set-Content', 'Add-Content', 'Clear-Content', 'Out-File', 'New-Item',
        'Copy-Item', 'Move-Item', 'Rename-Item', 'Set-Item', 'Set-ItemProperty', 'New-ItemProperty',
        'Remove-ItemProperty', 'Rename-ItemProperty', 'Expand-Archive', 'Set-Acl'
    )
    foreach ($c in $pathScoped) {
        if (Get-Command $c -ErrorAction SilentlyContinue) {
            Mock $c $block -ParameterFilter {
                -not (Test-KillerSafePath -Paths (@($Path) + @($LiteralPath) + @($FilePath) + @($Destination) + @($DestinationPath)))
            }
        }
    }

    $always = @(
        # processes, services, power
        'Start-Process', 'Stop-Process', 'Stop-Service', 'Start-Service', 'Restart-Service', 'Set-Service',
        'Restart-Computer', 'Stop-Computer', 'Rename-Computer', 'Add-Computer', 'Remove-Computer',
        # accounts and groups
        'New-LocalUser', 'Remove-LocalUser', 'Disable-LocalUser', 'Enable-LocalUser', 'Set-LocalUser',
        'Rename-LocalUser', 'Add-LocalGroupMember', 'Remove-LocalGroupMember',
        # scheduled tasks
        'Register-ScheduledTask', 'Unregister-ScheduledTask', 'Set-ScheduledTask', 'Start-ScheduledTask',
        'Stop-ScheduledTask', 'Disable-ScheduledTask', 'Enable-ScheduledTask',
        # BitLocker
        'Enable-BitLocker', 'Disable-BitLocker', 'Add-BitLockerKeyProtector', 'Remove-BitLockerKeyProtector',
        'Backup-BitLockerKeyProtector', 'BackupToAAD-BitLockerKeyProtector', 'Enable-BitLockerAutoUnlock',
        'Disable-BitLockerAutoUnlock', 'Suspend-BitLocker', 'Resume-BitLocker',
        # packages and apps
        'Remove-AppxPackage', 'Remove-AppxProvisionedPackage', 'Add-AppxPackage', 'Add-AppxProvisionedPackage',
        'Uninstall-Package', 'Install-Package', 'Install-Module',
        # network, firewall, Defender, printers, time
        'Set-NetConnectionProfile', 'Set-NetFirewallProfile', 'New-NetFirewallRule', 'Set-MpPreference',
        'Add-MpPreference', 'Remove-MpPreference', 'Update-MpSignature', 'Start-MpScan',
        'Add-Printer', 'Add-PrinterPort', 'Add-PrinterDriver', 'Remove-Printer', 'Remove-PrinterPort',
        'Set-DnsClientServerAddress',
        # HTTP
        'Invoke-WebRequest', 'Invoke-RestMethod', 'Start-BitsTransfer',
        # data
        'Clear-RecycleBin', 'Checkpoint-Computer', 'Enable-ComputerRestore', 'Disable-ComputerRestore'
    )
    foreach ($c in $always) {
        if (Get-Command $c -ErrorAction SilentlyContinue) { Mock $c $block }
    }

    $native = @(
        'manage-bde', 'schtasks', 'reg', 'cmdkey', 'w32tm', 'vssadmin', 'sc.exe', 'net', 'net1', 'netsh',
        'bcdedit', 'dism', 'sfc', 'msiexec', 'winget', 'choco', 'wusa', 'cleanmgr', 'UsoClient', 'MpCmdRun',
        'powercfg', 'icacls', 'takeown', 'cipher', 'wevtutil', 'shutdown', 'taskkill', 'wmic', 'robocopy',
        'gpupdate', 'dsregcmd', 'pnputil', 'rundll32', 'regsvr32', 'cmd', 'powershell', 'pwsh'
    )
    foreach ($c in $native) {
        if (Get-Command $c -CommandType Application -ErrorAction SilentlyContinue) { Mock $c $block }
    }
}
