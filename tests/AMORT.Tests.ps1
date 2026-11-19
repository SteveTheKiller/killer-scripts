#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# AMORT: entry point and exit codes, then findings 45 (owned processes only), 01 (restore point
# eligibility), 65 (repair skip) and 02 (rollback removal is opt-in).
# Every system boundary is a small Amort* wrapper or a mocked cmdlet. Nothing here touches the real machine.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath AMORT)
    $script:RealRemoveWindowsOldText = ${function:Remove-AmortWindowsOld}.ToString()

    # Runs Invoke-AmortMain and returns the exit code plus everything written to the output stream.
    function Invoke-AmortCapture {
        param([hashtable]$Arguments = @{})
        $out = & { try { Invoke-AmortMain @Arguments } catch { $_ } }
        $last = @($out)[-1]
        $code = $null
        if ($last -is [System.Management.Automation.ErrorRecord]) {
            if ($last.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw $last }
        }
        [pscustomobject]@{ Code = $code; Output = (@($out | Where-Object { $_ -is [string] }) -join "`n") }
    }

    function New-FakeProc {
        param([int]$ExitCode = 0, [int]$Id = 4000)
        $p = [pscustomobject]@{ Id = $Id; HasExited = $true; ExitCode = $ExitCode }
        $p | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { } -PassThru |
            Add-Member -MemberType ScriptMethod -Name Dispose -Value { } -PassThru |
            Add-Member -MemberType ScriptMethod -Name Kill -Value { } -PassThru
    }

    function New-FakeShadow {
        param([string]$Id, [int]$DaysOld, [string]$Volume = '\\?\Volume{c0ffee}\', [bool]$ClientAccessible = $true,
              [bool]$Persistent = $true, [string]$Provider = '{b5946137-7b9f-4925-af80-51abd60b20d5}', [string]$ExposedName = $null)
        [pscustomobject]@{
            ID = $Id; VolumeName = $Volume; InstallDate = (Get-Date).AddDays(-$DaysOld)
            ClientAccessible = $ClientAccessible; Persistent = $Persistent; ProviderID = $Provider; ExposedName = $ExposedName
        }
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'AMORT' {
    BeforeEach {
        $script:Elevated   = $true
        $script:DiskSize   = [int64](200GB)
        $script:Procs      = @()
        $script:Shadows    = @()
        $script:WinProcs   = @()
        $script:ExitCodes  = @{ RestoreHealth = 0; StartComponentCleanup = 0; sfc = 0 }
        $script:ToolCalls  = New-Object System.Collections.ArrayList
        $script:ToolExit   = @{}
        $script:ExistingPaths = @()
        $script:TiService  = [pscustomobject]@{ Name = 'TrustedInstaller'; Status = 'Stopped'; StartType = 'Manual' }
        $script:StartedProcs = New-Object System.Collections.ArrayList
        $script:StartedFiles = New-Object System.Collections.ArrayList

        Mock Test-AmortElevated { $script:Elevated }
        Mock Test-AmortConsole { $false }
        Mock Test-AmortRebootPending { $false }
        Mock Get-AmortCursorTop { -1 }
        Mock Get-AmortCursorLeft { 0 }
        Mock Set-AmortCursor { }
        Mock Write-HLine { }
        Mock Write-Host { }
        Mock Write-Progress { }
        Mock Clear-Host { }
        Mock Start-Sleep { }
        Mock Wait-Process { }
        Mock Optimize-Volume { }
        Mock New-AmortTempFile { Join-Path $TestDrive 'amort.tmp' }

        Mock Get-CimInstance { @() }
        Mock Get-CimInstance { [pscustomobject]@{ DeviceID = 'C:'; FreeSpace = [int64](100GB); Size = $script:DiskSize } } -ParameterFilter { $ClassName -eq 'Win32_LogicalDisk' }
        Mock Get-CimInstance { [pscustomobject]@{ Manufacturer = 'Contoso'; Model = 'Model 1' } } -ParameterFilter { $ClassName -eq 'Win32_ComputerSystem' }
        Mock Get-CimInstance { [pscustomobject]@{ Manufacturer = 'Contoso'; Product = 'Board' } } -ParameterFilter { $ClassName -eq 'Win32_BaseBoard' }
        Mock Get-CimInstance { [pscustomobject]@{ Caption = 'Windows 11 Pro' } } -ParameterFilter { $ClassName -eq 'Win32_OperatingSystem' }
        Mock Get-CimInstance { [pscustomobject]@{ DeviceID = '\\?\Volume{c0ffee}\' } } -ParameterFilter { $ClassName -eq 'Win32_Volume' }
        Mock Get-CimInstance { $script:Shadows } -ParameterFilter { $ClassName -eq 'Win32_ShadowCopy' }
        Mock Get-CimInstance { $script:WinProcs } -ParameterFilter { $ClassName -eq 'Win32_Process' }

        Mock Get-ItemProperty { $null }
        Mock Get-ItemProperty { [pscustomobject]@{ DisplayVersion = '25H2' } } -ParameterFilter { "$Path" -like '*Windows NT*' }
        Mock Get-ChildItem { @() }
        Mock Test-Path { $false }
        Mock Get-Process { @($script:Procs | Where-Object { $Name -contains $_.Name }) }
        Mock Get-Service { if ($Name -eq 'TrustedInstaller') { $script:TiService } else { [pscustomobject]@{ Name = $Name; Status = 'Running'; StartType = 'Automatic' } } }
        Mock Start-Service { }
        Mock Stop-Service { }
        Mock Set-Service { }
        Mock Stop-Process { }
        Mock Remove-Tree { }
        Mock Remove-AmortWindowsOld { }
        Mock Invoke-AmortTool {
            [void]$script:ToolCalls.Add("$FilePath $($ArgumentList -join ' ')")
            $key = (Split-Path $FilePath -Leaf)
            $code = if ($script:ToolExit.ContainsKey($key)) { $script:ToolExit[$key] } else { 0 }
            [pscustomobject]@{ ExitCode = $code; Output = '' }
        }
        Mock Start-AmortProcess {
            [void]$script:StartedProcs.Add($StartInfo.Arguments)
            [void]$script:StartedFiles.Add($StartInfo.FileName)
            $code = 0
            foreach ($k in $script:ExitCodes.Keys) { if ($StartInfo.Arguments -like "*$k*") { $code = $script:ExitCodes[$k] } }
            New-FakeProc -ExitCode $code
        }
    }

    Describe 'Entry-point convention' {
        It 'defines Invoke-AmortMain and Exit-KillerScript without running anything on dot-source' {
            Get-Command Invoke-AmortMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
            Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        }

        It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
            { Exit-KillerScript 3010 } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
        }

        It 'console helpers never throw without a console' {
            { Test-AmortConsole; Get-AmortCursorTop; Get-AmortCursorLeft; Set-AmortCursor 0 5 } | Should -Not -Throw
        }
    }

    Describe 'Exit codes and final status line' {
        It 'exits 2 and changes nothing when not elevated' {
            $script:Elevated = $false
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 2
            $r.Output | Should -Match '\[FAIL\] AMORT not elevated.*\(exit 2\)'
            Should -Invoke Invoke-AmortTool -Times 0 -Exactly
            Should -Invoke Start-AmortProcess -Times 0 -Exactly
        }

        It 'exits 0 with a PASS line on a healthy run' {
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            $r.Output | Should -Match '\[PASS\] AMORT complete \(exit 0\)'
            $script:StartedProcs.Count | Should -Be 3
        }

        It 'a dry run exits 0, says so, and runs no tool or repair' {
            $r = Invoke-AmortCapture @{ DryRun = $true }
            $r.Code | Should -Be 0
            $r.Output | Should -Match 'dry run complete.*\(exit 0\)'
            Should -Invoke Invoke-AmortTool -Times 0 -Exactly
            Should -Invoke Start-AmortProcess -Times 0 -Exactly
        }

        It 'exits 3 and names the step when DISM fails' {
            $script:ExitCodes['RestoreHealth'] = 87
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 3
            $r.Output | Should -Match '\[WARN\] AMORT partial:.*DISM RestoreHealth failed.*\(exit 3\)'
        }

        It 'exits 3010 when DISM asks for a reboot and nothing failed' {
            $script:ExitCodes['StartComponentCleanup'] = 3010
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 3010
            $r.Output | Should -Match 'reboot required \(exit 3010\)'
        }

        It 'exits 3 and says repair was skipped when a reboot is pending' {
            Mock Test-AmortRebootPending { $true }
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'reboot pending'
            $script:StartedProcs.Count | Should -Be 0
        }

        It 'exits 2 when the C: volume size cannot be read' {
            $script:DiskSize = 0
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 2
            $r.Output | Should -Match '\[FAIL\].*precondition.*\(exit 2\)'
        }

        It 'exits 1 with a FAIL line on an unexpected error' {
            Mock Remove-OldFiles { throw 'boom' }
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 1
            $r.Output | Should -Match '\[FAIL\] AMORT stopped on an error: boom \(exit 1\)'
        }
    }

    Describe 'Finding 45: only work this run started is ever stopped' {
        It 'leaves independent TiWorker and DISM processes alone after the owned repair exits' {
            # Servicing that starts after the preflight check belongs to Windows, not to this run.
            $script:GetProcCalls = 0
            Mock Get-Process {
                $script:GetProcCalls++
                if ($script:GetProcCalls -le 1) { return @() }
                @(
                    [pscustomobject]@{ Name = 'TiWorker'; Id = 7001 },
                    [pscustomobject]@{ Name = 'DISM'; Id = 7002 }
                ) | Where-Object { $Name -contains $_.Name }
            }
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            Should -Invoke Stop-Process -Times 0 -Exactly
        }

        It 'Stop-AmortOwnedTree stops the owned process and its descendants, and nothing else' {
            $script:WinProcs = @(
                [pscustomobject]@{ ProcessId = 100; ParentProcessId = 1;   Name = 'cmd.exe' },
                [pscustomobject]@{ ProcessId = 101; ParentProcessId = 100; Name = 'dism.exe' },
                [pscustomobject]@{ ProcessId = 102; ParentProcessId = 101; Name = 'DismHost.exe' },
                [pscustomobject]@{ ProcessId = 200; ParentProcessId = 1;   Name = 'dism.exe' },
                [pscustomobject]@{ ProcessId = 300; ParentProcessId = 1;   Name = 'TiWorker.exe' }
            )
            Stop-AmortOwnedTree -ProcessId 100
            Should -Invoke Stop-Process -Times 1 -Exactly -ParameterFilter { $Id -eq 100 }
            Should -Invoke Stop-Process -Times 1 -Exactly -ParameterFilter { $Id -eq 101 }
            Should -Invoke Stop-Process -Times 1 -Exactly -ParameterFilter { $Id -eq 102 }
            Should -Invoke Stop-Process -Times 0 -Exactly -ParameterFilter { $Id -in 200, 300 }
        }
    }

    Describe 'Finding 01: only explicitly eligible restore points are trimmed' {
        BeforeEach {
            $script:Deleted = New-Object System.Collections.ArrayList
            $script:Sticky = @()
            # The CIM view of shadow copies reflects what vssadmin has deleted so far.
            Mock Get-CimInstance { @($script:Shadows | Where-Object { $script:Deleted -notcontains $_.ID }) } -ParameterFilter { $ClassName -eq 'Win32_ShadowCopy' }
            Mock Invoke-AmortTool {
                [void]$script:ToolCalls.Add("$FilePath $($ArgumentList -join ' ')")
                $sh = $ArgumentList | Where-Object { $_ -like '/shadow=*' }
                if ($sh) {
                    $id = $sh.Substring(8)
                    if ($script:ToolExit.ContainsKey('vssadmin.exe')) { return [pscustomobject]@{ ExitCode = $script:ToolExit['vssadmin.exe']; Output = '' } }
                    if ($script:Sticky -notcontains $id) { [void]$script:Deleted.Add($id) }
                }
                [pscustomobject]@{ ExitCode = 0; Output = '' }
            }
        }

        It 'never deletes copies that are not System Restore points, however old' {
            $script:Shadows = @(
                (New-FakeShadow -Id '{r1}' -DaysOld 1),
                (New-FakeShadow -Id '{r2}' -DaysOld 5),
                (New-FakeShadow -Id '{r3}' -DaysOld 9),
                (New-FakeShadow -Id '{backup}' -DaysOld 100 -ClientAccessible $false),
                (New-FakeShadow -Id '{hw}' -DaysOld 101 -Provider '{11111111-2222-3333-4444-555555555555}'),
                (New-FakeShadow -Id '{mounted}' -DaysOld 102 -ExposedName 'Z:\'),
                (New-FakeShadow -Id '{transient}' -DaysOld 103 -Persistent $false),
                (New-FakeShadow -Id '{otherdisk}' -DaysOld 104 -Volume '\\?\Volume{d00d}\')
            )
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            @($script:Deleted) | Should -Be @('{r3}')
            ($script:ToolCalls -join "`n") | Should -Not -Match 'backup|hw\}|mounted|transient|otherdisk'
            $r.Output | Should -Match 'PASS'
        }

        It 'passes the documented volume argument and keeps the newest eligible copies' {
            $script:Shadows = @(
                (New-FakeShadow -Id '{r1}' -DaysOld 1),
                (New-FakeShadow -Id '{r2}' -DaysOld 5),
                (New-FakeShadow -Id '{r3}' -DaysOld 9),
                (New-FakeShadow -Id '{r4}' -DaysOld 12)
            )
            [void](Invoke-AmortCapture)
            @($script:Deleted | Sort-Object) | Should -Be @('{r3}', '{r4}')
            $line = @($script:ToolCalls | Where-Object { $_ -like '*vssadmin*{r3}*' })[0]
            $line | Should -Match '/for=C:'
            $line | Should -Match 'delete shadows'
        }

        It 'reports a partial result when vssadmin fails and when a copy survives its delete' {
            $script:Shadows = @((New-FakeShadow -Id '{r1}' -DaysOld 1), (New-FakeShadow -Id '{r2}' -DaysOld 5), (New-FakeShadow -Id '{r3}' -DaysOld 9))
            $script:ToolExit['vssadmin.exe'] = 2
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'restore point \{r3\} could not be deleted'

            $script:ToolExit.Remove('vssadmin.exe')
            $script:Sticky = @('{r3}')
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'restore point \{r3\} could not be deleted'
        }

        It 'deletes nothing when no copy is eligible' {
            $script:Shadows = @((New-FakeShadow -Id '{a}' -DaysOld 50 -ClientAccessible $false), (New-FakeShadow -Id '{b}' -DaysOld 60 -ClientAccessible $false), (New-FakeShadow -Id '{c}' -DaysOld 70 -ClientAccessible $false))
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            $script:Deleted.Count | Should -Be 0
            ($script:ToolCalls -join "`n") | Should -Not -Match 'vssadmin'
        }

        It 'a dry run estimates only the eligible copies and deletes nothing' {
            Mock Get-CimInstance { [pscustomobject]@{ UsedSpace = [int64](10GB); Volume = [pscustomobject]@{ DeviceID = '\\?\Volume{c0ffee}\' } } } -ParameterFilter { $ClassName -eq 'Win32_ShadowStorage' }
            $script:Shadows = @(
                (New-FakeShadow -Id '{r1}' -DaysOld 1), (New-FakeShadow -Id '{r2}' -DaysOld 5),
                (New-FakeShadow -Id '{r3}' -DaysOld 9), (New-FakeShadow -Id '{r4}' -DaysOld 12),
                (New-FakeShadow -Id '{b1}' -DaysOld 100 -ClientAccessible $false), (New-FakeShadow -Id '{b2}' -DaysOld 101 -ClientAccessible $false),
                (New-FakeShadow -Id '{b3}' -DaysOld 102 -ClientAccessible $false), (New-FakeShadow -Id '{b4}' -DaysOld 103 -ClientAccessible $false)
            )
            $r = Invoke-AmortCapture @{ DryRun = $true }
            $r.Code | Should -Be 0
            $script:Deleted.Count | Should -Be 0
            # 4 eligible, keep 2, so half of the shared storage. A count over all 8 would say 2.5 GB.
            $script:EstYieldBytes | Should -Be ([int64](5GB))
        }
    }

    Describe 'Finding 65: TrustedInstaller is left alone whenever repair is skipped' {
        BeforeEach {
            $script:TiService = [pscustomobject]@{ Name = 'TrustedInstaller'; Status = 'Stopped'; StartType = 'Disabled' }
        }

        It 'does not enable or start TrustedInstaller when a reboot is pending' {
            Mock Test-AmortRebootPending { $true }
            [void](Invoke-AmortCapture)
            Should -Invoke Set-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' }
            Should -Invoke Start-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' }
        }

        It 'does not touch TrustedInstaller when Windows servicing is already active' {
            Mock Get-Process { @([pscustomobject]@{ Name = 'TiWorker'; Id = 7001 }) } -ParameterFilter { $Name -contains 'TiWorker' }
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 3
            Should -Invoke Set-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' }
            Should -Invoke Start-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' }
            $script:StartedProcs.Count | Should -Be 0
        }

        It 'does not touch TrustedInstaller when a pending file rename skips repair' {
            Mock Get-ItemProperty { [pscustomobject]@{ PendingFileRenameOperations = @('a', 'b') } } -ParameterFilter { $Name -eq 'PendingFileRenameOperations' }
            [void](Invoke-AmortCapture)
            Should -Invoke Set-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' }
            Should -Invoke Start-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' }
        }

        It 'still enables and starts TrustedInstaller when repair will run' {
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            Should -Invoke Set-Service -Times 1 -Exactly -ParameterFilter { $Name -eq 'TrustedInstaller' -and $StartupType -eq 'Manual' }
            Should -Invoke Start-Service -Times 1 -ParameterFilter { $Name -eq 'TrustedInstaller' }
            $script:StartedProcs.Count | Should -Be 3
        }
    }

    Describe 'Finding 02: removing the rollback window is opt-in and refused while Windows is busy' {
        BeforeEach {
            Mock Test-Path { $true } -ParameterFilter { "$Path" -eq 'C:\Windows.old' }
            Mock Get-PathSize { [int64]0 }
            Mock Get-PathSize { [int64](30GB) } -ParameterFilter { $Path -eq 'C:\Windows.old' }
            Remove-Variable -Name RemoveRollback -Scope Global -ErrorAction SilentlyContinue
        }
        AfterEach { Remove-Variable -Name RemoveRollback -Scope Global -ErrorAction SilentlyContinue }

        It 'keeps Windows.old and the rollback data by default, and tells the technician how to remove it' {
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            Should -Invoke Remove-AmortWindowsOld -Times 0 -Exactly
            ($script:ToolCalls -join "`n") | Should -Not -Match 'Remove-OSUninstall'
            ($script:Notes -join "`n") | Should -Match 'Windows.old was kept.*-RemoveRollback'
        }

        It 'a dry run without the switch does not count Windows.old as reclaimable' {
            [void](Invoke-AmortCapture @{ DryRun = $true })
            $script:EstYieldBytes | Should -Be 0
            Should -Invoke Remove-AmortWindowsOld -Times 0 -Exactly
        }

        It 'a dry run with the switch estimates Windows.old and still removes nothing' {
            [void](Invoke-AmortCapture @{ DryRun = $true; RemoveRollback = $true })
            $script:EstYieldBytes | Should -Be ([int64](30GB))
            Should -Invoke Remove-AmortWindowsOld -Times 0 -Exactly
        }

        It 'removes Windows.old only when the switch is given and Windows is idle' {
            $r = Invoke-AmortCapture @{ RemoveRollback = $true }
            $r.Code | Should -Be 0
            Should -Invoke Remove-AmortWindowsOld -Times 1 -Exactly
        }

        It 'treats a global RemoveRollback of true as the switch, for consoles that cannot pass arguments' {
            $global:RemoveRollback = $true
            Test-AmortRemoveRollbackRequested -Switch $false | Should -BeTrue
            Remove-Variable -Name RemoveRollback -Scope Global
            Test-AmortRemoveRollbackRequested -Switch $false | Should -BeFalse
            Test-AmortRemoveRollbackRequested -Switch $true | Should -BeTrue
        }

        It 'refuses while Windows servicing is active' {
            Mock Get-Process { @([pscustomobject]@{ Name = 'TiWorker'; Id = 7001 }) } -ParameterFilter { $Name -contains 'TiWorker' }
            $r = Invoke-AmortCapture @{ RemoveRollback = $true }
            Should -Invoke Remove-AmortWindowsOld -Times 0 -Exactly
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'Windows.old was not removed: Windows servicing is active'
        }

        It 'refuses while a reboot is pending' {
            Mock Test-AmortRebootPending { $true }
            $r = Invoke-AmortCapture @{ RemoveRollback = $true }
            Should -Invoke Remove-AmortWindowsOld -Times 0 -Exactly
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'Windows.old was not removed: a reboot is pending'
        }

        It 'refuses while a Windows upgrade is running' {
            Mock Get-Process { @([pscustomobject]@{ Name = 'SetupHost'; Id = 7003 }) } -ParameterFilter { $Name -contains 'SetupHost' }
            $r = Invoke-AmortCapture @{ RemoveRollback = $true }
            Should -Invoke Remove-AmortWindowsOld -Times 0 -Exactly
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'Windows.old was not removed: a Windows upgrade is running'
        }

        It 'strips the uninstall protection only inside the explicit removal, and checks the result' {
            # The real function, captured before the per-test mock replaced it, with its own boundaries mocked.
            Mock Start-Process { $null }
            Mock Test-Path { $false } -ParameterFilter { "$Path" -like 'C:*' }
            & ([scriptblock]::Create($script:RealRemoveWindowsOldText))
            Should -Invoke Invoke-AmortTool -Times 1 -Exactly -ParameterFilter { $FilePath -eq 'DISM.exe' -and ($ArgumentList -contains '/Remove-OSUninstall') }
        }
    }

    Describe 'Native tools on a 32-bit host' {
        It 'Get-AmortNativePath returns the plain name on a 64-bit host' {
            Mock Test-AmortWow64 { $false }
            Get-AmortNativePath 'vssadmin.exe' | Should -Be 'vssadmin.exe'
        }

        It 'Get-AmortNativePath goes through Sysnative on a 32-bit host so System32 is not redirected' {
            Mock Test-AmortWow64 { $true }
            Get-AmortNativePath 'vssadmin.exe' | Should -Be (Join-Path $env:windir 'Sysnative\vssadmin.exe')
        }

        It 'launches repair through the native cmd.exe on a 32-bit host' {
            Mock Test-AmortWow64 { $true }
            $r = Invoke-AmortCapture
            $r.Code | Should -Be 0
            $script:StartedFiles.Count | Should -Be 3
            @($script:StartedFiles | Where-Object { $_ -ne (Join-Path $env:windir 'Sysnative\cmd.exe') }).Count | Should -Be 0
        }
    }

    Describe 'Finding 02 follow-up: rollback staging data and the real pending-reboot state' {
        BeforeEach {
            Mock Test-Path { $true } -ParameterFilter { "$Path" -eq 'C:\Windows.old' -or "$LiteralPath" -in 'C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS', 'C:\$GetCurrent' }
            Mock Get-PathSize { [int64]0 }
        }

        It 'keeps $WINDOWS.~BT and $WINDOWS.~WS while Windows.old exists and rollback removal was not requested' {
            [void](Invoke-AmortCapture)
            Should -Invoke Remove-Tree -Times 0 -Exactly -ParameterFilter { $Path -in 'C:\$WINDOWS.~BT', 'C:\$WINDOWS.~WS' }
            Should -Invoke Remove-Tree -Times 1 -Exactly -ParameterFilter { $Path -eq 'C:\$GetCurrent' }
        }

        It 'removes them when rollback removal is requested and allowed' {
            [void](Invoke-AmortCapture @{ RemoveRollback = $true })
            Should -Invoke Remove-Tree -Times 1 -Exactly -ParameterFilter { $Path -eq 'C:\$WINDOWS.~BT' }
            Should -Invoke Remove-Tree -Times 1 -Exactly -ParameterFilter { $Path -eq 'C:\$WINDOWS.~WS' }
        }

        It 'Test-AmortRebootPending reads the 64-bit registry view and sees both pending-reboot keys' {
            $src = ${function:Test-AmortRebootPending}.ToString()
            $src | Should -Match 'Registry64'
            $src | Should -Match 'Component Based Servicing.RebootPending'
            $src | Should -Match 'WindowsUpdate.Auto Update.RebootRequired'
        }
    }
}
