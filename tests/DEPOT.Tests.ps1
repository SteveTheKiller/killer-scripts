#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# DEPOT tests. Every system call, download and prompt is mocked. Nothing here touches the real
# machine or the network, and DEPOT.ps1 is never started in a child process.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath DEPOT)
}

AfterAll {
    Disable-KillerTestMode
}

Describe 'Entry-point convention' {
    It 'defines Invoke-DepotMain and Exit-KillerScript without running anything on dot-source' {
        Get-Command Invoke-DepotMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
        { Exit-KillerScript 3010 } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
    }
}

Describe 'Exit codes at the entry point' {
    BeforeEach {
        Mock Write-DepotRunLog { }
        Mock Start-Sleep { }
        Mock Read-Host { '' }
        Mock Clear-Host { }
        Mock Get-Process { }
        Mock Start-Process { }
        Mock Restart-Computer { }
        $script:DepotScriptPath = $null
    }

    It 'exits 2 when not elevated and there is no script file to relaunch (pasted or -EncodedCommand)' {
        Mock Test-DepotElevated { $false }
        { Invoke-DepotMain } | Should -Throw 'KILLERSCRIPT_EXIT:2'
        Should -Invoke Start-Process -Times 0 -Exactly
    }

    It 'exits 2 when the elevated relaunch fails' {
        Mock Test-DepotElevated { $false }
        Mock Start-Process { throw 'The operation was canceled by the user' }
        $script:DepotScriptPath = 'C:\Scripts\DEPOT.ps1'
        { Invoke-DepotMain } | Should -Throw 'KILLERSCRIPT_EXIT:2'
    }

    It 'exits 3010 when the run ends with a reboot required' {
        Mock Test-DepotElevated { $true }
        foreach ($f in 'Invoke-DepotPreflight', 'Invoke-DepotInteractiveSetup', 'Show-DepotBanner', 'Invoke-DepotWindowsUpdate',
                       'Install-DepotApps', 'Invoke-DepotCleanup', 'Invoke-DepotPrivacy') {
            Mock $f { }
        }
        Mock Invoke-DepotWindowsUpdate { $script:RebootRequired = $true }
        { Invoke-DepotMain } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
    }
}

Describe 'Finding 12: admin purge needs a verified survivor' {
    BeforeEach {
        $sid = { param($s) [Security.Principal.SecurityIdentifier]"S-1-5-21-111-222-333-$s" }
        # The built-in Administrator has been renamed to Owner, so only the SID identifies it.
        $script:Users = @(
            [pscustomobject]@{ Name = 'Owner';   SID = (& $sid 500);  Enabled = $true },
            [pscustomobject]@{ Name = 'oldtech'; SID = (& $sid 1001); Enabled = $true }
        )
        $script:Members = @(
            [pscustomobject]@{ Name = 'PC\Owner';   SID = (& $sid 500);  ObjectClass = 'User' },
            [pscustomobject]@{ Name = 'PC\oldtech'; SID = (& $sid 1001); ObjectClass = 'User' },
            [pscustomobject]@{ Name = 'CONTOSO\dadmin'; SID = [Security.Principal.SecurityIdentifier]'S-1-5-21-999-888-777-1105'; ObjectClass = 'User' }
        )
        $script:Choices = New-Object System.Collections.Queue
        $script:NewSid = (& $sid 1002)
        Reset-DepotState
        Mock Clear-Host { }
        Mock Read-DepotChoice { $script:Choices.Dequeue() }
        Mock Read-Host {
            if ($AsSecureString) { $s = New-Object System.Security.SecureString; 'pw'.ToCharArray() | ForEach-Object { $s.AppendChar($_) }; $s } else { 'techadmin' }
        }
        Mock Get-LocalUser {
            if ($SID) { $script:Users | Where-Object { $_.SID.Value -eq $SID.Value } }
            elseif ($Name) { $script:Users | Where-Object { $_.Name -eq $Name } }
            else { $script:Users }
        }
        Mock Get-LocalGroupMember { $script:Members }
        Mock New-LocalUser {
            $u = [pscustomobject]@{ Name = $Name; SID = $script:NewSid; Enabled = $true }
            $script:Users += $u
            $u
        }
        Mock Add-LocalGroupMember {
            $u = $script:Users | Where-Object { $_.Name -eq $Member -or $_.SID.Value -eq "$Member" } | Select-Object -First 1
            $script:Members += [pscustomobject]@{ Name = "PC\$($u.Name)"; SID = $u.SID; ObjectClass = 'User' }
        }
        Mock Disable-LocalUser { $u = $script:Users | Where-Object { ($SID -and $_.SID.Value -eq $SID.Value) -or ($Name -and $_.Name -eq $Name) }; if ($u) { $u.Enabled = $false } }
        Mock Remove-LocalUser {
            $script:Users = @($script:Users | Where-Object { -not ($SID -and $_.SID.Value -eq $SID.Value) -and -not ($Name -and $_.Name -eq $Name) })
            $script:Members = @($script:Members | Where-Object { $_.SID.Value -ne $SID.Value })
        }
        Mock Rename-Computer { }
    }

    It 'refuses to disable the built-in Administrator or delete anyone when the operator declines to create an admin' {
        'N', 'N', 'Y' | ForEach-Object { $script:Choices.Enqueue($_) }
        Invoke-DepotInteractiveSetup 6>$null 3>$null
        Should -Invoke Disable-LocalUser -Times 0 -Exactly
        Should -Invoke Remove-LocalUser -Times 0 -Exactly
        @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -BeGreaterThan 0
    }

    It 'refuses to disable or delete anyone when creating the new admin fails' {
        Mock New-LocalUser { throw 'The password does not meet the password policy requirements.' }
        'N', 'Y', 'Y' | ForEach-Object { $script:Choices.Enqueue($_) }
        Invoke-DepotInteractiveSetup 6>$null 3>$null
        Should -Invoke Disable-LocalUser -Times 0 -Exactly
        Should -Invoke Remove-LocalUser -Times 0 -Exactly
    }

    It 'refuses the purge when the new account exists but could not be added to Administrators' {
        Mock Add-LocalGroupMember { throw 'Access is denied.' }
        'N', 'Y', 'Y' | ForEach-Object { $script:Choices.Enqueue($_) }
        Invoke-DepotInteractiveSetup 6>$null 3>$null
        Should -Invoke Disable-LocalUser -Times 0 -Exactly
        Should -Invoke Remove-LocalUser -Times 0 -Exactly
    }

    It 'disables the renamed built-in Administrator by SID and deletes other local admins by SID once the new admin is verified' {
        'N', 'Y', 'Y' | ForEach-Object { $script:Choices.Enqueue($_) }
        Invoke-DepotInteractiveSetup 6>$null 3>$null
        Should -Invoke Disable-LocalUser -Times 1 -Exactly -ParameterFilter { $SID.Value -eq 'S-1-5-21-111-222-333-500' }
        Should -Invoke Remove-LocalUser -Times 1 -Exactly -ParameterFilter { $SID.Value -eq 'S-1-5-21-111-222-333-1001' }
        # not the survivor, and not the domain account that Remove-LocalUser cannot delete
        Should -Invoke Remove-LocalUser -Times 0 -Exactly -ParameterFilter { $SID.Value -eq 'S-1-5-21-111-222-333-1002' }
        Should -Invoke Remove-LocalUser -Times 0 -Exactly -ParameterFilter { $SID.Value -eq 'S-1-5-21-999-888-777-1105' }
        @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 0
    }

    It 'reports a failed deletion instead of printing Done' {
        Mock Remove-LocalUser { throw 'The user could not be removed.' }
        'N', 'Y', 'Y' | ForEach-Object { $script:Choices.Enqueue($_) }
        $out = Invoke-DepotInteractiveSetup 6>&1 3>$null | Out-String
        $out | Should -Not -Match 'Deleting account: PC\\oldtech\.\.\.\s*Done'
        (@($script:DepotResults | Where-Object { $_.Status -eq 'Failed' -and $_.Step -match 'oldtech' })).Count | Should -Be 1
    }

    It 'fails closed when the Administrators group cannot be read' {
        Mock Get-LocalGroupMember { throw 'One or more members could not be resolved.' }
        'N', 'Y', 'Y' | ForEach-Object { $script:Choices.Enqueue($_) }
        Invoke-DepotInteractiveSetup 6>$null 3>$null
        Should -Invoke Disable-LocalUser -Times 0 -Exactly
        Should -Invoke Remove-LocalUser -Times 0 -Exactly
    }
}

Describe 'Installers' {
    BeforeAll {
        function New-FakeProc {
            param($ExitCode = 0, [int]$ProcessId = 4242, [bool]$Exits = $true)
            $p = [pscustomobject]@{ Id = $ProcessId; ExitCode = $ExitCode; Handle = [IntPtr]::Zero; HasExited = $Exits }
            $p | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { param($ms) $this.HasExited } -Force
            $p
        }
    }

    BeforeEach {
        Reset-DepotState
        $script:SavedTemp = $script:DepotTemp
        $script:DepotTemp = $TestDrive
        Mock Start-Sleep { }
        Mock Get-Process { }
        Mock Get-Service { }
        Mock Get-ItemProperty { }
        # A filtered mock outranks the safety net's filtered Remove-Item guard, so this one is filtered too.
        Mock Remove-Item { } -ParameterFilter { $true }
        Mock Write-Host { }
        Mock Write-StatusTag { }
        Mock Get-DepotWebText { throw 'no network in tests' }
        Mock New-Object { throw 'no network in tests' } -ParameterFilter { $TypeName -match 'WebClient' }
        # Nothing is installed unless a test says so. TestDrive paths still reach the real Test-Path.
        Mock Test-Path { $false } -ParameterFilter { "$Path" -notlike "$TestDrive*" }
        Mock Get-Installer { $true }
        Mock Start-Process { New-FakeProc -ExitCode 0 }
        $global:LASTEXITCODE = 73
    }

    AfterEach {
        $script:DepotTemp = $script:SavedTemp
        $global:LASTEXITCODE = 0
    }

    Context 'Finding 35: the exit code comes from the waited process, not a stale LASTEXITCODE' {
        It '<Fn> reports the child exit code' -ForEach @(
            @{ Fn = 'Install-DepotTeams' },
            @{ Fn = 'Install-DepotChrome' },
            @{ Fn = 'Install-DepotAcrobat' },
            @{ Fn = 'Install-DepotZoom' },
            @{ Fn = 'Install-Depot7Zip' }
        ) {
            Mock Start-Process { New-FakeProc -ExitCode 17 }
            & $Fn
            Should -Invoke Write-StatusTag -Times 1 -Exactly -ParameterFilter { $Text -eq '[FAILED - exit 17]' }
        }

        It '<Fn> reports DONE for exit 0 even when LASTEXITCODE is stale and nonzero' -ForEach @(
            @{ Fn = 'Install-DepotTeams' },
            @{ Fn = 'Install-DepotChrome' },
            @{ Fn = 'Install-DepotAcrobat' },
            @{ Fn = 'Install-DepotZoom' },
            @{ Fn = 'Install-Depot7Zip' }
        ) {
            Mock Start-Process { New-FakeProc -ExitCode 0 }
            & $Fn
            Should -Invoke Write-StatusTag -Times 1 -Exactly -ParameterFilter { $Text -eq '[DONE]' }
        }

        It "returns -2, not a cancel, when the child exit code cannot be read" {
            Mock Start-Process { New-FakeProc -ExitCode $null }
            (Start-Installer -FilePath (Join-Path $TestDrive "x.exe") -ArgumentList "/S") | Should -Be -2
        }

        It 'runs the Office setup through the waiting helper and reports its exit code' {
            Mock Get-Installer { $true }
            Mock Start-Process {
                if ("$FilePath" -match 'ODT\.exe$') {
                    New-Item -ItemType Directory -Path (Join-Path $TestDrive 'ODT') -Force | Out-Null
                    Set-Content -LiteralPath (Join-Path $TestDrive 'ODT\setup.exe') -Value 'stub'
                    New-FakeProc -ExitCode 0
                } else { New-FakeProc -ExitCode 17 }
            }
            Install-DepotOffice
            Should -Invoke Write-StatusTag -Times 1 -Exactly -ParameterFilter { $Text -eq '[FAILED - exit 17]' }
        }
    }
}

Describe 'Finding 51: MSI Office is left for ODT RemoveMSI to remove' {
    BeforeAll {
        function New-FakeProc {
            param($ExitCode = 0, [int]$ProcessId = 4242)
            $p = [pscustomobject]@{ Id = $ProcessId; ExitCode = $ExitCode; Handle = [IntPtr]::Zero; HasExited = $true }
            $p | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { param($ms) $true } -Force
            $p
        }
    }

    BeforeEach {
        Reset-DepotState
        $script:SavedTemp = $script:DepotTemp
        $script:DepotTemp = $TestDrive
        Mock Start-Sleep { }
        Mock Get-Process { }
        Mock Get-Service { }
        Mock Get-ItemProperty { }
        Mock Remove-Item { } -ParameterFilter { $true }
        Mock Write-Host { }
        Mock Write-StatusTag { }
        Mock Get-DepotWebText { throw 'no network in tests' }
        Mock New-Object { throw 'no network in tests' } -ParameterFilter { $TypeName -match 'WebClient' }
        Mock Test-Path { $false } -ParameterFilter { "$Path" -notlike "$TestDrive*" }
        Mock Get-Installer { $true }
        Mock Start-Process {
            if ("$FilePath" -match 'ODT\.exe$') {
                New-Item -ItemType Directory -Path (Join-Path $TestDrive 'ODT') -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $TestDrive 'ODT\setup.exe') -Value 'stub'
            }
            New-FakeProc -ExitCode 0
        }
    }

    AfterEach {
        $script:DepotTemp = $script:SavedTemp
        Remove-Item -LiteralPath (Join-Path $TestDrive 'ODT') -Recurse -Force -ErrorAction SilentlyContinue
    }

    It 'does not delete Program Files\Microsoft Office or the Office 16.0 keys when an MSI Office product is installed' {
        Mock Get-ItemProperty {
            [pscustomobject]@{ DisplayName = 'Microsoft Office Professional Plus 2016'; WindowsInstaller = 1 }
        }
        Install-DepotOffice
        Should -Invoke Remove-Item -Times 0 -Exactly -ParameterFilter { "$Path" -match 'Microsoft Office|\\Office\\16\.0|ClickToRun|AppVISV' }
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter { "$FilePath" -match 'setup\.exe$' }
        (Get-Content -LiteralPath (Join-Path $TestDrive 'ODT\config.xml') -Raw) | Should -Match '<RemoveMSI/>'
    }

    It 'treats an Office1x WINWORD.EXE outside the Click-to-Run root as MSI Office even without an Uninstall entry' {
        Mock Test-Path { $true } -ParameterFilter { "$Path" -like '*Microsoft Office\Office1?\WINWORD.EXE' }
        Install-DepotOffice
        Should -Invoke Remove-Item -Times 0 -Exactly -ParameterFilter { "$Path" -match 'Microsoft Office|\\Office\\16\.0|ClickToRun|AppVISV' }
    }

    It 'still scrubs leftover Click-to-Run remnants when no MSI Office is present' {
        Mock Get-ItemProperty {
            [pscustomobject]@{ DisplayName = 'Microsoft Office 16 Click-to-Run Extensibility Component'; WindowsInstaller = 1 }
            [pscustomobject]@{ DisplayName = 'Contoso Agent'; WindowsInstaller = 1 }
        }
        Install-DepotOffice
        Should -Invoke Remove-Item -Times 1 -Exactly -ParameterFilter { "$Path" -match 'Office\\ClickToRun$' }
        Should -Invoke Start-Process -Times 1 -Exactly -ParameterFilter { "$FilePath" -match 'setup\.exe$' }
    }
}

Describe 'Finding 52: the update cache is not wiped while USO is running' {
    BeforeEach {
        Reset-DepotState
        Mock Write-Host { }
        Mock Write-StatusTag { }
        Mock Stop-Service { }
        Mock Start-Service { }
        Mock Remove-Item { } -ParameterFilter { $true }
        Mock Clear-RecycleBin { }
        Mock Get-Service { [pscustomobject]@{ Name = 'wuauserv'; Status = 'Running' } }
    }

    It 'does not stop wuauserv or delete SoftwareDistribution\Download during cleanup' {
        Invoke-DepotCleanup
        Should -Invoke Stop-Service -Times 0 -Exactly -ParameterFilter { "$Name" -eq 'wuauserv' }
        Should -Invoke Remove-Item -Times 0 -Exactly -ParameterFilter { "$Path" -match 'SoftwareDistribution' }
    }

    It 'still purges the temp folders and the recycle bin' {
        Invoke-DepotCleanup
        Should -Invoke Remove-Item -Times 1 -Exactly -ParameterFilter { "$Path" -like 'C:\Windows\Temp\*' }
        Should -Invoke Clear-RecycleBin -Times 1 -Exactly
    }
}

Describe 'Finding 13: only DEPOT-owned processes are stopped' {
    BeforeAll {
        function New-FakeProc {
            param($ExitCode = 0, [int]$ProcessId = 4242, [bool]$Exits = $true)
            $p = [pscustomobject]@{ Id = $ProcessId; ExitCode = $ExitCode; Handle = [IntPtr]::Zero; HasExited = $Exits }
            $p | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { param($ms) $this.HasExited } -Force
            $p
        }
        function New-Row {
            param([int]$ProcessId, [int]$ParentId, [string]$Name, [string]$Path = '', [string]$Cmd = '')
            [pscustomobject]@{ ProcessId = $ProcessId; ParentProcessId = $ParentId; Name = $Name; ExecutablePath = $Path; CommandLine = $Cmd }
        }
    }

    BeforeEach {
        Reset-DepotState
        $script:Killed = @()
        $script:Rows = @(
            (New-Row 100 4   'chrome.exe'   'C:\Program Files\Google\Chrome\Application\chrome.exe' '"chrome.exe"'),
            (New-Row 101 4   'msiexec.exe'  'C:\Windows\System32\msiexec.exe' 'msiexec.exe /i "C:\Other\LineOfBusiness.msi"'),
            (New-Row 102 4   'TrustedInstaller.exe' 'C:\Windows\servicing\TrustedInstaller.exe' ''),
            (New-Row 103 4   '7zFM.exe'     'C:\Program Files\7-Zip\7zFM.exe' '"7zFM.exe"'),
            (New-Row 200 4   'setup.exe'    'C:\Windows\Temp\ODT\setup.exe' '"C:\Windows\Temp\ODT\setup.exe" /configure x.xml'),
            (New-Row 201 4   'msiexec.exe'  'C:\Windows\System32\msiexec.exe' 'msiexec.exe /i "C:\Windows\Temp\ChromeEnterprise.msi" /qn'),
            (New-Row 500 4   'TeamsSetup.exe' 'C:\Windows\Temp\TeamsSetup.exe' ''),
            (New-Row 501 500 'child.exe'    'C:\Windows\Temp\x\child.exe' ''),
            (New-Row 502 501 'grandchild.exe' 'C:\Windows\System32\grandchild.exe' '')
        )
        Mock Write-Host { }
        Mock Start-Sleep { }
        Mock Get-DepotProcessInfo { $script:Rows }
        Mock Get-Process {
            $script:Rows | Where-Object { -not $Name -or $_.Name -replace '\.exe$', '' -in $Name } |
                ForEach-Object { [pscustomobject]@{ Name = ($_.Name -replace '\.exe$', ''); Path = $_.ExecutablePath; Id = $_.ProcessId } }
        }
        Mock Stop-Process {
            if ($Id) { $script:Killed += $Id } elseif ($InputObject) { $script:Killed += $InputObject.Id } elseif ($Name) { $script:Killed += "name:$Name" }
        }
        Mock Test-DepotEscPressed { $false }
        Mock Start-Process { New-FakeProc -ExitCode 0 -ProcessId 500 -Exits $false }
    }

    It 'the startup cleanup stops only processes running from DEPOT temp files' {
        Stop-DepotLeftoverProcess
        ($script:Killed | Sort-Object) -join ',' | Should -Be '200,201,500'
    }

    It 'cancelling with ESC stops the launched process tree and nothing else' {
        Mock Test-DepotEscPressed { $true }
        $r = Start-Installer -FilePath 'C:\Windows\Temp\TeamsSetup.exe' -ArgumentList '/silent'
        $r | Should -BeNullOrEmpty
        ($script:Killed | Sort-Object) -join ',' | Should -Be '500,501,502'
    }

    It 'a timeout stops the launched process tree, not just the root' {
        $r = Start-Installer -FilePath 'C:\Windows\Temp\TeamsSetup.exe' -ArgumentList '/silent' -TimeoutMinutes 0
        $r | Should -Be -1
        ($script:Killed | Sort-Object) -join ',' | Should -Be '500,501,502'
    }

    It 'the Office and Chrome steps do not stop processes by name' {
        $script:SavedTemp = $script:DepotTemp
        $script:DepotTemp = $TestDrive
        try {
            Mock Test-Path { $false } -ParameterFilter { "$Path" -notlike "$TestDrive*" }
            Mock Get-Service { }
            Mock Get-ItemProperty { }
            Mock Get-Installer { $false }
            Mock Write-StatusTag { }
            Mock Remove-Item { } -ParameterFilter { $true }
            Install-DepotOffice
            Install-DepotChrome
            Should -Invoke Stop-Process -Times 0 -Exactly
        } finally { $script:DepotTemp = $script:SavedTemp }
    }
}

Describe 'Finding 14: unattended runs, honest results and exit codes' {
    BeforeEach {
        Reset-DepotState
        Mock Start-Sleep { }
        Mock Write-Host { }
        Mock Write-HLine { }
        Mock Write-StatusTag { }
        Mock Clear-Host { }
        Mock Read-Host { throw 'a prompt was shown' }
        Mock Restart-Computer { }
        Mock Start-Process { }
        Mock Remove-Item { } -ParameterFilter { $true }
        Mock Test-Path { $false } -ParameterFilter { "$Path" -notlike "$TestDrive*" }
        Mock Write-DepotRunLog { }
        $script:DepotScriptPath = $null
    }

    Context 'no console, no technician' {
        It 'Read-DepotChoice answers N without reading a key when the host is not interactive' {
            Mock Test-DepotInteractive { $false }
            Read-DepotChoice 'Rename? ' | Should -Be 'N'
        }

        It 'the interactive setup does nothing destructive when the host is not interactive' {
            Mock Test-DepotInteractive { $false }
            Mock Invoke-DepotRename { }
            Mock New-DepotLocalAdmin { }
            Mock Invoke-DepotAdminPurge { }
            Invoke-DepotInteractiveSetup
            Should -Invoke Invoke-DepotRename -Times 0 -Exactly
            Should -Invoke New-DepotLocalAdmin -Times 0 -Exactly
            Should -Invoke Invoke-DepotAdminPurge -Times 0 -Exactly
        }

        It 'a console that cannot be cleared does not stop the run' {
            Mock Test-DepotInteractive { $false }
            Mock Clear-Host { throw 'The handle is invalid' }
            { Invoke-DepotInteractiveSetup } | Should -Not -Throw
        }

        It 'the final prompt is skipped and nothing restarts when the host is not interactive' {
            Mock Test-DepotInteractive { $false }
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:0'
            Should -Invoke Read-Host -Times 0 -Exactly
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }

        It 'a narrow width does not make the footer padding negative' {
            Mock Test-DepotInteractive { $false }
            $w = $script:Width
            $script:Width = 30
            try { { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:0' } finally { $script:Width = $w }
        }
    }

    Context 'results decide the exit code' {
        BeforeEach { Mock Test-DepotInteractive { $false } }

        It 'exits 3 when some steps worked and one failed' {
            Add-DepotResult 'Chrome' 'Done'
            Add-DepotResult 'Zoom' 'Failed' 'download'
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:3'
        }

        It 'exits 1 when every step that ran failed' {
            Add-DepotResult 'Chrome' 'Failed' 'exit 1603'
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:1'
        }

        It 'a failure outranks a pending reboot' {
            Add-DepotResult 'Chrome' 'Done'
            Add-DepotResult 'Zoom' 'Failed' 'download'
            $script:RebootRequired = $true
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:3'
        }

        It 'exits 3010 for a clean run that needs a reboot, and 0 for a clean run that does not' {
            Add-DepotResult 'Chrome' 'Done'
            $script:RebootRequired = $true
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
            $script:RebootRequired = $false
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:0'
        }

        It 'skipped and already-installed steps are not failures' {
            Add-DepotResult 'Chrome' 'AlreadyInstalled'
            Add-DepotResult 'Zoom' 'Skipped' 'ESC'
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:0'
        }

        It 'writes a plain final line whose code matches the exit code' {
            Add-DepotResult 'Chrome' 'Done'
            Add-DepotResult 'Zoom' 'Failed' 'download'
            Mock Write-Output { }
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:3'
            Should -Invoke Write-Output -Times 1 -Exactly -ParameterFilter { "$InputObject" -match 'exit code 3\b' -and "$InputObject" -match 'FAILURES' }
        }

        It 'the footer only says COMPLETE for a clean run' {
            Add-DepotResult 'Zoom' 'Failed' 'download'
            Mock Write-Host { } -ParameterFilter { "$Object" -match 'SEQUENCE COMPLETE' }
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:1'
            Should -Invoke Write-Host -Times 0 -Exactly -ParameterFilter { "$Object" -match 'SEQUENCE COMPLETE' }
        }

        It 'writes the run log' {
            Add-DepotResult 'Zoom' 'Failed' 'download'
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:1'
            Should -Invoke Write-DepotRunLog -Times 1 -Exactly
        }
    }

    Context 'self-delete only after a good run' {
        BeforeEach {
            Mock Test-DepotInteractive { $false }
            $script:DepotScriptPath = Join-Path $TestDrive 'DEPOT.ps1'
            Set-Content -LiteralPath $script:DepotScriptPath -Value '# x'
        }

        It 'does not delete the script after a failed run, so the technician can rerun it' {
            Add-DepotResult 'Zoom' 'Failed' 'download'
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:1'
            Should -Invoke Start-Process -Times 0 -Exactly
        }

        It 'deletes the script after a clean run' {
            Add-DepotResult 'Zoom' 'Done'
            { Complete-Depot } | Should -Throw 'KILLERSCRIPT_EXIT:0'
            Should -Invoke Start-Process -Times 1 -Exactly
        }
    }

    Context 'steps record what really happened' {
        It 'Show-Result records Failed for a failing exit code and Done for success' {
            $script:DepotStep = 'Zoom'
            Show-Result 1603
            Show-Result 0
            Show-Result $null
            @($script:DepotResults | ForEach-Object { $_.Status }) -join ',' | Should -Be 'Failed,Done,Skipped'
            @($script:DepotResults | ForEach-Object { $_.Step } | Select-Object -Unique) | Should -Be 'Zoom'
        }

        It 'Show-Result with 3010 marks a reboot as required' {
            $script:DepotStep = 'Teams'
            Show-Result 3010
            $script:RebootRequired | Should -BeTrue
        }

        It 'a Windows Update trigger that fails records Failed and does not claim a reboot' {
            Mock Stop-Service { }
            Mock Start-Service { }
            Mock Start-Process { throw 'UsoClient not found' }
            Invoke-DepotWindowsUpdate
            $script:RebootRequired | Should -BeFalse
            @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 1
        }

        It 'a Windows Update trigger that works records Done and a reboot' {
            Mock Stop-Service { }
            Mock Start-Service { }
            Mock Start-Process { }
            Invoke-DepotWindowsUpdate
            $script:RebootRequired | Should -BeTrue
            @($script:DepotResults | Where-Object { $_.Status -eq 'Done' }).Count | Should -Be 1
        }
    }

    Context 'cleanup after a failed step keeps the evidence' {
        BeforeEach {
            Mock Clear-RecycleBin { }
        }

        It 'does not purge C:\Windows\Temp when a step failed' {
            Add-DepotResult 'Zoom' 'Failed' 'download'
            Invoke-DepotCleanup
            Should -Invoke Remove-Item -Times 0 -Exactly -ParameterFilter { "$Path" -like 'C:\Windows\Temp*' }
        }

        It 'purges C:\Windows\Temp after a clean run' {
            Add-DepotResult 'Zoom' 'Done'
            Invoke-DepotCleanup
            Should -Invoke Remove-Item -Times 1 -ParameterFilter { "$Path" -like 'C:\Windows\Temp*' }
        }
    }
}

Describe 'Finding 42: per-user hives are loaded, verified and unloaded honestly' {
    BeforeAll {
        function New-ProfileRow { param($Sid, $Path) [pscustomobject]@{ PSChildName = $Sid; ProfileImagePath = $Path } }
        $script:SidA = 'S-1-5-21-1111111111-2222222222-3333333333-1001'
        $script:SidB = 'S-1-5-21-1111111111-2222222222-3333333333-1002'
        $script:SidAad = 'S-1-12-1-1111111111-2222222222-3333333333-4444444444'
    }

    BeforeEach {
        Reset-DepotState
        Mock Start-Sleep { }
        Mock Write-Host { }
        $script:Rows  = @()
        $script:Known = @()
        $script:RegCalls = @()
        $script:LoadExit = 0
        $script:LoadFailFor = $null
        $script:UnloadExit = 0
        Mock Get-ItemProperty { $script:Rows }
        Mock Test-Path { "$Path" -in $script:Known } -ParameterFilter { "$Path" -notlike "$TestDrive*" }
        Mock Invoke-DepotReg {
            $script:RegCalls += ,($Arguments -join ' ')
            if ($Arguments[0] -eq 'load') {
                $bad = ($script:LoadExit -ne 0) -or ($script:LoadFailFor -and "$($Arguments[1])" -like "*$($script:LoadFailFor)")
                if (-not $bad) { $script:Known += "Registry::HKEY_USERS\$($Arguments[1] -replace '^HKU\\','')" }
                return [pscustomobject]@{ ExitCode = $(if ($bad) { 1 } else { 0 }); Output = @() }
            }
            if ($Arguments[0] -eq 'unload') { return [pscustomobject]@{ ExitCode = $script:UnloadExit; Output = @() } }
            [pscustomobject]@{ ExitCode = 0; Output = @() }
        }
    }

    It 'only real user SIDs are accepted: local and AAD user SIDs yes, built-ins and look-alikes no' {
        $script:Rows = @(
            (New-ProfileRow $script:SidA 'C:\Users\a'),
            (New-ProfileRow $script:SidAad 'C:\Users\aad'),
            (New-ProfileRow 'S-1-5-18' 'C:\Windows\system32\config\systemprofile'),
            (New-ProfileRow 'S-1-5-21-1-2-3-1001.bak' 'C:\Users\bak'),
            (New-ProfileRow 'xS-1-5-21-1-2-3-1001' 'C:\Users\x'),
            (New-ProfileRow 'S-1-5-21-1-2-3' 'C:\Users\short')
        )
        $script:Known = @('C:\Users\a\NTUSER.DAT', 'C:\Users\aad\NTUSER.DAT', 'C:\Windows\system32\config\systemprofile\NTUSER.DAT', 'C:\Users\bak\NTUSER.DAT', 'C:\Users\x\NTUSER.DAT', 'C:\Users\short\NTUSER.DAT')
        $sids = @(Get-DepotProfileHive | ForEach-Object { $_.SID })
        $sids -join ',' | Should -Be "$($script:SidA),$($script:SidAad)"
    }

    It 'an already loaded hive (a logged-on user) is written directly and never loaded or unloaded' {
        $script:Rows = @(New-ProfileRow $script:SidA 'C:\Users\a')
        $script:Known = @('C:\Users\a\NTUSER.DAT', "Registry::HKEY_USERS\$($script:SidA)")
        Mount-DepotUserHive
        $script:UserHives.Count | Should -Be 1
        $script:UserHives[0].Mount | Should -Be "Registry::HKEY_USERS\$($script:SidA)"
        Dismount-DepotUserHive
        $script:RegCalls | Should -BeNullOrEmpty
    }

    It 'a hive that loads is counted and unloaded exactly once' {
        $script:Rows = @(New-ProfileRow $script:SidA 'C:\Users\a')
        $script:Known = @('C:\Users\a\NTUSER.DAT')
        Mount-DepotUserHive
        $script:UserHives.Count | Should -Be 1
        Dismount-DepotUserHive
        @($script:RegCalls | Where-Object { $_ -like 'unload*' }).Count | Should -Be 1
        @($script:RegCalls | Where-Object { $_ -like 'load*' }).Count | Should -Be 1
    }

    It 'a reg load that fails is not counted, not unloaded and is recorded as Failed' {
        $script:Rows = @(New-ProfileRow $script:SidA 'C:\Users\a')
        $script:Known = @('C:\Users\a\NTUSER.DAT')
        $script:LoadExit = 1
        Mount-DepotUserHive
        $script:UserHives.Count | Should -Be 0
        Dismount-DepotUserHive
        @($script:RegCalls | Where-Object { $_ -like 'unload*' }).Count | Should -Be 0
        @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 1
    }

    It 'a reg load that says success but leaves no mounted key is not counted' {
        $script:Rows = @(New-ProfileRow $script:SidA 'C:\Users\a')
        $script:Known = @('C:\Users\a\NTUSER.DAT')
        Mock Invoke-DepotReg { [pscustomobject]@{ ExitCode = 0; Output = @() } }
        Mount-DepotUserHive
        $script:UserHives.Count | Should -Be 0
    }

    It 'an unload that keeps failing is recorded as Failed' {
        $script:Rows = @(New-ProfileRow $script:SidA 'C:\Users\a')
        $script:Known = @('C:\Users\a\NTUSER.DAT')
        Mount-DepotUserHive
        $script:UnloadExit = 1
        Dismount-DepotUserHive
        @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' -and $_.Detail -match 'unload' }).Count | Should -Be 1
    }

    It 'the privacy step reports how many profiles it really reached' {
        $script:Rows = @((New-ProfileRow $script:SidA 'C:\Users\a'), (New-ProfileRow $script:SidB 'C:\Users\b'))
        $script:Known = @('C:\Users\a\NTUSER.DAT', 'C:\Users\b\NTUSER.DAT')
        Mock Set-RegValue { }
        Mock Set-ItemProperty { } -ParameterFilter { $true }
        Mock Get-Service { }
        Mock New-Item { } -ParameterFilter { $true }
        Mock Get-Item { } -ParameterFilter { $true }
        $script:LoadFailFor = '1002'
        Invoke-DepotPrivacy
        Should -Invoke Write-Host -Times 1 -Exactly -ParameterFilter { "$Object" -match 'applied to 1 of 2 user profile' }
    }
}

Describe 'Finding 71: rename and OneDrive report what really happened' {
    BeforeAll {
        function New-FakeProc {
            param($ExitCode = 0, [int]$ProcessId = 4242, [bool]$Exits = $true)
            $p = [pscustomobject]@{ Id = $ProcessId; ExitCode = $ExitCode; Handle = [IntPtr]::Zero; HasExited = $Exits }
            $p | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { param($ms) $this.HasExited } -Force
            $p
        }
    }

    BeforeEach {
        Reset-DepotState
        $script:SavedTemp = $script:DepotTemp
        $script:DepotTemp = $TestDrive
        $script:SavedVerify = $script:OneDriveVerifySeconds
        $script:OneDriveVerifySeconds = 0
        Mock Start-Sleep { }
        Mock Write-Host { }
        Mock Write-StatusTag { }
        Mock Get-Process { }
        Mock Get-Service { }
        Mock Get-ItemProperty { }
        Mock Remove-Item { } -ParameterFilter { $true }
        Mock Rename-Computer { }
        Mock Read-Host { 'NEWPC01' }
        $script:Known = @()
        Mock Test-Path { "$Path" -in $script:Known } -ParameterFilter { "$Path" -notlike "$TestDrive*" }
        Mock Get-Installer { $true }
        Mock Start-Process { New-FakeProc -ExitCode 0 }
    }

    AfterEach {
        $script:DepotTemp = $script:SavedTemp
        $script:OneDriveVerifySeconds = $script:SavedVerify
    }

    Context 'rename' {
        It 'a rename that throws is reported as failed, not Done, and sets no reboot flag' {
            Mock Rename-Computer { throw 'Access is denied' }
            Invoke-DepotRename
            @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 1
            $script:RebootRequired | Should -BeFalse
            Should -Invoke Write-Host -Times 0 -Exactly -ParameterFilter { "$Object" -match '^ Done' }
        }

        It 'a rename that works is recorded Done and requires a reboot' {
            Invoke-DepotRename
            Should -Invoke Rename-Computer -Times 1 -Exactly -ParameterFilter { $NewName -eq 'NEWPC01' }
            @($script:DepotResults | Where-Object { $_.Status -eq 'Done' }).Count | Should -Be 1
            $script:RebootRequired | Should -BeTrue
        }

        It 'an invalid computer name is refused before Rename-Computer is called' {
            Mock Read-Host { 'bad name!' }
            Invoke-DepotRename
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
    }

    Context 'OneDrive' {
        It 'installs through Start-Installer and announces it once' {
            Mock Write-Installing { }
            $script:Known = @("$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe")
            # not installed before the run: hide the file until the installer has "run"
            $script:Known = @()
            Mock Start-Installer { $script:Known = @("$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"); 0 }
            Install-DepotOneDrive
            Should -Invoke Write-Installing -Times 1 -Exactly
        }

        It 'exit 0 without OneDrive.exe on disk is a failure' {
            Mock Start-Installer { 0 }
            Install-DepotOneDrive
            @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 1
            @($script:DepotResults | Where-Object { $_.Status -eq 'Done' }).Count | Should -Be 0
        }

        It 'exit 0 with OneDrive.exe on disk is Done' {
            Mock Start-Installer { $script:Known = @("$env:ProgramFiles\Microsoft OneDrive\OneDrive.exe"); 0 }
            Install-DepotOneDrive
            @($script:DepotResults | Where-Object { $_.Status -eq 'Done' }).Count | Should -Be 1
            @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 0
        }

        It 'finds the per-machine OneDrive.exe through ProgramW6432 when a 32-bit host reports Program Files (x86)' {
            $saved = $env:ProgramW6432
            try {
                $env:ProgramW6432 = 'C:\PW6432'
                Mock Start-Installer { $script:Known = @('C:\PW6432\Microsoft OneDrive\OneDrive.exe'); 0 }
                Install-DepotOneDrive
            } finally { $env:ProgramW6432 = $saved }
            @($script:DepotResults | Where-Object { $_.Status -eq 'Done' }).Count | Should -Be 1
        }

        It 'a failing installer exit code is Failed whether or not a file exists' {
            Mock Start-Installer { 1603 }
            Install-DepotOneDrive
            @($script:DepotResults | Where-Object { $_.Status -eq 'Failed' }).Count | Should -Be 1
        }
    }
}

Describe 'Banner survives a host without local account cmdlets' {
    It 'Show-DepotBanner does not throw when Get-LocalGroupMember is unavailable (32-bit host)' {
        Mock Write-Host { }
        Mock Write-HLine { }
        Mock Clear-Host { }
        Mock Get-LocalGroupMember { throw 'The term is not recognized' }
        { Show-DepotBanner } | Should -Not -Throw
    }

    It 'Show-DepotBanner asks for the Administrators group by SID, not by its localized name' {
        Mock Write-Host { }
        Mock Write-HLine { }
        Mock Clear-Host { }
        Mock Get-LocalGroupMember { @() }
        Show-DepotBanner
        Should -Invoke Get-LocalGroupMember -Times 1 -Exactly -ParameterFilter { $SID -and "$SID" -eq 'S-1-5-32-544' }
    }
}
