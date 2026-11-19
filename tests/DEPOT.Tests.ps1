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
