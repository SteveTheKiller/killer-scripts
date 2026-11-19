#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# MACE: entry point, scope selection (16a), selected-profile isolation (16b) and OneDrive data (38).
# Every system boundary is a small Mace* wrapper or a mocked cmdlet. Nothing here touches the real machine.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath MACE)

    # Runs Invoke-MaceMain and returns the exit code plus everything written to the output stream.
    function Invoke-MaceCapture {
        $out = & { try { Invoke-MaceMain } catch { $_ } }
        $last = @($out)[-1]
        $code = $null
        if ($last -is [System.Management.Automation.ErrorRecord]) {
            if ($last.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw $last }
        }
        [pscustomobject]@{ Code = $code; Output = (@($out | Where-Object { $_ -is [string] }) -join "`n") }
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'MACE' {
    BeforeEach {
        $script:UsersRoot = Join-Path $TestDrive 'Users'
        $script:AliceDir  = Join-Path $UsersRoot 'alice'
        $script:BobDir    = Join-Path $UsersRoot 'bob'
        $script:AliceSid  = 'S-1-5-21-1-1-1-1001'
        $script:BobSid    = 'S-1-5-21-1-1-1-1002'
        foreach ($d in $AliceDir, $BobDir) {
            foreach ($rel in 'OneDrive', 'OneDrive - Contoso', 'AppData\Local\Microsoft\OneDrive', 'AppData\Local\Microsoft\Teams') {
                New-Item -ItemType Directory -Path (Join-Path $d $rel) -Force | Out-Null
            }
            Set-Content -LiteralPath (Join-Path $d 'OneDrive\unsynced.txt') -Value 'only copy'
        }

        $script:Removed   = @()
        $script:Stopped   = @()
        $script:Started   = @()
        $script:Appx      = @()
        $script:CmdKey    = @()
        $script:Winget    = @()
        $script:SavedUserProfile = $env:USERPROFILE
        $script:SavedModulePath  = $env:PSModulePath
        $env:USERPROFILE = $AliceDir
        Remove-Variable -Name MaceScope -Scope Global -ErrorAction SilentlyContinue

        Mock Test-MaceElevated { $true }
        Mock Test-MaceInteractive { $false }
        Mock Test-MaceConsole { $false }
        Mock Get-MaceCursorTop { -1 }
        Mock Get-MaceCursorLeft { -1 }
        Mock Set-MaceCursor { }
        Mock Write-HLine { }
        Mock Write-Host { }
        Mock Start-Sleep { }
        Mock Wait-Process { }
        Mock Write-MaceFile { }
        # Never scan the real machine's Office folders; TestDrive paths use the real scan.
        Mock Find-MaceProtectedFiles { @() } -ParameterFilter { "$Path" -notlike "$script:TestRoot*" }
        Mock Read-MaceKey { $null }
        Mock Get-MaceProfileDirs { Get-ChildItem -LiteralPath $script:UsersRoot -Directory }
        Mock Get-MaceProfileList {
            @(
                [pscustomobject]@{ PSChildName = $script:AliceSid; ProfileImagePath = $script:AliceDir },
                [pscustomobject]@{ PSChildName = $script:BobSid;   ProfileImagePath = $script:BobDir }
            )
        }
        Mock Resolve-MaceAccount { if ($Sid -eq $script:AliceSid) { 'HOST\alice' } elseif ($Sid -eq $script:BobSid) { 'HOST\bob' } }
        Mock Get-MaceProcess {
            if ($Name -eq 'OneDrive') {
                @([pscustomobject]@{ Id = 101; UserName = 'HOST\alice' }, [pscustomobject]@{ Id = 202; UserName = 'HOST\bob' })
            } else { @() }
        }
        Mock Stop-Process { $script:Stopped += $Id }
        Mock Get-MaceUninstallEntries {
            @(
                [pscustomobject]@{ DisplayName = 'Microsoft Office Professional Plus'; UninstallString = 'MsiExec.exe /X{11111111-2222-3333-4444-555555555555}' },
                [pscustomobject]@{ DisplayName = 'Teams Machine-Wide Installer'; UninstallString = 'MsiExec.exe /X{AAAAAAAA-2222-3333-4444-555555555555}' }
            )
        }
        Mock Test-MaceWinget { $true }
        Mock Invoke-MaceWinget { $script:Winget += ($Arguments -join ' ') }
        Mock Get-MaceAppx {
            if ($Name -eq 'Microsoft.OutlookForWindows') { @([pscustomobject]@{ PackageFullName = 'olk' }) } else { @() }
        }
        Mock Remove-MaceAppx { $script:Appx += ,@{ Name = $Name; Sid = $Sid } }
        Mock Remove-MaceProvisioned { }
        Mock Start-Process { $script:Started += "$FilePath $ArgumentList" }
        # A filter is needed so this mock outranks the safety net's filtered Remove-Item mock.
        Mock Remove-Item { $script:Removed += @(if ($LiteralPath) { $LiteralPath } else { $Path }) } -ParameterFilter { $true }
        # Machine locations always look present so the tests prove they are left alone, not just absent.
        $script:TestRoot = $TestDrive
        Mock Test-Path { $true } -ParameterFilter {
            "$Path" -notlike "$script:TestRoot*" -and (
                "$Path" -like 'C:\Program Files*' -or "$Path" -like 'C:\ProgramData*' -or "$Path" -like 'HKLM:*' -or
                "$Path" -like "$env:SYSTEMROOT*" -or "$Path" -like "$env:LOCALAPPDATA\Microsoft\OneDrive\*" -or "$Path" -like "$env:PROGRAMDATA*")
        }
        Mock cmdkey {
            if ($args[0] -eq '/list') { 'Target: MicrosoftOffice16_Data:SSPI:alice@contoso.com' }
            else { $script:CmdKey += ($args -join ' ') }
        }
    }

    AfterEach {
        $env:USERPROFILE = $script:SavedUserProfile
        $env:PSModulePath = $script:SavedModulePath
        Remove-Variable -Name MaceScope -Scope Global -ErrorAction SilentlyContinue
    }

    Context 'Entry point' {
        It 'defines Invoke-MaceMain without running it when dot-sourced in test mode' {
            Get-Command Invoke-MaceMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
            $script:Removed.Count | Should -Be 0
        }

        It 'exits 2 and changes nothing when not elevated' {
            Mock Test-MaceElevated { $false }
            $global:MaceScope = 'AllUsers'
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 2
            $r.Output | Should -Match 'MACE RESULT: NOT RUN \(exit 2\)'
            $script:Stopped.Count | Should -Be 0
            $script:Removed.Count | Should -Be 0
        }
    }

    Context 'Finding 38: OneDrive sync root' {
        It '38: never removes the OneDrive sync root or an org OneDrive folder, but still removes the client files' {
            $global:MaceScope = 'AllUsers'
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 0
            foreach ($d in $AliceDir, $BobDir) {
                $script:Removed | Should -Not -Contain (Join-Path $d 'OneDrive')
                $script:Removed | Should -Not -Contain (Join-Path $d 'OneDrive - Contoso')
                $script:Removed | Should -Contain (Join-Path $d 'AppData\Local\Microsoft\OneDrive')
            }
            Test-Path -LiteralPath (Join-Path $AliceDir 'OneDrive\unsynced.txt') | Should -BeTrue
        }
    }

    Context 'Finding 16a: scope selection' {
        It '16a: runs all users when non-interactive and no scope is set' {
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 0
            $script:Stopped | Should -Contain 101
            $script:Stopped | Should -Contain 202
        }

        It '16a: Enter at the menu runs all users' {
            Mock Test-MaceInteractive { $true }
            Mock Read-MaceKey { [pscustomobject]@{ VirtualKeyCode = 13; Character = [char]13 } }
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 0
            $script:Stopped | Should -Contain 101
            $script:Stopped | Should -Contain 202
        }

        It '16a: stops with exit 2 before changing anything when an interactive key is invalid' {
            Mock Test-MaceInteractive { $true }
            Mock Read-MaceKey { [pscustomobject]@{ VirtualKeyCode = 88; Character = 'x' } }
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 2
            $script:Stopped.Count | Should -Be 0
            $script:Removed.Count | Should -Be 0
        }

        It '16a: stops with exit 2 before changing anything when $MaceScope names no profile' {
            $global:MaceScope = 'nobody'
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 2
            $script:Stopped.Count | Should -Be 0
            $script:Removed.Count | Should -Be 0
        }

        It '16a: ESC aborts with exit 0 and changes nothing' {
            Mock Test-MaceInteractive { $true }
            Mock Read-MaceKey { [pscustomobject]@{ VirtualKeyCode = 27; Character = [char]27 } }
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 0
            $r.Output | Should -Match 'aborted'
            $script:Stopped.Count | Should -Be 0
            $script:Removed.Count | Should -Be 0
        }

        It '16a: A at the menu runs all users' {
            Mock Test-MaceInteractive { $true }
            Mock Read-MaceKey { [pscustomobject]@{ VirtualKeyCode = 65; Character = 'A' } }
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 0
            $r.Output | Should -Match 'MACE RESULT: SUCCESS \(exit 0\)'
            $script:Stopped | Should -Contain 202
        }

        It '16a: $MaceScope = AllUsers runs all users' {
            $global:MaceScope = 'AllUsers'
            $r = Invoke-MaceCapture
            $r.Code | Should -Be 0
            $script:Stopped | Should -Contain 202
        }
    }

    Context 'Finding 16b: selected profile isolation' {
        It '16b: stops only the selected user''s processes' {
            $global:MaceScope = 'bob'
            $null = Invoke-MaceCapture
            $script:Stopped | Should -Contain 202
            $script:Stopped | Should -Not -Contain 101
        }

        It '16b: leaves machine-wide apps, folders, keys and winget alone for one profile' {
            $global:MaceScope = 'bob'
            $null = Invoke-MaceCapture
            @($script:Removed | Where-Object { $_ -notlike "$UsersRoot*" }).Count | Should -Be 0
            @($script:Started | Where-Object { $_ -match 'msiexec|setup64|OneDriveSetup' }).Count | Should -Be 0
            $script:Winget.Count | Should -Be 0
        }

        It '16b: removes AppX for the selected profile by its SID, never for all users' {
            $global:MaceScope = 'bob'
            $null = Invoke-MaceCapture
            @($script:Appx).Count | Should -BeGreaterThan 0
            foreach ($a in $script:Appx) { $a.Sid | Should -Be $script:BobSid }
        }

        It '16b: does not clear the calling account''s credentials for a different profile and reports partial success' {
            $global:MaceScope = 'bob'
            $r = Invoke-MaceCapture
            $script:CmdKey.Count | Should -Be 0
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'MACE RESULT: PARTIAL \(exit 3\)'
        }

        It '16b: clears credentials and exits 0 when the selected profile is the calling account' {
            $global:MaceScope = 'alice'
            $r = Invoke-MaceCapture
            $script:CmdKey.Count | Should -BeGreaterThan 0
            $r.Code | Should -Be 0
            $script:Stopped | Should -Contain 101
            $script:Stopped | Should -Not -Contain 202
        }

        It '16b: all users still performs the machine-wide removal' {
            $global:MaceScope = 'AllUsers'
            $null = Invoke-MaceCapture
            @($script:Removed | Where-Object { $_ -like 'C:\Program Files*' }).Count | Should -BeGreaterThan 0
            @($script:Started | Where-Object { $_ -match 'msiexec' }).Count | Should -BeGreaterThan 0
        }
    }

    Context 'User data is never deleted' {
        BeforeEach {
            foreach ($d in $AliceDir, $BobDir) {
                foreach ($rel in 'Documents', 'Desktop', 'Pictures') {
                    New-Item -ItemType Directory -Path (Join-Path $d $rel) -Force | Out-Null
                    Set-Content -LiteralPath (Join-Path $d "$rel\keep.txt") -Value 'user file'
                }
                $ol = Join-Path $d 'AppData\Roaming\Microsoft\Outlook'
                New-Item -ItemType Directory -Path $ol -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $ol 'archive.pst') -Value 'mail archive'
                Set-Content -LiteralPath (Join-Path $ol 'cache.ost') -Value 'cache'
                Set-Content -LiteralPath (Join-Path $ol 'old.nk2') -Value 'autocomplete'
            }
            $script:Protected = @()
            foreach ($d in $AliceDir, $BobDir) {
                $script:Protected += @('Documents', 'Desktop', 'Pictures', 'OneDrive', 'OneDrive - Contoso' | ForEach-Object { Join-Path $d $_ })
                $script:Protected += (Join-Path $d 'AppData\Roaming\Microsoft\Outlook\archive.pst')
            }
        }

        It 'no delete target is, contains or sits inside Documents, Desktop, Pictures, a OneDrive root or a .pst' {
            $global:MaceScope = 'AllUsers'
            $null = Invoke-MaceCapture
            $script:Removed.Count | Should -BeGreaterThan 0
            foreach ($prot in $script:Protected) {
                $hits = @($script:Removed | Where-Object { $_ -eq $prot -or $prot -like "$_\*" -or $_ -like "$prot\*" })
                $hits.Count | Should -Be 0 -Because "$prot must survive"
            }
        }

        It 'still wipes the other Outlook data around a .pst and reports partial success' {
            $global:MaceScope = 'alice'
            $r = Invoke-MaceCapture
            $ol = Join-Path $AliceDir 'AppData\Roaming\Microsoft\Outlook'
            $script:Removed | Should -Contain (Join-Path $ol 'old.nk2')
            $script:Removed | Should -Contain (Join-Path $ol 'cache.ost')
            $script:Removed | Should -Not -Contain $ol
            $script:Removed | Should -Not -Contain (Join-Path $ol 'archive.pst')
            $r.Code | Should -Be 3
            $r.Output | Should -Match 'archive\.pst'
        }

        It 'removes the whole tree when it holds no .pst' {
            $global:MaceScope = 'alice'
            [IO.File]::Delete((Join-Path $AliceDir 'AppData\Roaming\Microsoft\Outlook\archive.pst'))
            $null = Invoke-MaceCapture
            $script:Removed | Should -Contain (Join-Path $AliceDir 'AppData\Roaming\Microsoft\Outlook')
        }
    }

    Context 'Shell folder repair' {
        BeforeEach {
            $script:Hive = 'TestRegistry:\MaceHive'
            $script:USF = "$Hive\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
            $script:SF  = "$Hive\Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders"
            New-Item -Path $USF -Force | Out-Null
            New-Item -Path $SF -Force | Out-Null
            $script:UserRegKeys = @()
            $script:UserRunValues = @()
            $script:ShellFixes = @(
                @{ Name = 'Personal';    Default = '%USERPROFILE%\Documents'; Dir = 'Documents' },
                @{ Name = 'Desktop';     Default = '%USERPROFILE%\Desktop';   Dir = 'Desktop'   },
                @{ Name = 'My Pictures'; Default = '%USERPROFILE%\Pictures';  Dir = 'Pictures'  }
            )
        }

        It 'leaves a folder alone when its OneDrive target exists' {
            $target = Join-Path $AliceDir 'OneDrive\Documents'
            New-Item -ItemType Directory -Path $target -Force | Out-Null
            Set-ItemProperty -Path $USF -Name Personal -Value $target
            Invoke-PerUserRegistry -HiveRoot $Hive -ProfilePath $AliceDir
            (Get-ItemProperty -Path $USF).Personal | Should -Be $target
            Test-Path (Join-Path $AliceDir 'Documents') | Should -BeFalse
        }

        It 'repairs a folder whose OneDrive target is missing' {
            $target = Join-Path $AliceDir 'OneDrive\Desktop'
            Set-ItemProperty -Path $USF -Name Desktop -Value $target
            Invoke-PerUserRegistry -HiveRoot $Hive -ProfilePath $AliceDir
            (Get-ItemProperty -Path $USF).Desktop | Should -Be '%USERPROFILE%\Desktop'
            Test-Path (Join-Path $AliceDir 'Desktop') | Should -BeTrue
        }

        It 'does not point the user at a new empty folder while files sit in another OneDrive folder' {
            $missing = Join-Path $AliceDir 'OneDrive\Pictures'
            $other = Join-Path $AliceDir 'OneDrive - Contoso\Pictures'
            New-Item -ItemType Directory -Path $other -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $other 'p.jpg') -Value 'x'
            Set-ItemProperty -Path $USF -Name 'My Pictures' -Value $missing
            Invoke-PerUserRegistry -HiveRoot $Hive -ProfilePath $AliceDir
            (Get-ItemProperty -Path $USF).'My Pictures' | Should -Be $missing
            Test-Path (Join-Path $AliceDir 'Pictures') | Should -BeFalse
        }
    }
}
