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

}
