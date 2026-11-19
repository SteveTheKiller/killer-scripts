#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# ORCA tests. Every process, package, registry, credential and winget call is mocked or goes
# to a fake profile tree under TestDrive. Nothing here touches the real machine.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    # Commands ORCA can reach that the shared safety net does not list. Blocked unless a test mocks them.
    $block = { throw "UNMOCKED DESTRUCTIVE CALL: $($MyInvocation.MyCommand.Name) was not mocked by this test." }
    foreach ($c in 'Install-PackageProvider', 'Import-Module') {
        if (Get-Command $c -ErrorAction SilentlyContinue) { Mock $c $block }
    }
    Enable-KillerTestMode
    . (Get-KillerScriptPath ORCA)

    function New-OrcaKey {
        param([string]$Key = 'Enter', [string]$Char = "`r")
        [pscustomobject]@{ Key = [ConsoleKey]$Key; KeyChar = [char]$Char }
    }

    # Runs Invoke-OrcaMain, capturing every stream as text and the KILLERSCRIPT_EXIT code it throws.
    function Invoke-OrcaRun {
        param([switch]$Silent)
        $script:Code = $null
        $script:Lines = New-Object System.Collections.ArrayList
        try {
            Invoke-OrcaMain -Silent:$Silent *>&1 | ForEach-Object { [void]$script:Lines.Add("$_") }
        } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $script:Code = [int]$Matches[1] } else { throw }
        }
        $script:Text = ($script:Lines -join "`n")
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'ORCA' {
    BeforeEach {
        # TestDrive is shared by every test in the block, so start each test from a clean tree.
        foreach ($leaf in 'Users', 'Temp') { $d = Join-Path $TestDrive $leaf; if ([IO.Directory]::Exists($d)) { [IO.Directory]::Delete($d, $true) } }
        $script:UsersRoot = Join-Path $TestDrive 'Users'
        $script:TempRoot  = Join-Path $TestDrive 'Temp'
        New-Item -ItemType Directory -Path $script:TempRoot -Force | Out-Null
        foreach ($u in 'alice', 'bob') {
            $out = Join-Path $script:UsersRoot "$u\AppData\Local\Microsoft\Outlook"
            New-Item -ItemType Directory -Path $out -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $out "$u@contoso.com.ost") -Value "synthetic OST for $u"
        }

        $script:Elevated   = $true
        $script:Active     = 'alice'
        $script:Keys       = New-Object System.Collections.Queue
        $script:Stopped    = @()
        $script:Provisioned = @([pscustomobject]@{ DisplayName = 'Microsoft.OutlookForWindows'; PackageName = 'Microsoft.OutlookForWindows_1.0_neutral_~_8wekyb3d8bbwe' })
        $script:UserPkgs   = @([pscustomobject]@{ Name = 'Microsoft.OutlookForWindows'; PackageFullName = 'Microsoft.OutlookForWindows_1.0_x64__8wekyb3d8bbwe' })
        $script:Started    = @()

        Mock Test-OrcaElevated { $script:Elevated }
        Mock Get-OrcaUsersRoot { $script:UsersRoot }
        Mock Get-OrcaTempRoot { $script:TempRoot }
        Mock Test-OrcaWingetPresent { $true }
        Mock Get-ActiveUser { $script:Active }
        Mock Read-OrcaKey { if ($script:Keys.Count -gt 0) { $script:Keys.Dequeue() } else { throw 'no console' } }
        Mock Read-Host { throw 'Read-Host must not be reached in this test' }
        Mock Clear-Host { }
        Mock Write-HLine { }
        Mock Write-Host { }
        Mock Start-Sleep { }
        Mock Get-Process { @() }
        Mock Stop-Process { }
        Mock Start-Process { $script:Started += ,@{ File = $FilePath; Args = $ArgumentList } }
        Mock Get-AppxPackage { $script:UserPkgs }
        Mock Remove-AppxPackage { }
        Mock Get-AppxProvisionedPackage { $script:Provisioned }
        Mock Add-AppxProvisionedPackage { }
        Mock Remove-AppxProvisionedPackage { }
        Mock cmdkey { $global:LASTEXITCODE = 0 }
        Mock reg { $global:LASTEXITCODE = 0 }
    }

    Context 'Entry point and exit codes' {
        It 'loads without running (test mode guard)' {
            Get-Command Invoke-OrcaMain -ErrorAction Stop | Should -Not -BeNullOrEmpty
            Get-Command Exit-KillerScript -ErrorAction Stop | Should -Not -BeNullOrEmpty
        }

        It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode' {
            { Exit-KillerScript 3010 } | Should -Throw '*KILLERSCRIPT_EXIT:3010*'
        }

        It 'a silent All Users run that works exits 0 and says so' {
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 0
            $script:Text | Should -Match 'ORCA RESULT: exit code 0 '
        }

        It 'Esc at the profile prompt exits 2 and changes nothing' {
            $script:Keys.Enqueue((New-OrcaKey -Key Escape -Char "`e"))
            Invoke-OrcaRun
            $script:Code | Should -Be 2
            $script:Text | Should -Match 'ORCA RESULT: exit code 2 '
            Assert-MockCalled Stop-Process -Times 0 -Exactly
        }

        It 'a failed New Outlook reinstall exits 1 with a FAILED result line' {
            $script:Provisioned = @()
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 1
            $script:Text | Should -Match 'ORCA RESULT: exit code 1 '
        }

        It 'an unexpected error exits 1 instead of leaking a raw exception' {
            Mock Get-ActiveUser { throw 'boom' }
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 1
            $script:Text | Should -Match 'ORCA RESULT: exit code 1 .*boom'
        }

        It 'a non-elevated run with no determinable user exits 2' {
            $script:Elevated = $false
            Mock Get-ActiveUser { $null }
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 2
        }
    }

    Context 'OST handling (18a, decision D2)' {
        BeforeEach {
            Mock Get-Date { [datetime]'2026-11-18T10:05:09' }
            # Classic only, Delete OSTs = Y, Purge extensions = N. Profile choice is set by each test.
            $script:AnswersAfterProfile = @('2', 'y', 'n')
            function Set-OrcaAnswers { param([string]$Profile)
                $script:Keys.Enqueue((New-OrcaKey -Key D0 -Char $Profile))
                foreach ($a in $script:AnswersAfterProfile) { $script:Keys.Enqueue((New-OrcaKey -Key A -Char $a)) }
            }
            $script:AliceOst = Join-Path $script:UsersRoot 'alice\AppData\Local\Microsoft\Outlook\alice@contoso.com.ost'
            $script:BobOst   = Join-Path $script:UsersRoot 'bob\AppData\Local\Microsoft\Outlook\bob@contoso.com.ost'
        }

        It 'renames the OST to NAME.ost.bak-TIMESTAMP and keeps its content' {
            Set-OrcaAnswers '0'
            Invoke-OrcaRun
            $bak = "$script:AliceOst.bak-20261118-100509"
            Test-Path -LiteralPath $script:AliceOst | Should -BeFalse
            Test-Path -LiteralPath $bak | Should -BeTrue
            (Get-Content -LiteralPath $bak -Raw).Trim() | Should -Be 'synthetic OST for alice'
            $script:Code | Should -Be 0
        }

        It 'leaves the other profile OST alone when one profile is selected' {
            Set-OrcaAnswers '0'
            Invoke-OrcaRun
            Test-Path -LiteralPath $script:BobOst | Should -BeTrue
        }

        It 'backs up every selected profile when All Users is chosen' {
            Set-OrcaAnswers 'A'
            Invoke-OrcaRun
            Test-Path -LiteralPath "$script:BobOst.bak-20261118-100509" | Should -BeTrue
            Test-Path -LiteralPath "$script:AliceOst.bak-20261118-100509" | Should -BeTrue
        }

        It 'reports the backup path in the results, not a deletion' {
            Mock Write-Host { $script:HostText += ($Object -join ' ') + "`n" }
            $script:HostText = ''
            Set-OrcaAnswers '0'
            Invoke-OrcaRun
            $script:HostText | Should -Match 'ost\.bak-20261118-100509'
            $script:HostText | Should -Not -Match 'Deleted OSTs'
        }

        It 'does not touch an earlier .ost.bak-* file' {
            $old = "$script:AliceOst.bak-20250101-000000"
            Set-Content -LiteralPath $old -Value 'older backup'
            Set-OrcaAnswers '0'
            Invoke-OrcaRun
            (Get-Content -LiteralPath $old -Raw).Trim() | Should -Be 'older backup'
            @(Get-ChildItem -LiteralPath (Split-Path $script:AliceOst) -Filter '*.bak-*').Count | Should -Be 2
        }

        It 'exits 3 and says so when the OST is locked and cannot be preserved' {
            Set-OrcaAnswers '0'
            $lock = [System.IO.File]::Open($script:AliceOst, 'Open', 'Read', 'None')
            try { Invoke-OrcaRun } finally { $lock.Dispose() }
            Test-Path -LiteralPath $script:AliceOst | Should -BeTrue
            $script:Code | Should -Be 3
            $script:Text | Should -Match 'ORCA RESULT: exit code 3 .*OST'
        }
    }
}
