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
}
