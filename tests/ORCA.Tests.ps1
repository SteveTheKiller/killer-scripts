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
    # The real profile reader, kept before any test mocks Get-OrcaProfile.
    $script:RealGetOrcaProfile = ${function:Get-OrcaProfile}
    $script:RealReadOrcaKey = ${function:Read-OrcaKey}
    $script:RealReadOrcaLine = ${function:Read-OrcaLine}

    function New-OrcaKey {
        param([string]$Key = 'Enter', [string]$Char = "`r")
        [pscustomobject]@{ Key = [ConsoleKey]$Key; KeyChar = [char]$Char }
    }

    # Runs Invoke-OrcaMain, capturing every stream as text and the KILLERSCRIPT_EXIT code it throws.
    function Invoke-OrcaRun {
        param([switch]$Silent, [switch]$FullSignOut)
        $script:Code = $null
        $script:Lines = New-Object System.Collections.ArrayList
        try {
            Invoke-OrcaMain -Silent:$Silent -FullSignOut:$FullSignOut *>&1 | ForEach-Object { [void]$script:Lines.Add("$_") }
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
            Set-Content -LiteralPath (Join-Path $script:UsersRoot "$u\NTUSER.DAT") -Value 'synthetic hive'
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
        Mock Get-OrcaProfile {
            @(
                [pscustomobject]@{ Name = 'alice'; Account = 'CONTOSO\alice'; Sid = 'S-1-5-21-1000-2000-3000-1001'; Path = (Join-Path $script:UsersRoot 'alice') },
                [pscustomobject]@{ Name = 'bob'; Account = 'CONTOSO\bob'; Sid = 'S-1-5-21-1000-2000-3000-1002'; Path = (Join-Path $script:UsersRoot 'bob') }
            )
        }
        Mock Get-OrcaCurrentSid { 'S-1-5-21-1000-2000-3000-1001' }
        Mock Get-OrcaTempRoot { $script:TempRoot }
        Mock Test-OrcaWingetPresent { $true }
        Mock Get-ActiveUser { $script:Active }
        Mock Read-OrcaKey { if ($script:Keys.Count -gt 0) { $script:Keys.Dequeue() } else { throw 'no console' } }
        Mock Read-Host { throw 'Read-Host must not be reached in this test' }
        # No console: a line prompt gets no answer, which means no launch and no profile choice.
        Mock Read-OrcaLine { $null }
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
        # winget download writes a package into the requested folder so All Users provisioning has something to add.
        Mock Invoke-OrcaWinget {
            if ($Arguments -like 'download*') {
                $m = [regex]::Match($Arguments, '--download-directory "([^"]+)"')
                New-Item -ItemType Directory -Path $m.Groups[1].Value -Force | Out-Null
                Set-Content -LiteralPath (Join-Path $m.Groups[1].Value 'Outlook.msix') -Value 'synthetic package'
            }
            0
        }
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

        It 'Read-OrcaKey reads the console and does not call itself' {
            $body = $script:RealReadOrcaKey.ToString()
            $body | Should -Match '\[Console\]::ReadKey'
            $body | Should -Not -Match '(?m)^\s*Read-OrcaKey\s*$'
        }

        It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode' {
            { Exit-KillerScript 3010 } | Should -Throw '*KILLERSCRIPT_EXIT:3010*'
        }

        It 'a silent run over only the current user that works exits 0 and says so' {
            Mock Get-OrcaProfile {
                @([pscustomobject]@{ Name = 'alice'; Account = 'CONTOSO\alice'; Sid = 'S-1-5-21-1000-2000-3000-1001'; Path = (Join-Path $script:UsersRoot 'alice') })
            }
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

    Context 'Process termination (18b)' {
        BeforeEach {
            function New-OrcaProc { param([string]$Name, [int]$Id, [string]$User, [string]$Cmd = '')
                [pscustomobject]@{ Name = $Name; Id = $Id; UserName = $User; SessionId = 1; CommandLine = $Cmd }
            }
            $olkCmd   = '"C:\Program Files\WindowsApps\Microsoft.OutlookForWindows_1.0\olk.exe" --webview'
            $teamsCmd = '"C:\Program Files\WindowsApps\MSTeams_1.0\msedgewebview2.exe" --user-data-dir=C:\Users\alice\AppData\Local\Packages\MSTeams'
            $script:Procs = @(
                (New-OrcaProc 'OUTLOOK' 100 'CONTOSO\alice'),
                (New-OrcaProc 'OUTLOOK' 101 'CONTOSO\bob'),
                (New-OrcaProc 'olk' 200 'CONTOSO\alice'),
                (New-OrcaProc 'olk' 201 'CONTOSO\bob'),
                (New-OrcaProc 'msedgewebview2' 300 'CONTOSO\alice' $olkCmd),
                (New-OrcaProc 'msedgewebview2' 301 'CONTOSO\alice' $teamsCmd),
                (New-OrcaProc 'msedgewebview2' 302 'CONTOSO\bob' $olkCmd),
                (New-OrcaProc 'WebViewHost' 310 'CONTOSO\alice' $olkCmd)
            )
            Mock Get-Process {
                $script:Procs | Where-Object { $Name -contains $_.Name }
            }
            Mock Get-CimInstance {
                $id = [int](([string]$Filter) -replace '\D', '')
                $script:Procs | Where-Object { $_.Id -eq $id } | ForEach-Object { [pscustomobject]@{ ProcessId = $_.Id; CommandLine = $_.CommandLine } }
            }
            Mock Stop-Process { foreach ($x in (@($InputObject | ForEach-Object { $_.Id }) + @($Id))) { if ($null -ne $x) { $script:Stopped += $x } } }
            # Profile 0 (alice), Both, no OST reset, no extension purge.
            foreach ($a in @(@('D0', '0'), @('D3', '3'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
        }

        It 'stops only the selected user Outlook and Outlook WebView processes' {
            Invoke-OrcaRun
            ($script:Stopped | Sort-Object) -join ',' | Should -Be '100,200,300,310'
        }

        It 'never stops another user Outlook or an unrelated WebView2 app' {
            Invoke-OrcaRun
            $script:Stopped | Should -Not -Contain 101
            $script:Stopped | Should -Not -Contain 201
            $script:Stopped | Should -Not -Contain 301
            $script:Stopped | Should -Not -Contain 302
        }

        It 'Classic mode leaves New Outlook processes running' {
            $script:Keys.Clear()
            foreach ($a in @(@('D0', '0'), @('D2', '2'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            ($script:Stopped | Sort-Object) -join ',' | Should -Be '100'
        }
    }

    Context 'Profile identity, credentials and sign-out scope (19)' {
        BeforeEach {
            $script:AliceSid  = 'S-1-5-21-1000-2000-3000-1001'
            $script:BobSid    = 'S-1-5-21-1000-2000-3000-1002'
            # alice lives in a folder that does not match her account name.
            [IO.Directory]::Delete((Join-Path $script:UsersRoot 'alice'), $true)
            $script:AliceHome = Join-Path $script:UsersRoot 'alice.CONTOSO'
            $script:BobHome   = Join-Path $script:UsersRoot 'bob'
            $aliceOut = Join-Path $script:AliceHome 'AppData\Local\Microsoft\Outlook'
            New-Item -ItemType Directory -Path $aliceOut -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $aliceOut 'alice@contoso.com.ost') -Value 'synthetic OST for alice'
            foreach ($h in $script:AliceHome, $script:BobHome) {
                Set-Content -LiteralPath (Join-Path $h 'NTUSER.DAT') -Value 'synthetic hive'
                foreach ($rel in 'Microsoft\IdentityCache', 'Microsoft\OneAuth', 'Microsoft\TokenBroker\Cache', 'Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy') {
                    New-Item -ItemType Directory -Path (Join-Path $h "AppData\Local\$rel") -Force | Out-Null
                }
            }
            $script:AppxRemoved = @()
            $script:RegCalls    = @()
            $script:CredDeletes = @()
            Mock Get-OrcaProfile {
                @(
                    [pscustomobject]@{ Name = 'alice'; Sid = $script:AliceSid; Path = $script:AliceHome },
                    [pscustomobject]@{ Name = 'bob';   Sid = $script:BobSid;   Path = $script:BobHome }
                )
            }
            Mock Get-OrcaCurrentSid { $script:AliceSid }
            Mock Remove-AppxPackage { $script:AppxRemoved += $User }
            Mock reg { $script:RegCalls += ($args -join ' '); $global:LASTEXITCODE = 0 }
            Mock cmdkey {
                $global:LASTEXITCODE = 0
                if ($args[0] -eq '/list') { '    Target: LegacyGeneric:target=MicrosoftOffice16_Data:SSPI:alice@contoso.com'; '    Target: LegacyGeneric:target=Unrelated' }
                elseif ("$($args[0])" -like '/delete:*') { $script:CredDeletes += "$($args[0])" }
            }
        }

        It 'removes the New Outlook package by profile SID, not by a name guess' {
            foreach ($a in @(@('A', 'A'), @('D3', '3'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            ($script:AppxRemoved | Sort-Object) -join ',' | Should -Be "$script:AliceSid,$script:BobSid"
        }

        It 'loads NTUSER.DAT from the resolved profile folder, not from the account name' {
            foreach ($a in @(@('A', 'A'), @('D2', '2'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            @($script:RegCalls | Where-Object { $_ -like 'load *' -and $_ -like '*alice.CONTOSO\NTUSER.DAT' }).Count | Should -BeGreaterThan 0
            @($script:RegCalls | Where-Object { $_ -like 'unload *' }).Count | Should -BeGreaterThan 0
        }

        It 'renames the OST under the resolved profile folder when one profile is selected' {
            foreach ($a in @(@('D0', '0'), @('D2', '2'), @('Y', 'y'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            @(Get-ChildItem -LiteralPath (Join-Path $script:AliceHome 'AppData\Local\Microsoft\Outlook') -Filter '*.ost.bak-*').Count | Should -Be 1
        }

        It 'clears Credential Manager entries only for the identity that runs the script' {
            foreach ($a in @(@('A', 'A'), @('D2', '2'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            $script:CredDeletes.Count | Should -Be 1
            $script:CredDeletes[0] | Should -Match 'MicrosoftOffice16_Data'
        }

        It 'exits 3 and names the profile whose Credential Manager entries could not be cleared' {
            foreach ($a in @(@('A', 'A'), @('D2', '2'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match 'ORCA RESULT: exit code 3 .*Credential Manager.*bob'
        }

        It 'keeps the shared Windows account broker caches by default and clears the Office ones' {
            Invoke-OrcaRun -Silent
            foreach ($h in $script:AliceHome, $script:BobHome) {
                Test-Path -LiteralPath (Join-Path $h 'AppData\Local\Microsoft\IdentityCache') | Should -BeFalse
                Test-Path -LiteralPath (Join-Path $h 'AppData\Local\Microsoft\OneAuth') | Should -BeFalse
                Test-Path -LiteralPath (Join-Path $h 'AppData\Local\Microsoft\TokenBroker\Cache') | Should -BeTrue
                Test-Path -LiteralPath (Join-Path $h 'AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy') | Should -BeTrue
            }
        }

        It '-FullSignOut also clears the Windows account broker caches' {
            Invoke-OrcaRun -Silent -FullSignOut
            foreach ($h in $script:AliceHome, $script:BobHome) {
                Test-Path -LiteralPath (Join-Path $h 'AppData\Local\Microsoft\TokenBroker\Cache') | Should -BeFalse
                Test-Path -LiteralPath (Join-Path $h 'AppData\Local\Packages\Microsoft.AAD.BrokerPlugin_cw5n1h2txyewy') | Should -BeFalse
            }
        }
    }

    Context 'New Outlook reinstall verification (53, 54)' {
        BeforeEach {
            $script:AliceSid = 'S-1-5-21-1000-2000-3000-1001'
            $script:Started = @()
            $script:WingetCalls = @()
            $script:Order = @()
            $script:MakeMsix = $true
            $script:AddThrows = $false
            Mock Invoke-OrcaWinget {
                $script:WingetCalls += $Arguments
                if ($Arguments -like 'download*' -and $script:MakeMsix) {
                    $m = [regex]::Match($Arguments, '--download-directory "([^"]+)"')
                    New-Item -ItemType Directory -Path $m.Groups[1].Value -Force | Out-Null
                    Set-Content -LiteralPath (Join-Path $m.Groups[1].Value 'Outlook.msix') -Value 'synthetic package'
                }
                0
            }
            Mock Remove-AppxProvisionedPackage { $script:Order += 'remove'; $script:Provisioned = @() }
            Mock Add-AppxProvisionedPackage {
                $script:Order += 'add'
                if ($script:AddThrows) { throw 'add failed' }
                $script:Provisioned = @([pscustomobject]@{ DisplayName = 'Microsoft.OutlookForWindows'; PackageName = 'Microsoft.OutlookForWindows_2.0_neutral_~_8wekyb3d8bbwe' })
            }
        }

        It 'All Users: a failed download with an old package present exits 1 and changes nothing (53)' {
            $script:MakeMsix = $false
            Mock Get-OrcaProfile {
                @([pscustomobject]@{ Name = 'alice'; Account = 'CONTOSO\alice'; Sid = 'S-1-5-21-1000-2000-3000-1001'; Path = (Join-Path $script:UsersRoot 'alice') })
            }
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 1
            $script:Text | Should -Match 'ORCA RESULT: exit code 1 '
            Assert-MockCalled Add-AppxProvisionedPackage -Times 0 -Exactly
            Assert-MockCalled Remove-AppxProvisionedPackage -Times 0 -Exactly
        }

        It 'All Users: a good download removes the old package, then provisions the new one (53)' {
            Mock Get-OrcaProfile {
                @([pscustomobject]@{ Name = 'alice'; Account = 'CONTOSO\alice'; Sid = 'S-1-5-21-1000-2000-3000-1001'; Path = (Join-Path $script:UsersRoot 'alice') })
            }
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 0
            ($script:Order -join ',') | Should -Be 'remove,add'
        }

        It 'All Users: Add-AppxProvisionedPackage failing exits 1 (53)' {
            $script:AddThrows = $true
            Mock Get-OrcaProfile {
                @([pscustomobject]@{ Name = 'alice'; Account = 'CONTOSO\alice'; Sid = 'S-1-5-21-1000-2000-3000-1001'; Path = (Join-Path $script:UsersRoot 'alice') })
            }
            Invoke-OrcaRun -Silent
            $script:Code | Should -Be 1
            $script:Text | Should -Match 'ORCA RESULT: exit code 1 '
        }

        It 'an admin repairing a different single user skips the install and exits 3 (54)' {
            foreach ($a in @(@('D1', '1'), @('D1', '1'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            $script:Code | Should -Be 3
            $script:Text | Should -Match 'New Outlook was not reinstalled for bob'
            Assert-MockCalled Start-Process -Times 0 -Exactly -ParameterFilter { "$ArgumentList" -like '*winget install*' }
            Assert-MockCalled Invoke-OrcaWinget -Times 0 -Exactly
        }

        It 'a single user run for the invoking user still installs New Outlook (54)' {
            foreach ($a in @(@('D0', '0'), @('D1', '1'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Invoke-OrcaRun
            $script:Code | Should -Be 0
            Assert-MockCalled Invoke-OrcaWinget -Times 1 -Exactly -ParameterFilter { $Arguments -like 'install*' }
        }
    }

    Context 'Prompt fallbacks without a key reader (70)' {
        BeforeEach {
            $script:Started = @()
            $script:OldUser = $env:USERNAME
            $env:USERNAME = 'alice'
            $script:Outlook = Join-Path $TestDrive 'OUTLOOK.EXE'
            Set-Content -LiteralPath $script:Outlook -Value 'synthetic exe'
            Mock Get-ItemProperty { [pscustomobject]@{ '(default)' = $script:Outlook } } -ParameterFilter { "$Path" -like '*App Paths*' }
        }
        AfterEach { $env:USERNAME = $script:OldUser }

        It 'a typed launch choice in the line fallback is honoured and launches Classic Outlook' {
            # Profile alice, Both types, no OST reset, no extension purge. The launch menu then has no key reader.
            foreach ($a in @(@('D0', '0'), @('D3', '3'), @('N', 'n'), @('N', 'n'))) { $script:Keys.Enqueue((New-OrcaKey -Key $a[0] -Char $a[1])) }
            Mock Read-Host { '1' }
            Mock Read-OrcaLine { '1' }
            Invoke-OrcaRun
            @($script:Started | Where-Object { $_.File -eq $script:Outlook }).Count | Should -Be 1
        }

        It 'a run with no console and no answers exits 2 instead of failing' {
            Mock Read-OrcaKey { throw 'no console' }
            Mock Read-OrcaLine { $null }
            Invoke-OrcaRun
            $script:Code | Should -Be 2
            $script:Text | Should -Match 'ORCA RESULT: exit code 2 '
            Assert-MockCalled Stop-Process -Times 0 -Exactly
        }

        It 'Read-OrcaLine returns null when Read-Host has no console' {
            Mock Read-Host { throw 'no console' }
            & $script:RealReadOrcaLine | Should -BeNullOrEmpty
        }
    }

    Context 'Get-OrcaProfile' {        BeforeEach {
            New-Item -Path 'TestRegistry:\ProfileList' -Force | Out-Null
            $script:LiveHome = Join-Path $TestDrive 'Users\alice.CONTOSO'
            New-Item -ItemType Directory -Path $script:LiveHome -Force | Out-Null
            function Add-OrcaProfileKey { param([string]$Sid, [string]$Path)
                New-Item -Path "TestRegistry:\ProfileList\$Sid" -Force | Out-Null
                New-ItemProperty -Path "TestRegistry:\ProfileList\$Sid" -Name ProfileImagePath -Value $Path -PropertyType ExpandString | Out-Null
            }
            Add-OrcaProfileKey 'S-1-5-21-1000-2000-3000-1001' $script:LiveHome
            Add-OrcaProfileKey 'S-1-5-18' 'C:\Windows\system32\config\systemprofile'
            Add-OrcaProfileKey 'S-1-5-21-1000-2000-3000-1003' (Join-Path $TestDrive 'Users\gone')
            Mock Resolve-OrcaAccountName { if ($Sid -like '*-1001') { 'alice' } else { $null } }
        }

        It 'binds each real user SID to its profile folder and account name' {
            $p = @(& $script:RealGetOrcaProfile -ListPath 'TestRegistry:\ProfileList')
            $p.Count | Should -Be 1
            $p[0].Name | Should -Be 'alice'
            $p[0].Sid  | Should -Be 'S-1-5-21-1000-2000-3000-1001'
            $p[0].Path | Should -Be $script:LiveHome
        }

        It 'skips service profiles and profiles whose folder is gone' {
            $p = @(& $script:RealGetOrcaProfile -ListPath 'TestRegistry:\ProfileList')
            $p.Sid | Should -Not -Contain 'S-1-5-18'
            $p.Sid | Should -Not -Contain 'S-1-5-21-1000-2000-3000-1003'
        }
    }
}
