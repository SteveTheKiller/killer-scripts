#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# W.U.R.S.A. tests. Every system boundary is mocked, so nothing here touches the real machine.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath 'WURSA')

    # Throwing guards for every boundary the safety net does not know about. A test that
    # reaches one of these without its own mock fails loudly instead of touching the host.
    foreach ($fn in 'Invoke-WursaNative', 'Invoke-WursaChoco', 'Save-WursaFile', 'Expand-WursaZip',
                    'Invoke-WursaScriptFile', 'Invoke-WursaRestorePoint', 'Register-WursaProductService',
                    'New-WursaUpdateSession', 'New-WursaUpdateColl', 'Get-WursaPowerLineStatus') {
        $name = $fn
        Mock $name { throw "UNMOCKED WURSA BOUNDARY: $name" }.GetNewClosure()
    }
    Mock New-Object { throw 'UNMOCKED New-Object (COM or WebClient)' } -ParameterFilter { $ComObject -or ($TypeName -match 'WebClient') }
    Mock Start-Job { throw 'UNMOCKED Start-Job' }
    Mock Start-Sleep { }
    Mock Set-ExecutionPolicy { }
    Mock Clear-Host { }

    function Get-ExitCode {
        param([scriptblock]$Block)
        try { & $Block | Out-Null } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(-?\d+)') { return [int]$Matches[1] }
            throw
        }
        return $null
    }

    $script:SidSystem = 'S-1-5-18'
    $script:SidAdmins = 'S-1-5-32-544'
    $script:SidUsers  = 'S-1-5-32-545'
    $script:SidUser   = 'S-1-5-21-1000-2000-3000-1001'

    function New-FakeAcl {
        param([string]$OwnerSid, [string[]]$AllowSids = @(), [bool]$Protected = $true)
        $acl = New-Object System.Security.AccessControl.DirectorySecurity
        $acl.SetOwner([System.Security.Principal.SecurityIdentifier]$OwnerSid)
        $acl.SetAccessRuleProtection($Protected, $false)
        foreach ($s in $AllowSids) {
            $rule = New-Object System.Security.AccessControl.FileSystemAccessRule(([System.Security.Principal.SecurityIdentifier]$s), 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
            $acl.AddAccessRule($rule)
        }
        $acl
    }

    # Mocks the whole machine for a Main-level run. Build 26100 is a 24H2 box that needs the upgrade.
    function Set-MainMocks {
        param([int]$Build = 26100, [string]$Version = '24H2')
        $script:MainBuild = $Build
        $script:MainVersion = $Version
        Mock Test-WursaElevated { $true }
        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_Battery' -or $Class -eq 'Win32_Battery') { return $null }
            if ($ClassName -eq 'Win32_Processor' -or $Class -eq 'Win32_Processor') { return [pscustomobject]@{ Architecture = 9 } }
            [pscustomobject]@{ Caption = 'Windows 11 Pro'; BuildNumber = "$script:MainBuild"; OSArchitecture = '64-bit'; LastBootUpTime = (Get-Date).AddDays(-1) }
        }
        Mock Get-ItemProperty { [pscustomobject]@{ DisplayVersion = $script:MainVersion; CurrentBuild = "$script:MainBuild"; UBR = 9300; InstallLanguage = '0409'; CompositionEditionID = 'Professional'; EditionID = 'Professional'; ProductName = 'Windows 11 Pro' } }
        Mock Get-ItemProperty { $null } -ParameterFilter { $Path -like '*Auto Update*RebootRequired' }
        Mock Test-Path { $false }
        Mock Invoke-WursaRestorePoint { @{ Result = 'OK'; Message = '' } }
        Mock Register-WursaProductService { }
        Mock Invoke-WursaWindowsUpdatePass { [pscustomobject]@{ Status = 'UpToDate'; Found = 0; Succeeded = 0; Failed = 0; RebootRequired = $false; Error = $null; Searcher = $null; FeatureOffered = $false } }
        Mock Invoke-WursaThirdPartyUpdate { [pscustomobject]@{ Status = 'Skipped'; Succeeded = 0; Failed = 0; RebootRequired = $false } }
        Mock Restart-Service { }
        $script:said = New-Object System.Collections.ArrayList
        Mock Write-Host { [void]$script:said.Add([string]$Object) }
        Mock Write-HLine { }
    }

    # The runner is a script the SYSTEM task executes later. These helpers pull one of its functions
    # out so a test can dot-source it, and dot-sourcing must happen in the test's own scope.
    function Get-RunnerFunctionText {
        param([string]$Name)
        $text = Get-WursaRunnerBody
        $errs = $null; $toks = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$toks, [ref]$errs)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq $Name }, $true)
        if (-not $fn) { throw "runner has no function $Name" }
        $fn.Extent.Text
    }
}

Describe 'Entry point convention' {
    It 'defines Invoke-WursaMain and does not run it when dot-sourced in test mode' {
        Get-Command Invoke-WursaMain -CommandType Function | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -CommandType Function | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws the exit sentinel in test mode' {
        { Exit-KillerScript 3 } | Should -Throw '*KILLERSCRIPT_EXIT:3*'
    }

    It 'ends with the guarded call and has no bare exit statement outside Exit-KillerScript' {
        $tokens = $null; $errors = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Get-KillerScriptPath 'WURSA'), [ref]$tokens, [ref]$errors)
        $exits = $ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExitStatementAst] }, $true)
        foreach ($e in $exits) {
            $fn = $e.Parent
            while ($fn -and $fn -isnot [System.Management.Automation.Language.FunctionDefinitionAst]) { $fn = $fn.Parent }
            $fn.Name | Should -Be 'Exit-KillerScript'
        }
    }

    It 'exits 2 when not elevated' {
        Mock Test-WursaElevated { $false }
        Mock Write-Warning { }
        Get-ExitCode { Invoke-WursaMain } | Should -Be 2
    }

    It 'exits 0 on a healthy -NoUpgrade -No3rdParty run and prints a status line that agrees' {
        $NoUpgrade = [switch]$true
        $No3rdParty = [switch]$true
        Mock Test-WursaElevated { $true }
        Mock Get-CimInstance { [pscustomobject]@{ Caption = 'Windows 11 Pro'; BuildNumber = '26300'; Architecture = 9; OSArchitecture = '64-bit' } }
        Mock Get-ItemProperty { [pscustomobject]@{ DisplayVersion = '26H2'; CurrentBuild = '26300'; UBR = 1; InstallLanguage = '0409'; CompositionEditionID = 'Professional'; EditionID = 'Professional'; ProductName = 'Windows 11 Pro' } }
        Mock Get-ItemProperty { $null } -ParameterFilter { $Path -like '*Auto Update*RebootRequired' }
        Mock Test-Path { $false }
        Mock Invoke-WursaRestorePoint { @{ Result = 'OK'; Message = '' } }
        Mock Register-WursaProductService { }
        Mock Invoke-WursaWindowsUpdatePass { [pscustomobject]@{ Status = 'UpToDate'; Found = 0; Succeeded = 0; Failed = 0; RebootRequired = $false; Error = $null; Searcher = $null; FeatureOffered = $false } }
        Mock Invoke-WursaThirdPartyUpdate { [pscustomobject]@{ Status = 'Skipped'; Succeeded = 0; Failed = 0; RebootRequired = $false } }
        $script:said = New-Object System.Collections.ArrayList
        Mock Write-Host { [void]$script:said.Add([string]$Object) }
        Get-ExitCode { Invoke-WursaMain } | Should -Be 0
        ($script:said -join "`n") | Should -Match 'STATUS: UPDATE PASS COMPLETE'
    }

    It 'turns an unexpected error into exit 1 and keeps the exit sentinel intact' {
        Mock Test-WursaElevated { throw 'boom' }
        Mock Write-Host { }
        Get-ExitCode { Invoke-WursaMain } | Should -Be 1
    }
}

Describe 'Final status mapping' {
    It 'maps <State> / <Ipu> to exit <Code> with a matching line' -ForEach @(
        @{ NoUp = $false; Build = 26300; State = 'CURRENT';       Ipu = '';                       Reboot = $false; Code = 0;    Like = '*SYSTEM CURRENT*' }
        @{ NoUp = $true;  Build = 26100; State = 'SKIPPED';       Ipu = '';                       Reboot = $true;  Code = 3010; Like = '*REBOOT REQUIRED*' }
        @{ NoUp = $true;  Build = 26100; State = 'SKIPPED';       Ipu = '';                       Reboot = $false; Code = 0;    Like = '*UPDATE PASS COMPLETE*' }
        @{ NoUp = $false; Build = 26100; State = 'REBOOT_REQUIRED'; Ipu = '';                     Reboot = $false; Code = 3010; Like = '*REBOOT REQUIRED TO FINISH*' }
        @{ NoUp = $false; Build = 26100; State = 'BLOCKED_BATTERY'; Ipu = '';                     Reboot = $false; Code = 1;    Like = '*NOT COMPLETE*' }
        @{ NoUp = $false; Build = 26100; State = 'NOT_STARTED';   Ipu = '';                       Reboot = $false; Code = 1;    Like = '*NOT COMPLETE*' }
        @{ NoUp = $false; Build = 26100; State = 'DISPATCHED';    Ipu = 'REBOOT_REQUIRED 3010 x'; Reboot = $false; Code = 3010; Like = '*REBOOT REQUIRED TO FINISH*' }
        @{ NoUp = $false; Build = 26100; State = 'DISPATCHED';    Ipu = 'BLOCK_LTSC x';           Reboot = $false; Code = 1;    Like = '*NOT COMPLETE*' }
        @{ NoUp = $false; Build = 26100; State = 'DISPATCHED';    Ipu = 'RUNNING x';              Reboot = $false; Code = 0;    Like = '*IN PROGRESS*' }
        @{ NoUp = $false; Build = 26100; State = 'UNKNOWN';       Ipu = 'WU_HANDOFF x';           Reboot = $false; Code = 0;    Like = '*IN PROGRESS*' }
        @{ NoUp = $false; Build = 26100; State = 'UNKNOWN';       Ipu = 'FAILED URL x';           Reboot = $false; Code = 1;    Like = '*NOT COMPLETE*' }
        @{ NoUp = $false; Build = 26100; State = 'UNKNOWN';       Ipu = '';                       Reboot = $false; Code = 1;    Like = '*NOT COMPLETE*' }
    ) {
        $r = Resolve-WursaFinalStatus -NoUpgrade $NoUp -FinalBuild $Build -FinalVersion '25H2' -FeatureUpgradeState $State -IpuState $Ipu -RebootPending $Reboot
        $r.Code | Should -Be $Code
        ($r.Lines | ForEach-Object { $_.Text }) -join "`n" | Should -BeLike $Like
    }
}

Describe 'UI helpers without a console' {
    It 'Write-StepUpdate and Write-SubResult do not throw when there is no cursor' {
        Mock Get-WursaCursorTop { -1 }
        Mock Write-Host { }
        { Write-StepUpdate '[1/5] Test...'; Write-StepUpdate -Success -CustomInfo '(ok)'; Write-SubResult '[OK]' 'Green' } | Should -Not -Throw
    }
}

Describe 'Staging directory ownership (finding 48)' {
    BeforeEach {
        $script:Acls = @{}
        $script:SetAclCalls = New-Object System.Collections.ArrayList
        Mock Get-Acl {
            $p = if ($LiteralPath) { [string]$LiteralPath } else { [string]$Path }
            if ($script:Acls.ContainsKey($p)) { return $script:Acls[$p] }
            # A directory created by an elevated admin: owned by Administrators, inherits Users create rights.
            New-FakeAcl -OwnerSid $SidAdmins -AllowSids @($SidSystem, $SidAdmins, $SidUsers) -Protected $false
        }
        Mock Set-Acl {
            $p = if ($LiteralPath) { [string]$LiteralPath } else { [string]$Path }
            $script:Acls[$p] = $AclObject
            [void]$script:SetAclCalls.Add(@{ Path = $p; Acl = $AclObject })
        }
        # A move carries the ACLs that were registered for the old path to the new one.
        Mock Move-Item {
            $src = [string]$LiteralPath; $dst = [string]$Destination
            foreach ($k in @($script:Acls.Keys)) {
                if ($k -eq $src -or $k.StartsWith($src + '')) {
                    $script:Acls[$dst + $k.Substring($src.Length)] = $script:Acls[$k]
                    $script:Acls.Remove($k)
                }
            }
            [System.IO.Directory]::Move($src, $dst)
        }
        $script:Dir = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
    }

    It 'creates a missing directory with a protected admin-only ACL before anything is written' {
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeTrue
        Test-Path -LiteralPath $script:Dir | Should -BeTrue
        $script:SetAclCalls.Count | Should -Be 1
        $acl = $script:SetAclCalls[0].Acl
        $acl.AreAccessRulesProtected | Should -BeTrue
        $sids = @($acl.GetAccessRules($true, $true, [System.Security.Principal.SecurityIdentifier]) | ForEach-Object { $_.IdentityReference.Value } | Sort-Object -Unique)
        $sids | Should -Be @($SidAdmins, $SidSystem | Sort-Object)
    }

    It 'reclaims a directory a standard user pre-created by moving it aside and creating a fresh one' {
        New-Item -ItemType Directory -Path $script:Dir | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'Invoke-IPU.ps1') -Value 'planted'
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidUser -AllowSids @($SidUser, $SidSystem) -Protected $false
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeTrue
        $r.MovedAside | Should -Not -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path $r.MovedAside 'Invoke-IPU.ps1') | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Dir 'Invoke-IPU.ps1') | Should -BeFalse
    }

    It 'treats an admin-owned directory that still inherits Users permissions as untrusted' {
        New-Item -ItemType Directory -Path $script:Dir | Out-Null
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidAdmins -AllowSids @($SidSystem, $SidAdmins, $SidUsers) -Protected $false
        (Get-WursaDirectoryProblem -Path $script:Dir) | Should -Not -BeNullOrEmpty
    }

    It 'reuses a directory that is already owned by an admin with a protected admin-only ACL' {
        New-Item -ItemType Directory -Path $script:Dir | Out-Null
        Set-Content -LiteralPath (Join-Path $script:Dir 'keep.txt') -Value 'x'
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidSystem -AllowSids @($SidSystem, $SidAdmins) -Protected $true
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeTrue
        $r.MovedAside | Should -BeNullOrEmpty
        Test-Path -LiteralPath (Join-Path $script:Dir 'keep.txt') | Should -BeTrue
        $script:SetAclCalls.Count | Should -Be 0
    }

    It 'refuses when an untrusted directory cannot be moved aside' {
        New-Item -ItemType Directory -Path $script:Dir | Out-Null
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidUser -AllowSids @($SidUser) -Protected $false
        Mock Move-Item { throw 'in use' }
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeFalse
        $r.Reason | Should -Match 'move'
    }

    It 'refuses when the freshly created directory turns out to be owned by someone else (race)' {
        Mock Set-Acl { }
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidUser -AllowSids @($SidUser) -Protected $false
        # The directory does not exist yet, so the creation succeeds, but the owner read back is the attacker.
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeFalse
    }

    It 'keeps a trusted restore record when it moves an old directory aside' {
        New-Item -ItemType Directory -Path $script:Dir | Out-Null
        $rec = Join-Path $script:Dir 'au_restore.txt'
        Set-Content -LiteralPath $rec -Value 'NoAutoUpdate=1'
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidAdmins -AllowSids @($SidSystem, $SidAdmins, $SidUsers) -Protected $false
        $script:Acls[$rec] = New-FakeAcl -OwnerSid $SidAdmins -AllowSids @($SidSystem, $SidAdmins) -Protected $false
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeTrue
        Get-Content -LiteralPath (Join-Path $script:Dir 'au_restore.txt') | Should -Be 'NoAutoUpdate=1'
    }

    It 'does not carry over a restore record owned by an untrusted user' {
        New-Item -ItemType Directory -Path $script:Dir | Out-Null
        $rec = Join-Path $script:Dir 'au_restore.txt'
        Set-Content -LiteralPath $rec -Value 'NoAutoUpdate=1'
        $script:Acls[$script:Dir] = New-FakeAcl -OwnerSid $SidAdmins -AllowSids @($SidSystem, $SidAdmins, $SidUsers) -Protected $false
        $script:Acls[$rec] = New-FakeAcl -OwnerSid $SidUser -AllowSids @($SidUser) -Protected $false
        $r = Initialize-WursaWorkDir -Path $script:Dir
        $r.Ready | Should -BeTrue
        Test-Path -LiteralPath (Join-Path $script:Dir 'au_restore.txt') | Should -BeFalse
    }
}

Describe 'Staging directory ownership in a full run (finding 48)' {
    It 'does not open Windows Update policy or dispatch when the staging directory cannot be trusted' {
        $InplaceUpgrade = [switch]$true
        $No3rdParty = [switch]$true
        $NoUpgrade = [switch]$false
        Set-MainMocks
        $script:writes = New-Object System.Collections.ArrayList
        Mock Initialize-WursaWorkDir { [pscustomobject]@{ Ready = $false; Reason = 'owner S-1-5-21-1-2-3-1001'; MovedAside = $null } }
        Mock Set-ItemProperty { [void]$script:writes.Add($Name) } -ParameterFilter { $true }
        Mock New-Item { throw 'must not create anything' }
        Mock Invoke-WursaNative { throw 'must not register a task' }
        Get-ExitCode { Invoke-WursaMain } | Should -Be 1
        $script:writes.Count | Should -Be 0
        ($script:said -join "`n") | Should -Match 'BLOCKED_WORKDIR'
    }
}

Describe 'Installation Assistant verification (finding 36)' {
    BeforeEach {
        $script:Pub = '^CN=Microsoft Corporation(,|$)'
        $script:Exe = Join-Path $TestDrive ([guid]::NewGuid().ToString('N') + '.exe')
    }

    It 'accepts a valid Microsoft signature' {
        Set-Content -LiteralPath $script:Exe -Value 'x'
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US' } } }
        Test-WursaSignedBinary -Path $script:Exe -PublisherPattern $script:Pub | Should -BeTrue
    }

    It 'rejects <Why>' -ForEach @(
        @{ Why = 'an unsigned file';           Status = 'NotSigned';    Subject = '' }
        @{ Why = 'a tampered signature';       Status = 'HashMismatch'; Subject = 'CN=Microsoft Corporation' }
        @{ Why = 'a valid but foreign signer'; Status = 'Valid';        Subject = 'CN=Evil Corp, O=Evil Corp' }
        @{ Why = 'a lookalike publisher';      Status = 'Valid';        Subject = 'CN=Microsoft Corporation Support Ltd' }
    ) {
        Set-Content -LiteralPath $script:Exe -Value 'x'
        $sub = $Subject; $st = $Status
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = $st; SignerCertificate = [pscustomobject]@{ Subject = $sub } } }.GetNewClosure()
        Test-WursaSignedBinary -Path $script:Exe -PublisherPattern $script:Pub | Should -BeFalse
    }

    It 'rejects a missing file and a pinned digest that does not match' {
        Test-WursaSignedBinary -Path (Join-Path $TestDrive 'nope.exe') -PublisherPattern $script:Pub | Should -BeFalse
        Set-Content -LiteralPath $script:Exe -Value 'x'
        Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Microsoft Corporation' } } }
        Test-WursaSignedBinary -Path $script:Exe -PublisherPattern $script:Pub -Sha256 ('0' * 64) | Should -BeFalse
    }

    Context 'the runner function' {
        BeforeEach {
            $script:Work = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $script:Work | Out-Null
            $IPU_WorkDir = $script:Work
            $IPU_StatusFile = Join-Path $script:Work 'ipu_status.txt'
            $IPU_IAExe = Join-Path $script:Work 'Windows11InstallationAssistant.exe'
            $IPU_IAArgs = '/QuietInstall'
            $IPU_IAUrl = 'https://example.invalid/ia.exe'
            $IPU_IASha256 = ''
            $IPU_IAPublisher = '^CN=Microsoft Corporation(,|$)'
            . ([scriptblock]::Create((Get-RunnerFunctionText 'Set-Status')))
            . ([scriptblock]::Create((Get-RunnerFunctionText 'Invoke-IPUInstallationAssistant')))
            function Register-IPUPostRebootVerifier { $true }
            function Invoke-IPUWindowsUpdate { param($Reason) 'WU' }
            $script:Launched = 0
            Mock Get-BitLockerVolume { $null }
            Mock Get-Process { $null }
            Mock Start-Process { $script:Launched++; [pscustomobject]@{ ExitCode = 0 } }
            Mock Invoke-WebRequest { Set-Content -LiteralPath $OutFile -Value 'downloaded' }
        }

        It 'does not launch an existing copy whose signature is not valid' {
            Set-Content -LiteralPath $IPU_IAExe -Value 'planted'
            Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null } }
            Invoke-IPUInstallationAssistant -Reason 'test' | Out-Null
            $script:Launched | Should -Be 0
            (Get-Content -LiteralPath $IPU_StatusFile) | Should -Match '^FAILED IA_SIGNATURE'
        }

        It 'does not launch a fresh download whose signer is not Microsoft' {
            Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Somebody Else' } } }
            Invoke-IPUInstallationAssistant -Reason 'test' | Out-Null
            $script:Launched | Should -Be 0
            (Get-Content -LiteralPath $IPU_StatusFile) | Should -Match '^FAILED IA_SIGNATURE'
        }

        It 'launches a copy with a valid Microsoft signature' {
            Set-Content -LiteralPath $IPU_IAExe -Value 'real'
            Mock Get-AuthenticodeSignature { [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation' } } }
            Invoke-IPUInstallationAssistant -Reason 'test' | Out-Null
            $script:Launched | Should -Be 1
        }
    }

    It 'embeds the verifier into the generated runner so the shipped code is the tested code' {
        $text = New-WursaRunnerScript -Config '$IPU_WorkDir = 1'
        $errs = $null; $toks = $null
        $ast = [System.Management.Automation.Language.Parser]::ParseInput($text, [ref]$toks, [ref]$errs)
        $errs.Count | Should -Be 0
        $names = @($ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Test-WursaSignedBinary'
        $names | Should -Contain 'Invoke-IPUInstallationAssistant'
        @($text.ToCharArray() | Where-Object { [int]$_ -gt 127 }).Count | Should -Be 0
    }
}

Describe 'Third-party catalog and Chocolatey (findings 39, 40, 30b)' {
    BeforeEach {
        $script:Lines = New-Object System.Collections.ArrayList
        Mock Write-Host { [void]$script:Lines.Add([string]$Object) }
        Mock Write-StepUpdate { }
        Mock Write-SubResult { param($Tag, $Color) [void]$script:Lines.Add([string]$Tag) }
        Mock Get-Command { [pscustomobject]@{ Name = 'choco' } } -ParameterFilter { $Name -eq 'choco' }
        Mock Get-ItemProperty { $null }
        Mock Get-Process { $null }
        Mock Get-WursaThirdPartyCatalog { @(@{ Name = 'Zoom'; Id = 'zoom'; Args = @("--install-arguments='/quiet /norestart'"); Process = 'Zoom' }) }
        $script:Calls = New-Object System.Collections.ArrayList
        $script:ChocoVersion = '2.4.1'
        $script:ListExit = 0
        $script:UpgradeExit = 0
        $script:SelfExit = 0
        $script:UpgradeText = 'Chocolatey upgraded 1/1 packages.'
        Mock Invoke-WursaChoco {
            [void]$script:Calls.Add([string[]]$Arguments)
            switch ($Arguments[0]) {
                '--version' { return [pscustomobject]@{ ExitCode = 0; Output = @($script:ChocoVersion) } }
                'list'      { return [pscustomobject]@{ ExitCode = $script:ListExit; Output = @('zoom|5.0') } }
                'upgrade'   {
                    if ($Arguments[1] -eq 'chocolatey') { return [pscustomobject]@{ ExitCode = $script:SelfExit; Output = @('Chocolatey upgraded 0/1 packages.') } }
                    return [pscustomobject]@{ ExitCode = $script:UpgradeExit; Output = @($script:UpgradeText) }
                }
            }
        }
        function Get-UpgradeCall { $script:Calls | Where-Object { $_[0] -eq 'upgrade' -and $_[1] -ne 'chocolatey' } | Select-Object -First 1 }
    }

    It 'defines the real catalog with whitespace-free Ids and array Args' {
        $cat = & (Get-Command Get-WursaThirdPartyCatalog -CommandType Function).ScriptBlock
        $cat.Count | Should -BeGreaterThan 10
        foreach ($e in $cat) {
            $e.Id | Should -Not -BeNullOrEmpty
            $e.Id | Should -Not -Match '\s'
            ($e.ContainsKey('ChocoID')) | Should -BeFalse
            @($e.Args) | ForEach-Object { $_ | Should -BeOfType [string] }
        }
        $chrome = $cat | Where-Object { $_.Id -eq 'googlechrome' }
        @($chrome.Args) | Should -Be @("--install-arguments='--system-level'", '--ignore-checksums')
    }

    It 'passes the package id and its arguments to choco as separate argv elements' {
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        $call = Get-UpgradeCall
        $call[0] | Should -Be 'upgrade'
        $call[1] | Should -Be 'zoom'
        $call | Should -Contain "--install-arguments='/quiet /norestart'"
        $call | Should -Contain '-y'
        ($call | Where-Object { $_ -match '^zoom\s' }) | Should -BeNullOrEmpty
    }

    It 'lists local packages with --local-only on Chocolatey 1.x' {
        $script:ChocoVersion = '1.4.0'
        Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x' | Out-Null
        $list = $script:Calls | Where-Object { $_[0] -eq 'list' } | Select-Object -First 1
        $list | Should -Contain '--local-only'
    }

    It 'lists local packages without --local-only on Chocolatey 2.x' {
        $script:ChocoVersion = '2.4.1'
        Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x' | Out-Null
        $list = $script:Calls | Where-Object { $_[0] -eq 'list' } | Select-Object -First 1
        $list | Should -Not -Contain '--local-only'
    }

    It 'treats a failed inventory as unverified: no upgrades, a warning, and a counted failure' {
        $script:ListExit = 1
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        Get-UpgradeCall | Should -BeNullOrEmpty
        $r.Failed | Should -BeGreaterThan 0
        $r.Status | Should -Be 'InventoryUnverified'
        ($script:Lines -join "`n") | Should -Match 'inventory'
    }

    It 'treats an unreadable Chocolatey version as an unverified inventory' {
        $script:ChocoVersion = 'garbage'
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        Get-UpgradeCall | Should -BeNullOrEmpty
        $r.Status | Should -Be 'InventoryUnverified'
        $r.Failed | Should -BeGreaterThan 0
    }

    It 'reports a failed app upgrade as FAILED and counts it, not as already updated' {
        $script:UpgradeExit = 1
        $script:UpgradeText = 'Chocolatey upgraded 0/1 packages. 1 packages failed.'
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        $r.Failed | Should -Be 1
        ($script:Lines -join "`n") | Should -Match '\[FAILED\]'
        # The only ALREADY UPDATED line belongs to the Chocolatey self-upgrade, not to the failed app.
        @($script:Lines | Where-Object { $_ -eq '[ALREADY UPDATED]' }).Count | Should -Be 1
        ($script:Lines -join "`n") | Should -Not -Match 'ADOPTED'
    }

    It 'treats exit 3010 and 1641 as success with a reboot request' -ForEach @(@{ Code = 3010 }, @{ Code = 1641 }) {
        $script:UpgradeExit = $Code
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        $r.Failed | Should -Be 0
        $r.Succeeded | Should -Be 1
        $r.RebootRequired | Should -BeTrue
    }

    It 'does not adopt a non-choco app when the upgrade command failed' {
        $script:UpgradeExit = 1
        Mock Get-ItemProperty { [pscustomobject]@{ DisplayName = 'Zoom Workplace' } }
        Mock Invoke-WursaChoco {
            [void]$script:Calls.Add([string[]]$Arguments)
            switch ($Arguments[0]) {
                '--version' { [pscustomobject]@{ ExitCode = 0; Output = @('2.4.1') } }
                'list'      { [pscustomobject]@{ ExitCode = 0; Output = @() } }
                'upgrade'   { if ($Arguments[1] -eq 'chocolatey') { [pscustomobject]@{ ExitCode = 0; Output = @() } } else { [pscustomobject]@{ ExitCode = 1; Output = @('failed') } } }
            }
        }
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        $r.Failed | Should -Be 1
        ($script:Lines -join "`n") | Should -Not -Match 'ADOPTED'
    }

    It 'reports a failed Chocolatey self-upgrade instead of calling it already updated' {
        $script:SelfExit = 1
        $r = Invoke-WursaThirdPartyUpdate -Skip $false -AutoConfirm $true -AutoFlag 'x'
        $r.Failed | Should -BeGreaterThan 0
        ($script:Lines -join "`n") | Should -Match '\[FAILED\]'
    }
}

Describe 'Chocolatey bootstrap (finding 66)' {
    BeforeEach {
        $script:Base = Join-Path $TestDrive ([guid]::NewGuid().ToString('N'))
        New-Item -ItemType Directory -Path $script:Base | Out-Null
        Mock Write-Host { }
        $script:ScriptRan = $false
        $script:Saved = $null
        Mock New-WursaSecureTempDir { $script:Base }
        Mock Save-WursaFile { $script:Saved = $OutFile }
        Mock Expand-WursaZip { }
        Mock Get-ChildItem { [pscustomobject]@{ FullName = (Join-Path $script:Base 'tools\chocolateyInstall\choco.exe') } } -ParameterFilter { $Filter -eq 'choco.exe' }
        Mock Get-ChildItem { [pscustomobject]@{ FullName = (Join-Path $script:Base 'tools\chocolateyInstall.ps1') } } -ParameterFilter { $Filter -eq 'chocolateyInstall.ps1' }
        Mock Invoke-WursaScriptFile { $script:ScriptRan = $true }
        Mock Get-Command { if ($script:ScriptRan) { [pscustomobject]@{ Name = 'choco' } } else { $null } } -ParameterFilter { $Name -eq 'choco' }
    }

    It 'downloads into a private directory, never into the shared TEMP folder' {
        Mock Test-WursaSignedBinary { $true }
        Install-WursaChocolatey | Should -BeTrue
        $script:Saved | Should -BeLike "$script:Base*"
    }

    It 'refuses to run the installer when choco.exe is not signed by Chocolatey Software' {
        Mock Test-WursaSignedBinary { $false }
        Install-WursaChocolatey | Should -BeFalse
        $script:ScriptRan | Should -BeFalse
    }

    It 'checks the signer against the Chocolatey publisher and runs the installer when valid' {
        Mock Test-WursaSignedBinary { $true }
        Install-WursaChocolatey | Should -BeTrue
        $script:ScriptRan | Should -BeTrue
        Should -Invoke Test-WursaSignedBinary -ParameterFilter { $PublisherPattern -match 'Chocolatey Software' }
    }

    It 'does not download anything when a private directory cannot be created' {
        Mock New-WursaSecureTempDir { $null }
        Install-WursaChocolatey | Should -BeFalse
        Should -Invoke Save-WursaFile -Times 0
    }

    It 'removes its private directory afterwards' {
        Mock Test-WursaSignedBinary { $true }
        Install-WursaChocolatey | Out-Null
        Test-Path -LiteralPath $script:Base | Should -BeFalse
    }
}

Describe 'Secure temp directory (finding 66)' {
    It 'creates a unique protected directory for each call and returns its path' {
        Mock Initialize-WursaWorkDir { [pscustomobject]@{ Ready = $true; Reason = ''; MovedAside = $null } }
        $a = New-WursaSecureTempDir -Base 'C:\ProgramData\WURSA'
        $b = New-WursaSecureTempDir -Base 'C:\ProgramData\WURSA'
        $a | Should -Not -Be $b
        $a | Should -BeLike 'C:\ProgramData\WURSA\*'
        Should -Invoke Initialize-WursaWorkDir -Times 2
    }

    It 'returns nothing when the directory cannot be trusted' {
        Mock Initialize-WursaWorkDir { [pscustomobject]@{ Ready = $false; Reason = 'nope'; MovedAside = $null } }
        New-WursaSecureTempDir -Base 'C:\ProgramData\WURSA' | Should -BeNullOrEmpty
    }
}

Describe 'Windows Update pass outcomes (finding 30a)' {
    BeforeEach {
        $script:Lines = New-Object System.Collections.ArrayList
        Mock Write-Host { [void]$script:Lines.Add([string]$Object) }
        Mock Write-StepUpdate { if ($CustomInfo) { [void]$script:Lines.Add([string]$CustomInfo) } }
        Mock Write-SubResult { [void]$script:Lines.Add([string]$Tag) }
        Mock Write-WursaUpdateLines { [void]$script:Lines.Add([string]$Tag) }
        Mock Write-Progress { }
        Mock Get-WursaCursorTop { -1 }
        Mock Get-WursaRebootPending { $false }
        Mock Test-Path { $false }
        Mock New-WursaUpdateColl { New-Object System.Collections.ArrayList }
        $script:SearchBlock = { param($q) [pscustomobject]@{ Updates = @() } }
        $script:Results = @{}
        function script:New-FakeSession {
            $searcher = [pscustomobject]@{}
            $searcher | Add-Member -MemberType ScriptMethod -Name Search -Value { param($q) & $script:SearchBlock $q }
            $session = [pscustomobject]@{}
            $session | Add-Member -MemberType ScriptMethod -Name CreateUpdateSearcher -Value { $searcher }.GetNewClosure()
            $session
        }
        Mock New-WursaUpdateSession { New-FakeSession }
        Mock Install-WursaUpdate { $script:Results[$Update.Title] }
    }

    It 'reports a scan error as a failed scan, not as up to date' {
        $script:SearchBlock = { param($q) throw 'WU search exploded' }
        $r = Invoke-WursaWindowsUpdatePass
        $r.Status | Should -Be 'ScanFailed'
        $r.Failed | Should -BeGreaterThan 0
        ($script:Lines -join "`n") | Should -Match '\[FAILED\]'
        ($script:Lines -join "`n") | Should -Not -Match 'System Up to Date'
    }

    It 'aggregates failures, successes, warnings and reboot flags across every update' {
        $mk = { param($t) [pscustomobject]@{ Title = $t } }
        $script:SearchBlock = { param($q) if ($q -match 'Driver') { [pscustomobject]@{ Updates = @() } } else { [pscustomobject]@{ Updates = @((& $mk 'KB1'), (& $mk 'KB2'), (& $mk 'KB3'), (& $mk 'KB4')) } } }
        $script:Results = @{
            'KB1' = [pscustomobject]@{ Tag = '[SUCCESS]';   Color = 'Green';  Outcome = 'Success'; RebootRequired = $false }
            'KB2' = [pscustomobject]@{ Tag = '[REBOOT REQ]'; Color = 'Yellow'; Outcome = 'Reboot';  RebootRequired = $true }
            'KB3' = [pscustomobject]@{ Tag = '[FAILED 0x80070005]'; Color = 'Red'; Outcome = 'Failed'; RebootRequired = $false }
            'KB4' = [pscustomobject]@{ Tag = '[SUCCESS*]';  Color = 'Yellow'; Outcome = 'Warn';    RebootRequired = $false }
        }
        $r = Invoke-WursaWindowsUpdatePass
        $r.Found | Should -Be 4
        $r.Succeeded | Should -Be 3
        $r.Failed | Should -Be 1
        $r.RebootRequired | Should -BeTrue
    }

    It 'counts an install that throws as a failure and keeps going' {
        $mk = { param($t) [pscustomobject]@{ Title = $t } }
        $script:SearchBlock = { param($q) if ($q -match 'Driver') { [pscustomobject]@{ Updates = @() } } else { [pscustomobject]@{ Updates = @((& $mk 'KB1'), (& $mk 'KB2')) } } }
        $script:Results = @{ 'KB2' = [pscustomobject]@{ Tag = '[SUCCESS]'; Color = 'Green'; Outcome = 'Success'; RebootRequired = $false } }
        Mock Install-WursaUpdate { if ($Update.Title -eq 'KB1') { throw 'boom' } $script:Results[$Update.Title] }
        $r = Invoke-WursaWindowsUpdatePass
        $r.Failed | Should -Be 1
        $r.Succeeded | Should -Be 1
    }
}

Describe 'Final status with failures and reboots (finding 30a)' {
    It 'maps failures <Failed>/successes <Succ>, reboot <Reboot>, state <State>/<Ipu> to exit <Code>' -ForEach @(
        @{ NoUp = $true;  Build = 26100; State = 'SKIPPED';   Ipu = '';                Failed = 1; Succ = 2; Reboot = $false; Code = 3;    Like = '*INCOMPLETE*' }
        @{ NoUp = $true;  Build = 26100; State = 'SKIPPED';   Ipu = '';                Failed = 1; Succ = 0; Reboot = $false; Code = 1;    Like = '*INCOMPLETE*' }
        @{ NoUp = $true;  Build = 26100; State = 'SKIPPED';   Ipu = '';                Failed = 1; Succ = 1; Reboot = $true;  Code = 3010; Like = '*REBOOT REQUIRED*' }
        @{ NoUp = $false; Build = 26300; State = 'CURRENT';   Ipu = '';                Failed = 1; Succ = 0; Reboot = $false; Code = 1;    Like = '*INCOMPLETE*' }
        @{ NoUp = $false; Build = 26300; State = 'CURRENT';   Ipu = '';                Failed = 0; Succ = 3; Reboot = $true;  Code = 3010; Like = '*REBOOT REQUIRED*' }
        @{ NoUp = $false; Build = 26300; State = 'CURRENT';   Ipu = '';                Failed = 0; Succ = 3; Reboot = $false; Code = 0;    Like = '*SYSTEM CURRENT*' }
        @{ NoUp = $false; Build = 26100; State = 'DISPATCHED'; Ipu = 'RUNNING x';       Failed = 0; Succ = 1; Reboot = $true;  Code = 0;    Like = '*IN PROGRESS*' }
        @{ NoUp = $false; Build = 26100; State = 'DISPATCHED'; Ipu = 'FAILED URL x';    Failed = 1; Succ = 1; Reboot = $true;  Code = 1;    Like = '*NOT COMPLETE*' }
        @{ NoUp = $false; Build = 26100; State = 'REBOOT_REQUIRED'; Ipu = '';          Failed = 1; Succ = 1; Reboot = $false; Code = 3010; Like = '*REBOOT REQUIRED TO FINISH*' }
    ) {
        $r = Resolve-WursaFinalStatus -NoUpgrade $NoUp -FinalBuild $Build -FinalVersion '25H2' -FeatureUpgradeState $State -IpuState $Ipu -RebootPending $Reboot -FailedCount $Failed -SucceededCount $Succ
        $r.Code | Should -Be $Code
        ($r.Lines | ForEach-Object { $_.Text }) -join "`n" | Should -BeLike $Like
    }
}

Describe 'Failures reach the exit code in a full run (finding 30a)' {
    BeforeEach { Set-MainMocks -Build 26300 -Version '26H2' }

    It 'exits 3 when some updates installed and one failed' {
        Mock Invoke-WursaWindowsUpdatePass { [pscustomobject]@{ Status = 'Installed'; Found = 2; Succeeded = 1; Failed = 1; RebootRequired = $false; Error = $null; Searcher = $null; FeatureOffered = $false } }
        $NoUpgrade = $true; $No3rdParty = $true
        Get-ExitCode { Invoke-WursaMain } | Should -Be 3
    }

    It 'exits 1 when the scan itself failed' {
        Mock Invoke-WursaWindowsUpdatePass { [pscustomobject]@{ Status = 'ScanFailed'; Found = 0; Succeeded = 0; Failed = 1; RebootRequired = $false; Error = 'x'; Searcher = $null; FeatureOffered = $false } }
        $NoUpgrade = $true; $No3rdParty = $true
        Get-ExitCode { Invoke-WursaMain } | Should -Be 1
    }

    It 'counts third-party failures and a third-party reboot request' {
        Mock Invoke-WursaThirdPartyUpdate { [pscustomobject]@{ Status = 'Done'; Succeeded = 1; Failed = 1; RebootRequired = $false } }
        $NoUpgrade = $true; $No3rdParty = $false; $Update3rdParty = $true
        Get-ExitCode { Invoke-WursaMain } | Should -Be 3
        Mock Invoke-WursaThirdPartyUpdate { [pscustomobject]@{ Status = 'Done'; Succeeded = 1; Failed = 0; RebootRequired = $true } }
        Get-ExitCode { Invoke-WursaMain } | Should -Be 3010
    }
}
