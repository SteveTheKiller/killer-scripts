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
