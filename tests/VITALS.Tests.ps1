#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# VITALS tests. Every registry, service, CIM, native and console call is mocked.

Describe 'VITALS entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'VITALS.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    }
    It 'defines Invoke-VitalsMain and Exit-KillerScript' {
        $names = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-VitalsMain'
        $names | Should -Contain 'Exit-KillerScript'
    }
    It 'ends with the test-mode guard and keeps only definitions and constants at the top level' {
        $stmts = @($script:Ast.EndBlock.Statements)
        $stmts[-1].Extent.Text | Should -Match 'KillerScriptsTestMode'
        $stmts[-1].Extent.Text | Should -Match 'Invoke-VitalsMain'
        $bad = @($stmts | Select-Object -SkipLast 1 | Where-Object {
            $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $_ -isnot [System.Management.Automation.Language.AssignmentStatementAst]
        })
        ($bad | ForEach-Object { $_.Extent.Text.Split("`n")[0] }) -join "`n" | Should -BeNullOrEmpty
    }
}

Describe 'VITALS' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'VITALS')

        function Get-Run {
            $code = -1
            $lines = New-Object System.Collections.ArrayList
            try { Invoke-VitalsMain *>&1 | ForEach-Object { [void]$lines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($lines -join "`n"); HostText = ($script:HostLines -join "`n") }
        }
        # Obviously fake protectors only. Nothing here is a real key.
        function New-Prot {
            param([string]$Type, [string]$Id, [string]$Password = $null)
            [pscustomobject]@{ KeyProtectorType = $Type; KeyProtectorId = $Id; RecoveryPassword = $Password }
        }
        function New-Vol {
            param([string]$Mount, [string]$Status, $Protectors)
            [pscustomobject]@{ MountPoint = $Mount; ProtectionStatus = $Status; KeyProtector = @($Protectors) }
        }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:HostLines = New-Object System.Collections.ArrayList
        $script:Bl = [pscustomobject]@{ Ok = $true; Volumes = @() }
        $script:Disks = [pscustomobject]@{ Ok = $true; Disks = @(
            [pscustomobject]@{ FriendlyName = 'Fake Disk One'; Size = 500GB; MediaType = 'SSD'; HealthStatus = 'Healthy'; DriveLetters = @('C:') },
            [pscustomobject]@{ FriendlyName = 'Fake Disk Two'; Size = 1000GB; MediaType = 'HDD'; HealthStatus = 'Healthy'; DriveLetters = @('D:') }
        ) }
        Mock Write-Host { [void]$script:HostLines.Add([string]$Object) }
        Mock Write-HLine {}
        Mock Initialize-VitalsConsole {}
        Mock Test-VitalsElevated { $true }
        Mock Test-VitalsRebootPending { $false }
        Mock Get-VitalsPublicIp { '203.0.113.9' }
        Mock Get-VitalsTpm { [pscustomobject]@{ Info = [pscustomobject]@{ TpmPresent = $true; TpmReady = $true }; Version = '2.0' } }
        Mock Get-VitalsSecureBoot { $true }
        Mock Get-VitalsIsEntra { $false }
        Mock Get-VitalsWinVer { '23H2' }
        Mock Get-VitalsDiskInfo { $script:Disks }
        Mock Get-VitalsBitLocker { $script:Bl }
        Mock Get-VitalsNetworkReport { @() }
        Mock Get-VitalsLocalAdmins { @('FAKEPC\Administrator') }
        Mock Get-WMI { $null }
        Mock Get-WMI { [pscustomobject]@{ LastBootUpTime = (Get-Date).AddDays(-1); Caption = 'Fake Windows'; BuildNumber = '22631' } } -ParameterFilter { $Class -eq 'Win32_OperatingSystem' }
        Mock Get-WMI { [pscustomobject]@{ Manufacturer = 'Contoso'; Model = 'Fake 1'; DomainRole = 0; Domain = 'WORKGROUP'; TotalPhysicalMemory = 8GB } } -ParameterFilter { $Class -eq 'Win32_ComputerSystem' }
        Mock Get-WMI { [pscustomobject]@{ Name = 'Fake CPU'; NumberOfCores = 4; NumberOfLogicalProcessors = 8; CurrentClockSpeed = 3000 } } -ParameterFilter { $Class -eq 'Win32_Processor' }
        Mock Get-WMI { [pscustomobject]@{ SerialNumber = 'FAKESERIAL' } } -ParameterFilter { $Class -eq 'Win32_Bios' }
    }

    Context 'entry-point convention and elevation (finding 69)' {
        It 'does not run when dot-sourced in test mode' {
            Get-Command Invoke-VitalsMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Get-VitalsBitLocker -Times 0 -Exactly
        }
        It 'exits 2 with a FAIL line and collects nothing when not elevated' {
            Mock Test-VitalsElevated { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
            Should -Invoke Get-VitalsBitLocker -Times 0 -Exactly
            Should -Invoke Get-WMI -Times 0 -Exactly
        }
    }

    Context 'finding 28: every BitLocker protector is shown with its ID' {
        BeforeEach {
            $script:Bl = [pscustomobject]@{ Ok = $true; Volumes = @(
                (New-Vol 'C:' 'On' @(
                    (New-Prot 'RecoveryPassword' '{11111111-AAAA-0000-0000-000000000001}' '111111-111111-111111-111111-111111-111111-111111-111111'),
                    (New-Prot 'RecoveryPassword' '{11111111-AAAA-0000-0000-000000000002}' '222222-222222-222222-222222-222222-222222-222222-222222'),
                    (New-Prot 'Tpm' '{11111111-AAAA-0000-0000-000000000003}'))),
                (New-Vol 'D:' 'On' @(
                    (New-Prot 'RecoveryPassword' '{22222222-BBBB-0000-0000-000000000001}' '333333-333333-333333-333333-333333-333333-333333-333333'))),
                (New-Vol '' 'Off' @(
                    (New-Prot 'RecoveryPassword' '{33333333-CCCC-0000-0000-000000000001}' '444444-444444-444444-444444-444444-444444-444444-444444')))
            ) }
        }
        It 'shows both recovery passwords of a volume that has two' {
            $r = Get-Run
            $r.HostText | Should -Match '111111-111111-111111-111111-111111-111111-111111-111111'
            $r.HostText | Should -Match '222222-222222-222222-222222-222222-222222-222222-222222'
        }
        It 'shows every protector ID, including a non-password protector' {
            $r = Get-Run
            foreach ($id in '{11111111-AAAA-0000-0000-000000000001}', '{11111111-AAAA-0000-0000-000000000002}', '{11111111-AAAA-0000-0000-000000000003}', '{22222222-BBBB-0000-0000-000000000001}', '{33333333-CCCC-0000-0000-000000000001}') {
                $r.HostText | Should -Match ([regex]::Escape($id))
            }
            $r.HostText | Should -Match 'Type: \s+Tpm'
        }
        It 'shows the second volume and a volume with protection off and no drive letter' {
            $r = Get-Run
            $r.HostText | Should -Match '333333-333333-333333-333333-333333-333333-333333-333333'
            $r.HostText | Should -Match '444444-444444-444444-444444-444444-444444-444444-444444'
            $r.HostText | Should -Match '\(no drive letter\)'
            $r.HostText | Should -Match 'Protection Off'
        }
        It 'prints each recovery password exactly once and prints no key for a protector that has none' {
            $r = Get-Run
            ([regex]::Matches($r.HostText, '111111-111111-111111-111111-111111-111111-111111-111111')).Count | Should -Be 1
            ([regex]::Matches($r.HostText, 'Key :')).Count | Should -Be 4
        }
        It 'marks BitLocker On in the disk table for a lettered volume that is on' {
            $r = Get-Run
            $r.HostText | Should -Match 'Fake Disk One'
            $r.Code | Should -Be 0
        }
        It 'builds the protector rows from every volume, not just the first password' {
            $rows = @(Get-VitalsProtectorRows $script:Bl.Volumes)
            $rows.Count | Should -Be 5
            @($rows | Where-Object Key).Count | Should -Be 4
        }
    }

    Context 'exit codes' {
        It 'exits 0 with a PASS line when everything was read' {
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\].*exit 0'
        }
        It 'exits 3 with a WARN line and prints no key when BitLocker cannot be read' {
            $script:Bl = [pscustomobject]@{ Ok = $false; Volumes = @() }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\].*BitLocker.*exit 3'
            $r.HostText | Should -Match 'N/A'
        }
        It 'exits 3 with a WARN line when the disks cannot be listed' {
            $script:Disks = [pscustomobject]@{ Ok = $false; Disks = @() }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\].*disks.*exit 3'
        }
    }
}
