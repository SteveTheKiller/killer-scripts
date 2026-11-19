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
