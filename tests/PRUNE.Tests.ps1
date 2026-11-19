#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# PRUNE tests. Every profile, CIM, scheduled task, registry, key and file call is mocked, and every
# path is a fixture string. Nothing here touches a real profile or deletes anything real.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath PRUNE)

    function New-FakeProfile {
        param([string]$Sid, [string]$Path, [bool]$Loaded = $false, [bool]$Special = $false, [int]$DaysOld = 400)
        [pscustomobject]@{ SID = $Sid; LocalPath = $Path; Loaded = $Loaded; Special = $Special; LastUseTime = (Get-Date).AddDays(-$DaysOld).ToUniversalTime() }
    }

    function Get-Run {
        # Runs Main and returns the exit code plus everything written to the output stream.
        param([string]$Username)
        $code = -1
        $script:RunLines = New-Object System.Collections.ArrayList
        try {
            if ($Username) { Invoke-PruneMain -Username $Username -SkipSizes *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } }
            else { Invoke-PruneMain -SkipSizes *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } }
        } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        [pscustomobject]@{ Code = $code; Text = ($script:RunLines -join "`n") }
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'Entry-point convention' {
    It 'defines Invoke-PruneMain and Exit-KillerScript without running anything on dot-source' {
        Get-Command Invoke-PruneMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
        { Exit-KillerScript 3010 } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
    }
}

Describe 'Exit codes and final status line' {
    BeforeEach {
        $script:Fake = [pscustomobject]@{
            Profiles = @(
                (New-FakeProfile -Sid 'S-1-5-21-1-2-3-1001' -Path 'C:\Users\alice'),
                (New-FakeProfile -Sid 'S-1-5-21-1-2-3-1002' -Path 'C:\Users\bob' -DaysOld 500)
            )
            Services = @()
            Gone     = New-Object System.Collections.ArrayList
        }
        $script:Removed = New-Object System.Collections.ArrayList
        $script:FailSids = @()
        $script:Elevated = $true
        $script:Selection = '1'
        $script:Confirm1 = 'Y'
        $script:Confirm2 = 'DELETE'

        Mock Test-PruneElevated { $script:Elevated }
        Mock Clear-Host { }
        Mock Get-PruneCursorTop { -1 }
        Mock Set-PruneCursor { $false }
        Mock Start-Sleep { }
        Mock Get-Process { }
        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_UserProfile') {
                $live = @($script:Fake.Profiles | Where-Object { $script:Fake.Gone -notcontains $_.SID })
                if ($Filter -match "SID='([^']+)'") { $sid = $Matches[1]; return @($live | Where-Object { $_.SID -eq $sid }) }
                return $live
            }
            if ($ClassName -eq 'Win32_Service') { return @($script:Fake.Services) }
        }
        Mock Remove-CimInstance {
            if ($script:FailSids -contains $InputObject.SID) { throw 'Access is denied' }
            [void]$script:Removed.Add($InputObject.SID)
            [void]$script:Fake.Gone.Add($InputObject.SID)
        } -RemoveParameterType InputObject
        Mock Test-Path { $false }
        Mock Resolve-SIDToAccount { param($SID) switch ($SID) { 'S-1-5-21-1-2-3-1001' { "$env:COMPUTERNAME\alice" } 'S-1-5-21-1-2-3-1002' { "$env:COMPUTERNAME\bob" } default { $null } } }
        Mock Get-AccountStatus { 'Local' }
        Mock Read-PruneLine {
            if ($Prompt -like '*Selection*') { return $script:Selection }
            if ($Prompt -like '*Delete*') { return $script:Confirm1 }
            if ($Prompt -like '*Confirm*') { return $script:Confirm2 }
        }
        Mock Read-PruneKey { $null }
    }

    It 'exits 2 with a FAIL line and touches nothing when not elevated' {
        $script:Elevated = $false
        $r = Get-Run
        $r.Code | Should -Be 2
        $r.Text | Should -Match '\[FAIL\] PRUNE not elevated.*exit 2'
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }

    It 'exits 0 and says so when the selected profile is removed and verified' {
        $r = Get-Run
        $r.Code | Should -Be 0
        $r.Text | Should -Match '\[PASS\] PRUNE complete: 1 profile\(s\) removed and verified \(exit 0\)'
        $script:Removed | Should -Contain 'S-1-5-21-1-2-3-1002'
    }

    It 'exits 1, not 0, when the only removal fails' {
        $script:FailSids = @('S-1-5-21-1-2-3-1002')
        $r = Get-Run
        $r.Code | Should -Be 1
        $r.Text | Should -Match '\[FAIL\] PRUNE failed.*exit 1'
    }

    It 'exits 3 when one removal succeeds and another fails' {
        $script:Selection = '1,2'
        $script:FailSids = @('S-1-5-21-1-2-3-1001')
        $r = Get-Run
        $r.Code | Should -Be 3
        $r.Text | Should -Match '\[WARN\] PRUNE partial.*exit 3'
    }

    It 'exits 1 when the record is removed but every folder remains' {
        Mock Test-Path { $true }
        $script:Selection = '1,2'
        $r = Get-Run
        $r.Code | Should -Be 1
        $r.Text | Should -Match 'folder remains'
        $r.Text | Should -Not -Match '\[PASS\]'
    }

    It 'exits 2 and removes nothing when there is no console to answer the prompts' {
        Mock Read-PruneLine { $null }
        $r = Get-Run
        $r.Code | Should -Be 2
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }

    It 'exits 2 when the confirmation word is wrong' {
        $script:Confirm2 = 'delete'
        $r = Get-Run
        $r.Code | Should -Be 2
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }

    It 'exits 2 when -Username matches nothing' {
        $r = Get-Run -Username 'nobody'
        $r.Code | Should -Be 2
        $r.Text | Should -Match 'No profile found matching'
    }

    It 'exits 2 when every profile is live and there is no console to investigate with' {
        foreach ($p in $script:Fake.Profiles) { $p.Loaded = $true }
        $r = Get-Run
        $r.Code | Should -Be 2
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }
}
