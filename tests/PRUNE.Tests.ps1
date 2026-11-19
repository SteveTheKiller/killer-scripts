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

Describe 'Finding 47: disabled and failed account lookups' {
    BeforeAll {
        $script:LocalSid = 'S-1-5-21-111-222-333-1005'
        $script:LocalName = "$env:COMPUTERNAME\olduser"
    }

    It 'reports Disabled for a disabled local account' {
        Mock Get-LocalUser { [pscustomobject]@{ Name = 'olduser'; Enabled = $false } }
        Get-AccountStatus -SID $script:LocalSid -AccountName $script:LocalName | Should -Be 'Disabled'
    }

    It 'reports Local for an enabled local account' {
        Mock Get-LocalUser { [pscustomobject]@{ Name = 'olduser'; Enabled = $true } }
        Get-AccountStatus -SID $script:LocalSid -AccountName $script:LocalName | Should -Be 'Local'
    }

    It 'reports a failed lookup as LookupFailed, never as Local' {
        Mock Get-LocalUser { throw 'The service is not running' }
        Get-AccountStatus -SID $script:LocalSid -AccountName $script:LocalName | Should -Be 'LookupFailed'
    }

    It 'looks the account up by SID so a similar name cannot answer for it' {
        Mock Get-LocalUser { [pscustomobject]@{ Name = 'olduser'; Enabled = $false } } -ParameterFilter { $SID -and $SID.Value -eq 'S-1-5-21-111-222-333-1005' }
        Mock Get-LocalUser { [pscustomobject]@{ Name = 'olduser2'; Enabled = $true } }
        Get-AccountStatus -SID $script:LocalSid -AccountName $script:LocalName | Should -Be 'Disabled'
    }

    It 'reports a domain account as Domain without querying local users' {
        Mock Get-LocalUser { throw 'must not be called' }
        Get-AccountStatus -SID 'S-1-5-21-900-800-700-1100' -AccountName 'CONTOSO\jdoe' | Should -Be 'Domain'
    }
}

Describe 'Finding 22: service and task owners are matched by exact identity' {
    BeforeAll {
        $script:BobSid     = 'S-1-5-21-111-222-333-1010'
        $script:BobbySid   = 'S-1-5-21-111-222-333-1011'
        $script:Bob2Sid    = 'S-1-5-21-111-222-333-1012'
        $script:CorpBobSid = 'S-1-5-21-900-800-700-1500'
        $script:Target = [pscustomobject]@{ SID = $script:BobSid; AccountName = "$env:COMPUTERNAME\bob"; DisplayName = "$env:COMPUTERNAME\bob"; LocalPath = 'C:\Users\bob' }
    }

    BeforeEach {
        # Names that merely contain "bob" map to different accounts. Only the SID decides.
        Mock Resolve-PrunePrincipalSid {
            param($Name)
            switch ($Name) {
                'bob'                { $script:BobSid }
                ".\bob"              { $script:BobSid }
                "$env:COMPUTERNAME\bob" { $script:BobSid }
                'bobby'              { $script:BobbySid }
                ".\bobby"            { $script:BobbySid }
                'bob2'               { $script:Bob2Sid }
                'CORP\bob'           { $script:CorpBobSid }
                'LocalSystem'        { 'S-1-5-18' }
                'S-1-5-21-111-222-333-1010' { $Name }
                default              { $null }
            }
        }
        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_Service') {
                @(
                    [pscustomobject]@{ Name = 'SvcBob';    StartName = '.\bob';        State = 'Running'; DisplayName = 'Bob service' },
                    [pscustomobject]@{ Name = 'SvcBobby';  StartName = '.\bobby';      State = 'Running'; DisplayName = 'Bobby service' },
                    [pscustomobject]@{ Name = 'SvcCorp';   StartName = 'CORP\bob';     State = 'Running'; DisplayName = 'Corp Bob service' },
                    [pscustomobject]@{ Name = 'SvcSystem'; StartName = 'LocalSystem';  State = 'Running'; DisplayName = 'System service' },
                    [pscustomobject]@{ Name = 'SvcGone';   StartName = 'robobob';      State = 'Stopped'; DisplayName = 'Unresolvable owner' }
                )
            }
        }
        Mock Get-ScheduledTask {
            @(
                [pscustomobject]@{ TaskName = 'BobTask';   TaskPath = '\'; State = 'Ready'; Principal = [pscustomobject]@{ UserId = 'bob' } },
                [pscustomobject]@{ TaskName = 'BobbyTask'; TaskPath = '\'; State = 'Ready'; Principal = [pscustomobject]@{ UserId = 'bobby' } },
                [pscustomobject]@{ TaskName = 'Bob2Task';  TaskPath = '\'; State = 'Ready'; Principal = [pscustomobject]@{ UserId = 'bob2' } },
                [pscustomobject]@{ TaskName = 'CorpTask';  TaskPath = '\'; State = 'Ready'; Principal = [pscustomobject]@{ UserId = 'CORP\bob' } },
                [pscustomobject]@{ TaskName = 'SidTask';   TaskPath = '\'; State = 'Ready'; Principal = [pscustomobject]@{ UserId = 'S-1-5-21-111-222-333-1010' } }
            )
        }
    }

    It 'lists only the services that run as exactly this account' {
        $o = Find-PruneOwnedItems -Target $script:Target
        @($o.Services | ForEach-Object Name) | Should -Be @('SvcBob')
    }

    It 'lists only the tasks that run as exactly this account, including one that names the SID' {
        $o = Find-PruneOwnedItems -Target $script:Target
        @($o.Tasks | ForEach-Object TaskName | Sort-Object) | Should -Be @('BobTask', 'SidTask')
    }

    It 'unregisters only the exact-owner tasks when the technician agrees' {
        $script:keys = New-Object System.Collections.Queue
        $script:keys.Enqueue([pscustomobject]@{ Key = [ConsoleKey]::Y; KeyChar = 'y' })
        $script:keys.Enqueue([pscustomobject]@{ Key = [ConsoleKey]::N; KeyChar = 'n' })
        Mock Read-PruneKey { if ($script:keys.Count -gt 0) { $script:keys.Dequeue() } else { $null } }
        Mock Unregister-ScheduledTask { }
        $null = Invoke-PruneInvestigation -Target $script:Target
        Should -Invoke Unregister-ScheduledTask -Times 2 -Exactly
        Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'BobTask' }
        Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'SidTask' }
    }

    It 'leaves a task alone when its owner cannot be resolved' {
        Mock Get-ScheduledTask { @([pscustomobject]@{ TaskName = 'Mystery'; TaskPath = '\'; State = 'Ready'; Principal = [pscustomobject]@{ UserId = 'nosuchuser' } }) }
        $o = Find-PruneOwnedItems -Target $script:Target
        @($o.Tasks).Count | Should -Be 0
    }
}

Describe 'Finding 22: principal resolution' {
    It 'resolves a SID string to itself without a directory lookup' {
        Resolve-PrunePrincipalSid -Name 'S-1-5-21-111-222-333-1010' | Should -Be 'S-1-5-21-111-222-333-1010'
    }
}

Describe 'Finding 68: Unresolved is reported apart from Orphaned' {
    BeforeAll {
        $script:MachineSid = 'S-1-5-21-111-222-333'
    }

    BeforeEach {
        Mock Get-LocalUser { [pscustomobject]@{ Name = 'Administrator'; Enabled = $true; SID = (New-Object System.Security.Principal.SecurityIdentifier('S-1-5-21-111-222-333-500')) } }
    }

    It 'labels a SID from this machine that no longer resolves as Orphaned' {
        Get-AccountStatus -SID 'S-1-5-21-111-222-333-1020' -AccountName $null | Should -Be 'Orphaned'
    }

    It 'labels a domain SID that does not resolve (laptop off VPN) as Unresolved, never Orphaned' {
        Get-AccountStatus -SID 'S-1-5-21-900-800-700-1500' -AccountName $null | Should -Be 'Unresolved'
    }

    It 'labels an Entra ID SID that does not resolve as Unresolved' {
        Get-AccountStatus -SID 'S-1-12-1-1111-2222-3333-4444' -AccountName $null | Should -Be 'Unresolved'
    }

    It 'labels the SID Unresolved, not Orphaned, when this machine SID cannot be determined' {
        Mock Get-LocalUser { throw 'The service is not running' }
        Get-AccountStatus -SID 'S-1-5-21-111-222-333-1020' -AccountName $null | Should -Be 'Unresolved'
    }

    It 'flags an Unresolved profile in the delete list as unverified, not as orphaned' {
        $script:Fake = [pscustomobject]@{
            Profiles = @((New-FakeProfile -Sid 'S-1-5-21-900-800-700-1500' -Path 'C:\Users\vpnuser'))
            Gone     = New-Object System.Collections.ArrayList
        }
        Mock Test-PruneElevated { $true }
        Mock Clear-Host { }
        Mock Get-PruneCursorTop { -1 }
        Mock Set-PruneCursor { $false }
        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_UserProfile') { return @($script:Fake.Profiles) }
        }
        Mock Remove-CimInstance { }  -RemoveParameterType InputObject
        Mock Test-Path { $false }
        Mock Resolve-SIDToAccount { $null }
        Mock Read-PruneLine {
            if ($Prompt -like '*Selection*') { return '1' }
            return $null
        }
        $r = Get-Run
        $r.Text | Should -Match 'unresolved'
        $r.Text | Should -Not -Match '\[orphaned'
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }
}

Describe 'Finding 23: removal rechecks the live record and surfaces partial results' {
    BeforeAll {
        function New-Selected {
            param([string]$Sid = 'S-1-5-21-1-2-3-1002', [string]$Path = 'C:\Users\bob', [string]$Name = 'bob', [bool]$FolderExists = $true)
            [pscustomobject]@{ SID = $Sid; LocalPath = $Path; DisplayName = "$env:COMPUTERNAME\$Name"; AccountName = "$env:COMPUTERNAME\$Name"; FolderExists = $FolderExists; Loaded = $false }
        }
    }

    BeforeEach {
        $script:Fresh = New-FakeProfile -Sid 'S-1-5-21-1-2-3-1002' -Path 'C:\Users\bob'
        $script:AllRecords = @()
        $script:Unloads = 0
        $script:ReloadedAfterUnload = $false
        $script:Active = 'someoneelse'
        $script:Removed = New-Object System.Collections.ArrayList
        $script:Deleted = New-Object System.Collections.ArrayList

        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_UserProfile') {
                if ($Filter -match "SID='([^']+)'") { if ($script:Fresh -and $script:Fresh.SID -eq $Matches[1]) { return $script:Fresh } else { return $null } }
                return @($script:AllRecords)
            }
        }
        Mock Remove-CimInstance { [void]$script:Removed.Add($InputObject.SID) } -RemoveParameterType InputObject
        Mock Remove-Item { [void]$script:Deleted.Add([string]$LiteralPath) } -ParameterFilter { $LiteralPath -like 'C:\Users\*' }
        Mock Test-Path { $true }
        Mock Get-PruneProfilesRoot { 'C:\Users' }
        Mock Test-PruneReparsePoint { $false }
        Mock Start-Sleep { }
        Mock Get-ActiveUser { $script:Active }
        Mock Invoke-PruneRegUnload { $script:Unloads++; if (-not $script:ReloadedAfterUnload) { $script:Fresh.Loaded = $false } }
    }

    It 'removes a profile that is still eligible' {
        $r = Remove-PruneProfile -Profile (New-Selected)
        $r.Removed | Should -BeTrue
        $script:Removed | Should -Contain 'S-1-5-21-1-2-3-1002'
    }

    It 'refuses a profile that became loaded after it was selected' {
        $script:Fresh.Loaded = $true
        $r = Remove-PruneProfile -Profile (New-Selected)
        $r.Removed | Should -BeFalse
        $r.Message | Should -Match 'loaded'
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
        Should -Invoke Invoke-PruneRegUnload -Times 0 -Exactly
    }

    It 'refuses a record that is now a special profile' {
        $script:Fresh.Special = $true
        $r = Remove-PruneProfile -Profile (New-Selected)
        $r.Removed | Should -BeFalse
        $r.Message | Should -Match 'special'
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }

    It 'refuses when the SID now points at a different profile path' {
        $script:Fresh.LocalPath = 'C:\Users\someoneelse'
        $r = Remove-PruneProfile -Profile (New-Selected)
        $r.Removed | Should -BeFalse
        $r.Message | Should -Match 'changed'
        Should -Invoke Remove-CimInstance -Times 0 -Exactly
    }

    It 'reports a failed removal instead of counting it' {
        Mock Remove-CimInstance { throw 'Access is denied' } -RemoveParameterType InputObject
        $r = Remove-PruneProfile -Profile (New-Selected)
        $r.Removed | Should -BeFalse
        $r.Message | Should -Match 'Access is denied'
    }

    It 'does not treat a failed profile lookup as the record being gone' {
        Mock Get-CimInstance { throw 'WMI repository is unavailable' }
        $r = Remove-PruneProfile -Profile (New-Selected)
        $r.Removed | Should -BeFalse
        Should -Invoke Remove-Item -Times 0 -Exactly
    }

    It 'rejects a SID that is not a SID before it reaches a CIM filter' {
        $r = Remove-PruneProfile -Profile (New-Selected -Sid "S-1-5-21-1' OR SID LIKE '%")
        $r.Removed | Should -BeFalse
        Should -Invoke Get-CimInstance -Times 0 -Exactly
    }

    Context 'when the profile record has disappeared' {
        BeforeEach { $script:Fresh = $null }

        It 'removes the leftover folder by literal path when it is a plain child of the profiles root' {
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Users\bob')
            $r.Removed | Should -BeTrue
            @($script:Deleted) | Should -Be @('C:\Users\bob')
        }

        It 'never deletes a wildcard path' {
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Users\*')
            $r.Removed | Should -BeFalse
            $script:Deleted.Count | Should -Be 0
        }

        It 'never deletes outside the profiles root' {
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Windows\System32')
            $r.Removed | Should -BeFalse
            $script:Deleted.Count | Should -Be 0
        }

        It 'never deletes the profiles root itself' {
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Users')
            $r.Removed | Should -BeFalse
            $script:Deleted.Count | Should -Be 0
        }

        It 'never deletes a folder another profile record still owns' {
            $script:AllRecords = @((New-FakeProfile -Sid 'S-1-5-21-1-2-3-1099' -Path 'C:\Users\bob'))
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Users\bob')
            $r.Removed | Should -BeFalse
            $script:Deleted.Count | Should -Be 0
        }

        It 'never deletes through a junction or symlink' {
            Mock Test-PruneReparsePoint { $true }
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Users\bob')
            $r.Removed | Should -BeFalse
            $script:Deleted.Count | Should -Be 0
        }

        It 'reports success without deleting when the folder is gone too' {
            Mock Test-Path { $false }
            $r = Remove-PruneProfile -Profile (New-Selected -Path 'C:\Users\bob')
            $r.Removed | Should -BeTrue
            $script:Deleted.Count | Should -Be 0
        }
    }

    Context 'when asked to release a held hive' {
        BeforeEach { $script:Fresh.Loaded = $true }

        It 'unloads a stale hive, rechecks, and removes the profile' {
            $r = Remove-PruneProfile -Profile (New-Selected) -ReleaseHive
            $script:Unloads | Should -Be 1
            $r.Removed | Should -BeTrue
            $script:Removed | Should -Contain 'S-1-5-21-1-2-3-1002'
        }

        It 'does not remove the profile when the hive is still held after the unload' {
            Mock Invoke-PruneRegUnload { $script:Unloads++ }
            $r = Remove-PruneProfile -Profile (New-Selected) -ReleaseHive
            $r.Removed | Should -BeFalse
            $r.Message | Should -Match 'held open'
            Should -Invoke Remove-CimInstance -Times 0 -Exactly
        }

        It 'does not remove the profile when it is loaded again right after the unload' {
            $script:ReloadedAfterUnload = $true
            $r = Remove-PruneProfile -Profile (New-Selected) -ReleaseHive
            $r.Removed | Should -BeFalse
            Should -Invoke Remove-CimInstance -Times 0 -Exactly
        }

        It 'never unloads the hive of the signed-in user' {
            $script:Active = 'bob'
            $r = Remove-PruneProfile -Profile (New-Selected) -ReleaseHive
            $r.Removed | Should -BeFalse
            $r.Message | Should -Match 'signed-in'
            $script:Unloads | Should -Be 0
            Should -Invoke Remove-CimInstance -Times 0 -Exactly
        }

        It 'never unloads a special profile' {
            $script:Fresh.Special = $true
            $r = Remove-PruneProfile -Profile (New-Selected) -ReleaseHive
            $r.Removed | Should -BeFalse
            $script:Unloads | Should -Be 0
        }
    }
}

Describe 'Finding 23: a profile that becomes loaded after selection ends the run honestly' {
    BeforeEach {
        $script:Fake = [pscustomobject]@{
            Profiles = @(
                (New-FakeProfile -Sid 'S-1-5-21-1-2-3-1001' -Path 'C:\Users\alice'),
                (New-FakeProfile -Sid 'S-1-5-21-1-2-3-1002' -Path 'C:\Users\bob' -DaysOld 500)
            )
            Gone = New-Object System.Collections.ArrayList
        }
        $script:Removed = New-Object System.Collections.ArrayList
        $script:LoadsAfterScan = @()
        $script:ScanDone = $false
        $script:Selection = '1,2'

        Mock Test-PruneElevated { $true }
        Mock Clear-Host { }
        Mock Get-PruneCursorTop { -1 }
        Mock Set-PruneCursor { $false }
        Mock Start-Sleep { }
        Mock Get-ActiveUser { $null }
        Mock Get-PruneProfilesRoot { 'C:\Users' }
        Mock Test-PruneReparsePoint { $false }
        Mock Test-Path { $false }
        Mock Resolve-SIDToAccount { param($SID) switch ($SID) { 'S-1-5-21-1-2-3-1001' { "$env:COMPUTERNAME\alice" } 'S-1-5-21-1-2-3-1002' { "$env:COMPUTERNAME\bob" } } }
        Mock Get-AccountStatus { 'Local' }
        Mock Get-CimInstance {
            if ($ClassName -eq 'Win32_UserProfile') {
                $live = @($script:Fake.Profiles | Where-Object { $script:Fake.Gone -notcontains $_.SID })
                if ($Filter -match "SID='([^']+)'") {
                    $sid = $Matches[1]
                    $hit = @($live | Where-Object { $_.SID -eq $sid })
                    foreach ($h in $hit) { if ($script:LoadsAfterScan -contains $h.SID) { $h.Loaded = $true } }
                    return $hit
                }
                return $live
            }
        }
        Mock Remove-CimInstance { [void]$script:Removed.Add($InputObject.SID); [void]$script:Fake.Gone.Add($InputObject.SID) } -RemoveParameterType InputObject
        Mock Read-PruneLine {
            if ($Prompt -like '*Selection*') { return $script:Selection }
            if ($Prompt -like '*Delete*') { return 'Y' }
            if ($Prompt -like '*Confirm*') { return 'DELETE' }
        }
        Mock Read-PruneKey { $null }
    }

    It 'skips the profile that became loaded, removes the other, and exits 3' {
        $script:LoadsAfterScan = @('S-1-5-21-1-2-3-1001')
        $r = Get-Run
        $r.Code | Should -Be 3
        $script:Removed | Should -Not -Contain 'S-1-5-21-1-2-3-1001'
        $script:Removed | Should -Contain 'S-1-5-21-1-2-3-1002'
        $r.Text | Should -Match 'loaded'
        $r.Text | Should -Match '\[WARN\] PRUNE partial'
    }

    It 'exits 1 and removes nothing when the only selected profile became loaded' {
        $script:Selection = '1'
        $script:LoadsAfterScan = @('S-1-5-21-1-2-3-1002')
        $r = Get-Run
        $r.Code | Should -Be 1
        $script:Removed.Count | Should -Be 0
        $r.Text | Should -Match '\[FAIL\] PRUNE failed'
    }
}
