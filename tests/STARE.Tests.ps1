#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# STARE tests. The scheduler, the COM scheduling service, the console and the file system ACLs are all
# mocked, so nothing here touches the machine. Keys and typed lines are scripted through two queues.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath 'STARE')

    function New-TestKey {
        # A console key press. 'ESC' is Escape, a digit or letter is that key.
        param([string]$Char)
        if ($Char -eq 'ESC') { return [pscustomobject]@{ Key = [ConsoleKey]::Escape; KeyChar = [char]27 } }
        $name = if ($Char -match '^[0-9]$') { "D$Char" } else { $Char.ToUpper() }
        [pscustomobject]@{ Key = [ConsoleKey]$name; KeyChar = [char]$Char }
    }

    function Set-Script {
        # Queues the key presses and typed lines the run will consume.
        param([string[]]$Keys = @(), [object[]]$Lines = @())
        $script:KeyQueue  = New-Object System.Collections.Queue
        $script:LineQueue = New-Object System.Collections.Queue
        foreach ($k in $Keys) { $script:KeyQueue.Enqueue((New-TestKey $k)) }
        foreach ($l in $Lines) { $script:LineQueue.Enqueue($l) }
    }

    function New-TestTask {
        param([string]$Name, [string]$Path = '\')
        [pscustomobject]@{ TaskName = $Name; TaskPath = $Path; State = 'Ready' }
    }

    function Get-Run {
        # Runs Main and returns the exit code, the output stream and everything written to the host.
        $code = -1
        $script:Said = New-Object System.Collections.ArrayList
        $out = New-Object System.Collections.ArrayList
        try { Invoke-StareMain *>&1 | ForEach-Object { [void]$out.Add([string]$_) } } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        [pscustomobject]@{ Code = $code; Output = ($out -join "`n"); Host = ($script:Said -join "`n"); All = ((@($out) + @($script:Said)) -join "`n") }
    }

    function New-FakeScheduleService {
        # A stand-in for the Schedule.Service COM object. It records every call it receives.
        param([switch]$FailRegister)
        $folder = [pscustomobject]@{ Calls = (New-Object System.Collections.ArrayList); Fail = [bool]$FailRegister }
        Add-Member -InputObject $folder -MemberType ScriptMethod -Name DeleteTask -Value {
            param($name, $flags) [void]$this.Calls.Add("DeleteTask:$name")
        }
        Add-Member -InputObject $folder -MemberType ScriptMethod -Name RegisterTaskDefinition -Value {
            param($name, $def, $flags, $user, $pw, $logon)
            [void]$this.Calls.Add("Register:${name}:$flags")
            if ($this.Fail) { throw 'simulated registration failure' }
        }
        $def = [pscustomobject]@{
            Settings  = [pscustomobject]@{ Enabled = $false; Hidden = $false; MultipleInstances = 0; ExecutionTimeLimit = '' }
            Principal = [pscustomobject]@{ UserId = ''; RunLevel = 0 }
            Triggers  = [pscustomobject]@{ }
            Actions   = [pscustomobject]@{ }
        }
        Add-Member -InputObject $def.Triggers -MemberType ScriptMethod -Name Create -Value {
            param($t) [pscustomobject]@{ StartBoundary = ''; Enabled = $false; MonthsOfYear = 0; RunOnLastDayOfMonth = $false; DaysOfMonth = 0 }
        }
        Add-Member -InputObject $def.Actions -MemberType ScriptMethod -Name Create -Value {
            param($t) [pscustomobject]@{ Path = ''; Arguments = '' }
        }
        $svc = [pscustomobject]@{ Folder = $folder; Def = $def }
        Add-Member -InputObject $svc -MemberType ScriptMethod -Name GetFolder -Value { param($p) $this.Folder }
        Add-Member -InputObject $svc -MemberType ScriptMethod -Name NewTask -Value { param($f) $this.Def }
        $svc
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'STARE' {
    BeforeEach {
        # Identity and console boundaries
        Mock Test-StareElevated { $true }
        Mock Test-StareConsole { $true }
        Mock Read-StareKey { $script:KeyQueue.Dequeue() }
        Mock Read-Esc { $script:LineQueue.Dequeue() }
        Mock Get-StareCursor { [pscustomobject]@{ X = 0; Y = 0 } }
        Mock Move-StareCursor {}
        Mock Clear-Host {}
        Mock Start-Sleep {}
        Mock Write-Host { [void]$script:Said.Add(([string]($Object -join ' '))) }
        # Scheduler boundary: a mutable list of tasks the mocks read and write
        $script:Sched = @()
        Mock Get-ScheduledTask {
            $r = @($script:Sched)
            if ($TaskName) { $r = @($r | Where-Object { $_.TaskName -eq $TaskName }) }
            if ($TaskPath) { $r = @($r | Where-Object { $_.TaskPath -eq $TaskPath }) }
            $r
        }
        Mock Get-ScheduledTaskInfo { $null }
        Mock Unregister-ScheduledTask {
            $script:Sched = @($script:Sched | Where-Object { -not ($_.TaskName -eq $TaskName -and ((-not $TaskPath) -or $_.TaskPath -eq $TaskPath)) })
        }
        Mock Register-ScheduledTask {}
        # New-ScheduledTaskAction and New-ScheduledTaskTrigger only build in-memory objects, so they run for real.
        Mock Get-StareScheduleService { $script:Fake }
        Set-Script
    }

    Context 'entry-point convention' {
        It 'defines Invoke-StareMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-StareMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Get-ScheduledTask -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'exits 0 and says so when the technician quits without changing anything' {
            Set-Script -Keys 'Q'
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Output | Should -Match 'STARE RESULT: exit code 0 - session ended, 0 change\(s\) applied'
        }
        It 'treats ESC on the list as quit' {
            Set-Script -Keys 'ESC'
            (Get-Run).Code | Should -Be 0
        }
        It 'exits 2 and touches nothing when not elevated' {
            Mock Test-StareElevated { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Output | Should -Match 'STARE RESULT: exit code 2 - not elevated'
            Should -Invoke Get-ScheduledTask -Times 0 -Exactly
            Should -Invoke Register-ScheduledTask -Times 0 -Exactly
        }
        It 'exits 2 without reading a key when there is no console to answer prompts' {
            Mock Test-StareConsole { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Output | Should -Match 'STARE RESULT: exit code 2 - .*needs a console'
            Should -Invoke Read-StareKey -Times 0 -Exactly
            Should -Invoke Get-ScheduledTask -Times 0 -Exactly
        }
        It 'exits 0 and reports one change after a task is created' {
            Set-Script -Keys 'C', '1', '4', 'X', 'Q' -Lines 'NightlyJob', 'shutdown.exe /r'
            $r = Get-Run
            $r.Code | Should -Be 0 -Because $r.Output
            $r.Output | Should -Match 'STARE RESULT: exit code 0 - session ended, 1 change\(s\) applied'
            Should -Invoke Register-ScheduledTask -Times 1 -Exactly
        }
        It 'exits 1 with an honest line when the only attempted change fails' {
            Mock Register-ScheduledTask { throw 'access denied' }
            Set-Script -Keys 'C', '1', '4', 'X', 'Q' -Lines 'NightlyJob', 'shutdown.exe /r'
            $r = Get-Run
            $r.Code | Should -Be 1 -Because $r.Output
            $r.Output | Should -Match 'STARE RESULT: exit code 1 - 1 change\(s\) failed'
        }
        It 'exits 3 when one change worked and another failed' {
            $script:Calls = 0
            Mock Register-ScheduledTask { $script:Calls++; if ($script:Calls -eq 2) { throw 'access denied' } }
            Set-Script -Keys 'C', '1', '4', 'X', 'C', '1', '4', 'X', 'Q' -Lines 'JobOne', 'shutdown.exe /r', 'JobTwo', 'shutdown.exe /r'
            $r = Get-Run
            $r.Code | Should -Be 3 -Because $r.Output
            $r.Output | Should -Match 'STARE RESULT: exit code 3 - partial: 1 change\(s\) applied, 1 failed'
        }
    }

    Context 'finding 60: delete targets one task by folder and name and reports the real result' {
        BeforeEach {
            # Two tasks share a name in different folders. The list sorts by name, then folder.
            $script:Sched = @((New-TestTask 'Backup' '\A\'), (New-TestTask 'Backup' '\B\'), (New-TestTask 'Zeta' '\'))
        }
        It 'unregisters only the task in the chosen folder and leaves the same-named one alone' {
            Set-Script -Keys 'D', 'Y', 'Q' -Lines '2'
            $r = Get-Run
            $r.Code | Should -Be 0 -Because $r.All
            Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly
            Should -Invoke Unregister-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'Backup' -and $TaskPath -eq '\B\' }
            @($script:Sched | Where-Object { $_.TaskName -eq 'Backup' -and $_.TaskPath -eq '\A\' }).Count | Should -Be 1
            $r.Host | Should -Match 'Deleted\.'
            $r.Output | Should -Match 'STARE RESULT: exit code 0 - session ended, 1 change\(s\) applied'
        }
        It 'shows the folder in the list and in the confirmation so duplicates can be told apart' {
            Set-Script -Keys 'D', 'N', 'Q' -Lines '1'
            $r = Get-Run
            $r.Host | Should -Match '\\A\\Backup'
            $r.Host | Should -Match "Delete '\\B\\Backup'|Delete '\\A\\Backup'"
        }
        It 'refuses task number 0 instead of resolving it to the last task' {
            Set-Script -Keys 'D', 'Q' -Lines '0', $null
            $r = Get-Run
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
            $r.Host | Should -Match 'Invalid Selection'
            @($script:Sched).Count | Should -Be 3
        }
        It 'refuses a number past the end of the list, including one too large for an integer' {
            Set-Script -Keys 'D', 'Q' -Lines '4', '99999999999', $null
            $r = Get-Run
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
            $r.Code | Should -Be 0 -Because $r.All
            @($script:Sched).Count | Should -Be 3
        }
        It 'deletes nothing when the confirmation is declined' {
            Set-Script -Keys 'D', 'N', 'Q' -Lines '2'
            $r = Get-Run
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
            $r.Host | Should -Match 'Cancelled'
            $r.Host | Should -Not -Match 'Deleted\.'
        }
        It 'does not print Deleted and exits 1 when the scheduler refuses the delete' {
            Mock Unregister-ScheduledTask { throw 'access is denied' }
            Set-Script -Keys 'D', 'Y', 'Q' -Lines '2'
            $r = Get-Run
            $r.Host | Should -Not -Match 'Deleted\.'
            $r.Host | Should -Match 'Delete failed.*access is denied'
            $r.Code | Should -Be 1 -Because $r.All
            $r.Output | Should -Match 'STARE RESULT: exit code 1 - 1 change\(s\) failed'
        }
        It 'reads the result back and does not print Deleted when the task is still registered' {
            Mock Unregister-ScheduledTask {}
            Set-Script -Keys 'D', 'Y', 'Q' -Lines '2'
            $r = Get-Run
            $r.Host | Should -Not -Match 'Deleted\.'
            $r.Host | Should -Match 'Delete failed.*still registered'
            $r.Code | Should -Be 1 -Because $r.All
        }
    }

    Context 'finding 25: a replacement never deletes the working task first' {
        BeforeEach {
            $script:Sched = @((New-TestTask 'Nightly' '\'), (New-TestTask 'Nightly' '\Other\'))
        }
        It 'keeps the existing task and exits 1 when the replacement fails (cmdlet path)' {
            Mock Register-ScheduledTask { throw 'the scheduler rejected the definition' }
            Set-Script -Keys 'C', '1', '4', 'X', 'Q' -Lines 'Nightly', 'shutdown.exe /r'
            $r = Get-Run
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
            @($script:Sched).Count | Should -Be 2
            $r.Host | Should -Match 'Error creating task.*rejected'
            $r.Code | Should -Be 1 -Because $r.All
        }
        It 'replaces an existing task in one validated call and touches no other folder (cmdlet path)' {
            Set-Script -Keys 'C', '1', '4', 'X', 'Q' -Lines 'Nightly', 'shutdown.exe /r'
            $r = Get-Run
            Should -Invoke Unregister-ScheduledTask -Times 0 -Exactly
            Should -Invoke Register-ScheduledTask -Times 1 -Exactly -ParameterFilter { $TaskName -eq 'Nightly' -and $Force }
            $r.Code | Should -Be 0 -Because $r.All
            $r.Host | Should -Match "Task 'Nightly' Created Successfully"
        }
        It 'does not delete the existing task before registering the monthly replacement (COM path)' {
            $script:Fake = New-FakeScheduleService -FailRegister
            Set-Script -Keys 'C', '1', '3', 'X', 'ESC', 'Q' -Lines 'Nightly', 'shutdown.exe /r', '15', '03:00'
            $r = Get-Run
            $script:Fake.Folder.Calls | Should -Not -Contain 'DeleteTask:Nightly'
            @($script:Fake.Folder.Calls | Where-Object { $_ -like 'DeleteTask*' }).Count | Should -Be 0
            $script:Fake.Folder.Calls | Should -Contain 'Register:Nightly:6'
            $r.Host | Should -Match 'ERR'
            $r.Code | Should -Be 1 -Because $r.All
        }
        It 'registers the monthly task with create-or-update and no delete when it works (COM path)' {
            $script:Fake = New-FakeScheduleService
            Set-Script -Keys 'C', '1', '3', 'X', 'Q' -Lines 'Nightly', 'shutdown.exe /r', '15', '03:00'
            $r = Get-Run
            @($script:Fake.Folder.Calls) | Should -Be @('Register:Nightly:6')
            $r.Code | Should -Be 0 -Because $r.All
            $r.Host | Should -Match "Task 'Nightly' Created Successfully"
        }
    }
}
