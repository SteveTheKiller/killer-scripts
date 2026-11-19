

Describe 'URT' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'URT')

        function Get-Run {
            # Runs Main and returns the exit code, the output stream and the captured host text.
            $code = -1
            $lines = New-Object System.Collections.ArrayList
            try { Invoke-UrtMain *>&1 | ForEach-Object { [void]$lines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($lines -join "`n"); HostText = ($script:HostLines -join "`n") }
        }
        function New-In {
            param([string]$Status = 'ok', $Value = $null)
            [pscustomobject]@{ Status = $Status; Value = $Value }
        }
        function New-Cs {
            param([int]$Role, [string]$Domain = 'WORKGROUP')
            [pscustomobject]@{ DomainRole = $Role; Domain = $Domain; PartOfDomain = ($Role -eq 1 -or $Role -eq 3) }
        }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:Answers = New-Object System.Collections.Queue
        $script:Asked = New-Object System.Collections.ArrayList
        $script:HostLines = New-Object System.Collections.ArrayList
        $script:Cs = New-Cs 0
        Mock Write-Host { [void]$script:HostLines.Add([string]$Object) }
        Mock Write-HLine {}
        Mock Write-StepUpdate {}
        Mock Clear-UrtScreen {}
        Mock Test-UrtElevated { $true }
        Mock Test-UrtInteractive { $true }
        Mock Get-UrtComputerSystem { $script:Cs }
        Mock Read-UrtInput {
            [void]$script:Asked.Add($Prompt)
            if ($script:Answers.Count -gt 0) { $script:Answers.Dequeue() } else { New-In 'none' }
        }
        Mock Rename-Computer {}
        Mock Restart-Computer {}
    }

    Context 'entry-point convention' {
        It 'does not run when dot-sourced in test mode' {
            Get-Command Invoke-UrtMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 5 } | Should -Throw 'KILLERSCRIPT_EXIT:5'
        }
    }

    Context 'not elevated and no console' {
        It 'exits 2 with a FAIL line when not elevated and asks nothing' {
            Mock Test-UrtElevated { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
            Should -Invoke Read-UrtInput -Times 0 -Exactly
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'exits 2 without asking anything when there is no console' {
            Mock Test-UrtInteractive { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*console.*exit 2'
            Should -Invoke Read-UrtInput -Times 0 -Exactly
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'exits 2 and renames nothing when the prompt cannot be answered' {
            $r = Get-Run
            $r.Code | Should -Be 2
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
    }

    Context 'finding 27: every DomainRole is classified' {
        It 'classifies all six roles and refuses anything else' {
            Get-UrtRoleKind 0 | Should -Be 'Workgroup'
            Get-UrtRoleKind 2 | Should -Be 'Workgroup'
            Get-UrtRoleKind 1 | Should -Be 'Domain'
            Get-UrtRoleKind 3 | Should -Be 'Domain'
            Get-UrtRoleKind 4 | Should -Be 'DomainController'
            Get-UrtRoleKind 5 | Should -Be 'DomainController'
            Get-UrtRoleKind 9 | Should -Be 'Unknown'
            Get-UrtRoleKind $null | Should -Be 'Unknown'
        }
        It 'refuses role <Role> before any prompt and never calls Rename-Computer' -ForEach @(@{ Role = 4 }, @{ Role = 5 }) {
            $script:Cs = New-Cs $Role 'corp.local'
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*domain controller.*exit 2'
            Should -Invoke Read-UrtInput -Times 0 -Exactly
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'exits 2 when the computer system cannot be read' {
            $script:Cs = $null
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\]'
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'renames a standalone computer (role <Role>) without credentials' -ForEach @(@{ Role = 0 }, @{ Role = 2 }) {
            $script:Cs = New-Cs $Role
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $script:Answers.Enqueue((New-In 'ok' ''))
            $r = Get-Run
            $r.Code | Should -Be 3010
            Should -Invoke Rename-Computer -Times 1 -Exactly -ParameterFilter { $NewName -eq 'NEWPC01' -and -not $DomainCredential }
        }
        It 'renames a domain member (role <Role>) with the typed domain credential' -ForEach @(@{ Role = 1 }, @{ Role = 3 }) {
            $script:Cs = New-Cs $Role 'corp.local'
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $script:Answers.Enqueue((New-In 'ok' 'CORP\admin'))
            $script:Answers.Enqueue((New-In 'ok' (ConvertTo-SecureString 'not-a-real-password' -AsPlainText -Force)))
            $script:Answers.Enqueue((New-In 'ok' ''))
            $r = Get-Run
            $r.Code | Should -Be 3010
            Should -Invoke Rename-Computer -Times 1 -Exactly -ParameterFilter { $NewName -eq 'NEWPC01' -and $DomainCredential.UserName -eq 'CORP\admin' }
        }
        It 'exits 2 and renames nothing when a domain member gives no user name' {
            $script:Cs = New-Cs 1 'corp.local'
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $script:Answers.Enqueue((New-In 'ok' ''))
            $r = Get-Run
            $r.Code | Should -Be 2
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
    }

    Context 'input handling' {
        It 'exits 0 and changes nothing when Esc is pressed at the name' {
            $script:Answers.Enqueue((New-In 'cancel'))
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'exits 0 and changes nothing when no name is typed' {
            $script:Answers.Enqueue((New-In 'ok' '   '))
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
        It 'exits 0 and changes nothing when the name matches the current name' {
            $script:Answers.Enqueue((New-In 'ok' $env:COMPUTERNAME))
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Rename-Computer -Times 0 -Exactly
        }
    }

    Context 'result and restart' {
        It 'exits 3010 with a PASS line when renamed and the restart question is answered with Enter' {
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $script:Answers.Enqueue((New-In 'ok' ''))
            $r = Get-Run
            $r.Code | Should -Be 3010
            $r.Text | Should -Match '\[PASS\].*NEWPC01.*restart.*exit 3010'
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }
        It 'does not restart when the restart question cannot be answered' {
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $r = Get-Run
            $r.Code | Should -Be 3010
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }
        It 'restarts only when RESTART is typed' {
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $script:Answers.Enqueue((New-In 'ok' 'RESTART'))
            $r = Get-Run
            $r.Code | Should -Be 0
            Should -Invoke Restart-Computer -Times 1 -Exactly
        }
        It 'exits 1 with a FAIL line and no further prompt when the rename fails' {
            Mock Rename-Computer { throw 'Access is denied' }
            $script:Answers.Enqueue((New-In 'ok' 'NEWPC01'))
            $r = Get-Run
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\].*Access is denied.*exit 1'
            $script:Asked.Count | Should -Be 1
            Should -Invoke Restart-Computer -Times 0 -Exactly
        }
    }
}
