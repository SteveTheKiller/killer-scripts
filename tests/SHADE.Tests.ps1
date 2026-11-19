#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# SHADE tests. Every registry, service and system call is mocked, so nothing here touches the machine.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath 'SHADE')

    function Get-Exit {
        # Runs Main and returns the code carried by Exit-KillerScript.
        try { Invoke-ShadeMain *>&1 | Out-Null } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { return [int]$Matches[1] }
            throw
        }
        -1
    }

    function Get-Run {
        # Runs Main and returns the exit code plus everything written to the output stream.
        $code = -1
        $script:RunLines = New-Object System.Collections.ArrayList
        try { Invoke-ShadeMain *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        [pscustomobject]@{ Code = $code; Text = ($script:RunLines -join "`n") }
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'SHADE' {
    BeforeEach {
        # Console and system boundaries
        Mock Write-StepUpdate {}
        Mock Write-HLine {}
        Mock Clear-Host {}
        Mock Write-Host {}
        Mock Get-CimInstance { $null }
        Mock Get-ItemProperty { [pscustomobject]@{ DisplayVersion = '23H2' } }
        # Identity boundaries
        Mock Test-IsAdmin { $true }
        # Registry boundaries
        Mock Test-IsSystem { $false }
        Mock Get-LoadedUserSids { @() }
        Mock Test-Path { $true }
        Mock New-Item {} -ParameterFilter { $true }
        Mock Set-ItemProperty {} -ParameterFilter { $true }
        # Service boundaries
        Mock Get-Service { [pscustomobject]@{ Name = $Name } }
        Mock Stop-Service {}
        Mock Set-Service {}
        Mock Invoke-ScDisable {}
        Mock Remove-Item {} -ParameterFilter { $true }
    }

    Context 'entry-point convention' {
        It 'defines Invoke-ShadeMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-ShadeMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Set-ItemProperty -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'exits 0 on a healthy run' {
            Get-Exit | Should -Be 0
        }
    }

    Context 'finding 69: not elevated' {
        It 'exits 2 and touches nothing when not elevated' {
            Mock Test-IsAdmin { $false }
            Get-Exit | Should -Be 2
            Should -Invoke Set-ItemProperty -Times 0 -Exactly
            Should -Invoke Stop-Service -Times 0 -Exactly
        }
    }

    Context 'finding 49: camera and microphone policy' {
        BeforeEach { $script:AppPrivacy = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AppPrivacy' }
        It 'sets LetAppsAccessCamera and LetAppsAccessMicrophone to 0 (user in control)' {
            Get-Exit | Should -Be 0
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter { $Path -eq $AppPrivacy -and $Name -eq 'LetAppsAccessCamera' -and $Value -eq 0 }
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter { $Path -eq $AppPrivacy -and $Name -eq 'LetAppsAccessMicrophone' -and $Value -eq 0 }
        }
        It 'never force-allows (1) camera or microphone for every app' {
            Get-Exit | Out-Null
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter { $Path -eq $AppPrivacy -and $Name -like 'LetAppsAccess*' -and $Value -ne 0 }
        }
    }

    Context 'finding 61: dmwappushservice' {
        It 'does not touch dmwappushservice at all' {
            Get-Exit | Should -Be 0
            Should -Invoke Get-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'dmwappushservice' }
            Should -Invoke Stop-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'dmwappushservice' }
            Should -Invoke Set-Service -Times 0 -Exactly -ParameterFilter { $Name -eq 'dmwappushservice' }
            Should -Invoke Invoke-ScDisable -Times 0 -Exactly -ParameterFilter { $Name -eq 'dmwappushservice' }
        }
        It 'still disables the other telemetry services' {
            Get-Exit | Out-Null
            foreach ($n in 'DiagTrack', 'WerSvc', 'PcaSvc', 'lfsvc') {
                Should -Invoke Set-Service -Times 1 -Exactly -ParameterFilter { $Name -eq $n -and $StartupType -eq 'Disabled' }
            }
        }
    }

    Context 'finding 24: identity, per-user scope and reported outcomes' {
        BeforeEach {
            $script:Ad = 'HKCU:\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo'
            $script:Sid1 = 'S-1-5-21-1111111111-2222222222-3333333333-1001'
            $script:Sid2 = 'S-1-5-21-1111111111-2222222222-3333333333-1002'
        }

        It 'writes per-user values to the executing HKCU when run as a normal admin' {
            Get-Exit | Should -Be 0
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter { $Path -eq $Ad -and $Name -eq 'Enabled' -and $Value -eq 0 }
        }

        It 'as SYSTEM, applies per-user values to every signed-in user hive and never to the SYSTEM HKCU' {
            Mock Test-IsSystem { $true }
            Mock Get-LoadedUserSids { @($Sid1, $Sid2) }
            Get-Exit | Should -Be 0
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter { $Path -eq "Registry::HKEY_USERS\$Sid1\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -and $Name -eq 'Enabled' }
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter { $Path -eq "Registry::HKEY_USERS\$Sid2\Software\Microsoft\Windows\CurrentVersion\AdvertisingInfo" -and $Name -eq 'Enabled' }
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter { $Path -like 'HKCU:*' }
        }

        It 'as SYSTEM, still writes the machine-wide HKLM policy once' {
            Mock Test-IsSystem { $true }
            Mock Get-LoadedUserSids { @($Sid1, $Sid2) }
            Get-Exit | Out-Null
            Should -Invoke Set-ItemProperty -Times 1 -Exactly -ParameterFilter { $Path -eq 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\AdvertisingInfo' -and $Name -eq 'DisabledByGroupPolicy' }
        }

        It 'as SYSTEM with nobody signed in, skips per-user values, says so, and exits 3' {
            Mock Test-IsSystem { $true }
            Mock Get-LoadedUserSids { @() }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'skipped'
            $r.Text | Should -Match 'signed-in'
            $r.Text | Should -Match '\(exit 3\)'
            Should -Invoke Set-ItemProperty -Times 0 -Exactly -ParameterFilter { $Path -like 'HKCU:*' -or $Path -like '*HKEY_USERS*' }
        }

        It 'does not swallow a failed registry write: it is reported and the exit is 3' {
            Mock Set-ItemProperty { throw 'access denied' } -ParameterFilter { $Name -eq 'DisableLocation' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'DisableLocation'
            $r.Text | Should -Match 'access denied'
            $r.Text | Should -Match '\(exit 3\)'
        }

        It 'does not swallow a failed service change: it is reported and the exit is 3' {
            Mock Set-Service { throw 'service locked' } -ParameterFilter { $Name -eq 'DiagTrack' }
            Mock Invoke-ScDisable { 5 } -ParameterFilter { $Name -eq 'DiagTrack' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match 'DiagTrack'
            $r.Text | Should -Match '\(exit 3\)'
        }

        It 'prints a final status line that agrees with exit 0 on a clean run' {
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\] SHADE.*\(exit 0\)'
            $r.Text | Should -Match 'Per-user settings'
        }
    }
}

Describe 'SHADE console guard' {
    BeforeEach {
        # Only the console boundaries throw. Write-StepUpdate and Write-Host stay real so the
        # step lines are observable, exactly as in a redirected RMM or LiveConnect session.
        Mock Get-ShadeConsoleValue { throw 'The handle is invalid.' }
        Mock Set-ShadeConsolePosition { throw 'The handle is invalid.' }
        Mock Set-ShadeConsoleUtf8 { throw 'The handle is invalid.' }
        Mock Clear-Host { throw 'The handle is invalid.' }
        Mock Get-CimInstance { $null }
        Mock Get-ItemProperty { [pscustomobject]@{ DisplayVersion = '23H2' } }
        Mock Test-IsAdmin { $true }
        Mock Test-IsSystem { $false }
        Mock Get-LoadedUserSids { @() }
        Mock Test-Path { $true }
        Mock New-Item {} -ParameterFilter { $true }
        Mock Set-ItemProperty {} -ParameterFilter { $true }
        Mock Get-Service { [pscustomobject]@{ Name = $Name } }
        Mock Stop-Service {}
        Mock Set-Service {}
        Mock Invoke-ScDisable {}
        Mock Remove-Item {} -ParameterFilter { $true }

        function Get-GuardRun {
            $code = -1
            $script:GuardErrors = 0
            $script:GuardLines = New-Object System.Collections.ArrayList
            try {
                Invoke-ShadeMain *>&1 | ForEach-Object {
                    if ($_ -is [System.Management.Automation.ErrorRecord]) { $script:GuardErrors++ }
                    [void]$script:GuardLines.Add([string]$_)
                }
            } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Lines = @($script:GuardLines); Errors = $script:GuardErrors }
        }
    }

    It 'prints every step line, writes no errors and keeps exit 0 and the PASS line' {
        $r = Get-GuardRun
        $r.Errors | Should -Be 0
        $r.Code | Should -Be 0
        $text = $r.Lines -join "`n"
        foreach ($n in 1..8) {
            $tag = '[{0:D2}/08]' -f $n
            $text | Should -Match ([regex]::Escape($tag))
        }
        @($r.Lines | Where-Object { $_ -match '\[SUCCESS\]' }).Count | Should -Be 8
        $text | Should -Not -Match 'handle is invalid'
        $r.Lines[-1] | Should -Match '\[PASS\] SHADE.*\(exit 0\)'
    }

    It 'keeps exit 3 and the FAIL line when a value fails and the console throws' {
        Mock Set-ItemProperty { throw 'access denied' } -ParameterFilter { $Name -eq 'DisableLocation' }
        $r = Get-GuardRun
        $r.Errors | Should -Be 0
        $r.Code | Should -Be 3
        $r.Lines[-1] | Should -Match '\(exit 3\)'
    }

    It 'still exits 2 without errors when not elevated and the console throws' {
        Mock Test-IsAdmin { $false }
        $r = Get-GuardRun
        $r.Code | Should -Be 2
        $r.Lines[-1] | Should -Match '\(exit 2\)'
    }
}