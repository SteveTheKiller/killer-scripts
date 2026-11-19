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
}
