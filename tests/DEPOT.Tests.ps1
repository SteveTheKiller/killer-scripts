#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# DEPOT tests. Every system call, download and prompt is mocked. Nothing here touches the real
# machine or the network, and DEPOT.ps1 is never started in a child process.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath DEPOT)
}

AfterAll {
    Disable-KillerTestMode
}

Describe 'Entry-point convention' {
    It 'defines Invoke-DepotMain and Exit-KillerScript without running anything on dot-source' {
        Get-Command Invoke-DepotMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
        { Exit-KillerScript 3010 } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
    }
}

Describe 'Exit codes at the entry point' {
    BeforeEach {
        Mock Start-Sleep { }
        Mock Read-Host { '' }
        Mock Clear-Host { }
        Mock Get-Process { }
        Mock Start-Process { }
        Mock Restart-Computer { }
        $script:DepotScriptPath = $null
    }

    It 'exits 2 when not elevated and there is no script file to relaunch (pasted or -EncodedCommand)' {
        Mock Test-DepotElevated { $false }
        { Invoke-DepotMain } | Should -Throw 'KILLERSCRIPT_EXIT:2'
        Should -Invoke Start-Process -Times 0 -Exactly
    }

    It 'exits 2 when the elevated relaunch fails' {
        Mock Test-DepotElevated { $false }
        Mock Start-Process { throw 'The operation was canceled by the user' }
        $script:DepotScriptPath = 'C:\Scripts\DEPOT.ps1'
        { Invoke-DepotMain } | Should -Throw 'KILLERSCRIPT_EXIT:2'
    }

    It 'exits 3010 when the run ends with a reboot required' {
        Mock Test-DepotElevated { $true }
        foreach ($f in 'Invoke-DepotPreflight', 'Invoke-DepotInteractiveSetup', 'Show-DepotBanner', 'Invoke-DepotWindowsUpdate',
                       'Install-DepotApps', 'Invoke-DepotCleanup', 'Invoke-DepotPrivacy') {
            Mock $f { }
        }
        Mock Invoke-DepotWindowsUpdate { $script:RebootRequired = $true }
        { Invoke-DepotMain } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
    }
}
