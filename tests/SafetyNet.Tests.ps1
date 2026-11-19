#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# The safety net must stop any unmocked destructive call before it reaches the host.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
}

Describe 'Safety net' {
    It 'blocks Remove-Item outside TestDrive' {
        { Remove-Item -Path 'C:\Windows\Temp\ks-safety-net-probe' -Force } | Should -Throw '*UNMOCKED DESTRUCTIVE CALL*'
    }

    It 'allows Remove-Item inside TestDrive' {
        $p = Join-Path $TestDrive 'probe.txt'
        Set-Content -LiteralPath $p -Value 'x'
        Remove-Item -LiteralPath $p -Force
        Test-Path -LiteralPath $p | Should -BeFalse
    }

    It 'blocks registry writes' {
        { Set-ItemProperty -Path 'HKLM:\SOFTWARE\KillerScriptsProbe' -Name x -Value 1 } | Should -Throw '*UNMOCKED DESTRUCTIVE CALL*'
    }

    It 'blocks account changes' {
        { Disable-LocalUser -Name 'Administrator' } | Should -Throw '*UNMOCKED DESTRUCTIVE CALL*'
    }

    It 'blocks native tools such as manage-bde' {
        { manage-bde -status C: } | Should -Throw '*UNMOCKED DESTRUCTIVE CALL*'
    }

    It 'blocks network calls' {
        { Invoke-WebRequest -Uri 'https://example.invalid' } | Should -Throw '*UNMOCKED DESTRUCTIVE CALL*'
    }

    It 'lets a test override a blocked command with its own mock' {
        Mock Disable-LocalUser { 'mocked' }
        Disable-LocalUser -Name 'Administrator' | Should -Be 'mocked'
    }
}
