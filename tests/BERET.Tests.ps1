#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# BERET tests. Every BitLocker cmdlet, manage-bde, scheduled task, clock and disk call is mocked.
# Nothing here touches a real volume, and BERET.ps1 is never started in a child process.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath BERET)

    # Fake machine state. Tests change these tables, the mocks read them.
    function Initialize-BeretWorld {
        $script:Vols = @{}
        $script:Vols['C:'] = New-BeretVolume 'C:' 'FullyDecrypted' 'None'
        $script:Joined = @{ Domain = $false; Entra = $false; Msa = $false }
        $script:Calls = New-Object System.Collections.ArrayList
        $script:KeyCounter = 0
        $script:Choice = 'N'
        $script:Dsreg = ''
        $script:Escrow = @{}
    }

    function New-BeretVolume {
        param([string]$Mount, [string]$Status, [string]$Method, [object[]]$Protectors = @(), [bool]$AutoUnlock = $false)
        [pscustomobject]@{
            MountPoint = $Mount; VolumeStatus = $Status; EncryptionMethod = $Method
            EncryptionPercentage = $(if ($Status -eq 'FullyDecrypted') { 0 } else { 100 })
            AutoUnlockEnabled = $AutoUnlock; KeyProtector = @($Protectors)
        }
    }

    function New-BeretProtector {
        param([string]$Type, [string]$Id = ('{' + [guid]::NewGuid().ToString().ToUpper() + '}'), [string]$Password = '')
        [pscustomobject]@{ KeyProtectorType = $Type; KeyProtectorId = $Id; RecoveryPassword = $Password }
    }

    function Invoke-BeretRun {
        # Runs Main, capturing every stream as text and the KILLERSCRIPT_EXIT code it throws.
        $script:Code = $null
        $script:Lines = New-Object System.Collections.ArrayList
        try {
            Invoke-BeretMain *>&1 | ForEach-Object { [void]$script:Lines.Add("$_") }
        } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $script:Code = [int]$Matches[1] } else { throw }
        }
        $script:Text = ($script:Lines -join "`n")
    }
}

AfterAll {
    Disable-KillerTestMode
}

Describe 'Entry-point convention' {
    It 'defines Invoke-BeretMain and Exit-KillerScript without running anything on dot-source' {
        Get-Command Invoke-BeretMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
        { Exit-KillerScript 3 } | Should -Throw 'KILLERSCRIPT_EXIT:3'
    }
}
