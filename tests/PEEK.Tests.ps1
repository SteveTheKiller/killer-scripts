#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# PEEK tests. PEEK is a read-only monitor, but every system call here is still mocked or replaced
# by a synthetic fixture: events, processes, Panther logs, compatibility reports and registry state.

Describe 'PEEK entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'PEEK.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
        $script:Source = [IO.File]::ReadAllText($path)
    }

    It 'defines Invoke-PEEKMain and a local Exit-KillerScript' {
        $names = @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-PEEKMain'
        $names | Should -Contain 'Exit-KillerScript'
    }

    It 'ends with the test-mode guard and has no bare exit statement outside Exit-KillerScript' {
        $Source | Should -Match '(?m)^if \(-not \$global:KillerScriptsTestMode\) \{ Invoke-PEEKMain \}\s*$'
        $exits = @($Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.ExitStatementAst] }, $true))
        $exits.Count | Should -Be 1
    }
}

Describe 'PEEK main' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'PEEK')

        function Get-PeekRun {
            # Runs Main and returns the exit code carried by Exit-KillerScript plus the output stream text.
            $code = -1
            $script:RunLines = New-Object System.Collections.ArrayList
            try { Invoke-PEEKMain *>&1 | ForEach-Object { [void]$script:RunLines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($script:RunLines -join "`n") }
        }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $Once = [switch]$true
        $RefreshSeconds = 1
        Mock Test-PEEKElevated { $true }
        Mock Get-PEEKBitLockerState { 'Unavailable' }
        Mock Show-PEEKSnapshot { }
        Mock Write-Footer { }
    }

    It 'loads without running when dot-sourced in test mode' {
        Get-Command Invoke-PEEKMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'exits 2 with a FAIL line and shows nothing when not elevated' {
        Mock Test-PEEKElevated { $false }
        $r = Get-PeekRun
        $r.Code | Should -Be 2
        $r.Text | Should -Match '\[FAIL\].*elevation'
        Should -Invoke Show-PEEKSnapshot -Times 0 -Exactly
    }

    It 'exits 0 with an OK line after a one-shot snapshot' {
        $r = Get-PeekRun
        $r.Code | Should -Be 0
        $r.Text | Should -Match '\[OK\].*exit 0'
        Should -Invoke Show-PEEKSnapshot -Times 1 -Exactly
    }

    It 'exits 1 with a FAIL line when the snapshot throws, never claiming success' {
        Mock Show-PEEKSnapshot { throw 'boom' }
        $r = Get-PeekRun
        $r.Code | Should -Be 1
        $r.Text | Should -Match '\[FAIL\].*boom.*exit 1'
        $r.Text | Should -Not -Match '\[OK\]'
    }

    It 'does not throw when there is no console to clear' {
        Mock Clear-Host { throw 'The handle is invalid.' }
        { Clear-PEEKScreen } | Should -Not -Throw
    }
}
