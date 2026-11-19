#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# Proves the entry-point convention every script uses behaves the same however an RMM launches it.
# The fixture follows the exact pattern shipped scripts use: Invoke-<Name>Main behind the
# KillerScriptsTestMode guard, and Exit-KillerScript for every exit.

BeforeAll {
    $script:Fixture = Join-Path $PSScriptRoot 'fixtures\GuardSample.ps1'
    $script:Ps51    = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $script:Ps51x86 = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'

    function Invoke-Host {
        param([string]$Exe, [string[]]$Arguments, [string]$StdIn)
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = $Exe
        $psi.Arguments = ($Arguments -join ' ')
        $psi.UseShellExecute = $false
        $psi.RedirectStandardOutput = $true
        $psi.RedirectStandardError = $true
        $psi.RedirectStandardInput = $true
        $p = [System.Diagnostics.Process]::Start($psi)
        if ($StdIn) { $p.StandardInput.Write($StdIn) }
        $p.StandardInput.Close()
        $out = $p.StandardOutput.ReadToEnd()
        $null = $p.StandardError.ReadToEnd()
        $p.WaitForExit()
        [pscustomobject]@{ Output = $out; ExitCode = $p.ExitCode }
    }
}

Describe 'Entry-point convention' {
    It 'runs Main and returns its exit code when launched with -File' {
        $r = Invoke-Host $Ps51 @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', "`"$Fixture`"")
        $r.Output | Should -Match 'GUARD-SAMPLE-RAN'
        $r.ExitCode | Should -Be 3
    }

    It 'runs Main when launched with -EncodedCommand' {
        $body = Get-Content -LiteralPath $Fixture -Raw
        $enc = [Convert]::ToBase64String([Text.Encoding]::Unicode.GetBytes($body))
        $r = Invoke-Host $Ps51 @('-NoProfile', '-NonInteractive', '-EncodedCommand', $enc)
        $r.Output | Should -Match 'GUARD-SAMPLE-RAN'
        $r.ExitCode | Should -Be 3
    }

    It 'runs Main when the text is pasted into a console with no interactive host' {
        $body = Get-Content -LiteralPath $Fixture -Raw
        $r = Invoke-Host $Ps51 @('-NoProfile', '-Command', '-') -StdIn ($body + "`r`n")
        $r.Output | Should -Match 'GUARD-SAMPLE-RAN'
    }

    It 'runs Main under the 32-bit Windows PowerShell host' -Skip:(-not (Test-Path (Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'))) {
        $r = Invoke-Host $Ps51x86 @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', "`"$Fixture`"")
        $r.Output | Should -Match 'GUARD-SAMPLE-RAN'
        $r.Output | Should -Match 'BITNESS=32'
        $r.ExitCode | Should -Be 3
    }

    It 'runs Main when invoked with the call operator, as the KillerScripts module does' {
        $cmd = "& '$Fixture'; exit `$LASTEXITCODE"
        $r = Invoke-Host $Ps51 @('-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-Command', "`"$cmd`"")
        $r.Output | Should -Match 'GUARD-SAMPLE-RAN'
        $r.ExitCode | Should -Be 3
    }

    It 'does not run Main when dot-sourced in test mode, and Exit-KillerScript throws instead of exiting' {
        $global:KillerScriptsTestMode = $true
        try {
            . $Fixture
            Get-Command Invoke-GuardSampleMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
            { Invoke-GuardSampleMain } | Should -Throw 'KILLERSCRIPT_EXIT:3'
        } finally {
            Remove-Variable -Name KillerScriptsTestMode -Scope Global -ErrorAction SilentlyContinue
        }
    }
}
