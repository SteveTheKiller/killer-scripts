#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# PRINT tests. Every printer, port, driver, spooler, network and console call is mocked.

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath 'PRINT')

    function New-Key {
        param([string]$Char, [ConsoleKey]$Key = [ConsoleKey]::NoName)
        New-Object System.ConsoleKeyInfo ([char]$Char), $Key, $false, $false, $false
    }
    function New-EscKey { New-Object System.ConsoleKeyInfo ([char]27), ([ConsoleKey]::Escape), $false, $false, $false }

    function Get-Run {
        # Runs Main and returns the exit code plus everything written to the output stream.
        param([switch]$Silent)
        $code = -1
        $lines = New-Object System.Collections.ArrayList
        try { Invoke-PrintMain -Silent:$Silent *>&1 | ForEach-Object { [void]$lines.Add([string]$_) } } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
        }
        [pscustomobject]@{ Code = $code; Text = ($lines -join "`n") }
    }
}

AfterAll { Disable-KillerTestMode }

Describe 'PRINT' {
    BeforeEach {
        $script:Keys      = New-Object System.Collections.Queue
        $script:Inputs    = New-Object System.Collections.Queue
        $script:KeyCalls  = 0
        $script:PrintFailures = 0; $script:PrintSuccesses = 0
        $script:HostLines = New-Object System.Collections.ArrayList

        # Console boundaries
        Mock Write-Host { [void]$script:HostLines.Add([string]$Object) }
        Mock Write-HLine {}
        Mock Write-Banner {}
        Mock Write-PrintFooter {}
        Mock Start-Sleep {}
        Mock Get-PrintCursorTop { 0 }
        Mock Set-PrintCursor {}
        Mock Test-PrintElevated { $true }
        Mock Test-PrintInteractive { $true }
        Mock Read-PrintKey {
            $script:KeyCalls++
            if ($script:KeyCalls -gt 60) { throw 'RUNAWAY LOOP: Read-PrintKey called more than 60 times' }
            if ($script:Keys.Count -gt 0) { $script:Keys.Dequeue() } else { New-EscKey }
        }
        Mock Read-Input { if ($script:Inputs.Count -gt 0) { $script:Inputs.Dequeue() } else { $null } }

        # Printer, port, driver, spooler and network boundaries
        Mock Get-Printer { @() }
        Mock Remove-Printer {}
        Mock Get-PrinterPort { $null }
        Mock Add-PrinterPort {}
        Mock Add-PrinterDriver {}
        Mock Add-Printer {}
        Mock Get-PrinterModel { 'Acme X100' }
        Mock Invoke-PrinterScan { @() }
        Mock Get-PrinterDriver {
            @(
                [pscustomobject]@{ Name = 'Zeta Z1' },
                [pscustomobject]@{ Name = 'Microsoft IPP Class Driver' },
                [pscustomobject]@{ Name = 'Acme X100 PCL' }
            )
        }
        Mock Get-NetRoute { @() }
        Mock Get-NetIPInterface { @() }
        Mock Get-NetIPAddress { @() }
    }

    Context 'entry-point convention' {
        It 'defines Invoke-PrintMain and does not run it when dot-sourced in test mode' {
            Get-Command Invoke-PrintMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Get-Printer -Times 0 -Exactly
        }
        It 'routes exits through Exit-KillerScript, which throws in test mode' {
            { Exit-KillerScript 7 } | Should -Throw 'KILLERSCRIPT_EXIT:7'
        }
        It 'exits 2 with a FAIL line when not elevated' {
            Mock Test-PrintElevated { $false }
            Mock Write-Warning {}
            $r = Get-Run -Silent
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
        }
        It 'lists printers in -Silent mode and exits 0 with a PASS line' {
            Mock Get-Printer { @([pscustomobject]@{ Name = 'A'; DriverName = 'D'; PortName = 'P' }) }
            $r = Get-Run -Silent
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\].*1 printer.*exit 0'
        }
        It 'exits 1 when the spooler cannot list printers in -Silent mode' {
            Mock Get-Printer { throw 'spooler down' }
            $r = Get-Run -Silent
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\].*spooler down.*exit 1'
        }
        It 'exits 0 when the interactive menu is quit without changes' {
            $script:Keys.Enqueue((New-Key 'q' ([ConsoleKey]::Q)))
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\]'
        }
        It 'exits 1 when the only change attempted failed' {
            Mock Get-Printer { @([pscustomobject]@{ Name = 'HP'; PortName = 'IP_1'; Default = $false }) }
            Mock Show-PrinterList { param([ref]$PrinterRef) $PrinterRef.Value = @([pscustomobject]@{ Name = 'HP'; PortName = 'IP_1' }) }
            Mock Remove-Printer { throw 'access denied' }
            foreach ($k in @((New-Key 'd' ([ConsoleKey]::D)), (New-Key 'y' ([ConsoleKey]::Y)), (New-Key 'q' ([ConsoleKey]::Q)))) { $script:Keys.Enqueue($k) }
            $script:Inputs.Enqueue('1')
            $r = Get-Run
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\].*exit 1'
        }
        It 'exits 3 when one change succeeded and another failed' {
            Mock Show-PrinterList { param([ref]$PrinterRef) $PrinterRef.Value = @([pscustomobject]@{ Name = 'HP'; PortName = 'IP_1' }) }
            $script:RemoveCalls = 0
            Mock Remove-Printer { $script:RemoveCalls++; if ($script:RemoveCalls -eq 2) { throw 'access denied' } }
            foreach ($k in @('d','y','d','y','q')) { $script:Keys.Enqueue((New-Key $k)) }
            $script:Inputs.Enqueue('1'); $script:Inputs.Enqueue('1')
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\].*exit 3'
        }
    }

    Context 'finding 59: no console and no -Silent' {
        It 'exits 2 with a clear error instead of redrawing forever' {
            Mock Test-PrintInteractive { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*-Silent.*exit 2'
            Should -Invoke Read-PrintKey -Times 0 -Exactly
        }
        It 'stops with exit 1 when console input disappears mid-session' {
            Mock Read-PrintKey { $script:KeyCalls++; if ($script:KeyCalls -gt 60) { throw 'RUNAWAY LOOP' }; $null }
            $r = Get-Run
            $r.Code | Should -Be 1
            $r.Text | Should -Match '\[FAIL\].*exit 1'
        }
    }
}
