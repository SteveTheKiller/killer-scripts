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

    Context 'finding 74: installed-driver menu' {
        BeforeEach {
            # Add menu 2 (manual IP), driver step C, then protocol 1 (RAW). Drivers sort to Acme, Microsoft, Zeta.
            foreach ($k in @('2','c')) { $script:Keys.Enqueue((New-Key $k)) }
            $script:Keys.Enqueue((New-Key '1'))
        }
        It '<Entry> keeps the suggested driver' -ForEach @(
            @{ Entry = 'an empty entry'; Raw = '' },
            @{ Entry = '0'; Raw = '0' },
            @{ Entry = 'an out-of-range number'; Raw = '99' },
            @{ Entry = 'non-numeric text'; Raw = 'abc' }
        ) {
            foreach ($i in @('10.0.0.5', $Raw, '')) { $script:Inputs.Enqueue($i) }
            Invoke-AddPrinterWizard
            Should -Invoke Add-Printer -Times 1 -Exactly -ParameterFilter { $DriverName -eq 'Acme X100 PCL' }
        }
        It 'picks the driver whose number was entered' {
            foreach ($i in @('10.0.0.5', '3', '')) { $script:Inputs.Enqueue($i) }
            Invoke-AddPrinterWizard
            Should -Invoke Add-Printer -Times 1 -Exactly -ParameterFilter { $DriverName -eq 'Zeta Z1' }
        }
    }

    Context 'finding 21: scan range' {
        It 'limits a /25 to its own 126 hosts' {
            $r = Get-PrintScanRange -IPAddress '192.168.1.130' -PrefixLength 25
            $r.Addresses.Count | Should -Be 126
            $r.Addresses[0]    | Should -Be '192.168.1.129'
            $r.Addresses[-1]   | Should -Be '192.168.1.254'
            $r.Addresses | Should -Not -Contain '192.168.1.5'
        }
        It 'covers all 510 hosts of a /23' {
            $r = Get-PrintScanRange -IPAddress '10.0.5.9' -PrefixLength 23
            $r.Addresses.Count | Should -Be 510
            $r.Addresses[0]    | Should -Be '10.0.4.1'
            $r.Addresses[-1]   | Should -Be '10.0.5.254'
        }
        It 'bounds a huge network to the /24 around the host' {
            $r = Get-PrintScanRange -IPAddress '10.20.30.40' -PrefixLength 16
            $r.Addresses.Count | Should -Be 254
            $r.Addresses[0]    | Should -Be '10.20.30.1'
            $r.Note            | Should -Match '/24'
        }
        It 'refuses a /31 and /32, which have nothing to scan' {
            Get-PrintScanRange -IPAddress '10.0.0.1' -PrefixLength 32 | Should -BeNullOrEmpty
            Get-PrintScanRange -IPAddress '10.0.0.1' -PrefixLength 31 | Should -BeNullOrEmpty
        }
        It 'uses the default-route interface, not the address with the longest prefix' {
            Mock Get-NetRoute {
                @(
                    [pscustomobject]@{ InterfaceIndex = 5; RouteMetric = 0;  DestinationPrefix = '0.0.0.0/0' },
                    [pscustomobject]@{ InterfaceIndex = 9; RouteMetric = 25; DestinationPrefix = '0.0.0.0/0' }
                )
            }
            Mock Get-NetIPInterface { @([pscustomobject]@{ InterfaceIndex = 5; InterfaceMetric = 10 }, [pscustomobject]@{ InterfaceIndex = 9; InterfaceMetric = 10 }) }
            Mock Get-NetIPAddress {
                @(
                    [pscustomobject]@{ InterfaceIndex = 9; IPAddress = '172.16.0.7';  PrefixLength = 30; PrefixOrigin = 'Manual' },
                    [pscustomobject]@{ InterfaceIndex = 5; IPAddress = '192.168.4.20'; PrefixLength = 25; PrefixOrigin = 'Dhcp' }
                )
            }
            $t = Get-PrintScanTarget
            $t.Addresses[0]  | Should -Be '192.168.4.1'
            $t.Addresses[-1] | Should -Be '192.168.4.126'
        }
        It 'returns nothing when there is no default route' {
            Get-PrintScanTarget | Should -BeNullOrEmpty
        }
        It 'hands the scan only the validated range from the wizard' {
            Mock Get-NetRoute { @([pscustomobject]@{ InterfaceIndex = 5; RouteMetric = 0; DestinationPrefix = '0.0.0.0/0' }) }
            Mock Get-NetIPInterface { @([pscustomobject]@{ InterfaceIndex = 5; InterfaceMetric = 10 }) }
            Mock Get-NetIPAddress { @([pscustomobject]@{ InterfaceIndex = 5; IPAddress = '192.168.1.130'; PrefixLength = 25; PrefixOrigin = 'Dhcp' }) }
            $script:Keys.Enqueue((New-Key '1'))
            Invoke-AddPrinterWizard
            Should -Invoke Invoke-PrinterScan -Times 1 -Exactly -ParameterFilter {
                $Addresses.Count -eq 126 -and ($Addresses -notcontains '192.168.1.5') -and ($Addresses -contains '192.168.1.254')
            }
        }
    }

    Context 'finding 41: protocol and port' {
        BeforeEach {
            Mock Test-PrintIppSupported { $true }
            $script:Keys.Enqueue((New-Key '2'))      # add menu: manual IP
            $script:Keys.Enqueue((New-Key 'a'))      # accept suggested driver
            foreach ($i in @('10.0.0.5', '')) { $script:Inputs.Enqueue($i) }
        }
        It 'creates a RAW port on 9100 named for the protocol' {
            $script:Keys.Enqueue((New-Key '1'))
            Invoke-AddPrinterWizard
            Should -Invoke Add-PrinterPort -Times 1 -Exactly -ParameterFilter { $PortNumber -eq 9100 -and $PrinterHostAddress -eq '10.0.0.5' }
            Should -Invoke Add-Printer -Times 1 -Exactly -ParameterFilter { $PortName -like 'IP_10.0.0.5*' }
        }
        It 'never builds a raw TCP port on 631 for IPP, and uses the native IPP URL instead' {
            $script:Keys.Enqueue((New-Key '2'))
            Invoke-AddPrinterWizard
            Should -Invoke Add-PrinterPort -Times 0 -Exactly
            Should -Invoke Add-Printer -Times 1 -Exactly -ParameterFilter { $IppURL -eq 'http://10.0.0.5:631/ipp/print' }
        }
        It 'fails honestly when this Windows build cannot add IPP printers' {
            Mock Test-PrintIppSupported { $false }
            $script:Keys.Enqueue((New-Key '2'))
            Invoke-AddPrinterWizard
            Should -Invoke Add-PrinterPort -Times 0 -Exactly
            Should -Invoke Add-Printer -Times 0 -Exactly
            ($script:HostLines -join "`n") | Should -Match 'IPP'
            $script:PrintFailures | Should -Be 1
        }
        It 'does not reuse a same-IP port that was made for another protocol' {
            Mock Get-PrinterPort { [pscustomobject]@{ Name = 'IP_10.0.0.5'; PrinterHostAddress = '10.0.0.5'; PortNumber = 631 } } -ParameterFilter { $Name -eq 'IP_10.0.0.5' }
            $script:Keys.Enqueue((New-Key '1'))
            Invoke-AddPrinterWizard
            Should -Invoke Add-PrinterPort -Times 1 -Exactly -ParameterFilter { $PortNumber -eq 9100 -and $Name -ne 'IP_10.0.0.5' }
            Should -Invoke Add-Printer -Times 1 -Exactly -ParameterFilter { $PortName -ne 'IP_10.0.0.5' }
        }
        It 'reuses a same-IP port that already matches RAW 9100' {
            Mock Get-PrinterPort { [pscustomobject]@{ Name = 'IP_10.0.0.5'; PrinterHostAddress = '10.0.0.5'; PortNumber = 9100 } } -ParameterFilter { $Name -eq 'IP_10.0.0.5' }
            $script:Keys.Enqueue((New-Key '1'))
            Invoke-AddPrinterWizard
            Should -Invoke Add-PrinterPort -Times 0 -Exactly
            Should -Invoke Add-Printer -Times 1 -Exactly -ParameterFilter { $PortName -eq 'IP_10.0.0.5' }
        }
    }
}
