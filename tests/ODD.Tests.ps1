#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# ODD tests. Every registry, service, CIM, native and console call is mocked.

Describe 'ODD entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'ODD.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    }
    It 'defines Invoke-OddMain and Exit-KillerScript' {
        $names = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-OddMain'
        $names | Should -Contain 'Exit-KillerScript'
    }
    It 'ends with the test-mode guard and keeps only definitions and constants at the top level' {
        $stmts = @($script:Ast.EndBlock.Statements)
        $stmts[-1].Extent.Text | Should -Match 'KillerScriptsTestMode'
        $stmts[-1].Extent.Text | Should -Match 'Invoke-OddMain'
        $bad = @($stmts | Select-Object -SkipLast 1 | Where-Object {
            $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $_ -isnot [System.Management.Automation.Language.AssignmentStatementAst]
        })
        ($bad | ForEach-Object { $_.Extent.Text.Split("`n")[0] }) -join "`n" | Should -BeNullOrEmpty
    }
}

Describe 'ODD' {
    BeforeAll {
        . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
        Initialize-KillerSafetyNet
        Enable-KillerTestMode
        . (Get-KillerScriptPath 'ODD')

        function Get-Run {
            $code = -1
            $lines = New-Object System.Collections.ArrayList
            try { Invoke-OddMain *>&1 | ForEach-Object { [void]$lines.Add([string]$_) } } catch {
                if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $code = [int]$Matches[1] } else { throw }
            }
            [pscustomobject]@{ Code = $code; Text = ($lines -join "`n"); Host = @($script:HostLines) }
        }
        function New-Pnp {
            param([string]$Name, [string]$Id, [string]$Service = 'usbaudio', [int]$Error = 0)
            [pscustomobject]@{ Name = $Name; DeviceID = $Id; Service = $Service; ClassGuid = '{4d36e96c-e325-11ce-bfc1-08002be10318}'; ConfigManagerErrorCode = $Error }
        }
        function New-Ep {
            param([string]$Guid, [string]$Name, [string]$Parent, [int64]$State = 1, [int]$Level = 0)
            [pscustomobject]@{ Guid = $Guid; Name = $Name; ParentId = $Parent; State = $State; Level = $Level; Format = $null; Exclusive = $null }
        }
        # Green device rows only: they carry a status column, the header lines do not.
        function Get-Green { param($Lines) @($Lines | Where-Object { $_ -like 'Green|*' -and $_ -match '\s(Healthy|Disabled|Error \d+)\s' }) }
    }
    AfterAll { Disable-KillerTestMode }

    BeforeEach {
        $script:HostLines = New-Object System.Collections.ArrayList
        $script:Pnp = @()
        $script:EpOut = @()
        $script:EpIn = @()
        $script:CoreOut = [pscustomobject]@{ Ok = $true; Id = $null }
        $script:CoreIn = [pscustomobject]@{ Ok = $true; Id = $null }
        $script:Svc = [pscustomobject]@{ Status = 'Running' }
        Mock Write-Host { [void]$script:HostLines.Add("$ForegroundColor|$Object") }
        Mock Write-HLine {}
        Mock Initialize-OddConsole {}
        Mock Test-OddElevated { $true }
        Mock Get-OddAudioService { $script:Svc }
        Mock Get-OddPnpEntities { $script:Pnp }
        Mock Get-OddDriverRows { @() }
        Mock Get-OddEndpoints { $script:EpOut } -ParameterFilter { $Branch -eq 'Render' }
        Mock Get-OddEndpoints { $script:EpIn } -ParameterFilter { $Branch -eq 'Capture' }
        Mock Get-OddDefaultEndpoint { $script:CoreOut } -ParameterFilter { $Flow -eq 'Render' }
        Mock Get-OddDefaultEndpoint { $script:CoreIn } -ParameterFilter { $Flow -eq 'Capture' }
    }

    Context 'entry-point convention and elevation' {
        It 'does not run when dot-sourced in test mode' {
            Get-Command Invoke-OddMain -CommandType Function | Should -Not -BeNullOrEmpty
            Should -Invoke Get-OddEndpoints -Times 0 -Exactly
        }
        It 'exits 2 with a FAIL line and reads nothing when not elevated' {
            Mock Test-OddElevated { $false }
            $r = Get-Run
            $r.Code | Should -Be 2
            $r.Text | Should -Match '\[FAIL\].*exit 2'
            Should -Invoke Get-OddEndpoints -Times 0 -Exactly
        }
    }

    Context 'finding 17: the default is matched by identity, not by name tokens' {
        It 'highlights exactly one of two USB devices whose names share every token' {
            $alpha = 'USB\VID_1111&PID_0001&MI_00\6&AAAA&0&0000'
            $beta  = 'USB\VID_2222&PID_0002&MI_00\6&BBBB&0&0000'
            $script:Pnp = @((New-Pnp 'USB Audio Device Alpha' $alpha), (New-Pnp 'USB Audio Device Beta' $beta))
            $script:EpOut = @((New-Ep 'aaaa0000-0000-0000-0000-000000000001' 'USB Audio Device Alpha' $alpha), (New-Ep 'bbbb0000-0000-0000-0000-000000000002' 'USB Audio Device Beta' $beta))
            $script:CoreOut = [pscustomobject]@{ Ok = $true; Id = 'bbbb0000-0000-0000-0000-000000000002' }
            $r = Get-Run
            $green = @(Get-Green $r.Host)
            $green.Count | Should -Be 1
            $green[0] | Should -Match 'USB Audio Device Beta'
            ($r.Host | Where-Object { $_ -like 'White|USB Audio Device Alpha*' }).Count | Should -Be 1
        }
        It 'highlights only the device whose identity matches when two devices have the identical name' {
            $one = 'USB\VID_1111&PID_0001&MI_00\6&AAAA&0&0000'
            $two = 'USB\VID_1111&PID_0001&MI_00\6&CCCC&0&0000'
            $script:Pnp = @((New-Pnp 'USB Audio Device' $one), (New-Pnp 'USB Audio Device' $two))
            $script:EpOut = @((New-Ep 'aaaa0000-0000-0000-0000-000000000001' 'USB Audio Device' $one), (New-Ep 'cccc0000-0000-0000-0000-000000000003' 'USB Audio Device' $two))
            $script:CoreOut = [pscustomobject]@{ Ok = $true; Id = 'cccc0000-0000-0000-0000-000000000003' }
            $r = Get-Run
            @(Get-Green $r.Host).Count | Should -Be 1
        }
        It 'highlights one device when two endpoints share one parent and the default is either' {
            $hd = 'HDAUDIO\FUNC_01&VEN_10EC&DEV_0892\5&337FEF56&1&0001'
            $script:Pnp = @((New-Pnp 'High Definition Audio Device' $hd 'HdAudAddService'), (New-Pnp 'Other Definition Audio Device' 'HDAUDIO\FUNC_01&VEN_1002\7&1&2&0001' 'AtiHDAudioService'))
            $script:EpOut = @((New-Ep 'aaaa0000-0000-0000-0000-000000000001' 'Speakers' $hd), (New-Ep 'aaaa0000-0000-0000-0000-000000000002' 'Headphones' $hd))
            $script:CoreOut = [pscustomobject]@{ Ok = $true; Id = 'aaaa0000-0000-0000-0000-000000000002' }
            $r = Get-Run
            $green = @(Get-Green $r.Host)
            $green.Count | Should -Be 1
            $green[0] | Should -Match 'High Definition Audio Device'
            $r.Host -join "`n" | Should -Match 'Default Output'
        }
        It 'does not highlight the output default in the input list' {
            $mic = 'USB\VID_0BDA&PID_4937&MI_00\7&3AB243A&1&0000'
            $script:Pnp = @((New-Pnp 'USB2.0 Microphone' $mic))
            $script:EpOut = @((New-Ep 'aaaa0000-0000-0000-0000-000000000001' 'Speakers' $mic))
            $script:CoreOut = [pscustomobject]@{ Ok = $true; Id = 'aaaa0000-0000-0000-0000-000000000001' }
            $script:CoreIn = [pscustomobject]@{ Ok = $true; Id = $null }
            $r = Get-Run
            # The microphone is listed in both Output and Input sections but the input has no default.
            @(Get-Green $r.Host).Count | Should -BeLessOrEqual 1
        }
        It 'shows not detected and highlights nothing when Windows reports no default endpoint' {
            $dev = 'USB\VID_1111&PID_0001&MI_00\6&AAAA&0&0000'
            $script:Pnp = @((New-Pnp 'USB Audio Device' $dev))
            $script:EpOut = @((New-Ep 'aaaa0000-0000-0000-0000-000000000001' 'USB Audio Device' $dev))
            $r = Get-Run
            @(Get-Green $r.Host).Count | Should -Be 0
            ($r.Host | Where-Object { $_ -like 'Yellow|`[not detected`]' }).Count | Should -BeGreaterThan 0
        }
        It 'labels the default as inferred, and still matches by identity, when Core Audio cannot be queried' {
            $alpha = 'USB\VID_1111&PID_0001&MI_00\6&AAAA&0&0000'
            $beta  = 'USB\VID_2222&PID_0002&MI_00\6&BBBB&0&0000'
            $script:Pnp = @((New-Pnp 'USB Audio Device Alpha' $alpha), (New-Pnp 'USB Audio Device Beta' $beta))
            # State 268435457 has the active bit with high flag bits, State 4 is not present.
            $script:EpOut = @((New-Ep 'aaaa0000-0000-0000-0000-000000000001' 'USB Audio Device Alpha' $alpha 4 9), (New-Ep 'bbbb0000-0000-0000-0000-000000000002' 'USB Audio Device Beta' $beta 268435457 3))
            $script:CoreOut = [pscustomobject]@{ Ok = $false; Id = $null }
            $r = Get-Run
            $r.Host -join "`n" | Should -Match 'USB Audio Device Beta \(inferred\)'
            $green = @(Get-Green $r.Host)
            $green.Count | Should -Be 1
            $green[0] | Should -Match 'Beta'
        }
    }

    Context 'quality and exit codes' {
        It 'reads sample rate, bit depth and exclusive mode from the endpoint' {
            $bytes = New-Object byte[] 40
            [BitConverter]::GetBytes([int]48000).CopyTo($bytes, 12)
            [BitConverter]::GetBytes([int16]24).CopyTo($bytes, 22)
            $q = Get-AudioQualityInfo ([pscustomobject]@{ Format = $bytes; Exclusive = 1 })
            $q.SampleRate | Should -Be '48 kHz'
            $q.BitDepth | Should -Be '24-bit'
            $q.ExclusiveMode | Should -Be 'Allowed'
            (Get-AudioQualityInfo $null).SampleRate | Should -Be 'N/A'
        }
        It 'exits 0 with a PASS line when the service runs and no device is in error' {
            $script:Pnp = @((New-Pnp 'USB Audio Device' 'USB\X\1'), (New-Pnp 'Off Device' 'USB\X\2' 'usbaudio' 22))
            $r = Get-Run
            $r.Code | Should -Be 0
            $r.Text | Should -Match '\[PASS\].*exit 0'
        }
        It 'exits 3 with a WARN line when the audio service is stopped' {
            $script:Svc = [pscustomobject]@{ Status = 'Stopped' }
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\].*not running.*exit 3'
        }
        It 'exits 3 with a WARN line when a listed device is in error' {
            $script:Pnp = @((New-Pnp 'Broken Device' 'USB\X\1' 'usbaudio' 10))
            $r = Get-Run
            $r.Code | Should -Be 3
            $r.Text | Should -Match '\[WARN\].*error.*exit 3'
        }
    }
}
