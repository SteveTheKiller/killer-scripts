#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# BEAM tests. Graph is a fake that lives in this file: every Invoke-RestMethod, download, listener,
# browser launch and packager call is mocked. Nothing here touches the network or a real tenant, and
# BEAM.ps1 is never started in a child process (the generated installer wrapper is, with a fake setup).

BeforeAll {
    . (Join-Path $PSScriptRoot 'TestHelpers.ps1')
    Initialize-KillerSafetyNet
    Enable-KillerTestMode
    . (Get-KillerScriptPath BEAM)
    Add-Type -AssemblyName System.IO.Compression
    Add-Type -AssemblyName System.IO.Compression.FileSystem

    # ---- fake Graph --------------------------------------------------------------------------
    function New-FakeGraph {
        $script:G = @{
            Apps = New-Object System.Collections.ArrayList
            Groups = New-Object System.Collections.ArrayList
            Members = @{}
            Entra = New-Object System.Collections.ArrayList
            Managed = New-Object System.Collections.ArrayList
            Calls = New-Object System.Collections.ArrayList
            Fail = New-Object System.Collections.ArrayList
            FileState = @{}
            Seq = 0
        }
    }

    function Add-FakeDevice {
        # One enrolled device: an Intune managed device plus its Entra device object.
        param([string]$Name, [switch]$NoIntune, [switch]$NoEntra)
        $devId = [guid]::NewGuid().ToString()
        if (-not $NoIntune) {
            [void]$script:G.Managed.Add([pscustomobject]@{ id = "md-$Name"; deviceName = $Name; operatingSystem = 'Windows'; osVersion = '10.0.26100'; azureADDeviceId = $devId })
        }
        if (-not $NoEntra) {
            [void]$script:G.Entra.Add([pscustomobject]@{ id = "ent-$Name"; displayName = $Name; deviceId = $devId })
        }
    }

    function Add-FakeFailure {
        # Mode 'error' writes a non-terminating error (how Invoke-RestMethod fails in practice); 'throw' throws.
        param([string]$Method, [string]$Pattern, [string]$Mode = 'error', [string]$Message = 'HTTP 500 fake failure')
        [void]$script:G.Fail.Add([pscustomobject]@{ Method = $Method; Pattern = $Pattern; Mode = $Mode; Message = $Message })
    }

    function Merge-FakeObject {
        param($Target, $Source)
        foreach ($p in $Source.PSObject.Properties) { $Target | Add-Member -NotePropertyName $p.Name -NotePropertyValue $p.Value -Force }
    }

    function Invoke-FakeGraph {
        param($Uri, $Method, $Body)
        $m = if ($Method) { "$Method".ToUpper() } else { 'GET' }
        $text = $null
        if ($Body -is [byte[]]) { $text = [Text.Encoding]::UTF8.GetString($Body) } elseif ($Body) { $text = "$Body" }
        $obj = $null
        if ($text -and $text.TrimStart().StartsWith('{')) { $obj = $text | ConvertFrom-Json }
        [void]$script:G.Calls.Add([pscustomobject]@{ Method = $m; Uri = "$Uri"; Text = $text; Body = $obj })
        foreach ($f in $script:G.Fail) {
            if ($f.Method -eq $m -and "$Uri" -match $f.Pattern) {
                if ($f.Mode -eq 'throw') { throw $f.Message }
                Write-Error $f.Message
                return
            }
        }
        $u = "$Uri" -replace '^https://graph\.microsoft\.com/(beta|v1\.0)', ''
        if ($u -match '^/deviceAppManagement/mobileApps\?\$top=1') { return [pscustomobject]@{ value = @() } }
        if ($u -match "^/deviceAppManagement/mobileApps\?\`$filter=displayName eq '(.*)'$") {
            $n = $Matches[1] -replace "''", "'"
            return [pscustomobject]@{ value = @($script:G.Apps | Where-Object { $_.displayName -eq $n }) }
        }
        if ($u -match '^/deviceAppManagement/mobileApps$' -and $m -eq 'POST') {
            $script:G.Seq++
            $app = $obj
            $app | Add-Member -NotePropertyName id -NotePropertyValue "app-$($script:G.Seq)" -Force
            $app | Add-Member -NotePropertyName assignments -NotePropertyValue @() -Force
            [void]$script:G.Apps.Add($app)
            return $app
        }
        if ($u -match '^/deviceAppManagement/mobileApps/([^/]+)/microsoft\.graph\.win32LobApp/contentVersions$' -and $m -eq 'POST') {
            $script:G.Seq++
            return [pscustomobject]@{ id = "cv-$($script:G.Seq)" }
        }
        if ($u -match '/contentVersions/([^/]+)/files$' -and $m -eq 'POST') {
            $script:G.Seq++
            $fid = "file-$($script:G.Seq)"
            $script:G.FileState[$fid] = 'azureStorageUriRequestSuccess'
            return [pscustomobject]@{ id = $fid }
        }
        if ($u -match '/files/([^/]+)/commit$' -and $m -eq 'POST') { $script:G.FileState[$Matches[1]] = 'commitFileSuccess'; return }
        if ($u -match '/files/([^/]+)$' -and $m -eq 'GET') {
            return [pscustomobject]@{ id = $Matches[1]; uploadState = $script:G.FileState[$Matches[1]]; azureStorageUri = 'https://fake.blob.invalid/c?sig=1' }
        }
        if ($u -match '^/deviceAppManagement/mobileApps/([^/]+)/assign$' -and $m -eq 'POST') {
            $app = $script:G.Apps | Where-Object { $_.id -eq $Matches[1] } | Select-Object -First 1
            $app | Add-Member -NotePropertyName assignments -NotePropertyValue @($obj.mobileAppAssignments) -Force
            return
        }
        if ($u -match '^/deviceAppManagement/mobileApps/([^/]+)$') {
            $app = $script:G.Apps | Where-Object { $_.id -eq $Matches[1] } | Select-Object -First 1
            if ($m -eq 'PATCH') { Merge-FakeObject $app $obj; return }
            return $app
        }
        if ($u -match "^/groups\?\`$filter=displayName eq '(.*)'$") {
            $n = $Matches[1] -replace "''", "'"
            return [pscustomobject]@{ value = @($script:G.Groups | Where-Object { $_.displayName -eq $n }) }
        }
        if ($u -match '^/groups$' -and $m -eq 'POST') {
            $script:G.Seq++
            $obj | Add-Member -NotePropertyName id -NotePropertyValue "grp-$($script:G.Seq)" -Force
            [void]$script:G.Groups.Add($obj)
            $script:G.Members[$obj.id] = New-Object System.Collections.ArrayList
            return $obj
        }
        if ($u -match '^/groups/([^/]+)/members/\$ref$' -and $m -eq 'POST') {
            $mid = ($obj.'@odata.id' -split '/')[-1]
            if ($script:G.Members[$Matches[1]] -contains $mid) { throw 'One or more added object references already exist' }
            [void]$script:G.Members[$Matches[1]].Add($mid)
            return
        }
        if ($u -match '^/groups/([^/]+)/members' -and $m -eq 'GET') {
            return [pscustomobject]@{ value = @($script:G.Members[$Matches[1]] | ForEach-Object { [pscustomobject]@{ id = $_ } }) }
        }
        if ($u -match "^/devices\?\`$filter=(displayName|deviceId) eq '(.*)'$") {
            $prop = $Matches[1]; $val = $Matches[2] -replace "''", "'"
            return [pscustomobject]@{ value = @($script:G.Entra | Where-Object { $_.$prop -eq $val }) }
        }
        if ($u -match "^/deviceManagement/managedDevices\?\`$filter=deviceName eq '(.*)'$") {
            $val = $Matches[1] -replace "''", "'"
            return [pscustomobject]@{ value = @($script:G.Managed | Where-Object { $_.deviceName -eq $val }) }
        }
        if ($u -match '^/deviceManagement/managedDevices/([^/]+)/syncDevice$') { return }
        throw "FAKE GRAPH: unhandled $m $Uri"
    }

    function Get-FakeCalls {
        param([string]$Method, [string]$Pattern)
        @($script:G.Calls | Where-Object { (-not $Method -or $_.Method -eq $Method) -and (-not $Pattern -or $_.Uri -match $Pattern) })
    }

    # ---- fake packager: writes a real .intunewin zip, like IntuneWinAppUtil does ----------------
    function New-FakeIntunewin {
        param([string]$SourceDir, [string]$SetupFile, [string]$OutDir)
        $name = [IO.Path]::GetFileNameWithoutExtension($SetupFile) + '.intunewin'
        $path = Join-Path $OutDir $name
        $xml = '<ApplicationInfo><UnencryptedContentSize>2048</UnencryptedContentSize><EncryptionInfo>' +
               '<EncryptionKey>ek</EncryptionKey><MacKey>mk</MacKey><InitializationVector>iv</InitializationVector><Mac>mac</Mac>' +
               '<ProfileIdentifier>ProfileVersion1</ProfileIdentifier><FileDigest>fd</FileDigest><FileDigestAlgorithm>SHA256</FileDigestAlgorithm>' +
               '</EncryptionInfo></ApplicationInfo>'
        $zip = [System.IO.Compression.ZipFile]::Open($path, 'Create')
        try {
            foreach ($e in @(@('IntuneWinPackage/Metadata/Detection.xml', $xml), @('IntuneWinPackage/Contents/IntunePackage.intunewin', ('X' * 64)))) {
                $entry = $zip.CreateEntry($e[0])
                $sw = New-Object IO.StreamWriter($entry.Open())
                $sw.Write($e[1]); $sw.Dispose()
            }
        } finally { $zip.Dispose() }
    }

    function New-FakeInstaller {
        param([string]$Name = 'setup.exe')
        $dir = Join-Path $TestDrive 'source'
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
        $p = Join-Path $dir $Name
        [IO.File]::WriteAllBytes($p, [byte[]](0x4D, 0x5A, 0x90, 0x00, 1, 2, 3, 4))
        $p
    }

    function Get-Decoded {
        param([string]$Base64, [string]$Encoding = 'UTF8')
        [Text.Encoding]::$Encoding.GetString([Convert]::FromBase64String($Base64))
    }

    function Get-AppInstallScript {
        # The PowerShell the SYSTEM install command would run, decoded from the app's installCommandLine.
        param($App)
        $b64 = ($App.installCommandLine -split '-EncodedCommand ')[-1].Trim()
        Get-Decoded $b64 'Unicode'
    }

    function Get-AppDetectScript {
        param($App)
        Get-Decoded $App.rules[0].scriptContent 'UTF8'
    }

    function Get-MarkerPathFromScript {
        param([string]$Script)
        ([regex]::Match($Script, "C:\\ProgramData\\BEAM\\[^'\s]+\.marker")).Value
    }

    function Invoke-BeamRun {
        # Runs Main, capturing every stream as text and the KILLERSCRIPT_EXIT code it throws.
        $script:Code = $null
        $script:Lines = New-Object System.Collections.ArrayList
        try {
            Invoke-BeamMain *>&1 | ForEach-Object { [void]$script:Lines.Add("$_") }
        } catch {
            if ($_.Exception.Message -match 'KILLERSCRIPT_EXIT:(\d+)') { $script:Code = [int]$Matches[1] } else { throw }
        }
        $script:Text = ($script:Lines -join "`n")
    }

    function Invoke-Deploy {
        # Calls the publishing function with sensible defaults for one local installer.
        param([string]$AppName = 'Contoso Agent', [string[]]$Targets = @('PC01'), [string]$RunId = '20261118120000-aaaa1111',
              [string]$InstallArgs = '/S', [bool]$IsMsi = $false, [string]$LocalPath, [string]$Url, [string]$SetupFile = 'setup.exe')
        if (-not $LocalPath -and -not $Url) { $LocalPath = New-FakeInstaller $SetupFile }
        Invoke-IntuneInstall -Targets $Targets -Url $Url -LocalPath $LocalPath -SetupFile $SetupFile -IsMsi $IsMsi `
            -InstallArgs $InstallArgs -AppName $AppName -Publisher 'BEAM' -TenantId '11111111-2222-3333-4444-555555555555' -RunId $RunId
    }
}

AfterAll {
    Disable-KillerTestMode
}

Describe 'BEAM' {
BeforeEach {
    New-FakeGraph
    $script:Log = New-Object System.Collections.ArrayList
    $script:IsInteractive = $false
    $script:BeamToolPath = 'C:\Fake\IntuneWinAppUtil.exe'
    $script:Reads = New-Object System.Collections.Queue
    $script:PackagerCalls = New-Object System.Collections.ArrayList
    $script:Uploads = New-Object System.Collections.ArrayList
    Mock Write-Host { }
    Mock Write-HLine { }
    Mock Clear-Host { }
    Mock Start-Sleep { }
    Mock Write-StepUpdate { [void]$script:Log.Add("[$Status] $Message") }
    Mock Write-Detail { [void]$script:Log.Add("<$Label> $Value ($Status)") }
    $script:WorkRoot = Join-Path $TestDrive ('w' + [guid]::NewGuid().ToString('N').Substring(0, 8))
    Mock Get-BeamWorkRoot { $script:WorkRoot }
    Mock Invoke-RestMethod { Invoke-FakeGraph -Uri $Uri -Method $Method -Body $Body }
    Mock Get-GraphToken { 'fake-token' }
    Mock Initialize-BeamTool { 'C:\Fake\IntuneWinAppUtil.exe' }
    Mock Send-AzureStorageBlock { [void]$script:Uploads.Add($SasUri) }
    Mock Invoke-BeamPackager {
        [void]$script:PackagerCalls.Add([pscustomobject]@{ SourceDir = $SourceDir; OutDir = $OutDir; SetupFile = $SetupFile; Files = @(Get-ChildItem -LiteralPath $SourceDir -File | ForEach-Object Name) })
        New-FakeIntunewin -SourceDir $SourceDir -SetupFile $SetupFile -OutDir $OutDir
    }
    Mock Read-Host { if ($script:Reads.Count -gt 0) { $script:Reads.Dequeue() } else { '' } }
}

Context 'Entry-point convention' {
    It 'defines Invoke-BeamMain and Exit-KillerScript without running anything on dot-source' {
        Get-Command Invoke-BeamMain -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
        Get-Command Exit-KillerScript -ErrorAction SilentlyContinue | Should -Not -BeNullOrEmpty
    }

    It 'Exit-KillerScript throws KILLERSCRIPT_EXIT in test mode instead of exiting' {
        { Exit-KillerScript 3010 } | Should -Throw 'KILLERSCRIPT_EXIT:3010'
    }
}

Context 'Exit codes at the entry point' {
    BeforeEach {
        foreach ($r in 'https://files.invalid/setup.exe', '', '', '/S', 'PC01', 'admin@contoso.invalid') { $script:Reads.Enqueue($r) }
        Mock Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](0x4D, 0x5A, 0x90, 0x00, 1, 2, 3, 4)) }
        Add-FakeDevice 'PC01'
        Mock Invoke-RestMethod {
            if ("$Uri" -match 'openid-configuration') { return [pscustomobject]@{ issuer = 'https://sts.windows.net/11111111-2222-3333-4444-555555555555/' } }
            Invoke-FakeGraph -Uri $Uri -Method $Method -Body $Body
        }
    }

    It 'exits 0 and says so when every target is queued' {
        Invoke-BeamRun
        $script:Code | Should -Be 0
        $script:Text | Should -Match 'BEAM RESULT: exit code 0 -'
    }

    It 'exits 2 when there is no console to read the inputs from' {
        Mock Read-Host { throw 'Read-Host : Windows PowerShell is in NonInteractive mode.' }
        Invoke-BeamRun
        $script:Code | Should -Be 2
        $script:Text | Should -Match 'BEAM RESULT: exit code 2 -'
        Should -Invoke Invoke-RestMethod -Times 0 -Exactly
    }

    It 'exits 2 when no hostname target is left after dropping IPs' {
        $script:Reads.Clear()
        foreach ($r in 'https://files.invalid/setup.exe', '', '', '/S', '10.0.0.5', 'admin@contoso.invalid') { $script:Reads.Enqueue($r) }
        Invoke-BeamRun
        $script:Code | Should -Be 2
        $script:Text | Should -Match 'BEAM RESULT: exit code 2 -'
    }

    It 'exits 1 when authentication fails' {
        Mock Get-GraphToken { $null }
        Invoke-BeamRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BEAM RESULT: exit code 1 -'
    }
}

Context 'Finding 03b: the install wrapper stamps the marker only on an accepted result' {
    BeforeAll {
        function Invoke-InstallWrapper {
            # Runs the generated SYSTEM install command in a child Windows PowerShell with a fake setup.cmd
            # that exits with $SetupExit. The marker folder is redirected into TestDrive.
            param($App, [int]$SetupExit)
            $dir = Join-Path $TestDrive ('run' + [guid]::NewGuid().ToString('N'))
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            Set-Content -LiteralPath (Join-Path $dir 'setup.cmd') -Value "@exit /b $SetupExit" -Encoding ASCII
            $orig = Get-AppInstallScript $App
            $pd = Join-Path $dir 'ProgramData'
            $marker = (Get-MarkerPathFromScript $orig).Replace('C:\ProgramData\BEAM', (Join-Path $pd 'BEAM'))
            $scr = $orig.Replace('C:\ProgramData\BEAM', (Join-Path $pd 'BEAM'))
            $file = Join-Path $dir 'wrap.ps1'
            Set-Content -LiteralPath $file -Value $scr -Encoding ASCII
            $psi = New-Object System.Diagnostics.ProcessStartInfo
            $psi.FileName = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
            $psi.Arguments = "-NoProfile -NonInteractive -ExecutionPolicy Bypass -File `"$file`""
            $psi.WorkingDirectory = $dir
            $psi.UseShellExecute = $false; $psi.RedirectStandardOutput = $true; $psi.RedirectStandardError = $true
            $p = [System.Diagnostics.Process]::Start($psi)
            $null = $p.StandardOutput.ReadToEnd(); $null = $p.StandardError.ReadToEnd()
            $p.WaitForExit()
            $text = $null
            if (Test-Path -LiteralPath $marker) { $text = (Get-Content -LiteralPath $marker -Raw).Trim() }
            [pscustomobject]@{ ExitCode = $p.ExitCode; MarkerExists = (Test-Path -LiteralPath $marker); MarkerText = $text }
        }
    }

    BeforeEach {
        Invoke-Deploy -RunId 'RUN-ID-0001' -SetupFile 'setup.cmd' | Out-Null
        $script:App = $script:G.Apps[0]
    }

    It 'does not stamp the marker when the installer fails, and returns the installer exit code' {
        $r = Invoke-InstallWrapper -App $script:App -SetupExit 5
        $r.MarkerExists | Should -BeFalse
        $r.ExitCode | Should -Be 5
    }

    It 'does not stamp the marker for the retry code 1618' {
        $r = Invoke-InstallWrapper -App $script:App -SetupExit 1618
        $r.MarkerExists | Should -BeFalse
        $r.ExitCode | Should -Be 1618
    }

    It 'stamps this run ID and exits 0 when the installer succeeds' {
        $r = Invoke-InstallWrapper -App $script:App -SetupExit 0
        $r.MarkerText | Should -Be 'RUN-ID-0001'
        $r.ExitCode | Should -Be 0
    }

    It 'stamps the marker and keeps 3010 when the installer needs a reboot' {
        $r = Invoke-InstallWrapper -App $script:App -SetupExit 3010
        $r.MarkerText | Should -Be 'RUN-ID-0001'
        $r.ExitCode | Should -Be 3010
    }

    It 'treats 1641 (restart initiated) as a success that needs a reboot, reported as 3010' {
        $r = Invoke-InstallWrapper -App $script:App -SetupExit 1641
        $r.MarkerText | Should -Be 'RUN-ID-0001'
        $r.ExitCode | Should -Be 3010
    }
}
}
