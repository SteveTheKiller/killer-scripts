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
            NextFileState = 'azureStorageUriRequestSuccess'
            CommitState = 'commitFileSuccess'
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

    function Get-FakeFilterValue {
        # Graph rejects an OData string literal with an unescaped single quote.
        param([string]$Raw)
        if ($Raw -match "(?<!')'(?!')") { throw 'HTTP 400 Bad Request: invalid filter' }
        $Raw -replace "''", "'"
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
                if ($f.Mode -eq 'noop') { return }
                Write-Error $f.Message
                return
            }
        }
        $u = "$Uri" -replace '^https://graph\.microsoft\.com/(beta|v1\.0)', ''
        if ($u -match '^/deviceAppManagement/mobileApps\?\$top=1') { return [pscustomobject]@{ value = @() } }
        if ($u -match "^/deviceAppManagement/mobileApps\?\`$filter=displayName eq '(.*)'$") {
            $n = Get-FakeFilterValue $Matches[1]
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
            $script:G.FileState[$fid] = $script:G.NextFileState
            return [pscustomobject]@{ id = $fid }
        }
        if ($u -match '/files/([^/]+)/commit$' -and $m -eq 'POST') { $script:G.FileState[$Matches[1]] = $script:G.CommitState; return }
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
            $n = Get-FakeFilterValue $Matches[1]
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
            $prop = $Matches[1]; $val = Get-FakeFilterValue $Matches[2]
            return [pscustomobject]@{ value = @($script:G.Entra | Where-Object { $_.$prop -eq $val }) }
        }
        if ($u -match "^/deviceManagement/managedDevices\?\`$filter=deviceName eq '(.*)'$") {
            $val = Get-FakeFilterValue $Matches[1]
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
              [string]$InstallArgs = '/S', [bool]$IsMsi = $false, [string]$LocalPath, [string]$Url, [string]$SetupFile = 'setup.exe', [switch]$KeepWork)
        if (-not $KeepWork) { $script:WorkRoot = Join-Path $TestDrive ('w' + [guid]::NewGuid().ToString('N').Substring(0, 8)) }
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
    $script:Clock = [DateTime]::UtcNow
    Mock Get-BeamUtcNow { $script:Clock = $script:Clock.AddSeconds(5); $script:Clock }
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

Context 'Finding 03a: each app has its own detection marker' {
    It 'gives different apps different marker files, used by both the install command and detection' {
        Invoke-Deploy -AppName 'Alpha Agent' | Out-Null
        Invoke-Deploy -AppName 'Beta Agent' | Out-Null
        $a = $script:G.Apps | Where-Object { $_.displayName -eq 'Alpha Agent (BEAM)' }
        $b = $script:G.Apps | Where-Object { $_.displayName -eq 'Beta Agent (BEAM)' }
        $ma = Get-MarkerPathFromScript (Get-AppDetectScript $a)
        $mb = Get-MarkerPathFromScript (Get-AppDetectScript $b)
        $ma | Should -Not -BeNullOrEmpty
        $mb | Should -Not -BeNullOrEmpty
        $ma | Should -Not -Be $mb
        $ma | Should -Not -Be 'C:\ProgramData\BEAM\BEAM.marker'
        (Get-MarkerPathFromScript (Get-AppInstallScript $a)) | Should -Be $ma
        (Get-MarkerPathFromScript (Get-AppInstallScript $b)) | Should -Be $mb
    }

    It 'keeps names that sanitize to the same text apart' {
        Invoke-Deploy -AppName 'Tool:One' | Out-Null
        Invoke-Deploy -AppName 'Tool?One' | Out-Null
        $a = $script:G.Apps | Where-Object { $_.displayName -eq 'Tool:One (BEAM)' }
        $b = $script:G.Apps | Where-Object { $_.displayName -eq 'Tool?One (BEAM)' }
        (Get-MarkerPathFromScript (Get-AppDetectScript $a)) | Should -Not -Be (Get-MarkerPathFromScript (Get-AppDetectScript $b))
    }

    It 'reuses the same marker file when the same app is deployed again, with the new run ID' {
        Invoke-Deploy -AppName 'Alpha Agent' -RunId 'RUN-1' | Out-Null
        $first = Get-MarkerPathFromScript (Get-AppDetectScript $script:G.Apps[0])
        Invoke-Deploy -AppName 'Alpha Agent' -RunId 'RUN-2' | Out-Null
        $script:G.Apps.Count | Should -Be 1
        $detect = Get-AppDetectScript $script:G.Apps[0]
        (Get-MarkerPathFromScript $detect) | Should -Be $first
        $detect | Should -Match 'RUN-2'
    }
}

Context 'Finding 50: a failed Graph call is a failure, not an OK' {
    BeforeEach { Add-FakeDevice 'PC01' }

    It 'reports failure and never prints OK when committing the app version fails' {
        Add-FakeFailure -Method PATCH -Pattern 'mobileApps/app-\d+$'
        $r = Invoke-Deploy
        $r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Not -Match 'OK\] App version committed'
        $script:G.Apps[0].PSObject.Properties.Name | Should -Not -Contain 'committedContentVersion'
    }

    It 'reports failure and never marks anything assigned when the assignment call fails' {
        Add-FakeFailure -Method POST -Pattern '/assign$'
        $r = Invoke-Deploy
        $r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Not -Match 'OK\] App assigned'
    }

    It 'reports failure when the file commit call fails' {
        Add-FakeFailure -Method POST -Pattern '/commit$'
        $r = Invoke-Deploy
        $r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Not -Match 'OK\] File committed'
    }

    It 'exits 1 from the entry point and does not claim any target is queued when assignment fails' {
        foreach ($x in 'https://files.invalid/setup.exe', '', '', '/S', 'PC01', 'admin@contoso.invalid') { $script:Reads.Enqueue($x) }
        Mock Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](0x4D, 0x5A, 0x90, 0x00, 1, 2, 3, 4)) }
        Mock Invoke-RestMethod {
            if ("$Uri" -match 'openid-configuration') { return [pscustomobject]@{ issuer = 'https://sts.windows.net/11111111-2222-3333-4444-555555555555/' } }
            Invoke-FakeGraph -Uri $Uri -Method $Method -Body $Body
        }
        Add-FakeFailure -Method POST -Pattern '/assign$'
        Invoke-BeamRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BEAM RESULT: exit code 1 -'
    }
}

Context 'Finding 64: a failed or timed-out Graph wait is reported, not fatal' {
    BeforeEach { Add-FakeDevice 'PC01' }

    It 'returns failure instead of throwing when Azure Storage URI allocation fails' {
        $script:G.NextFileState = 'azureStorageUriRequestFailed'
        $r = $null
        { $script:r = Invoke-Deploy } | Should -Not -Throw
        $script:r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Match 'FAIL\].*azureStorageUriRequestFailed'
    }

    It 'returns failure instead of throwing when the commit never completes before the deadline' {
        $script:G.CommitState = 'commitFilePending'
        { $script:r = Invoke-Deploy } | Should -Not -Throw
        $script:r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Match 'FAIL\].*Timeout'
    }

    It 'returns failure instead of throwing when polling the file state gets an HTTP error' {
        Add-FakeFailure -Method GET -Pattern '/files/file-\d+$' -Mode throw -Message 'The remote server returned an error: (503)'
        { $script:r = Invoke-Deploy } | Should -Not -Throw
        $script:r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Match 'FAIL\].*503'
    }

    It 'still ends with a per-target FAIL, a summary line and exit 1 from the entry point' {
        foreach ($x in 'https://files.invalid/setup.exe', '', '', '/S', 'PC01', 'admin@contoso.invalid') { $script:Reads.Enqueue($x) }
        Mock Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](0x4D, 0x5A, 0x90, 0x00, 1, 2, 3, 4)) }
        Mock Invoke-RestMethod {
            if ("$Uri" -match 'openid-configuration') { return [pscustomobject]@{ issuer = 'https://sts.windows.net/11111111-2222-3333-4444-555555555555/' } }
            Invoke-FakeGraph -Uri $Uri -Method $Method -Body $Body
        }
        $script:G.NextFileState = 'azureStorageUriRequestFailed'
        Invoke-BeamRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BEAM RESULT: exit code 1 -'
        ($script:Log -join "`n") | Should -Match '\[FAIL\] PC01'
    }
}

Context 'Finding 46: every run packages from its own staging folder' {
    BeforeEach { Add-FakeDevice 'PC01' }

    It 'packages only the staged installer, never a stale file left by an earlier run' {
        $legacy = Join-Path $script:WorkRoot 'IntuneWork\src\Contoso Agent'
        New-Item -ItemType Directory -Path $legacy -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $legacy 'old-installer.exe') -Value 'MZ stale'
        Invoke-Deploy -RunId 'RUN-AAAA' -KeepWork | Out-Null
        $call = $script:PackagerCalls[0]
        $call.Files | Should -Be @('setup.exe')
        $call.SourceDir | Should -Match 'RUN-AAAA'
        $call.OutDir | Should -Match 'RUN-AAAA'
    }

    It 'gives two runs whose app names sanitize to the same text separate folders' {
        Invoke-Deploy -AppName 'Tool:One' -RunId 'RUN-1' -KeepWork | Out-Null
        Invoke-Deploy -AppName 'Tool?One' -RunId 'RUN-2' -KeepWork | Out-Null
        $script:PackagerCalls.Count | Should -Be 2
        $script:PackagerCalls[0].SourceDir | Should -Not -Be $script:PackagerCalls[1].SourceDir
        $script:PackagerCalls[0].OutDir | Should -Not -Be $script:PackagerCalls[1].OutDir
    }

    It 'refuses to reuse a staging folder that already exists for the same run ID' {
        $taken = Join-Path $script:WorkRoot 'IntuneWork\run-RUN-TAKEN'
        New-Item -ItemType Directory -Path $taken -Force | Out-Null
        $r = Invoke-Deploy -RunId 'RUN-TAKEN' -KeepWork
        $r.Published | Should -BeFalse
        $script:PackagerCalls.Count | Should -Be 0
        Test-Path -LiteralPath $taken | Should -BeTrue
    }

    It 'removes its own staging folder when the run ends, and nothing else' {
        $other = Join-Path $script:WorkRoot 'IntuneWork\run-OTHER'
        New-Item -ItemType Directory -Path $other -Force | Out-Null
        Invoke-Deploy -RunId 'RUN-CLEAN' -KeepWork | Out-Null
        Test-Path -LiteralPath (Join-Path $script:WorkRoot 'IntuneWork\run-RUN-CLEAN') | Should -BeFalse
        Test-Path -LiteralPath $other | Should -BeTrue
    }

    It 'rejects a staging manifest with an extra or missing file' {
        $d = Join-Path $TestDrive 'manifest'
        New-Item -ItemType Directory -Path $d -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $d 'setup.exe') -Value 'MZ'
        Test-BeamStagingManifest -Dir $d -SetupFile 'setup.exe' | Should -BeTrue
        Set-Content -LiteralPath (Join-Path $d 'extra.dll') -Value 'x'
        Test-BeamStagingManifest -Dir $d -SetupFile 'setup.exe' | Should -BeFalse
        Remove-Item -LiteralPath (Join-Path $d 'extra.dll') -Force
        Remove-Item -LiteralPath (Join-Path $d 'setup.exe') -Force
        Test-BeamStagingManifest -Dir $d -SetupFile 'setup.exe' | Should -BeFalse
    }

    It 'reports a corrupt package as a failed run, not a crash' {
        Mock Invoke-BeamPackager { Set-Content -LiteralPath (Join-Path $OutDir 'bad.intunewin') -Value 'not a zip' }
        $r = $null
        { $script:r = Invoke-Deploy } | Should -Not -Throw
        $script:r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Match 'FAIL\].*Unexpected error'
    }
}

Context 'Finding 05: a rerun keeps the live deployment coherent until the new content is validated' {
    BeforeEach {
        Add-FakeDevice 'PC01'
        Invoke-Deploy -RunId 'RUN-1' | Out-Null
        $script:Live = $script:G.Apps[0]
        $script:Before = @{ Cmd = $script:Live.installCommandLine; Rules = ($script:Live.rules | ConvertTo-Json -Depth 5 -Compress); Version = $script:Live.committedContentVersion }
        $script:G.Calls.Clear()
        $script:Log.Clear()
    }

    It 'sets up the first run as a committed, detectable app' {
        $script:Before.Version | Should -Not -BeNullOrEmpty
        (Get-AppDetectScript $script:Live) | Should -Match 'RUN-1'
    }

    It 'leaves the live command, detection and content untouched when the upload fails' {
        Mock Send-AzureStorageBlock { throw 'storage unreachable' }
        $r = Invoke-Deploy -RunId 'RUN-2'
        $r.Published | Should -BeFalse
        $script:Live.installCommandLine | Should -Be $script:Before.Cmd
        ($script:Live.rules | ConvertTo-Json -Depth 5 -Compress) | Should -Be $script:Before.Rules
        $script:Live.committedContentVersion | Should -Be $script:Before.Version
        (Get-FakeCalls -Method PATCH -Pattern 'mobileApps/app-\d+$').Count | Should -Be 0
    }

    It 'leaves the live deployment untouched when the file commit fails' {
        $script:G.CommitState = 'commitFileFailed'
        $r = Invoke-Deploy -RunId 'RUN-2'
        $r.Published | Should -BeFalse
        $script:Live.installCommandLine | Should -Be $script:Before.Cmd
        ($script:Live.rules | ConvertTo-Json -Depth 5 -Compress) | Should -Be $script:Before.Rules
    }

    It 'leaves the live deployment untouched and reports failure when the publishing patch is rejected' {
        Add-FakeFailure -Method PATCH -Pattern 'mobileApps/app-\d+$'
        $r = Invoke-Deploy -RunId 'RUN-2'
        $r.Published | Should -BeFalse
        $script:Live.installCommandLine | Should -Be $script:Before.Cmd
        ($script:Live.rules | ConvertTo-Json -Depth 5 -Compress) | Should -Be $script:Before.Rules
        $script:Live.committedContentVersion | Should -Be $script:Before.Version
    }

    It 'does not report published when the patch is accepted but did not take effect' {
        Add-FakeFailure -Method PATCH -Pattern 'mobileApps/app-\d+$' -Mode noop
        $r = Invoke-Deploy -RunId 'RUN-2'
        $r.Published | Should -BeFalse
        ($script:Log -join "`n") | Should -Not -Match 'OK\] App version committed'
    }

    It 'publishes command, detection and content together in one patch after the file is committed' {
        $r = Invoke-Deploy -RunId 'RUN-2'
        $r.Published | Should -BeTrue
        $patches = @(Get-FakeCalls -Method PATCH -Pattern 'mobileApps/app-\d+$')
        $patches.Count | Should -Be 1
        $commitIdx = [array]::IndexOf(@($script:G.Calls | ForEach-Object { $_.Uri }), (Get-FakeCalls -Method POST -Pattern '/commit$')[0].Uri)
        $patchIdx = [array]::IndexOf(@($script:G.Calls), $patches[0])
        $patchIdx | Should -BeGreaterThan $commitIdx
        $script:Live.committedContentVersion | Should -Not -Be $script:Before.Version
        (Get-AppDetectScript $script:Live) | Should -Match 'RUN-2'
        (Get-AppInstallScript $script:Live) | Should -Match 'RUN-2'
    }

    It 'reuses the app instead of creating a duplicate when its name contains an apostrophe' {
        Invoke-Deploy -AppName "O'Neil Tool" -RunId 'RUN-A' | Out-Null
        Invoke-Deploy -AppName "O'Neil Tool" -RunId 'RUN-B' | Out-Null
        @($script:G.Apps | Where-Object { $_.displayName -eq "O'Neil Tool (BEAM)" }).Count | Should -Be 1
    }

    It 'stops instead of creating a duplicate app when the existing-app lookup fails' {
        Add-FakeFailure -Method GET -Pattern 'mobileApps\?\$filter' -Mode throw -Message 'HTTP 429 Too Many Requests'
        $r = Invoke-Deploy -RunId 'RUN-2'
        $r.Published | Should -BeFalse
        $script:G.Apps.Count | Should -Be 1
        ($script:Log -join "`n") | Should -Match 'FAIL\].*429'
    }
}

Context 'Finding 04: each run targets its own group and only verified devices are queued' {
    BeforeEach { $script:Log.Clear() }

    It 'never inherits the members of an older group that has the same base name' {
        Add-FakeDevice 'PC01'
        $old = [pscustomobject]@{ id = 'grp-old'; displayName = 'BEAM-Contoso Agent' }
        [void]$script:G.Groups.Add($old)
        $script:G.Members['grp-old'] = New-Object System.Collections.ArrayList
        [void]$script:G.Members['grp-old'].Add('ent-STALE')
        $r = Invoke-Deploy
        $r.Published | Should -BeTrue
        $assign = Get-FakeCalls -Method POST -Pattern '/assign$'
        $gid = $assign[0].Body.mobileAppAssignments[0].target.groupId
        $gid | Should -Not -Be 'grp-old'
        @($script:G.Members[$gid]) | Should -Be @('ent-PC01')
        @($script:G.Members['grp-old']) | Should -Be @('ent-STALE')
    }

    It 'creates a different group for each run, named with its run ID' {
        Add-FakeDevice 'PC01'
        Invoke-Deploy -RunId 'RUN-1' | Out-Null
        Invoke-Deploy -RunId 'RUN-2' | Out-Null
        $script:G.Groups.Count | Should -Be 2
        ($script:G.Groups | ForEach-Object displayName) | Should -Contain 'BEAM-Contoso Agent-RUN-1'
        ($script:G.Groups | ForEach-Object displayName) | Should -Contain 'BEAM-Contoso Agent-RUN-2'
    }

    It 'keeps the group mailNickname within 64 ASCII characters even for a long run ID' {
        Add-FakeDevice 'PC01'
        Invoke-Deploy -AppName ('A' * 120) -RunId '20261118120000-aaaa1111' | Out-Null
        $g = $script:G.Groups[0]
        $g.mailNickname.Length | Should -BeLessOrEqual 64
        $g.mailNickname | Should -Match '^[\x21-\x7E]+$'
        $g.mailNickname | Should -Match '20261118120000-aaaa1111$'
    }

    It 'matches the Entra device by the managed device azureADDeviceId, not by the first name match' {
        [void]$script:G.Entra.Add([pscustomobject]@{ id = 'ent-STALE'; displayName = 'PC01'; deviceId = [guid]::NewGuid().ToString() })
        Add-FakeDevice 'PC01'
        $r = Invoke-Deploy
        $r.Published | Should -BeTrue
        $gid = (Get-FakeCalls -Method POST -Pattern '/assign$')[0].Body.mobileAppAssignments[0].target.groupId
        @($script:G.Members[$gid]) | Should -Be @('ent-PC01')
    }

    It 'marks a target FAIL as ambiguous when two Intune devices share its name' {
        Add-FakeDevice 'PC01'
        [void]$script:G.Managed.Add([pscustomobject]@{ id = 'md-dup'; deviceName = 'PC01'; operatingSystem = 'Windows'; osVersion = '1'; azureADDeviceId = [guid]::NewGuid().ToString() })
        $r = Invoke-Deploy
        $r.Targets['PC01'] | Should -Be 'FAIL'
        ($script:Log -join "`n") | Should -Match 'ambiguous'
        (Get-FakeCalls -Method POST -Pattern 'members/\$ref').Count | Should -Be 0
    }

    It 'marks a target FAIL when its device cannot be found anywhere' {
        $r = Invoke-Deploy -Targets @('GHOST')
        $r.Targets['GHOST'] | Should -Be 'FAIL'
        $r.Published | Should -BeFalse
        (Get-FakeCalls -Method POST -Pattern '/assign$').Count | Should -Be 0
    }

    It 'marks a target FAIL, not queued, when adding it to the group fails' {
        Add-FakeDevice 'PC01'
        Add-FakeFailure -Method POST -Pattern 'members/\$ref' -Mode throw -Message 'HTTP 403 Forbidden'
        $r = Invoke-Deploy
        $r.Targets['PC01'] | Should -Be 'FAIL'
        $r.Published | Should -BeFalse
        (Get-FakeCalls -Method POST -Pattern '/assign$').Count | Should -Be 0
        ($script:Log -join "`n") | Should -Match '403'
    }

    It 'marks a target FAIL when the add is accepted but the group read-back does not list the device' {
        Add-FakeDevice 'PC01'
        Add-FakeFailure -Method POST -Pattern 'members/\$ref' -Mode noop
        $r = Invoke-Deploy
        $r.Targets['PC01'] | Should -Be 'FAIL'
        ($script:Log -join "`n") | Should -Match 'not listed|not a member'
    }

    It 'marks a target FAIL when the Entra device lookup is denied' {
        Add-FakeDevice 'PC01'
        Add-FakeFailure -Method GET -Pattern '/devices\?' -Mode throw -Message 'HTTP 403 Insufficient privileges'
        $r = Invoke-Deploy
        $r.Targets['PC01'] | Should -Be 'FAIL'
    }

    It 'reports each target on its own: one queued, one missing' {
        Add-FakeDevice 'PC01'
        $r = Invoke-Deploy -Targets @('PC01', 'GHOST')
        $r.Published | Should -BeTrue
        $r.Targets['PC01'] | Should -Be 'Intune queued'
        $r.Targets['GHOST'] | Should -Be 'FAIL'
        $gid = (Get-FakeCalls -Method POST -Pattern '/assign$')[0].Body.mobileAppAssignments[0].target.groupId
        @($script:G.Members[$gid]) | Should -Be @('ent-PC01')
    }

    It 'still queues a verified target when only the sync trigger fails' {
        Add-FakeDevice 'PC01'
        Add-FakeFailure -Method POST -Pattern 'syncDevice' -Mode throw -Message 'HTTP 500'
        $r = Invoke-Deploy
        $r.Targets['PC01'] | Should -Be 'Intune queued'
        ($script:Log -join "`n") | Should -Match 'Sync.*WARN'
    }

    It 'marks every target FAIL when the assignment itself fails' {
        Add-FakeDevice 'PC01'
        Add-FakeFailure -Method POST -Pattern '/assign$'
        $r = Invoke-Deploy
        $r.Targets['PC01'] | Should -Be 'FAIL'
        $r.Published | Should -BeFalse
    }
}

Context 'Finding 04: the summary and exit code follow the per-target results' {
    BeforeEach {
        Mock Invoke-WebRequest { [IO.File]::WriteAllBytes($OutFile, [byte[]](0x4D, 0x5A, 0x90, 0x00, 1, 2, 3, 4)) }
        Mock Invoke-RestMethod {
            if ("$Uri" -match 'openid-configuration') { return [pscustomobject]@{ issuer = 'https://sts.windows.net/11111111-2222-3333-4444-555555555555/' } }
            Invoke-FakeGraph -Uri $Uri -Method $Method -Body $Body
        }
    }
    It 'exits 3 and lists the failed target when only some targets were queued' {
        Add-FakeDevice 'PC01'
        foreach ($r in 'https://files.invalid/setup.exe', '', '', '/S', 'PC01,GHOST', 'admin@contoso.invalid') { $script:Reads.Enqueue($r) }
        Invoke-BeamRun
        $script:Code | Should -Be 3
        $script:Text | Should -Match 'BEAM RESULT: exit code 3 -'
        $script:Text | Should -Match 'queued for 1 of 2'
        (@($script:Log) -join "`n") | Should -Match '\[FAIL\] GHOST'
        (@($script:Log) -join "`n") | Should -Not -Match '\[FAIL\] PC01'
    }
    It 'exits 1 and queues nothing when no target could be added to the group' {
        foreach ($r in 'https://files.invalid/setup.exe', '', '', '/S', 'GHOST', 'admin@contoso.invalid') { $script:Reads.Enqueue($r) }
        Invoke-BeamRun
        $script:Code | Should -Be 1
        $script:Text | Should -Match 'BEAM RESULT: exit code 1 -'
        (Get-FakeCalls -Method POST -Pattern '/assign$').Count | Should -Be 0
    }
}

Context 'Finding 62: the installer to be deployed is identified before it is packaged' {
    BeforeEach {
        Add-FakeDevice 'PC01'
        $script:Signature = [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Contoso Software, O=Contoso Ltd' } }
        Mock Get-AuthenticodeSignature { $script:Signature }
        $script:Fetched = New-Object System.Collections.ArrayList
        Mock Invoke-WebRequest {
            [void]$script:Fetched.Add("$Uri")
            if ("$Uri" -match 'landing') {
                [IO.File]::WriteAllText($OutFile, '<html><a href="https://cdn.invalid/real/setup.exe">Download</a></html>')
            } else {
                [IO.File]::WriteAllBytes($OutFile, [byte[]](0x4D, 0x5A, 0x90, 0x00, 1, 2, 3, 4))
            }
        }
    }

    It 'shows the source URL, the SHA256 and the Authenticode signer before packaging' {
        $r = Invoke-Deploy -Url 'https://files.invalid/setup.exe'
        $r.Published | Should -BeTrue
        $txt = $script:Log -join "`n"
        $txt | Should -Match 'https://files\.invalid/setup\.exe'
        $txt | Should -Match 'SHA256\)? ?[^\n]*[0-9A-Fa-f]{64}'
        $txt | Should -Match 'Contoso Ltd'
        $script:PackagerCalls.Count | Should -Be 1
    }

    It 'warns, but still deploys, when a URL the admin typed serves an unsigned installer' {
        $script:Signature = [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null }
        $r = Invoke-Deploy -Url 'https://files.invalid/setup.exe'
        $r.Published | Should -BeTrue
        ($script:Log -join "`n") | Should -Match 'NotSigned'
    }

    It 'refuses to package a link scraped from a page unless the technician confirms it' {
        $script:Reads.Enqueue('n')
        $r = Invoke-Deploy -Url 'https://files.invalid/landing'
        $r.Published | Should -BeFalse
        $script:PackagerCalls.Count | Should -Be 0
        $txt = $script:Log -join "`n"
        $txt | Should -Match 'https://cdn\.invalid/real/setup\.exe'
        $txt | Should -Match 'Contoso Ltd'
        (Get-FakeCalls -Method POST -Pattern 'mobileApps$').Count | Should -Be 0
    }

    It 'treats a blank answer as no for a scraped link' {
        $r = Invoke-Deploy -Url 'https://files.invalid/landing'
        $r.Published | Should -BeFalse
        $script:PackagerCalls.Count | Should -Be 0
    }

    It 'packages a scraped link after an explicit yes' {
        $script:Reads.Enqueue('y')
        $r = Invoke-Deploy -Url 'https://files.invalid/landing'
        $r.Published | Should -BeTrue
        $script:PackagerCalls.Count | Should -Be 1
        $script:Fetched | Should -Contain 'https://cdn.invalid/real/setup.exe'
    }

    It 'refuses a scraped link when there is no console to ask on' {
        Mock Read-Host { throw 'Read-Host : Windows PowerShell is in NonInteractive mode.' }
        $r = Invoke-Deploy -Url 'https://files.invalid/landing'
        $r.Published | Should -BeFalse
        $script:PackagerCalls.Count | Should -Be 0
    }
}

Context 'Finding 63: IntuneWinAppUtil is pinned to a release and its Microsoft signature is verified' {
    BeforeAll {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Get-KillerScriptPath BEAM), [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Initialize-BeamTool' }, $true)
        Invoke-Expression ($fn.Extent.Text -replace 'function Initialize-BeamTool', 'function Initialize-BeamToolReal')
    }
    BeforeEach {
        $script:Sig = [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Microsoft Corporation, O=Microsoft Corporation, L=Redmond, S=Washington, C=US' } }
        Mock Get-AuthenticodeSignature { $script:Sig }
        $script:Fetches = New-Object System.Collections.ArrayList
        Mock Invoke-WebRequest {
            [void]$script:Fetches.Add([pscustomobject]@{ Uri = "$Uri"; OutFile = $OutFile })
            [IO.File]::WriteAllBytes($OutFile, [byte[]](0x4D, 0x5A, 0x90, 0x00, 9, 9))
        }
        $script:ToolPath = Join-Path $script:WorkRoot 'IntuneWinAppUtil.exe'
    }

    It 'downloads from a pinned release tag, never from master' {
        $r = Initialize-BeamToolReal
        $r | Should -Be $script:ToolPath
        $script:Fetches.Count | Should -Be 1
        $script:Fetches[0].Uri | Should -Match '/raw/v\d+\.\d+\.\d+/IntuneWinAppUtil\.exe$'
        $script:Fetches[0].Uri | Should -Not -Match 'master'
    }

    It 'downloads to a temporary name and only moves it into place after the signature is verified' {
        Initialize-BeamToolReal | Out-Null
        $script:Fetches[0].OutFile | Should -Not -Be $script:ToolPath
        Test-Path -LiteralPath $script:ToolPath | Should -BeTrue
        @(Get-ChildItem -LiteralPath $script:WorkRoot -Filter '*.download*' -Force -ErrorAction SilentlyContinue).Count | Should -Be 0
    }

    It 'rejects and removes a download that is not signed' {
        $script:Sig = [pscustomobject]@{ Status = 'NotSigned'; SignerCertificate = $null }
        Initialize-BeamToolReal | Should -BeNullOrEmpty
        Test-Path -LiteralPath $script:ToolPath | Should -BeFalse
        @(Get-ChildItem -LiteralPath $script:WorkRoot -Force | Where-Object { $_.Name -ne 'IntuneWork' }).Count | Should -Be 0
    }

    It 'rejects a validly signed file whose signer is not Microsoft Corporation' {
        $script:Sig = [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=Evil Ltd, O=Evil Ltd, C=US' } }
        Initialize-BeamToolReal | Should -BeNullOrEmpty
        Test-Path -LiteralPath $script:ToolPath | Should -BeFalse
    }

    It 'rejects a signer that only mentions Microsoft Corporation in another field' {
        $script:Sig = [pscustomobject]@{ Status = 'Valid'; SignerCertificate = [pscustomobject]@{ Subject = 'CN=O=Microsoft Corporation, O=Evil Ltd, C=US' } }
        Initialize-BeamToolReal | Should -BeNullOrEmpty
    }

    It 'reuses a cached copy without downloading when it still verifies' {
        New-Item -ItemType Directory -Path $script:WorkRoot -Force | Out-Null
        [IO.File]::WriteAllBytes($script:ToolPath, [byte[]](0x4D, 0x5A, 1, 2))
        Initialize-BeamToolReal | Should -Be $script:ToolPath
        $script:Fetches.Count | Should -Be 0
    }

    It 'replaces a cached copy that no longer verifies by downloading once more' {
        New-Item -ItemType Directory -Path $script:WorkRoot -Force | Out-Null
        [IO.File]::WriteAllBytes($script:ToolPath, [byte[]](0x4D, 0x5A, 1, 2))
        $script:SigCalls = 0
        Mock Get-AuthenticodeSignature {
            $script:SigCalls++
            if ($script:SigCalls -eq 1) { [pscustomobject]@{ Status = 'HashMismatch'; SignerCertificate = $null } } else { $script:Sig }
        }
        Initialize-BeamToolReal | Should -Be $script:ToolPath
        $script:Fetches.Count | Should -Be 1
    }

    It 'fails closed and leaves no executable behind when the cached copy and the fresh download both fail' {
        New-Item -ItemType Directory -Path $script:WorkRoot -Force | Out-Null
        [IO.File]::WriteAllBytes($script:ToolPath, [byte[]](0x4D, 0x5A, 1, 2))
        $script:Sig = [pscustomobject]@{ Status = 'HashMismatch'; SignerCertificate = $null }
        Initialize-BeamToolReal | Should -BeNullOrEmpty
        Test-Path -LiteralPath $script:ToolPath | Should -BeFalse
    }

    It 'fails when the download itself fails' {
        Mock Invoke-WebRequest { throw 'The remote server returned an error: (404) Not Found.' }
        Initialize-BeamToolReal | Should -BeNullOrEmpty
    }
}

Context 'Finding 37: the redirect must return the state BEAM issued before the code is exchanged' {
    BeforeAll {
        $ast = [System.Management.Automation.Language.Parser]::ParseFile((Get-KillerScriptPath BEAM), [ref]$null, [ref]$null)
        $fn = $ast.Find({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Get-GraphToken' }, $true)
        Invoke-Expression ($fn.Extent.Text -replace 'function Get-GraphToken', 'function Get-GraphTokenReal')
    }
    BeforeEach {
        $script:TokenCalls = 0
        $script:AuthUrl = $null
        $script:Redirect = { param($state) "?code=abc123&state=$state&session_state=zzz" }
        $script:Clip = $null
        Mock Open-BeamListener { [pscustomobject]@{ Listener = $null; Port = 8400 } }
        Mock Start-BeamBrowser { $script:AuthUrl = $Url }
        Mock Receive-BeamRedirect {
            $s = ([regex]::Match($script:AuthUrl, '[?&]state=([^&]+)')).Groups[1].Value
            [pscustomobject]@{ Cancelled = $false; Query = (& $script:Redirect $s) }
        }
        Mock Get-BeamClipboardText { $script:Clip }
        Mock Invoke-RestMethod { $script:TokenCalls++; [pscustomobject]@{ access_token = 'tok' } }
    }

    It 'issues a state value in the authorization URL' {
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Out-Null
        $script:AuthUrl | Should -Match '[?&]state=[0-9a-f]{32}(&|$)'
    }

    It 'exchanges the code when the returned state matches' {
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -Be 'tok'
        $script:TokenCalls | Should -Be 1
    }

    It 'aborts without calling the token endpoint when the returned state differs' {
        $script:Redirect = { param($state) '?code=abc123&state=attacker' }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
        ($script:Log -join "`n") | Should -Match 'state'
    }

    It 'aborts without calling the token endpoint when the redirect carries no state at all' {
        $script:Redirect = { param($state) '?code=abc123' }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
    }

    It 'aborts when the state parameter is present but empty' {
        $script:Redirect = { param($state) '?code=abc123&state=' }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
    }

    It 'aborts when the state is repeated, even if one copy matches' {
        $script:Redirect = { param($state) "?code=abc123&state=$state&state=attacker" }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
    }

    It 'does not treat session_state as the state' {
        $script:Redirect = { param($state) "?code=abc123&session_state=$state" }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
    }

    It 'applies the same check to a redirect URL pasted from the clipboard' {
        $script:Redirect = { param($state) $null }
        $script:Clip = 'http://localhost:8400/?code=abc123'
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
    }

    It 'accepts a clipboard redirect that carries the issued state' {
        $script:Redirect = { param($state) $null }
        Mock Get-BeamClipboardText {
            $s = ([regex]::Match($script:AuthUrl, '[?&]state=([^&]+)')).Groups[1].Value
            "http://localhost:8400/?code=abc123&state=$s"
        }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -Be 'tok'
        $script:TokenCalls | Should -Be 1
    }

    It 'aborts on an error redirect without calling the token endpoint' {
        $script:Redirect = { param($state) "?error=access_denied&state=$state" }
        Get-GraphTokenReal -TenantId '11111111-2222-3333-4444-555555555555' | Should -BeNullOrEmpty
        $script:TokenCalls | Should -Be 0
    }
}}
