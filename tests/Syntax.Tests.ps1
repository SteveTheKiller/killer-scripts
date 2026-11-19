#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# Every shipped script must parse with zero errors. Run-Tests.ps1 runs this file under Windows
# PowerShell 5.1 as well as 7, so syntax newer than 5.1 fails here.

BeforeDiscovery {
    $repo = Split-Path -Parent $PSScriptRoot
    $script:Scripts = Get-ChildItem -LiteralPath $repo -Filter '*.ps1' -File |
        ForEach-Object { @{ Name = $_.Name; Path = $_.FullName } }
}

Describe 'Shipped script syntax' {
    It 'finds all 19 shipped scripts' {
        $repo = Split-Path -Parent $PSScriptRoot
        @(Get-ChildItem -LiteralPath $repo -Filter '*.ps1' -File).Count | Should -Be 19
    }

    It '<Name> parses with zero errors' -ForEach $Scripts {
        $tokens = $null; $errors = $null
        $null = [System.Management.Automation.Language.Parser]::ParseFile($Path, [ref]$tokens, [ref]$errors)
        ($errors | ForEach-Object { "$($_.Extent.StartLineNumber): $($_.Message)" }) -join "`n" | Should -BeNullOrEmpty
    }

    It '<Name> is saved with a UTF-8 byte order mark so Windows PowerShell 5.1 reads non-ASCII text correctly' -ForEach $Scripts {
        $bytes = [System.IO.File]::ReadAllBytes($Path)
        $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
        $nonAscii = @($bytes | Where-Object { $_ -gt 127 }).Count -gt 0
        if ($nonAscii) { $hasBom | Should -BeTrue }
    }
}
