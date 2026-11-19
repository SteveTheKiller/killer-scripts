<#
.SYNOPSIS
    Runs the killer-scripts test suite under Windows PowerShell 5.1 and PowerShell 7.
.DESCRIPTION
    Test-only. Needs Pester 5 installed for both editions. Shipped scripts never depend on it.
    Exit code 0 means every edition passed. 1 means at least one test failed, or that an
    edition or Pester was missing.
.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1
.EXAMPLE
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tests\Run-Tests.ps1 -Path .\tests\BERET.Tests.ps1
#>
param(
    [string[]]$Path = @($PSScriptRoot),
    [ValidateSet('Both', '5.1', '7')][string]$Edition = 'Both'
)

$hosts = @()
if ($Edition -in 'Both', '5.1') { $hosts += Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe' }
if ($Edition -in 'Both', '7') {
    $pw = Get-Command pwsh.exe -ErrorAction SilentlyContinue
    if ($pw) { $hosts += $pw.Source } else { Write-Host '[FAIL] pwsh.exe not found' ; exit 1 }
}

$pathList = ($Path | ForEach-Object { "'" + ((Resolve-Path -LiteralPath $_).Path -replace "'", "''") + "'" }) -join ','
$failed = $false
foreach ($h in $hosts) {
    Write-Host "=== $h ==="
    $cmd = "Import-Module Pester -MinimumVersion 5.0 -ErrorAction Stop; " +
           "`$c = New-PesterConfiguration; `$c.Run.Path = @($pathList); `$c.Run.Exit = `$true; " +
           "`$c.Output.Verbosity = 'Normal'; Invoke-Pester -Configuration `$c"
    & $h -NoProfile -ExecutionPolicy Bypass -Command $cmd
    if ($LASTEXITCODE -ne 0) { $failed = $true; Write-Host "[FAIL] $h exit $LASTEXITCODE" }
}
if ($failed) { exit 1 }
Write-Host '[PASS] all editions'
exit 0
