<#
    Inert fixture for HostModes.Tests.ps1. It uses the same entry-point convention as every
    shipped killer-script, but does nothing except print a marker and exit 3.
#>

function Exit-KillerScript {
    param([int]$Code)
    if ($global:KillerScriptsTestMode) { throw "KILLERSCRIPT_EXIT:$Code" }
    exit $Code
}

function Invoke-GuardSampleMain {
    Write-Output 'GUARD-SAMPLE-RAN'
    Write-Output ("BITNESS={0}" -f ([IntPtr]::Size * 8))
    Exit-KillerScript 3
}

if (-not $global:KillerScriptsTestMode) { Invoke-GuardSampleMain }
