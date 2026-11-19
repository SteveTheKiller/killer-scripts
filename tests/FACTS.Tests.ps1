#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }
# FACTS tests. Every registry, service, scheduler and console call is mocked.

BeforeDiscovery {
    $script:FactsPath = Join-Path (Split-Path -Parent $PSScriptRoot) 'FACTS.ps1'
}

Describe 'FACTS entry-point convention (static)' {
    BeforeAll {
        $path = Join-Path (Split-Path -Parent $PSScriptRoot) 'FACTS.ps1'
        $tokens = $null; $errors = $null
        $script:Ast = [System.Management.Automation.Language.Parser]::ParseFile($path, [ref]$tokens, [ref]$errors)
    }
    It 'defines Invoke-FactsMain and Exit-KillerScript' {
        $names = @($script:Ast.FindAll({ param($n) $n -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true) | ForEach-Object { $_.Name })
        $names | Should -Contain 'Invoke-FactsMain'
        $names | Should -Contain 'Exit-KillerScript'
    }
    It 'ends with the test-mode guard and keeps only definitions and constants at the top level' {
        $stmts = @($script:Ast.EndBlock.Statements)
        $stmts[-1].Extent.Text | Should -Match 'KillerScriptsTestMode'
        $stmts[-1].Extent.Text | Should -Match 'Invoke-FactsMain'
        $bad = @($stmts | Select-Object -SkipLast 1 | Where-Object {
            $_ -isnot [System.Management.Automation.Language.FunctionDefinitionAst] -and
            $_ -isnot [System.Management.Automation.Language.AssignmentStatementAst]
        })
        ($bad | ForEach-Object { $_.Extent.Text.Split("`n")[0] }) -join "`n" | Should -BeNullOrEmpty
    }
}
