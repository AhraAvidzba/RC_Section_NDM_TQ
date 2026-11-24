# ДЛЯ ТЕСТОВ. Проверяет один холодный HardShear Direct/Excel от запуска до Results и схемы.
param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
$taskTemp=$env:TEMP; $taskTmp=$env:TMP
Push-Location -LiteralPath $Root
try {
    . './Guard-IsolatedOffice.ps1'
    $prep=Get-Content -LiteralPath 'Preparation.json' -Raw | ConvertFrom-Json
    foreach($pin in $prep.inputs.psobject.Properties) {
        if((Get-FileHash -LiteralPath $pin.Name).Hash -cne $pin.Value){throw ('Changed input '+$pin.Name)}
    }
    $book='docs/regression/Diagnostic.xlsm'
    $compiled=Get-Content -LiteralPath 'docs/regression/CodeUpdate.json' -Raw | ConvertFrom-Json
    if(!$compiled.Compiled -or $compiled.ProductionChanged -or (Get-FileHash -LiteralPath $book).Hash -cne $compiled.WorkbookSHA256){throw 'Private compilation not verified.'}
    if(Test-Path -LiteralPath 'Native/PriorBranchHardShearDirect.txt'){throw 'Regression report already exists.'}
    Assert-IsolatedOfficeReady 'PriorBranchMaterialRegressions'
    $env:TEMP=(Resolve-Path -LiteralPath 'Temp').Path; $env:TMP=$env:TEMP
    $taskMacros='modTestCoupledElasticState.RunTestCoupledWorkbookTiming'
    & './tools/build_workbook/Run-FocusedKaniWatchdog.ps1' -SourceWorkbook $book -ReportPath 'Native/PriorBranchHardShearDirect.txt' `
        -Macro $taskMacros -MacroArgument1 Direct -MacroArgument2 Excel -TimeoutSeconds 180 -Mode No -FailFast
    if($LASTEXITCODE -ne 0){throw ('Material regressions exited '+$LASTEXITCODE)}
} finally {
    $env:TEMP=$taskTemp; $env:TMP=$taskTmp
    Pop-Location
}
