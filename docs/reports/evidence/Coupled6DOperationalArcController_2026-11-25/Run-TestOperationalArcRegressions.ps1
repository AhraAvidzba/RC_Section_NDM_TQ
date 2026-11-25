# ДЛЯ ТЕСТОВ. Проверяет затронутые плечи, огибающую и локализацию предела стали; полным набором приёмки не является.
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
    if(Test-Path -LiteralPath 'Native/OperationalArcRegressions.txt'){throw 'Regression report already exists.'}
    Assert-IsolatedOfficeReady 'OperationalArcPeakAndSteelRegressions'
    $env:TEMP=(Resolve-Path -LiteralPath 'Temp').Path; $env:TMP=$env:TEMP
    $taskMacros='modTestCoupledEnvelopeControl.RunTestCoupledEnvelopeControl,modTestCoupledElasticState.RunTestCoupledRetainedShoulders,modTestCoupledElasticState.RunTestCoupledDiscardedShoulder,modTestCoupledElasticState.RunTestCoupledDeformationSteelCrossing,modTestCoupledElasticState.RunTestCoupledSteelAtDesign'
    & './tools/build_workbook/Run-FocusedKaniWatchdog.ps1' -SourceWorkbook $book -ReportPath 'Native/OperationalArcRegressions.txt' `
        -Macro $taskMacros -TimeoutSeconds 600 -Mode No -FailFast
    if($LASTEXITCODE -ne 0){throw ('Regression exited '+$LASTEXITCODE)}
} finally {
    $env:TEMP=$taskTemp; $env:TMP=$taskTmp
    Pop-Location
}
