# ДЛЯ ТЕСТОВ. Один штатный запуск с читающими контрольными снимками; не используется как чистый замер производительности.
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
    $closed=Get-Content -LiteralPath 'CompileClosure.json' -Raw | ConvertFrom-Json
    if(!$compiled.Compiled -or $compiled.ProductionChanged -or $closed.State -cne 'Absent' -or
       (Get-FileHash -LiteralPath $book).Hash -cne $compiled.WorkbookSHA256){throw 'Private compilation not verified.'}
    if(Compare-Object ($compiled.Modules|Sort-Object) ($prep.directed_scope.actual_import_this_pass|Sort-Object)){
        throw 'Unexpected actual imports.'
    }
    if(Test-Path -LiteralPath 'Native/Stock.Checkpoints.csv'){throw 'Stock checkpoint already exists.'}
    Assert-IsolatedOfficeReady 'MixedHStockWithReadOnlyCheckpoints'
    $env:TEMP=(Resolve-Path -LiteralPath 'Temp').Path; $env:TMP=$env:TEMP
    & './tools/build_workbook/Run-FocusedKaniWatchdog.ps1' -SourceWorkbook $book -ReportPath 'Native/HBackendStock.txt' `
        -Macro 'modTestCoupledElasticState.RunTestStockHTrace' -MacroArgument1 $Root `
        -VerifyResultsReopen -TimeoutSeconds 600 -Mode No -FailFast
    if($LASTEXITCODE -ne 0){throw ('Stock exited '+$LASTEXITCODE)}
} finally {
    $env:TEMP=$taskTemp; $env:TMP=$taskTmp
    Pop-Location
}
