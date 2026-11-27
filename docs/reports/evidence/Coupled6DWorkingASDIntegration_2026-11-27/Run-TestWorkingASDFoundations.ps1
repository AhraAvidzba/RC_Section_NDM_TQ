# ДЛЯ ТЕСТОВ. Короткий направленный набор математических и стальных проверок фактической книги; полный расчётный/UI-прогон не запускается.
param([Parameter(Mandatory=$true)][string]$Root)
$ErrorActionPreference='Stop'
$previousTaskTemp=$env:TEMP;$previousTaskTmp=$env:TMP
Push-Location -LiteralPath $Root
try {
    . './Guard-IsolatedOffice.ps1'
    $proof=Get-Content -LiteralPath 'PreservedPackageProof.json' -Raw | ConvertFrom-Json
    $book='docs/regression/Diagnostic.xlsm'
    if((Get-FileHash -LiteralPath $book).Hash -cne $proof.saved_candidate_sha256){throw 'Candidate changed.'}
    if(Test-Path -LiteralPath 'Native/Foundations.txt'){throw 'Foundation report already exists.'}
    $reopen=Get-Content -LiteralPath 'Native/PreservedReopenProof.json' -Raw | ConvertFrom-Json
    if(!$reopen.Compiled -or !$reopen.NoSave -or !$reopen.ReadOnly){throw 'Preserved read-only compilation is not verified.'}
    $closure=Get-Content -LiteralPath 'Native/ReopenClosure.json' -Raw | ConvertFrom-Json
    if($closure.State -ne 'Absent'){throw 'Owned readonly compiler is not closed.'}
    Assert-IsolatedOfficeReady 'PreservedWorkingASDFoundations'
    $env:TEMP=(Resolve-Path -LiteralPath 'Temp').Path;$env:TMP=$env:TEMP
    $macros=@('modTestCoupledHybrid.RunTestCoupledHybrid','modTestCoupledHybrid.RunTestCoupledLeastSquaresTrust',
        'modTestCoupledConstraint.RunTestCoupledConstraint','modTestCoupledControlCondition.RunTestCoupledControlCondition',
        'modTestMaterialDiagrams.RunMaterialDiagramTests','modTestMaterialDiagrams.RunPlateauSlopeTests',
        'modTestMaterialDiagrams.RunAsymmetricSteelStrengthTests') -join ','
    & './tools/build_workbook/Run-FocusedKaniWatchdog.ps1' -SourceWorkbook $book -ReportPath 'Native/Foundations.txt' `
        -Macro $macros -TimeoutSeconds 600 -Mode No -FailFast
    if($LASTEXITCODE -ne 0){throw ('Directed foundations exited '+$LASTEXITCODE)}
} finally {
    $env:TEMP=$previousTaskTemp;$env:TMP=$previousTaskTmp
    Pop-Location
}
