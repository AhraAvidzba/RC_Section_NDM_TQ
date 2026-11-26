# ДЛЯ ТЕСТОВ: направленная компиляция и проверки рабочего спектрального материала.
param([ValidateSet('Compile','Core','Material','Tangent','MaterialContexts','Quadrature','MVT','ControlProbe','Trace44','Trace74','Stock')][string]$Phase='Compile')
$ErrorActionPreference='Stop'
$tempBefore=$env:TEMP;$tmpBefore=$env:TMP
$flowBefore=$env:NDM_MVT_FLOW_FOLDER;$modeBefore=$env:NDM_MVT_FLOW_MODE
Push-Location -LiteralPath $PSScriptRoot
try {
    . './Guard-IsolatedOffice.ps1'
    $prep=Get-Content -LiteralPath 'Preparation.json' -Raw | ConvertFrom-Json
    foreach ($file in $prep.inputs.psobject.Properties) {
        if ((Get-FileHash -LiteralPath $file.Name).Hash -cne $file.Value) {throw "Изменён вход $($file.Name)."}
    }
    $env:TEMP=(Resolve-Path -LiteralPath 'Temp').Path;$env:TMP=$env:TEMP
    $env:NDM_MVT_FLOW_FOLDER=$null;$env:NDM_MVT_FLOW_MODE=$null
    $book='docs/regression/Diagnostic.xlsm'
    if ($Phase -eq 'Compile') {
        if ((Test-Path -LiteralPath 'Compile.log') -or (Get-FileHash -LiteralPath $book).Hash -cne $prep.source_book_sha256) {throw 'Изменена чистая книга.'}
        Assert-IsolatedOfficeReady 'CompileWorkingASD'
        & './tools/build_workbook/Update-IsolatedOptimizerCode.ps1' -WorkbookPath $book -ModuleNames $prep.modules |
            Tee-Object -FilePath 'Compile.log'
        Save-OfficeClosure 'Compile.log' 'CompileClosure.json'
    } else {
        $compile=Get-Content -LiteralPath 'docs/regression/CodeUpdate.json' -Raw | ConvertFrom-Json
        $closure=Get-Content -LiteralPath 'CompileClosure.json' -Raw | ConvertFrom-Json
        if (!$compile.Compiled -or $compile.ProductionChanged -or $closure.State -cne 'Absent' -or
            $compile.ProductionSHA256 -cne $prep.source_book_sha256 -or
            $compile.WorkbookSHA256 -cne (Get-FileHash -LiteralPath $book).Hash -or
            (Compare-Object ($prep.modules|Sort-Object) ($compile.Modules|Sort-Object))) {throw 'Нет подтверждённой частной компиляции.'}
        $report='Native/'+$Phase+'.txt'
        if (Test-Path -LiteralPath $report) {throw 'Протокол уже существует.'}
        $legacy=Join-Path $env:WINDIR 'System32/WindowsPowerShell/v1.0/powershell.exe'
        $arguments=@('-NoProfile','-ExecutionPolicy','Bypass','-File','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
            '-SourceWorkbook',$book,'-ReportPath',$report,'-Mode','No','-FailFast')
        if ($Phase -eq 'Core') {
            $suite=@('RunCoupledNewtonTests','RunTestCoupledSteelLimit','RunTestCoupledSpatialTransactions','RunTestCoupledContinuation',
                'RunTestCoupledDeformationSteelCrossing','RunTestCoupledHoopCapacity','RunTestCoupledSteelClasses','RunTestCoupledSteelAtDesign')
            $arguments+=@('-Macro',(($suite|ForEach-Object {'modTestCoupledElasticState.'+$_}) -join ','),'-TimeoutSeconds','165')
        } elseif ($Phase -eq 'Material') {
            $arguments+=@('-Macro','modTestCoupledASDMaterial.RunTestCoupledASDMaterial','-MacroArgument1',
                (Resolve-Path -LiteralPath 'Data').Path,'-TimeoutSeconds','165')
        } elseif ($Phase -eq 'Tangent') {
            $arguments+=@('-Macro','modTestCoupledASDTangent.RunTestCoupledASDTangent','-MacroArgument1',
                (Resolve-Path -LiteralPath 'Data').Path,'-TimeoutSeconds','165')
        } elseif ($Phase -eq 'Quadrature') {
            $arguments+=@('-Macro','modTestCoupledQuadrature.RunTestCoupledQuadrature','-TimeoutSeconds','165')
        } elseif ($Phase -in @('MVT','ControlProbe','MaterialContexts','Trace44','Trace74')) {
            $env:NDM_MVT_FLOW_FOLDER=$PSScriptRoot
            $caseName=if ($Phase -eq 'MVT') {'MVTGaussMixed'} else {$Phase}
            $arguments+=@('-Macro','modTestMVT2ProductionASDTest.RunTestMVT2ProductionASDTest','-MacroArgument1',$caseName,'-TimeoutSeconds','240')
        } else {
            $arguments+=@('-Macro','modTestCoupledElasticState.RunTestCoupledWorkbookTiming','-MacroArgument1','Both',
                '-MacroArgument2','Excel','-TimeoutSeconds','240','-VerifyResultsReopen')
        }
        Assert-IsolatedOfficeReady ('VerifyWorkingASD'+$Phase)
        & $legacy @arguments
        $exitCode=$LASTEXITCODE
        if ($Phase -in @('MVT','ControlProbe','MaterialContexts','Trace44','Trace74')) {
            $fixtures=@(Get-Content -LiteralPath $report -Encoding UTF8 | Where-Object {$_ -like 'FIXTURE: *'})
            if ($fixtures.Count -ne 1) {throw 'Нет однозначной частной книги MVT2.'}
            $folder=Split-Path -Parent $fixtures[0].Substring('FIXTURE: '.Length)
            if (!$folder.StartsWith($env:TEMP+'\',[StringComparison]::OrdinalIgnoreCase)) {throw 'Снимок MVT2 вне временного каталога.'}
            foreach ($suffix in @('','.Tensor.csv','.Operators.csv','.Checkpoint.csv')) {
                $trace=Join-Path $folder ('MVT2GaussSpan.csv'+$suffix)
                if (Test-Path -LiteralPath $trace) {
                    $outputName=if ($suffix) {'Native/'+$Phase+$suffix} else {'Native/'+$Phase+'.csv'}
                    Copy-Item -LiteralPath $trace -Destination $outputName
                }
            }
        }
        if ($exitCode -ne 0) {throw "Направленный этап $Phase завершился с кодом $exitCode."}
    }
} finally {
    $env:TEMP=$tempBefore;$env:TMP=$tmpBefore
    $env:NDM_MVT_FLOW_FOLDER=$flowBefore;$env:NDM_MVT_FLOW_MODE=$modeBefore
    Pop-Location
}
