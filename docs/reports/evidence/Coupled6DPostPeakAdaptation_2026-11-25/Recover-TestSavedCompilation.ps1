# ДЛЯ ТЕСТОВ. Восстанавливает паспорт уже скомпилированной и сохранённой частной книги после отдельной ошибки закрытия COM.
param([Parameter(Mandatory=$true)][string]$Root,[Parameter(Mandatory=$true)][string]$Parent)
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $Root
try {
    $log=Get-Content -LiteralPath 'CompilePoint.log' -Raw
    foreach($marker in @('COMPILE_VBE_REFERENCES_RELEASED','COMPILE_BEFORE_SAVE','COMPILE_AFTER_SAVE_BEFORE_CLOSE')) {
        if(!$log.Contains($marker)){throw ('Не подтверждена фаза '+$marker)}
    }
    if($log.Contains('COMPILE_AFTER_CLOSE') -or $log.Contains('PRIVATE_CODE_READY')){throw 'Это не отдельная ошибка закрытия после сохранения.'}
    $failure=Get-Content -LiteralPath 'CompilePointFailure.json' -Raw|ConvertFrom-Json
    if($failure.Message -ne 'You cannot call a method on a null-valued expression.' -or $failure.ScriptStackTrace -notmatch 'line 201'){throw 'Другая ошибка компиляции.'}
    $source=Get-Content -LiteralPath 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'
    if($source[200].Trim() -cne '$taskCloseAnchor=$excel.Workbooks.Add()'){throw 'Ошибка не относится к созданию служебной книги после Save.'}
    $closure=Get-Content -LiteralPath 'CompilePointClosure.json' -Raw|ConvertFrom-Json
    if($closure.State -ne 'Absent'){throw 'Собственный процесс не закрыт.'}
    $worker=Get-Process -Id $closure.PID -ErrorAction SilentlyContinue
    if($worker -and $worker.StartTime.ToUniversalTime().Ticks -eq [long]$closure.StartTicksUTC){throw 'Собственный Excel ещё существует.'}
    $before=Get-FileHash -LiteralPath 'docs/regression/Diagnostic.xlsm'
    $original=Get-Content -LiteralPath 'ParentCodeUpdate.json' -Raw|ConvertFrom-Json
    $production=Get-FileHash -LiteralPath 'workbook/output/RC_Section_NDM.xlsm'
    if($production.Hash -cne $original.ProductionSHA256){throw 'Изменена книга выпуска.'}
    . './tools/build_workbook/GeometryExampleSupport.ps1'
    . './tools/build_workbook/SettingsCatalog.ps1'
    $areas=@(Get-WorkbookPrintAreas (Join-Path $Parent 'docs/regression/Diagnostic.xlsm'))
    Restore-WorkbookPrintAreas (Resolve-Path -LiteralPath 'docs/regression/Diagnostic.xlsm').Path $areas
    Copy-Item -LiteralPath 'docs/regression/CodeUpdate.json' -Destination 'CodeUpdateBeforeRecovery.json'
    [ordered]@{Workbook=(Resolve-Path -LiteralPath 'docs/regression/Diagnostic.xlsm').Path;
        WorkbookSHA256=(Get-FileHash -LiteralPath 'docs/regression/Diagnostic.xlsm').Hash;
        ProductionSHA256=$production.Hash;Compiled=$true;Modules=@('modTestCoupledElasticState');
        RetiredModules=@();EntryPointNames=@();TraceGeometry=$false;ProductionChanged=$false;
        CompilerPhaseAndSaveCompletedBeforeCleanupFailure=$true;OriginalCompilerInvocationStillFailed=$true;
        ExactOwnProcessClosureVerified=$true;MandatoryActual255SourceAuditStillRequired=$true}|
        ConvertTo-Json -Depth 5|Set-Content -LiteralPath 'docs/regression/CodeUpdate.json' -Encoding UTF8
    [ordered]@{Recovery='Compilation phase and Save completed; separate COM cleanup failed';
        OriginalFailurePreserved=$true;CompilerLogSHA256=(Get-FileHash -LiteralPath 'CompilePoint.log').Hash;
        CompilerScriptSHA256=(Get-FileHash -LiteralPath 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1').Hash;
        SavedBeforeCleanupFailureSHA256=$before.Hash;OwnProcessClosure=$closure;
        ActualSourceAuditAndNativeExecutionStillRequired=$true}|
        ConvertTo-Json -Depth 6|Set-Content -LiteralPath 'SavedCompilationRecovery.json' -Encoding UTF8
    Write-Output 'SAVED COMPILATION PHASE RECOVERED; ORIGINAL CLEANUP FAILURE PRESERVED; ACTUAL SOURCE AUDIT REQUIRED'
} finally {Pop-Location}
