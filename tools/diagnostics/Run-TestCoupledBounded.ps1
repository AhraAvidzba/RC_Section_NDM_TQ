# ==========================================================================
# ДЛЯ ТЕСТОВ. Запускает один расчёт только в собственной частной книге на D.
# Бюджет берётся из политики проекта и не увеличивается автоматически.
# Завершение процесса не заменяет независимую проверку результата.
# ==========================================================================
param(
    [Parameter(Mandatory=$true)][string]$CaseRoot,
    [string]$Macro='modTestSelectableMaterialCore.TestRunASDSpatialMainCapacity',
    [ValidateSet('Local','Capacity','Verification')][string]$Kind='Local'
)
$ErrorActionPreference='Stop'
$taskPolicy=Get-Content -LiteralPath (Join-Path $PSScriptRoot 'CoupledDiagnosticPolicy.json') -Raw | ConvertFrom-Json
if ($taskPolicy.automatic_budget_increase_allowed -or $taskPolicy.single_case_diagnostic_max_wall_seconds -ne 10 -or
        $taskPolicy.development_capacity_search_max_wall_seconds -ne 180 -or
        $taskPolicy.saved_history_verification_max_wall_seconds -ne 30) {
    throw 'Политика коротких диагностических расчётов изменена. Проверьте ограничение времени перед запуском.'
}
$taskMaximumSeconds=switch($Kind){
    'Capacity' {[int]$taskPolicy.development_capacity_search_max_wall_seconds}
    'Verification' {[int]$taskPolicy.saved_history_verification_max_wall_seconds}
    'Local' {[int]$taskPolicy.single_case_diagnostic_max_wall_seconds}
}
$taskAllowedRoot=[IO.Path]::GetFullPath('D:/CodexScratch/RC_Section_NDM_TQ')+[IO.Path]::DirectorySeparatorChar
$taskCase=[IO.Path]::GetFullPath((Resolve-Path -LiteralPath $CaseRoot).Path)
if (!$taskCase.StartsWith($taskAllowedRoot,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Диагностический расчёт разрешён только в собственной рабочей папке на диске D.'
}
if ($Macro -notmatch '^[A-Za-z_][A-Za-z0-9_]*\.Test[A-Za-z0-9_]+$') {
    throw 'Для диагностики укажите явную тестовую точку входа VBA.'
}
if ($Kind -eq 'Verification' -and $Macro -cne 'modTestCoupledProgressGuard.TestRunProgressGuard') {
    throw 'Режим проверки сохранённой истории допускает только проверенную точку входа без решения НДС.'
}
$taskBookRoot=Split-Path -Parent $taskCase
$taskBook=Join-Path $taskBookRoot 'docs/regression/Diagnostic.xlsm'
$taskCompiled=Get-Content -LiteralPath (Join-Path $taskBookRoot 'docs/regression/CodeUpdate.json') -Raw | ConvertFrom-Json
if (!$taskCompiled.Compiled -or (Get-FileHash -LiteralPath $taskBook).Hash -cne $taskCompiled.WorkbookSHA256) {
    throw 'Частная книга не соответствует подтверждённой сборке.'
}
$taskReport=Join-Path $taskCase 'Native/BoundedProbe.txt'
$taskRelativeReport=(Split-Path -Leaf $taskCase)+'/Native/BoundedProbe.txt'
if (Test-Path -LiteralPath $taskReport) { throw 'Отчёт ограниченной пробы уже существует. Его перезапись запрещена.' }
$taskOriginalTemp=$env:TEMP; $taskOriginalTmp=$env:TMP
Push-Location -LiteralPath $taskCase
try {
    . (Join-Path $taskBookRoot 'Guard-IsolatedOffice.ps1')
    Assert-IsolatedOfficeReady ('BoundedSingleCase'+$Kind)
    $env:TEMP=(Resolve-Path -LiteralPath (Join-Path $taskBookRoot 'Temp')).Path; $env:TMP=$env:TEMP
    [ordered]@{Macro=$Macro;Kind=$Kind;MaximumWallSeconds=$taskMaximumSeconds;AutomaticBudgetIncreaseAllowed=$false;BookSHA256=$taskCompiled.WorkbookSHA256;UTC=[DateTimeOffset]::UtcNow.ToString('o');NoPhysicalVerdictFromProcessExit=$true} |
        ConvertTo-Json | Set-Content -LiteralPath (Join-Path $taskCase 'Native/BoundedProbeContract.json') -Encoding utf8
    & (Join-Path $taskBookRoot 'tools/build_workbook/Run-FocusedKaniWatchdog.ps1') -SourceWorkbook $taskBook -ReportPath $taskRelativeReport `
        -Macro $Macro -MacroArgument1 $taskCase -TimeoutSeconds $taskMaximumSeconds -Mode No -FailFast
    if ($LASTEXITCODE -ne 0) { throw 'Короткая диагностическая проба остановлена. Физический предел по этому факту не установлен; смотрите сохранённую диагностику.' }
} finally {
    $env:TEMP=$taskOriginalTemp; $env:TMP=$taskOriginalTmp
    Pop-Location
}
