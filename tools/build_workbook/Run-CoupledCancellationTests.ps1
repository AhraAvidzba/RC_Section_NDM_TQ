# Вводит два диагностических источника Excel 18 только в отдельной копии.
# Системные клавиши, рабочая книга и её исходный код не изменяются.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$EvidenceDirectory,
    [ValidateSet('Cancellation','PublicationFailure')][string]$Verification = 'Cancellation'
)

$ErrorActionPreference = 'Stop'
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$regressionRoot = (Resolve-Path -LiteralPath (Join-Path $root 'docs/regression')).Path
$relativeDirectory = $EvidenceDirectory
$directory = [IO.Path]::GetFullPath((Join-Path $root $relativeDirectory))
if (-not $directory.StartsWith($regressionRoot+[IO.Path]::DirectorySeparatorChar, [StringComparison]::OrdinalIgnoreCase)) {
    throw 'Укажите новый каталог доказательств внутри docs/regression.'
}
if (Test-Path -LiteralPath $directory) { throw 'Каталог доказательств уже существует.' }
$null = New-Item -ItemType Directory -Path $directory
$source = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$target = Join-Path $directory 'Cancellation.xlsm'
Copy-Item -LiteralPath $source -Destination $target
$memory = Get-CimInstance Win32_PerfFormattedData_PerfOS_Memory
if ($memory.AvailableMBytes -lt 4096 -or ([double]$memory.CommitLimit-[double]$memory.CommittedBytes) -lt 8GB) {
    throw 'Недостаточно свободной памяти для запуска Excel.'
}

# Добавляет проверку в единственную подтверждённую строку модуля.
function Add-TestCancellationTrigger([object]$Project, [string]$Name, [string]$Anchor, [string]$Code, [bool]$After = $false) {
    $module = $Project.VBComponents.Item($Name).CodeModule
    try {
        $lines = [string]($module.Lines(1, $module.CountOfLines)) -split '\r\n|\n|\r'
        $positions = @(for ($i = 0; $i -lt $lines.Count; $i++) {
            if ($lines[$i].Trim() -ieq $Anchor) { $i + 1 } # Редактор VBA может изменить регистр имени.
        })
        if ($positions.Count -ne 1) { throw "Найдено строк для источника отмены в ${Name}: $($positions.Count)." }
        $module.InsertLines($positions[0] + [int]$After, $Code)
    } finally {
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($module)
    }
}

# Проверяет владение созданным процессом Excel по PID окна и времени запуска.
Add-Type 'using System; using System.Runtime.InteropServices; public static class CoupledCancelWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId); }'
$running = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
$excel = $null
$book = $null
$project = $null
$compile = $null
$owner = $null
$ownerStarted = $null
try {
    $excel = New-Object -ComObject Excel.Application
    [uint32]$excelId = 0
    [void][CoupledCancelWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelId)
    if ($excelId -eq 0 -or $excelId -in $running) { throw 'Не удалось выделить собственный процесс Excel.' }
    $owner = Get-Process -Id $excelId -ErrorAction Stop
    $ownerStarted = $owner.StartTime
    $null = $owner.Handle
    $excel.EnableEvents = $false
    $excel.DisplayAlerts = $false
    $excel.AutomationSecurity = 3
    $book = $excel.Workbooks.Open($target, 0, $false)
    if ($book.ReadOnly) { throw 'Книга проверки открыта только для чтения.' }
    $project = $book.VBProject
    $null = $project.VBComponents.Item('modTestCoupledCancellation')
    if ($Verification -eq 'Cancellation') {
        Add-TestCancellationTrigger $project 'modCoupledWorkbookInput' 'settings.LoadCoupledSettings workbook' @'
    If CStr(workbook.Names("rngLoadCombinations").RefersToRange.Cells(2, 1).Value2) = "TEST_SPATIAL_CANCEL_PREP" Then _
        Err.Raise 18, "TestCancellationPreparation", "Interruption at coupled input"
'@
        Add-TestCancellationTrigger $project 'CBatchSectionCalculator' 'Dim combined As CCombinationResult: Set combined = mResults(index)' @'
    If mInputs(index).CombinationID = "TEST_SPATIAL_CANCEL_SECOND" Then
        If index <> 2 Or mResults(1).StrengthResult.DirectState.Status <> "OK" Then _
            Err.Raise 5, "TestCancellationSecond", "The preceding combination did not finish"
        Err.Raise 18, "TestCancellationSecond", "Interruption at the second combination"
    End If
'@
    } else {
        $trigger = @'
    If Not prepared Is Nothing Then
        If CStr(workbook.Names("rngLoadCombinations").RefersToRange.Cells(2, 1).Value2) = "TEST_PUBLICATION_WRITE_FAILURE" Then _
            Err.Raise 5, "TestPublicationAfterBase", "TestPublicationAfterBase: injected after worksheet base write"
    End If
'@
        Add-TestCancellationTrigger $project 'modCalculationSnapshot' `
            'writer.RestoreSnapshotBase workbook, snapshot, fullFormatting, mode = "Excel"' $trigger $true
    }
    $excel.VBE.ActiveVBProject = $project
    $compile = $excel.VBE.CommandBars.FindControl(1, 578)
    if ($null -eq $compile) { throw 'Команда компиляции VBA недоступна.' }
    if ($compile.Enabled) { $compile.Execute() }
    if ($compile.Enabled) { throw 'Не удалось скомпилировать книгу проверки отмены.' }
    $book.Save()
} finally {
    if ($null -ne $ownerStarted) {
        $current = Get-Process -Id $owner.Id -ErrorAction SilentlyContinue
        if ($null -ne $current -and $current.ProcessName -eq 'EXCEL' -and $current.StartTime -eq $ownerStarted) {
            if ($null -ne $book) { $book.Close($false) }
            $excel.Quit()
        }
    }
    foreach ($object in @($compile, $project, $book, $excel)) {
        if ($null -ne $object) { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($object) }
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
if (-not $owner.WaitForExit(5000)) { throw 'Закрытие процесса Excel, подготовившего тестовую книгу, не подтверждено.' }
if ($sourceHash -cne (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash) {
    throw 'Хеш исходной книги изменился на этапе подготовки теста.'
}

$windowsPowerShell = Join-Path $env:SystemRoot 'System32/WindowsPowerShell/v1.0/powershell.exe'
$runner = Join-Path $PSScriptRoot 'Run-Audit03Watchdog.ps1'
$macro = if ($Verification -eq 'PublicationFailure') { 'RunTestResultPublicationFailure' } else { 'RunTestCoupledCancellation' }
$total = if ($Verification -eq 'PublicationFailure') { 'TOTAL_RESULT_PUBLICATION' } else { 'TOTAL_COUPLED_CANCELLATION' }
$failed = @()
foreach ($storage in @('Excel','File')) {
    $report = Join-Path $relativeDirectory ($storage+'.txt')
    & $windowsPowerShell -NoProfile -ExecutionPolicy Bypass -File $runner -SourceWorkbook $target `
        -Mode No -ReportPath $report -Macro ('modTestCoupledCancellation.'+$macro) -MacroArgument1 $storage
    $code = $LASTEXITCODE
    $closure = Get-Content -LiteralPath (Join-Path $root ($report+'.ProcessClosure.json')) -Raw -Encoding UTF8 | ConvertFrom-Json
    if (-not $closure.Passed) { throw "Закрытие собственного экземпляра Excel после $storage не подтверждено." }
    $expected = if ($Verification -eq 'PublicationFailure') {
        if ($storage -eq 'File') { 7 } else { 6 }
    } elseif ($storage -eq 'File') { 16 } else { 13 }
    if ($code -ne 0 -or (Get-Content -LiteralPath (Join-Path $root $report) -Encoding UTF8) `
            -cnotcontains "${total}: passed=$expected; failed=0") {
        $failed += $storage
    }
}
if ($failed.Count) { throw "Проверка $Verification не пройдена в режимах $($failed -join ', ')." }
