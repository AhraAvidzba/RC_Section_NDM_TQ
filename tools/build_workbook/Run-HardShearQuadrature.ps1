# Сравнивает численные варианты HardShear в отдельных экземплярах Excel.
# Рабочая книга и её исходники в эксперименте не изменяются.
param(
    [Parameter(Mandatory=$true)][string]$SourceWorkbook,
    [Parameter(Mandatory=$true)][string]$ReportDirectory,
    [ValidateSet(0,5,7,8)][int]$GaussOrder = 0,
    [string]$GradientPrototype = ''
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$root = (Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
$source = (Resolve-Path -LiteralPath $SourceWorkbook).Path
$prototype = $null
if ($GradientPrototype) {
    if ($GaussOrder -ne 0) { throw 'Change only one numerical factor in this experiment.' }
    $prototype = Get-Item -LiteralPath $GradientPrototype
    if ($prototype.Name -cne 'CCoupledNewtonSolver.cls') { throw 'Expected the isolated axial-gradient prototype.' }
}
$directory = [IO.Path]::GetFullPath((Join-Path $root $ReportDirectory))
$regression = (Resolve-Path -LiteralPath (Join-Path $root 'docs/regression')).Path + [IO.Path]::DirectorySeparatorChar
if (-not $directory.StartsWith($regression,[StringComparison]::OrdinalIgnoreCase)) {
    throw 'Use a private evidence directory under docs/regression.'
}
$null = New-Item -ItemType Directory -Path $directory -Force
$bookPath = Join-Path $directory 'HardShear.xlsm'
if (Test-Path -LiteralPath $bookPath) { throw 'The experiment workbook already exists.' }
Copy-Item -LiteralPath $source -Destination $bookPath
$sourceHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash

$modulePath = ''
if ($GaussOrder -ne 0) {
    & python (Join-Path $root 'tools/validate_hardshear_quadrature.py') --order $GaussOrder --output $directory
    if ($LASTEXITCODE -ne 0) { throw 'The independent quadrature calculation failed.' }
    $modulePath = Join-Path $directory 'CCoupledSectionQuadrature.cls'
} elseif ($prototype) {
    $modulePath = Join-Path $directory $prototype.Name
    Copy-Item -LiteralPath $prototype.FullName -Destination $modulePath
}
if ($modulePath) {
    $moduleName = [IO.Path]::GetFileNameWithoutExtension($modulePath)
    $moduleText = [IO.File]::ReadAllText($modulePath, [Text.Encoding]::UTF8)
    $body = [regex]::Match($moduleText, '(?ms)^Option Explicit.*').Value.Trim()
    if (-not $body) { throw 'The experiment source has no VBA body.' }
    $printAreas = @(Get-WorkbookPrintAreas $bookPath)
    $existingProcesses = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue | ForEach-Object Id)
    Add-Type 'using System; using System.Runtime.InteropServices; public static class QuadratureExcelOwner { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId); }'
    $excel = New-Object -ComObject Excel.Application
    $owned = $false
    $book = $null
    try {
        [uint32]$excelId = 0
        [void][QuadratureExcelOwner]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd, [ref]$excelId)
        if (-not $excelId -or $excelId -in $existingProcesses) { throw 'Could not identify the new Excel instance.' }
        $owned = $true
        $excel.EnableEvents = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 3
        $book = $excel.Workbooks.Open($bookPath, 0, $false)
        if ($book.ReadOnly) { throw 'Experiment workbook is read-only.' }
        $module = $book.VBProject.VBComponents.Item($moduleName).CodeModule
        $module.DeleteLines(1, $module.CountOfLines)
        $module.AddFromString(($body -split '\r?\n') -join [Environment]::NewLine)
        $excel.VBE.ActiveVBProject = $book.VBProject
        $compile = $excel.VBE.CommandBars.FindControl(1, 578)
        if ($null -eq $compile) { throw 'VBA compile command is unavailable.' }
        if ($compile.Enabled) { $compile.Execute() }
        if ($compile.Enabled) { throw 'Experiment VBA compilation failed.' }
        $book.Save()
    }
    finally {
        if ($owned) {
            if ($null -ne $book) { $book.Close($false) }
            $excel.Quit()
        }
        [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    Restore-WorkbookPrintAreas $bookPath $printAreas
}

$reportPath = Join-Path $ReportDirectory 'Native.txt'
$testArguments = @{
    SourceWorkbook = $bookPath; ReportPath = $reportPath
    Macro = 'modTestCoupledElasticState.RunTestCoupledWorkbookTiming'
    MacroArgument1 = 'Both'; MacroArgument2 = 'Excel'
    VerifyResultsReopen = $true; VerifyStatusReopen = $true
}
& (Join-Path $PSScriptRoot 'Run-Audit02ModeTests.ps1') @testArguments
$testExit = [int]$LASTEXITCODE
$report = Get-Content -LiteralPath (Join-Path $root $reportPath) -Encoding UTF8
$fixtureLine = $report | Where-Object { $_.StartsWith('FIXTURE: ') } | Select-Object -First 1
if ($fixtureLine) {
    $fixture = $fixtureLine.Substring('FIXTURE: '.Length)
    $execution = Join-Path (Split-Path -Parent $fixture) 'RC_Section_NDM_execution_report.txt'
    if (Test-Path -LiteralPath $execution) {
        Copy-Item -LiteralPath $execution -Destination (Join-Path $directory 'Native.execution.txt')
    }
}
if ($sourceHash -cne (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash) {
    throw 'Source workbook changed during the experiment.'
}
exit $testExit
