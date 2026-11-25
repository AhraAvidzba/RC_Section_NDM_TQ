# Updates source modules in an isolated workbook; never writes the production file.
param(
    [Parameter(Mandatory=$true)][string]$WorkbookPath,
    [string[]]$ModuleNames = @(),
    [string[]]$RetiredModules = @(),
    [string[]]$EntryPointNames = @(),
    [string[]]$ExampleFiles = @(),
    [string[]]$RecipeRegistration = @(),
    [switch]$TraceGeometry,
    [switch]$WorkbookEvents
)
$ErrorActionPreference='Stop'
$root=(Resolve-Path -LiteralPath (Join-Path $PSScriptRoot '../..')).Path
. (Join-Path $PSScriptRoot 'GeometryExampleSupport.ps1')
. (Join-Path $PSScriptRoot 'SettingsCatalog.ps1')
$path=(Resolve-Path -LiteralPath (Join-Path $root $WorkbookPath)).Path
$allowed=[IO.Path]::GetFullPath((Join-Path $root 'docs/regression'))+[IO.Path]::DirectorySeparatorChar
if (-not $path.StartsWith($allowed,[StringComparison]::OrdinalIgnoreCase)) {throw 'Only isolated evidence workbooks may be updated.'}
$printAreas=@(Get-WorkbookPrintAreas $path)
$production=Join-Path $root 'workbook/output/RC_Section_NDM.xlsm'
function Get-OpenWorkbookHash([string]$Path) {
    # Excel держит исходную книгу открытой. Разрешение совместного чтения
    # не меняет файл; сравнение до/после всё равно обнаруживает его сохранение.
    $stream=[IO.File]::Open($Path,[IO.FileMode]::Open,[IO.FileAccess]::Read,[IO.FileShare]::ReadWrite)
    $digest=[Security.Cryptography.SHA256]::Create()
    try {[BitConverter]::ToString($digest.ComputeHash($stream)).Replace('-','')}
    finally {$digest.Dispose();$stream.Dispose()}
}
$productionHash=Get-OpenWorkbookHash $production
if ($EntryPointNames.Count -and ($ModuleNames.Count -ne 1 -or $RetiredModules.Count -or $TraceGeometry)) {throw 'Entry-point instrumentation requires exactly one module and no retirement/trace.'}
foreach ($name in $EntryPointNames) {if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$') {throw ('Invalid entry-point name: '+$name)}}
$sourceFiles=@(Get-ChildItem -LiteralPath (Join-Path $root 'src'),(Join-Path $root 'tests') -Recurse -File | Where-Object Extension -in '.bas','.cls')
$sourceFiles+=@(Get-GeometryExampleFiles $root $ExampleFiles)
if ($RecipeRegistration.Count -and ('CSectionTypeRegistry' -notin $ModuleNames -or $RecipeRegistration[1] -notin $ModuleNames)) {throw 'Example registration must import both registry and recipe.'}
foreach ($name in $RetiredModules) {
    if ($name -notmatch '^[A-Za-z][A-Za-z0-9_]*$' -or $name -in $ModuleNames) {throw ('Invalid retirement: '+$name)}
    if (@($sourceFiles | Where-Object BaseName -eq $name).Count) {throw ('Retired source still exists: '+$name)}
}
Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class PrivateCodeWindow { [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId); }'
$excel=$null;$book=$null;$ownedPid=0;$ownedTicks=0;$taskCloseAnchor=$null
$preexisting=@(Get-Process Excel -ErrorAction SilentlyContinue)
try {
    $excel=New-Object -ComObject Excel.Application
    [uint32]$id=0;[void][PrivateCodeWindow]::GetWindowThreadProcessId([IntPtr]$excel.Hwnd,[ref]$id)
    if ($id -le 0 -or $id -in $preexisting.Id) {throw 'Excel compiler window is not a newly created process.'}
    $worker=Get-Process -Id $id -ErrorAction Stop
    if ($worker.ProcessName -ine 'EXCEL') {throw 'Compiler window has no Excel owner.'}
    $ownedPid=[int]$id;$ownedTicks=$worker.StartTime.ToUniversalTime().Ticks
    $excel.Visible=$false;$excel.DisplayAlerts=$false;$excel.EnableEvents=$false;$excel.AutomationSecurity=3
    Write-Output ('OWNED_EXCEL: id='+$ownedPid+'; utcStartTicks='+$ownedTicks)
    $book=$excel.Workbooks.Open($path,0,$false)
    if ($book.ReadOnly) {throw 'Private workbook is read-only.'}
    $updated=@()
    $removed=@()
    # ДЛЯ ТЕСТОВ: Excel может вернуть пустой проект до завершения загрузки VBE; это не ошибка численного кода.
    $book.Activate()
    $project=$null;$components=$null
    for ($projectProbe=0;$projectProbe -lt 20;$projectProbe++) {
        $project=$book.VBProject
        if ($project) {$components=$project.VBComponents}
        if ($components) {break}
        Start-Sleep -Milliseconds 250
    }
    if (!$components) {throw 'VBA project of the owned private workbook is unavailable after loading.'}
    $excel.VBE.ActiveVBProject=$project
    $componentNames=@($components | ForEach-Object {$_.Name})
    foreach ($name in $RetiredModules) {
        if ($name -notin $componentNames) {throw ('Retirement component not found: '+$name)}
        $component=$components.Item($name)
        if ($component.Type -notin 1,2) {throw ('Cannot retire a document/designer component: '+$name)}
        $components.Remove($component)
        $removed += $name
    }
    if ($RetiredModules.Count) {$componentNames=@($components | ForEach-Object {$_.Name})}
    foreach ($file in $sourceFiles) {
        if ($ModuleNames.Count -and $file.BaseName -notin $ModuleNames) {continue}
        $body=[regex]::Match([IO.File]::ReadAllText($file.FullName,[Text.Encoding]::UTF8),'(?ms)^Option Explicit.*').Value
        if (-not $body -and $file.BaseName -eq 'modSectionRecipeApiGuide') {
            $body=([IO.File]::ReadAllLines($file.FullName,[Text.Encoding]::UTF8) | Where-Object {$_ -notmatch '^Attribute VB_Name = '}) -join "`r`n"
            if (@($body -split "`r?`n" | Where-Object {$_.Trim() -and $_ -notmatch "^\s*'"}).Count) {throw 'API guide must contain comments only.'}
        }
        if (-not $body) {throw ('Source body missing: '+$file.FullName)}
        if ($file.BaseName -eq 'CSectionTypeRegistry') {$body=Add-GeometryExampleRegistration $body $RecipeRegistration}
        if ($TraceGeometry -and $file.BaseName -eq 'CSectionGeometryQuery') {
            $body=$body.Replace('Option Explicit',"Option Explicit`r`nPrivate mCutTrace As String")
            $body=$body.Replace('    Set ClipHalfPlane = MakeRegion',"    mCutTrace = TestCutTrace(raw, rawCount, nx, ny, bound, lineCuts, lineCount)`r`n    Set ClipHalfPlane = MakeRegion")
            $needle='If candidate = 0 Then GeometryError'
            if ([regex]::Matches($body,[regex]::Escape($needle)).Count -ne 1) {throw 'Geometry trace target mismatch.'}
            $body=$body.Replace($needle,'If candidate = 0 Then GeometryError TestFailedRegionTrace(edges, count, used) & ')
            $body += @'

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выводит использованные и оставшиеся рёбра неудачного обхода области только в частной диагностической копии.
Private Function TestFailedRegionTrace(ByRef edges() As TRegionEdge, ByVal count As Long, ByRef used() As Boolean) As String
    Dim i As Long, result As String
    result = mCutTrace & "REGION_TRACE; count=" & CStr(count) & vbCrLf
    For i = 1 To count
        result = result & "EDGE|" & CStr(i) & "|used=" & CStr(used(i)) & "|" & _
            CStr(edges(i).X1) & "|" & CStr(edges(i).Y1) & "|" & _
            CStr(edges(i).X2) & "|" & CStr(edges(i).Y2) & "|sweep=" & _
            CStr(edges(i).Sweep) & vbCrLf
    Next i
    TestFailedRegionTrace = result
End Function
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает пересечения и проверку середины интервалов отсечения для разбора построения частной области.
Private Function TestCutTrace(ByRef raw() As TRegionEdge, ByVal rawCount As Long, ByVal nx As Double, ByVal ny As Double, ByVal bound As Double, ByRef cuts() As Double, ByVal count As Long) As String
    Dim i As Long, s As Double, x As Double, y As Double, result As String
    result = "CUT_TRACE|nx=" & CStr(nx) & "|ny=" & CStr(ny) & "|bound=" & CStr(bound) & "|cuts=" & CStr(count) & vbCrLf
    For i = 1 To count
        result = result & "CUT|" & CStr(i) & "|s=" & CStr(cuts(i)) & vbCrLf
        If i < count Then
            s = (cuts(i) + cuts(i + 1)) * 0.5
            x = nx * bound + ny * s: y = ny * bound - nx * s
            result = result & "MID|" & CStr(x) & "|" & CStr(y) & "|inside=" & CStr(ContainsRaw(raw, rawCount, x + nx * QUERY_TOLERANCE * 8#, y + ny * QUERY_TOLERANCE * 8#, False)) & vbCrLf
        End If
    Next i
    TestCutTrace = result
End Function
'@
        }
        if ($file.BaseName -in $componentNames) {
            for($componentReadAttempt=0;$componentReadAttempt -lt 3;$componentReadAttempt++) {
                try {$component=$components.Item($file.BaseName);break}
                catch {
                    if($_.Exception.HResult -ne -2147417856 -or $componentReadAttempt -eq 2){throw}
                    Write-Output ('COMPONENT_READ_RETRY: '+$file.BaseName)
                    Start-Sleep -Milliseconds 250
                }
            }
        } else {
            $component=$components.Add($(if ($file.Extension -eq '.cls') {2} else {1}))
            $component.Name=$file.BaseName
            $componentNames+=,$file.BaseName
        }
        $module=$component.CodeModule
        if ($EntryPointNames.Count) {
            $original=[string]$module.Lines(1,$module.CountOfLines)
            $entries=@()
            foreach ($name in $EntryPointNames) {
                $declaration='(?m)^Public (?:Function|Sub) '+[regex]::Escape($name)+'\('
                if ([regex]::IsMatch($original,$declaration)) {throw ('Entry point already exists in baseline: '+$name)}
                $matches=[regex]::Matches($body,'(?ms)^Public (Function|Sub) '+[regex]::Escape($name)+'\(.*?^End \1\b')
                if ($matches.Count -ne 1) {throw ('Source entry point is not unique: '+$name)}
                $entries+=,"' FOR TESTS: baseline entry point only; existing methods retained.`r`n"+$matches[0].Value
            }
            $body=$entries -join "`r`n`r`n"
        } else {
            # ДЛЯ ТЕСТОВ: повторяется полная замена частного модуля, а не добавление к возможному частичному результату.
            for($replaceAttempt=0;$replaceAttempt -lt 3;$replaceAttempt++) {
                try {
                    $module=$components.Item($file.BaseName).CodeModule
                    if($module.CountOfLines){$module.DeleteLines(1,$module.CountOfLines)}
                    $module.AddFromString(($body -split '\r?\n') -join "`r`n")
                    break
                } catch {
                    if($_.Exception.HResult -ne -2147417856 -or $replaceAttempt -eq 2){throw}
                    Write-Output ('COMPLETE_MODULE_REPLACE_RETRY: '+$file.BaseName)
                    Start-Sleep -Milliseconds 250
                }
            }
        }
        if($EntryPointNames.Count){$module.AddFromString(($body -split '\r?\n') -join "`r`n")}
        $updated += $file.BaseName
    }
    if ($ModuleNames.Count -and @($ModuleNames | Where-Object {$_ -notin $updated}).Count) {throw 'Requested source module was not found.'}
    if ($WorkbookEvents) {
        # Обработчики из исходника полной сборки принадлежат документному модулю книги.
        $documentName=[string]$book.CodeName
        if ([string]::IsNullOrWhiteSpace($documentName)) {$documentName='ThisWorkbook'}
        $module=$components.Item($documentName).CodeModule
        $events=[IO.File]::ReadAllText((Join-Path $PSScriptRoot 'WorkbookHelpEvents.vba'),[Text.Encoding]::UTF8)
        if ($module.CountOfLines) {$module.DeleteLines(1,$module.CountOfLines)}
        $module.AddFromString(($events -split '\r?\n') -join "`r`n")
        $updated += $documentName
    }
    $excel.VBE.ActiveVBProject=$project
    $compile=$excel.VBE.CommandBars.FindControl(1,578)
    if ($null -eq $compile) {throw 'Compile command unavailable.'}
    if ($compile.Enabled) {$compile.Execute()}
    # ДЛЯ ТЕСТОВ: ожидание завершения команды ограничено; активная команда после него остаётся ошибкой сборки.
    for ($taskCompileProbe=0;$taskCompileProbe -lt 40 -and $compile.Enabled;$taskCompileProbe++) {
        Start-Sleep -Milliseconds 250
    }
    if ($compile.Enabled) {
        $pane=$excel.VBE.ActiveCodePane
        if ($null -ne $pane) {
            [int]$firstLine=0;[int]$firstColumn=0;[int]$lastLine=0;[int]$lastColumn=0
            $pane.GetSelection([ref]$firstLine,[ref]$firstColumn,[ref]$lastLine,[ref]$lastColumn)
            Write-Output ('COMPILE_LOCATION: module='+$pane.CodeModule.Parent.Name+'; line='+$firstLine+'; column='+$firstColumn)
            Write-Output $pane.CodeModule.Lines([Math]::Max(1,$firstLine-2),5)
        } else {Write-Output 'COMPILE_LOCATION_UNAVAILABLE: active code pane is null; compilation not accepted.'}
        throw 'Private VBA compilation failed or has not completed within the bounded wait.'
    }
    # ДЛЯ ТЕСТОВ: целевая книга закрывается при существующей другой книге только в этом собственном процессе Excel.
    $taskCloseAnchor=$excel.Workbooks.Add()
    $taskCloseAnchor.Saved=$true
    Write-Output 'COMPILE_OWN_EMPTY_CLOSE_ANCHOR_CREATED'

    # ДЛЯ ТЕСТОВ: проект уже скомпилирован; ссылки его объектов освобождаются до закрытия книги-владельца.
    foreach ($taskVbeObject in @($compile,$module,$component,$components,$project)) {
        if ($null -ne $taskVbeObject -and [Runtime.InteropServices.Marshal]::IsComObject($taskVbeObject)) {
            [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($taskVbeObject)
        }
    }
    $taskVbeObject=$null;$compile=$null;$module=$null;$component=$null;$components=$null;$project=$null
    Write-Output 'COMPILE_VBE_REFERENCES_RELEASED'
    Write-Output 'COMPILE_BEFORE_SAVE'
    $book.Save()
    Write-Output 'COMPILE_AFTER_SAVE_BEFORE_CLOSE'

    $book.Close($false);$book=$null
    Write-Output 'COMPILE_AFTER_CLOSE'
    Restore-WorkbookPrintAreas $path $printAreas
    Write-Output 'COMPILE_PRINT_AREAS_RESTORED'
    if ((Get-OpenWorkbookHash $production) -cne $productionHash) {throw 'Production changed during preparation.'}
    [ordered]@{Workbook=$path;WorkbookSHA256=(Get-FileHash -LiteralPath $path).Hash;ProductionSHA256=$productionHash;Compiled=$true;Modules=$updated;RetiredModules=$removed;EntryPointNames=$EntryPointNames;TraceGeometry=[bool]$TraceGeometry;ProductionChanged=$false} |
        ConvertTo-Json -Depth 4 | Set-Content -LiteralPath (Join-Path (Split-Path -Parent $path) 'CodeUpdate.json') -Encoding UTF8
    Write-Output ('PRIVATE_CODE_READY: modules='+$updated.Count+'; production unchanged=True')

} finally {
    # ДЛЯ ТЕСТОВ: первичная ошибка остаётся ошибкой сборки; очистка не блокируется ожиданием COM-финализаторов перед остановкой собственного PID.
    if ($book) {try {$book.Close($false)} catch {Write-Output ('OWN_CLOSE_WARNING: '+$_.Exception.Message)}}
    if ($excel -and $ownedPid) {try {$excel.Quit()} catch {Write-Output ('OWN_QUIT_WARNING: '+$_.Exception.Message)}}
    if ($ownedPid) {
        $owned=Get-Process -Id $ownedPid -ErrorAction SilentlyContinue
        if ($owned -and $owned.ProcessName -eq 'EXCEL' -and $owned.StartTime.ToUniversalTime().Ticks -eq $ownedTicks) {Stop-Process -Id $ownedPid -Force}
    }
    $book=$null;$compile=$null;$module=$null;$component=$null;$components=$null;$project=$null;$taskCloseAnchor=$null
    if ($excel) {try {[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)} catch {Write-Output ('OWN_RELEASE_WARNING: '+$_.Exception.Message)}}
    $excel=$null
    [GC]::Collect()
}
