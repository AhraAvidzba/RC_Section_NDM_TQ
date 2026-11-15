Attribute VB_Name = "modTestCoupledCancellation"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Проверяет отмену расчёта при подготовке и на втором сочетании. Сценарий
' запускается в изолированной книге с двумя введёнными тестовым раннером
' источниками ошибки 18. Обычная книга не содержит этих источников.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет предыдущий расчёт и повторяет запуск с отменой в двух местах.
Public Function RunTestCoupledCancellation(ByVal storage As String) As String
    Dim checks As Long, report As String
    On Error GoTo Failed
    Dim book As Workbook: Set book = ThisWorkbook
    Dim initialCancel As XlEnableCancelKey: initialCancel = book.Application.EnableCancelKey
    If InStr(1, book.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestCoupledCancellation", "Запустите тест из изолированной копии книги."
    If storage <> "Excel" And storage <> "File" Then _
        Err.Raise 5, "RunTestCoupledCancellation", "Выберите Excel или File."
    modTestPerformance.TestConfigurePerformanceFixture "NDS", 1, "Newton", "Rectangles", "Normal", "No"
    Dim settings As Range: Set settings = book.Names("rngSystemSettings").RefersToRange
    TestCancellationSetting(settings, "Results.SnapshotStorage").Value2 = storage
    TestCancellationSetting(settings, "Results.SnapshotName").Value2 = "TestCoupledCancellation_" & storage
    TestCancellationSetting(settings, "General.ExecutionReportEnabled").Value2 = "No"
    TestCancellationSetting(settings, "Plot.Enabled").Value2 = "No"

    Dim loads As Range: Set loads = book.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 1).Value2 = "TEST_CANCEL_BASE"
    FormSectionForWorkbook book
    RunSectionCalculationForWorkbook book
    TestCancellationRequire ReadLoadCasePropertyText(book, "TEST_CANCEL_BASE", "DirectStateStatus") = "OK", _
        storage & " | сохранён исходный расчёт", checks, report

    Dim geometry As Variant, properties As Variant, material As Variant, elements As Variant
    geometry = ReadPhysicalResultTable(book, "rngNDMSectionGeometry")
    properties = ReadPhysicalResultTable(book, "rngNDMSectionProperties")
    material = ReadPhysicalResultTable(book, "rngNDMMaterialDiagrams")
    elements = ReadPhysicalResultTable(book, "rngNDMElementResults")
    Dim summary As Range, strength As Range, summaryData As Variant, strengthData As Variant
    Dim summaryWriter As New CBatchResultWriter, strengthWriter As New CStrengthSummaryWriter
    Set summary = book.Names("rngBatchSummary").RefersToRange.Offset(5, 0).Resize(37, summaryWriter.RequiredSummaryColumns)
    Set strength = book.Names("rngStrengthSummaryAnchor").RefersToRange.Resize(30, strengthWriter.RequiredColumns)
    summaryData = summary.Value2: strengthData = strength.Value2

    Dim store As New CCalculationSnapshotStore, checksum As New CFileChecksum, path As String, hash As String
    If storage = "File" Then
        path = store.SnapshotPath(book): hash = checksum.FileSHA256(path)
        TestCancellationRequire Len(hash) = 64, storage & " | зафиксирован SHA-256 исходного снимка", checks, report
    End If

    Dim stage As Variant, number As Long, description As String, appState As Variant
    For Each stage In Array("Prepare", "Second")
        loads.Cells(2, 1).Value2 = "TEST_CANCEL_BASE"
        If CStr(stage) = "Prepare" Then loads.Cells(2, 1).Value2 = "TEST_SPATIAL_CANCEL_PREP"
        loads.Cells(2, 2).Value2 = 0.9 * CDbl(loads.Cells(2, 2).Value2)
        loads.Cells(3, 1).Value2 = "TEST_SPATIAL_CANCEL_SECOND"
        loads.Cells(3, 5).Value2 = loads.Cells(2, 5).Value2
        loads.Cells(3, 8).Value2 = 0.001 ' Ненулевой T включает пространственный маршрут подготовки и второго сочетания.
        book.Application.EnableCancelKey = xlInterrupt
        appState = TestCancellationAppState(book.Application)
        number = TestInvokeCancelledCalculation(book, description)

        Dim label As String: label = storage & " | " & CStr(stage)
        TestCancellationRequire number = 18 And InStr(1, description, "Расчёт отменён", vbBinaryCompare) > 0 And _
            InStr(1, summaryWriter.FreshnessNotice(book), "ранее сохранённые", vbBinaryCompare) > 0, _
            label & " | номер отмены, сообщение и отметка прежних результатов", checks, report
        TestCancellationRequire TestCancellationStateEqual(book.Application, appState), _
            label & " | настройки Excel восстановлены", checks, report
        TestCancellationRequire TestCalculationSnapshotReadReleased(), label & " | контекст снимка освобождён", checks, report
        TestCancellationRequire SnapshotMatricesEqual(summaryData, summary.Value2) And SnapshotMatricesEqual(strengthData, strength.Value2), _
            label & " | значения обеих сводок сохранены", checks, report
        TestCancellationRequire SnapshotMatricesEqual(geometry, ReadPhysicalResultTable(book, "rngNDMSectionGeometry")) And _
            SnapshotMatricesEqual(properties, ReadPhysicalResultTable(book, "rngNDMSectionProperties")), _
            label & " | RunID, свойства и геометрия сохранены", checks, report
        TestCancellationRequire SnapshotMatricesEqual(material, ReadPhysicalResultTable(book, "rngNDMMaterialDiagrams")) And _
            SnapshotMatricesEqual(elements, ReadPhysicalResultTable(book, "rngNDMElementResults")), _
            label & " | диаграммы и результаты элементов сохранены", checks, report
        If storage = "File" Then
            TestCancellationRequire hash = checksum.FileSHA256(path) And _
                Not CreateObject("Scripting.FileSystemObject").FileExists(path & ".part"), _
                label & " | файл снимка сохранён, временного файла нет", checks, report
        End If
    Next stage
    book.Application.EnableCancelKey = initialCancel
    RunTestCoupledCancellation = report & "TOTAL_COUPLED_CANCELLATION: passed=" & CStr(checks) & "; failed=0"
    Exit Function
Failed:
    RunTestCoupledCancellation = report & "FAIL: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_COUPLED_CANCELLATION: passed=" & CStr(checks) & "; failed=1"
    On Error Resume Next
    If Not book Is Nothing Then book.Application.EnableCancelKey = initialCancel
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет актуальность, ошибки и частичный результат с обычными событиями книги.
Public Function RunTestResultFreshness(ByVal storage As String) As String
    Dim checks As Long, report As String
    On Error GoTo Failed
    Dim book As Workbook: Set book = ThisWorkbook
    Dim initialEvents As Boolean: initialEvents = book.Application.EnableEvents
    If InStr(1, book.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestResultFreshness", "Запустите тест из изолированной копии книги."
    If storage <> "Excel" And storage <> "File" Then Err.Raise 5, "RunTestResultFreshness", "Выберите Excel или File."
    modTestPerformance.TestConfigurePerformanceFixture "NDS", 1, "Newton", "Rectangles", "Normal", "No"
    Dim settings As Range: Set settings = book.Names("rngSystemSettings").RefersToRange
    TestCancellationSetting(settings, "Results.SnapshotStorage").Value2 = storage
    TestCancellationSetting(settings, "Results.SnapshotName").Value2 = "TestResultFreshness_" & storage
    TestCancellationSetting(settings, "General.ExecutionReportEnabled").Value2 = "No"
    TestCancellationSetting(settings, "Plot.Enabled").Value2 = "No"
    Dim loads As Range: Set loads = book.Names("rngLoadCombinations").RefersToRange
    loads.Cells(2, 1).Value2 = "TEST_FRESH_BASE"
    Dim summary As New CBatchResultWriter
    FormSectionForWorkbook book
    TestCancellationRequire Len(summary.FreshnessNotice(book)) = 0, storage & " | геометрия без отметки расчёта", checks, report
    RunSectionCalculationForWorkbook book
    TestCancellationRequire Len(summary.FreshnessNotice(book)) = 0 And _
        Len(ReadLoadCasePropertyText(book, "ALL", RESULT_INPUT_STAMP)) > 64, _
        storage & " | сохранён согласованный исходный расчёт", checks, report

    Dim geometry As Variant, store As New CCalculationSnapshotStore, checksum As New CFileChecksum, path As String, hash As String
    geometry = ReadPhysicalResultTable(book, "rngNDMSectionGeometry")
    If storage = "File" Then path = store.SnapshotPath(book): hash = checksum.FileSHA256(path)
    book.Application.EnableEvents = True
    TestCancellationSetting(settings, "General.WorstCombinationCriterion").Value2 = "Cracks"
    TestCancellationSetting(settings, "Output.CombinationID").Value2 = "TEST_FRESH_BASE"
    TestChangeFreshnessVisualization book
    UpdateSectionResultsForWorkbook book
    TestCancellationRequire Len(summary.FreshnessNotice(book)) = 0, storage & " | выбор схемы и Worst не меняет актуальность", checks, report

    Dim previousQx As Variant: previousQx = loads.Cells(2, 9).Value2
    loads.Cells(2, 9).Value2 = 0.001 ' Qx меняется при неизменной исходной геометрии.
    TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Исходные данные изменены", vbBinaryCompare) = 1, _
        storage & " | SheetChange обнаруживает изменение Qx", checks, report
    UpdateSectionResultsForWorkbook book
    TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Исходные данные изменены", vbBinaryCompare) = 1 And _
        SnapshotMatricesEqual(geometry, ReadPhysicalResultTable(book, "rngNDMSectionGeometry")), _
        storage & " | ручное обновление сохраняет отметку и RunID", checks, report
    If storage = "File" Then TestCancellationRequire checksum.FileSHA256(path) = hash, storage & " | .ndm не изменён", checks, report
    loads.Cells(2, 9).Value2 = previousQx
    TestCancellationRequire Len(summary.FreshnessNotice(book)) = 0, storage & " | восстановление исходного Qx", checks, report

    book.Application.EnableEvents = False
    loads.Cells(2, 9).Value2 = 0.001
    RefreshVisibleResultFreshness book ' Та же проверка, которую вызывает Workbook_SheetActivate.
    TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Исходные данные изменены", vbBinaryCompare) = 1, _
        storage & " | сверка замечает ввод без события редактирования", checks, report
    loads.Cells(2, 9).Value2 = previousQx
    book.Application.EnableEvents = True

    loads.Cells(3, 1).Value2 = "TEST_FRESH_BASE"
    Dim reason As String, number As Long
    number = TestInvokeCancelledCalculation(book, reason)
    TestCancellationRequire number <> 0 And InStr(1, summary.FreshnessNotice(book), "Новый расчёт не выполнен", vbBinaryCompare) = 1, _
        storage & " | ошибка чтения нагрузок отмечает прежний результат", checks, report
    loads.Cells(3, 1).ClearContents
    UpdateSectionResultsForWorkbook book
    TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Новый расчёт не выполнен", vbBinaryCompare) = 1, _
        storage & " | обновление того же RunID сохраняет исход запуска", checks, report
    RunSectionCalculationForWorkbook book
    TestCancellationRequire Len(summary.FreshnessNotice(book)) = 0, storage & " | новый успешный запуск сбрасывает сообщение", checks, report

    loads.Cells(3, 1).Value2 = "TEST_FRESH_INVALID"
    loads.Cells(3, 2).Value2 = "invalid"
    RunSectionCalculationForWorkbook book
    TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Часть проверок не завершена", vbBinaryCompare) = 1 And _
        ReadLoadCasePropertyText(book, "ALL", RESULT_INCOMPLETE_COUNT) = "1", _
        storage & " | сохранён частичный расчёт с недоступной проверкой", checks, report
    UpdateSectionResultsForWorkbook book
    TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Часть проверок не завершена", vbBinaryCompare) = 1, _
        storage & " | отметка частичного расчёта восстановлена из того же снимка", checks, report

    If storage = "File" Then
        TestCancellationSetting(settings, "Results.SnapshotName").Value2 = "MissingFreshnessSnapshot"
        number = TestInvokeFreshnessUpdate(book)
        TestCancellationRequire number <> 0 And InStr(1, summary.FreshnessNotice(book), "Не удалось прочитать выбранный снимок", vbBinaryCompare) = 1, _
            storage & " | отсутствие файла отмечено без подмены листом", checks, report
        TestCancellationSetting(settings, "Results.SnapshotName").Value2 = "TestResultFreshness_" & storage
        UpdateSectionResultsForWorkbook book
        TestCancellationRequire InStr(1, summary.FreshnessNotice(book), "Часть проверок не завершена", vbBinaryCompare) = 1, _
            storage & " | выбранный файл снова определяет отметку", checks, report
    Else
        MarkResultAttemptStopped book, False, True
        number = TestInvokeFreshnessUpdate(book)
        TestCancellationRequire number <> 0 And InStr(1, summary.FreshnessNotice(book), "Вывод прерван", vbBinaryCompare) = 1, _
            storage & " | чтение незавершённого листа заблокировано", checks, report
    End If
    book.Application.EnableEvents = initialEvents
    RunTestResultFreshness = report & "TOTAL_RESULT_FRESHNESS: passed=" & CStr(checks) & "; failed=0"
    Exit Function
Failed:
    RunTestResultFreshness = report & "FAIL: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_RESULT_FRESHNESS: passed=" & CStr(checks) & "; failed=1"
    On Error Resume Next
    If Not book Is Nothing Then book.Application.EnableEvents = initialEvents
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отказ после записи базы или общей сводки.
' В Excel остаётся прошлый RunID, в File восстанавливается новый авторитетный снимок.
Public Function RunTestResultPublicationFailure(ByVal storage As String) As String
    Dim checks As Long, report As String
    On Error GoTo Failed
    Dim book As Workbook: Set book = ThisWorkbook
    If InStr(1, book.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestResultPublicationFailure", "Запустите тест из изолированной копии книги."
    If storage <> "Excel" And storage <> "File" Then Err.Raise 5, "RunTestResultPublicationFailure", "Выберите Excel или File."
    modTestPerformance.TestConfigurePerformanceFixture "NDS", 1, "Newton", "Rectangles", "Normal", "No"
    Dim settings As Range: Set settings = book.Names("rngSystemSettings").RefersToRange
    TestCancellationSetting(settings, "Results.SnapshotStorage").Value2 = storage
    TestCancellationSetting(settings, "Results.SnapshotName").Value2 = "TestPublicationFailure_" & storage
    TestCancellationSetting(settings, "General.ExecutionReportEnabled").Value2 = "No"
    TestCancellationSetting(settings, "Plot.Enabled").Value2 = "No"
    If storage = "File" Then
        TestCancellationSetting(settings, "Plot.Enabled").Value2 = "Yes"
        TestCancellationSetting(settings, "Plot.AutoUpdateAfterCalculation").Value2 = "Yes"
        TestChangeFreshnessVisualization book
        Dim plotState As Range
        Set plotState = book.Names("rngCalculationProfiles").RefersToRange.Columns(2).Find("Visualization.State", _
            LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
        plotState.Offset(0, 1).Value2 = "StrengthState"
    End If
    Dim loads As Range: Set loads = book.Names("rngLoadCombinations").RefersToRange
    Dim topLC As Range: Set topLC = book.Names("rngBatchSummary").RefersToRange.Offset(12, 0)
    loads.Cells(2, 1).Value2 = "TEST_PUBLICATION_BASE"
    FormSectionForWorkbook book
    RunSectionCalculationForWorkbook book
    TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_BASE", storage & " | исходная сводка сохранена", checks, report

    Dim store As New CCalculationSnapshotStore, checksum As New CFileChecksum, path As String, oldHash As String, currentHash As String
    Dim chart As Chart
    If storage = "File" Then
        path = store.SnapshotPath(book): oldHash = checksum.FileSHA256(path)
        Set chart = book.Worksheets.Item("Расчет").ChartObjects.Item("chtNDMSectionPlot").Chart
        TestCancellationSetting(settings, "Plot.AutoUpdateAfterCalculation").Value2 = "No"
    End If
    Dim sheet As Worksheet: Set sheet = topLC.Worksheet
    Dim reason As String, number As Long, summary As New CBatchResultWriter
    Dim userCell As Range, formulaCell As Range
    Set userCell = book.Names("rngBatchSummary").RefersToRange.Offset(0, summary.RequiredSummaryOutputColumns + 3)
    userCell.Value2 = "Пользовательское значение"
    book.Names.Add "TestPublicationUserValue", "='" & sheet.Name & "'!" & userCell.Address
    Set formulaCell = book.Names("rngStrengthSummaryAnchor").RefersToRange.Offset(0, 1)
    If storage = "Excel" Then formulaCell.Formula = "=TestPublicationUserValue"
    Dim initialCopyObjects As Boolean, copySettingChanged As Boolean, marker As Shape, originalShapes As Long, markerAnchor As String
    If storage = "Excel" Then
        initialCopyObjects = book.Application.CopyObjectsWithCells: copySettingChanged = True
        book.Application.CopyObjectsWithCells = True
        Set marker = sheet.Shapes.AddShape(msoShapeRectangle, CSng(formulaCell.Left + 1#), CSng(formulaCell.Top + 1#), 9!, 9!)
        marker.Name = "TestPublicationShape": marker.Placement = xlMoveAndSize
        markerAnchor = marker.TopLeftCell.Address: originalShapes = sheet.Shapes.Count
    End If
    sheet.Calculate
    Dim saved As Collection: Set saved = TestSavedPublicationTables(book)
    Dim oldNames As Long, oldStyles As Long, oldLinks As String, oldWindows As Long, oldBooks As Long, oldFill As Long
    oldNames = book.Names.Count: oldStyles = book.Styles.Count: oldLinks = TestPublicationLinks(book)
    oldWindows = book.Application.Windows.Count: oldBooks = book.Application.Workbooks.Count
    oldFill = topLC.Offset(0, 3).Interior.Color
    Dim appState As Variant: appState = TestCancellationAppState(book.Application)
    sheet.Protect UserInterfaceOnly:=False
    number = TestInvokeCancelledCalculation(book, reason)
    sheet.Unprotect
    TestCancellationRequire number <> 0 And InStr(reason, "защищ") > 0 And TestSamePublicationTables(saved, TestSavedPublicationTables(book)), _
        storage & " | защищённый лист отклонён до записи", checks, report
    If storage = "Excel" Then
        sheet.Protect UserInterfaceOnly:=True
        number = TestInvokeCancelledCalculation(book, reason)
        TestCancellationRequire number <> 0 And InStr(reason, "вернуть оформление") > 0 And _
            sheet.ProtectContents And sheet.ProtectionMode And TestSamePublicationTables(saved, TestSavedPublicationTables(book)), _
            storage & " | запрет записи сохраняет вид и защиту, разрешающую работу макросов", checks, report
        sheet.Unprotect
    End If
    If storage = "File" Then
        TestCancellationRequire checksum.FileSHA256(path) = oldHash, storage & " | защищённый лист сохраняет прежний файл", checks, report
        Dim fileSystem As Object, occupied As Object
        Set fileSystem = CreateObject("Scripting.FileSystemObject")
        Set occupied = fileSystem.CreateTextFile(path & ".part", False)
        occupied.Write "TestBeforePublication"
        occupied.Close
        number = TestInvokeCancelledCalculation(book, reason)
        Dim refusalPreserved As Boolean
        refusalPreserved = number <> 0 And InStr(1, reason, "Временный файл занят", vbBinaryCompare) > 0 And _
            checksum.FileSHA256(path) = oldHash And fileSystem.FileExists(path & ".part") And _
            InStr(1, summary.FreshnessNotice(book), "Новый расчёт не выполнен", vbBinaryCompare) = 1 And _
            ReadLoadCasePropertyText(book, "TEST_PUBLICATION_BASE", "DirectStateStatus") = "OK"
        fileSystem.DeleteFile path & ".part"
        TestCancellationRequire refusalPreserved, storage & " | отказ до замены файла сохраняет прежний снимок и его чтение", checks, report
    End If
    Dim failures As Variant, expected As Variant, i As Long, message As String
    failures = Array("TEST_PUBLICATION_WRITE_FAILURE", "TEST_PUBLICATION_SUMMARY_FAILURE")
    expected = Array("TestPublicationAfterBase", "TestPublicationDuringSummary")
    For i = 0 To UBound(failures)
        loads.Cells(2, 1).Value2 = failures(i)
        loads.Cells(2, 2).Value2 = 0.9 * CDbl(loads.Cells(2, 2).Value2) ' Меняет значения технических полей нового НДС.
        If storage = "File" Then
            If i = 1 Then sheet.Protect UserInterfaceOnly:=True
            message = RunSectionCalculationForWorkbook(book)
            currentHash = checksum.FileSHA256(path)
            TestCancellationRequire InStr(1, message, expected(i), vbBinaryCompare) > 0 And _
                InStr(message, "вывод восстановлен") > 0 And currentHash <> oldHash, _
                storage & " | автоматически восстановлен новый файл после " & expected(i), checks, report
            TestCancellationRequire topLC.Value2 = failures(i) And Len(summary.FreshnessNotice(book)) = 0 And _
                ReadLoadCasePropertyText(book, CStr(failures(i)), "DirectStateStatus") = "OK" And _
                InStr(TestResultPlotTitle(chart), "Схема требует обновления") = 1, _
                storage & " | сводка и файл согласованы, автоматическая перерисовка отключена", checks, report
            Set saved = TestSavedPublicationTables(book)
            UpdateSectionResultsForWorkbook book
            TestCancellationRequire checksum.FileSHA256(path) = currentHash And TestSamePublicationTables(saved, TestSavedPublicationTables(book)) And _
                InStr(TestResultPlotTitle(chart), CStr(failures(i))) > 0, _
                storage & " | повторный вывод воспроизводит все таблицы без записи файла", checks, report
            If i = 1 Then
                TestCancellationRequire sheet.ProtectContents And sheet.ProtectionMode, _
                    storage & " | восстановление File не изменило защиту, разрешающую работу макросов", checks, report
            End If
            oldHash = currentHash
        Else
            number = TestInvokeCancelledCalculation(book, reason)
            TestCancellationRequire number = 5 And InStr(1, reason, expected(i), vbBinaryCompare) > 0, _
                storage & " | введена ошибка " & expected(i), checks, report
            TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_BASE" And _
                TestSamePublicationTables(saved, TestSavedPublicationTables(book)) And _
                InStr(1, summary.FreshnessNotice(book), "Новый расчёт не выполнен", vbBinaryCompare) = 1, _
                storage & " | полностью восстановлены прежние сводки и технические таблицы | " & _
                Replace(reason, vbCrLf, " ") & " | отметка: " & summary.FreshnessNotice(book), checks, report
            Dim elements As Variant: elements = ReadElementResultTable(book, "TEST_PUBLICATION_BASE", "StrengthState")
            TestCancellationRequire HasResultTableRows(elements) And formulaCell.Formula = "=TestPublicationUserValue" And _
                topLC.Offset(0, 3).Interior.Color = oldFill And book.Names.Count = oldNames And _
                TestPublicationLinks(book) = oldLinks And book.Styles.Count = oldStyles, _
                storage & " | индекс, формула, цвет, ссылки и стили прежней книги сохранены", checks, report
            TestCancellationRequire book.Application.CopyObjectsWithCells And sheet.Shapes.Count = originalShapes And _
                sheet.Shapes.Item("TestPublicationShape").TopLeftCell.Address = markerAnchor, _
                storage & " | фигура сохранила привязку и параметр копирования восстановлен", checks, report
        End If
        TestCancellationRequire TestCancellationStateEqual(book.Application, appState) And TestCalculationSnapshotReadReleased() And _
            book.Application.Workbooks.Count = oldBooks And book.Application.Windows.Count = oldWindows And userCell.Value2 = "Пользовательское значение", _
            storage & " | интерфейс, контекст и сторонняя ячейка сохранены", checks, report
        If storage = "File" And i = 1 Then sheet.Unprotect
    Next i
    If storage = "File" Then
        loads.Cells(2, 1).Value2 = "TEST_PUBLICATION_REPAIR_FAILURE"
        number = TestInvokeCancelledCalculation(book, reason)
        currentHash = checksum.FileSHA256(path)
        TestCancellationRequire number = 5 And InStr(1, reason, "TestPublicationAfterBase", vbBinaryCompare) > 0 And _
            InStr(1, summary.FreshnessNotice(book), "Вывод прерван", vbBinaryCompare) = 1 And currentHash <> oldHash, _
            storage & " | при отказе восстановления сохраняется барьер незавершённого вида", checks, report
        UpdateSectionResultsForWorkbook book
        TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_REPAIR_FAILURE" And Len(summary.FreshnessNotice(book)) = 0 And _
            checksum.FileSHA256(path) = currentHash, storage & " | ручное обновление восстановило лист без перезаписи файла", checks, report
    Else
        loads.Cells(2, 1).Value2 = "TEST_PUBLICATION_RECOVERY"
        RunSectionCalculationForWorkbook book
        TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_RECOVERY" And Len(summary.FreshnessNotice(book)) = 0, _
            storage & " | следующий расчёт опубликовал новый RunID", checks, report
    End If
    If copySettingChanged Then book.Application.CopyObjectsWithCells = initialCopyObjects
    RunTestResultPublicationFailure = report & "TOTAL_RESULT_PUBLICATION: passed=" & CStr(checks) & "; failed=0"
    Exit Function
Failed:
    RunTestResultPublicationFailure = report & "FAIL: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_RESULT_PUBLICATION: passed=" & CStr(checks) & "; failed=1"
    On Error Resume Next
    If copySettingChanged Then book.Application.CopyObjectsWithCells = initialCopyObjects
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Снимает все расчётные блоки из физического листа, исключая локальную отметку A5.
Private Function TestSavedPublicationTables(ByVal book As Workbook) As Collection
    Dim tables As New Collection, key As Variant, data As Variant
    For Each key In Array("rngNDMSectionGeometry", "rngNDMSectionProperties", "rngNDMMaterialDiagrams", "rngNDMSectionContours", _
            "rngNDMElementResults", "rngNDMElementGeometry", "rngNDMSectionAnnotations", "rngNDMCoupledFields", "rngNDMElementResultIndex")
        data = ReadPhysicalResultTable(book, CStr(key))
        tables.Add data
    Next key
    Dim batch As New CBatchResultWriter, strength As New CStrengthSummaryWriter, cracks As New CCrackSummaryWriter, stability As New CStabilitySummaryWriter
    Dim target As Range
    Set target = book.Names("rngBatchSummary").RefersToRange.Resize(batch.RequiredSummaryRowsForWorkbook(book, 1), batch.RequiredSummaryOutputColumns)
    data = target.Value2: data(5, 1) = Empty: tables.Add data
    Set target = book.Names("rngStrengthSummaryAnchor").RefersToRange.Offset(-strength.HeaderRowsAboveAnchor, 0)
    tables.Add target.Resize(strength.RequiredRowsForWorkbook(book, 1), strength.RequiredColumns).Value2
    Set target = book.Names("rngCrackSummaryAnchor").RefersToRange.Offset(-cracks.HeaderRowsAboveAnchor, 0)
    tables.Add target.Resize(cracks.RequiredRowsForWorkbook(book, 1), cracks.RequiredColumns).Value2
    Set target = book.Names("rngStabilitySummaryAnchor").RefersToRange.Offset(-stability.HeaderRowsAboveAnchor, 0)
    tables.Add target.Resize(stability.RequiredRowsForWorkbook(book, 1), stability.RequiredColumns).Value2
    Set TestSavedPublicationTables = tables
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет содержимое каждого физического блока в том же порядке.
Private Function TestSamePublicationTables(ByVal expected As Collection, ByVal actual As Collection) As Boolean
    If actual.Count <> expected.Count Then Exit Function
    Dim i As Long, first As Variant, second As Variant
    For i = 1 To actual.Count
        first = expected.Item(i): second = actual.Item(i)
        If IsArray(first) <> IsArray(second) Then Exit Function
        If IsArray(first) Then
            If Not SnapshotMatricesEqual(first, second) Then Exit Function
        End If
    Next i
    TestSamePublicationTables = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Превращает список ссылок в сравнимый текст без раскрытия или изменения связей Excel.
Private Function TestPublicationLinks(ByVal book As Workbook) As String
    Dim links As Variant: links = book.LinkSources(xlExcelLinks)
    If IsArray(links) Then TestPublicationLinks = Join(links, vbLf)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет заголовок сохранённой схемы после успешного расчёта без перерисовки.
' Ручное обновление читает текущий RunID и сохраняет пользовательский объект.
Public Function RunTestResultPlotPublication(ByVal storage As String) As String
    Dim checks As Long, report As String
    On Error GoTo Failed
    Dim book As Workbook: Set book = ThisWorkbook
    If InStr(1, book.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestResultPlotPublication", "Запустите тест из изолированной копии книги."
    If storage <> "Excel" And storage <> "File" Then Err.Raise 5, "RunTestResultPlotPublication", "Выберите Excel или File."
    modTestPerformance.TestConfigurePerformanceFixture "NDS", 1, "Newton", "Rectangles", "Normal", "No"
    Dim settings As Range, loads As Range, profiles As Range
    Set settings = book.Names("rngSystemSettings").RefersToRange
    Set loads = book.Names("rngLoadCombinations").RefersToRange
    Set profiles = book.Names("rngCalculationProfiles").RefersToRange
    TestCancellationSetting(settings, "Results.SnapshotStorage").Value2 = storage
    TestCancellationSetting(settings, "Results.SnapshotName").Value2 = "TestPlotPublication_" & storage
    TestCancellationSetting(settings, "General.ExecutionReportEnabled").Value2 = "No"
    TestCancellationSetting(settings, "Output.CombinationMode").Value2 = "Worst"
    TestCancellationSetting(settings, "Plot.Enabled").Value2 = "Yes"
    TestCancellationSetting(settings, "Plot.AutoUpdateAfterCalculation").Value2 = "Yes"
    TestCancellationSetting(settings, "Plot.ResultGradient").Value2 = "Yes"
    TestChangeFreshnessVisualization book
    Dim row As Long
    For row = 1 To profiles.Rows.Count
        If profiles.Cells(row, 2).Value2 = "Visualization.State" Then
            profiles.Cells(row, 3).Value2 = "StrengthState"
            Exit For
        End If
    Next row
    TestCancellationRequire row <= profiles.Rows.Count, storage & " | задано прямое НДС для схемы", checks, report

    Dim chart As Chart, summary As New CBatchResultWriter, originalProperties As Variant, originalRunID As Double
    loads.Cells(2, 1).Value2 = "TEST_PLOT_RUN_A"
    FormSectionForWorkbook book
    RunSectionCalculationForWorkbook book
    Set chart = book.Worksheets.Item("Расчет").ChartObjects.Item("chtNDMSectionPlot").Chart
    TestCancellationRequire InStr(TestResultPlotTitle(chart), "TEST_PLOT_RUN_A") > 0, _
        storage & " | схема первого расчёта показана", checks, report
    originalProperties = ReadPhysicalResultTable(book, "rngNDMSectionProperties"): originalRunID = CDbl(originalProperties(2, 1))
    Dim userShape As Shape, retainedShape As Shape
    Set userShape = chart.Shapes.AddShape(msoShapeRectangle, 12#, 17#, 18#, 19#)
    userShape.Name = "TestUserPlotMarker"
    Set retainedShape = chart.Shapes.AddShape(msoShapeRectangle, 14#, 17#, 18#, 19#)
    retainedShape.Name = "NDMPlot_TestRetained"
    Dim userName As String, retainedName As String
    userName = userShape.Name: retainedName = retainedShape.Name

    TestCancellationSetting(settings, "Plot.AutoUpdateAfterCalculation").Value2 = "No"
    loads.Cells(2, 1).Value2 = "TEST_PLOT_RUN_B"
    RunSectionCalculationForWorkbook book
    Dim savedProperties As Variant: savedProperties = ReadPhysicalResultTable(book, "rngNDMSectionProperties")
    TestCancellationRequire CDbl(savedProperties(2, 1)) <> originalRunID And _
        ReadLoadCasePropertyText(book, "TEST_PLOT_RUN_B", "DirectStateStatus") = "OK" And _
        Len(summary.FreshnessNotice(book)) = 0, storage & " | новый расчёт сохранён и актуален", checks, report
    TestCancellationRequire InStr(TestResultPlotTitle(chart), "Схема требует обновления") = 1 And _
        chart.Shapes.Item(retainedName).Name = retainedName And chart.Shapes.Item(userName).Name = userName, _
        storage & " | сохранённый рисунок явно требует обновления", checks, report
    TestRefreshPublishedPlot book, chart, storage, "TEST_PLOT_RUN_B", userName, checks, report

    TestCancellationSetting(settings, "Plot.AutoUpdateAfterCalculation").Value2 = "Yes"
    TestCancellationSetting(settings, "Plot.ResultGradient").Value2 = "INVALID"
    loads.Cells(2, 1).Value2 = "TEST_PLOT_RUN_C"
    Dim message As String: message = RunSectionCalculationForWorkbook(book)
    TestCancellationRequire ReadLoadCasePropertyText(book, "TEST_PLOT_RUN_C", "DirectStateStatus") = "OK" And _
        Len(summary.FreshnessNotice(book)) = 0 And InStr(message, "Plot.ResultGradient") > 0, _
        storage & " | ошибка оформления сообщена, расчёт опубликован", checks, report
    TestCancellationRequire InStr(TestResultPlotTitle(chart), "Схема требует обновления") = 1 And _
        chart.Shapes.Item(userName).Name = userName, storage & " | схема после ошибки явно требует обновления", checks, report
    TestCancellationSetting(settings, "Plot.ResultGradient").Value2 = "Yes"
    TestRefreshPublishedPlot book, chart, storage, "TEST_PLOT_RUN_C", userName, checks, report

    RunTestResultPlotPublication = report & "TOTAL_RESULT_PLOT_PUBLICATION: passed=" & CStr(checks) & "; failed=0"
    Exit Function
Failed:
    RunTestResultPlotPublication = report & "FAIL: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_RESULT_PLOT_PUBLICATION: passed=" & CStr(checks) & "; failed=1"
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вручную восстанавливает текущую схему и проверяет неизменность снимка и счётчика решений.
Private Sub TestRefreshPublishedPlot(ByVal book As Workbook, ByVal chart As Chart, ByVal storage As String, _
        ByVal combination As String, ByVal userShapeName As String, ByRef checks As Long, ByRef report As String)
    Dim store As New CCalculationSnapshotStore, checksum As New CFileChecksum, path As String, hash As String
    If storage = "File" Then path = store.SnapshotPath(book): hash = checksum.FileSHA256(path)
    Dim properties As Variant: properties = ReadPhysicalResultTable(book, "rngNDMSectionProperties")
    Dim solveCount As Long: solveCount = SectionEquilibriumSolveCount()
    UpdateSectionResultsForWorkbook book
    TestCancellationRequire InStr(TestResultPlotTitle(chart), combination) > 0 And _
        chart.Shapes.Item(userShapeName).Name = userShapeName, storage & " | восстановлена текущая схема " & combination, checks, report
    TestCancellationRequire SnapshotMatricesEqual(properties, ReadPhysicalResultTable(book, "rngNDMSectionProperties")) And _
        solveCount = SectionEquilibriumSolveCount(), storage & " | обновление схемы сохраняет снимок и не решает НДС", checks, report
    If storage = "File" Then TestCancellationRequire checksum.FileSHA256(path) = hash, _
        storage & " | обновление схемы не меняет .ndm", checks, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает видимую надпись из собственного заголовка диаграммы.
Private Function TestResultPlotTitle(ByVal chart As Chart) As String
    TestResultPlotTitle = chart.Shapes.Item("NDMPlot_Title").TextFrame.Characters().Text
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Устанавливает отображаемую величину профиля по смысловому ключу строки.
Private Sub TestChangeFreshnessVisualization(ByVal book As Workbook)
    Dim profiles As Range: Set profiles = book.Names("rngCalculationProfiles").RefersToRange
    Dim row As Long
    For row = 1 To profiles.Rows.Count
        If CStr(profiles.Cells(row, 2).Value2) = "Visualization.Quantity" Then
            profiles.Cells(row, 3).Value2 = "Strain"
            Exit Sub
        End If
    Next row
    Err.Raise 5, "TestChangeFreshnessVisualization", "В таблице профилей не найден ключ Visualization.Quantity."
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает исходный отказ обычного обновления Results без диалогового окна.
Private Function TestInvokeFreshnessUpdate(ByVal book As Workbook) As Long
    On Error Resume Next
    Err.Clear
    UpdateSectionResultsForWorkbook book
    TestInvokeFreshnessUpdate = Err.Number
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает номер и текст обычного запуска, не подавляя проверяемую отмену.
Private Function TestInvokeCancelledCalculation(ByVal book As Workbook, ByRef description As String) As Long
    Dim ignored As String
    On Error Resume Next
    Err.Clear
    ignored = RunSectionCalculationForWorkbook(book)
    TestInvokeCancelledCalculation = Err.Number: description = Err.Description
    On Error GoTo 0
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считывает состояние Excel перед проверяемым запуском.
Private Function TestCancellationAppState(ByVal app As Application) As Variant
    TestCancellationAppState = Array(app.EnableCancelKey, app.ScreenUpdating, app.EnableEvents, app.DisplayAlerts, app.Calculation)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет все пять временно изменяемых свойств приложения.
Private Function TestCancellationStateEqual(ByVal app As Application, ByVal saved As Variant) As Boolean
    Dim actual As Variant, i As Long: actual = TestCancellationAppState(app)
    For i = LBound(saved) To UBound(saved)
        If actual(i) <> saved(i) Then Exit Function
    Next i
    TestCancellationStateEqual = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Получает значение настройки по ключу, независимо от перестановки строк.
Private Function TestCancellationSetting(ByVal settings As Range, ByVal key As String) As Range
    Dim row As Long
    For row = 2 To settings.Rows.Count
        If CStr(settings.Cells(row, 1).Value2) = key Then
            Set TestCancellationSetting = settings.Cells(row, 2)
            Exit Function
        End If
    Next row
    Err.Raise 5, "TestCancellationSetting", "Не найдена настройка " & key & "."
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет проверку в отчёте и прекращает запуск при первой ошибке.
Private Sub TestCancellationRequire(ByVal accepted As Boolean, ByVal label As String, ByRef checks As Long, ByRef report As String)
    If Not accepted Then Err.Raise 5, "TestCancellationRequire", label
    checks = checks + 1
    report = report & "PASS: " & label & vbCrLf
End Sub
