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
' Проверяет отказ после первой записи листа и штатное восстановление.
' Ошибка вводится только в изолированной книге направленным раннером.
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
    Dim loads As Range: Set loads = book.Names("rngLoadCombinations").RefersToRange
    Dim topLC As Range: Set topLC = book.Names("rngBatchSummary").RefersToRange.Offset(12, 0)
    loads.Cells(2, 1).Value2 = "TEST_PUBLICATION_BASE"
    FormSectionForWorkbook book
    RunSectionCalculationForWorkbook book
    TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_BASE", storage & " | исходная сводка сохранена", checks, report

    Dim store As New CCalculationSnapshotStore, checksum As New CFileChecksum, path As String, oldHash As String, currentHash As String
    If storage = "File" Then path = store.SnapshotPath(book): oldHash = checksum.FileSHA256(path)
    Dim properties As Variant: properties = ReadPhysicalResultTable(book, "rngNDMSectionProperties")
    Dim appState As Variant: appState = TestCancellationAppState(book.Application)
    Dim reason As String, number As Long, summary As New CBatchResultWriter
    If storage = "File" Then
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
    loads.Cells(2, 1).Value2 = "TEST_PUBLICATION_WRITE_FAILURE"
    loads.Cells(2, 2).Value2 = 0.9 * CDbl(loads.Cells(2, 2).Value2) ' Меняет входную подпись и свойства полей нового НДС.
    number = TestInvokeCancelledCalculation(book, reason)
    TestCancellationRequire number = 5 And InStr(1, reason, "TestPublicationAfterBase", vbBinaryCompare) > 0, _
        storage & " | ошибка введена после записи базы", checks, report
    TestCancellationRequire TestCancellationStateEqual(book.Application, appState) And TestCalculationSnapshotReadReleased(), _
        storage & " | настройки Excel и контекст чтения восстановлены", checks, report
    TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_BASE" And _
        Not SnapshotMatricesEqual(properties, ReadPhysicalResultTable(book, "rngNDMSectionProperties")) And _
        InStr(1, summary.FreshnessNotice(book), "Вывод прерван", vbBinaryCompare) = 1, _
        storage & " | смешанный вид явно отмечен как незавершённый", checks, report

    If storage = "File" Then
        currentHash = checksum.FileSHA256(path)
        TestCancellationRequire currentHash <> oldHash And _
            ReadLoadCasePropertyText(book, "TEST_PUBLICATION_WRITE_FAILURE", "DirectStateStatus") = "OK", _
            storage & " | авторитетный новый файл согласован и читается независимо от листа", checks, report
        UpdateSectionResultsForWorkbook book
        TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_WRITE_FAILURE" And Len(summary.FreshnessNotice(book)) = 0 And _
            checksum.FileSHA256(path) = currentHash, storage & " | ручное обновление восстановило лист без перезаписи файла", checks, report
    Else
        number = TestInvokeFreshnessUpdate(book)
        TestCancellationRequire number <> 0 And InStr(1, summary.FreshnessNotice(book), "Вывод прерван", vbBinaryCompare) = 1, _
            storage & " | схема не читает незавершённый Excel", checks, report
        loads.Cells(2, 1).Value2 = "TEST_PUBLICATION_RECOVERY"
        RunSectionCalculationForWorkbook book
        TestCancellationRequire topLC.Value2 = "TEST_PUBLICATION_RECOVERY" And Len(summary.FreshnessNotice(book)) = 0, _
            storage & " | повторный расчёт восстановил согласованный лист", checks, report
    End If
    RunTestResultPublicationFailure = report & "TOTAL_RESULT_PUBLICATION: passed=" & CStr(checks) & "; failed=0"
    Exit Function
Failed:
    RunTestResultPublicationFailure = report & "FAIL: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_RESULT_PUBLICATION: passed=" & CStr(checks) & "; failed=1"
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
