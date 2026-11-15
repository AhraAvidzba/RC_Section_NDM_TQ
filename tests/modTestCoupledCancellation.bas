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
        TestCancellationRequire number = 18 And InStr(1, description, "Расчёт отменён", vbBinaryCompare) > 0, _
            label & " | сохранён номер отмены и русское сообщение", checks, report
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
