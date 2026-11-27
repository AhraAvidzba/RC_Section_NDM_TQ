Attribute VB_Name = "modTestCoupledASDConfigInput"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: явные ограничения ASD из фактической таблицы Config
' ==========================================================================
' Проверяет адресные ошибки, отключённое растяжение и неприменимость полей CDP.
' Изменяет только частную тестовую книгу и восстанавливает значения таблицы.
Private mPassedTest As Long, mFailedTest As Long, mReportTest As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выполняет направленные проверки reader -> неизменяемый объект ограничений ASD.
Public Function RunTestCoupledASDConfigInput(ByVal folder As String) As String
    On Error GoTo Failed
    Dim table As Object: Set table = ThisWorkbook.Names("rngCoupledMaterialSettings").RefersToRange
    Dim compressionCell As Object, tensionCell As Object
    Set compressionCell = TestFindValueCell(table, "Coupled.ASDCompressionLimit")
    Set tensionCell = TestFindValueCell(table, "Coupled.ASDTensionLimit")
    Dim savedCompression As Variant, savedTension As Variant
    savedCompression = compressionCell.Formula: savedTension = tensionCell.Formula
    mPassedTest = 0: mFailedTest = 0: mReportTest = vbNullString
    TestLawValidation table
    TestValidCase table, compressionCell, tensionCell, Empty, 0#
    TestValidCase table, compressionCell, tensionCell, 0#, 0#
    TestValidCase table, compressionCell, tensionCell, 0.00015, 0.00015
    Dim value As Variant
    For Each value In Array(Empty, 0#, -1#, "bad", "TODO", CVErr(2007), 1.1E+100)
        compressionCell.Value2 = value: tensionCell.ClearContents
        TestInvalidCase table, "Coupled.ASDCompressionLimit", compressionCell
    Next value
    compressionCell.Value2 = 0.0035
    For Each value In Array(-1#, "bad", "TODO", CVErr(2015), 1.1E+100)
        tensionCell.Value2 = value
        TestInvalidCase table, "Coupled.ASDTensionLimit", tensionCell
    Next value
    compressionCell.Value2 = CVErr(2007): tensionCell.Value2 = CVErr(2015)
    Dim reader As New CSystemSettingsReader, limits As CCoupledASDStrainLimits
    reader.LoadCoupledSettings ThisWorkbook
    Set limits = ReadWorkbookASDStrainLimits(reader, "CDP")
    TestRecord limits Is Nothing, "CDP не читает неприменимые ячейки ASD с ошибками Excel"
    compressionCell.Value2 = 0.0035: tensionCell.ClearContents
    Dim tensionKey As Object: Set tensionKey = tensionCell.Offset(0, -1)
    tensionKey.Value2 = "Coupled.MissingTensionTest"
    Set reader = New CSystemSettingsReader
    reader.LoadCoupledSettings ThisWorkbook
    Dim number As Long, description As String
    On Error Resume Next
    Set limits = ReadWorkbookASDStrainLimits(reader, "ASD")
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo Failed
    tensionKey.Value2 = "Coupled.ASDTensionLimit"
    TestRecord number <> 0 And InStr(description, "rngCoupledMaterialSettings") > 0, "Отсутствующая структурная строка требует восстановления таблицы"
    compressionCell.Formula = savedCompression: tensionCell.Formula = savedTension
    RunTestCoupledASDConfigInput = mReportTest & "TOTAL_ASD_CONFIG_INPUT: passed=" & mPassedTest & "; failed=" & mFailedTest
    Exit Function
Failed:
    Dim reason As String: reason = Err.Source & " " & Err.Description
    On Error Resume Next
    If Not compressionCell Is Nothing Then compressionCell.Formula = savedCompression
    If Not tensionCell Is Nothing Then tensionCell.Formula = savedTension
    If Not tensionKey Is Nothing Then tensionKey.Value2 = "Coupled.ASDTensionLimit"
    RunTestCoupledASDConfigInput = "RUNTIME ERROR: " & reason
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит значение по ключу, чтобы проверки не зависели от физического номера строки.
Private Function TestFindValueCell(ByVal table As Object, ByVal key As String) As Object
    Dim index As Long
    For index = 2 To table.Rows.Count
        If CStr(table.Cells(index, 1).Value2) = key Then
            Set TestFindValueCell = table.Cells(index, 2)
            Exit Function
        End If
    Next index
    Err.Raise 5, "modTestCoupledASDConfigInput", "В тестовой книге отсутствует ключ " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет точное чтение активного сжатия и необязательного растяжения.
Private Sub TestValidCase(ByVal table As Object, ByVal compressionCell As Object, ByVal tensionCell As Object, _
        ByVal tensionValue As Variant, ByVal expectedTension As Double)
    compressionCell.Value2 = 0.0035: tensionCell.Value2 = tensionValue
    Dim reader As New CSystemSettingsReader, limits As CCoupledASDStrainLimits
    reader.LoadCoupledSettings ThisWorkbook
    Set limits = ReadWorkbookASDStrainLimits(reader, "ASD")
    TestRecord Not limits Is Nothing, "ASD получает неизменяемый объект ограничений"
    TestRecord limits.LimitStrain(True) = 0.0035, "Явный сжатый предел читается точно"
    TestRecord limits.LimitStrain(False) = expectedTension, "Растянутый предел или его отключение читается точно"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Требует ошибку с ключом и фактическим адресом активного неверного значения.
Private Sub TestInvalidCase(ByVal table As Object, ByVal key As String, ByVal inputCell As Object)
    Dim reader As New CSystemSettingsReader, limits As CCoupledASDStrainLimits
    reader.LoadCoupledSettings ThisWorkbook
    Dim number As Long, description As String
    On Error Resume Next
    Set limits = ReadWorkbookASDStrainLimits(reader, "ASD")
    number = Err.Number: description = Err.Description: Err.Clear
    On Error GoTo 0
    TestRecord number <> 0, "Некорректное активное ограничение отклоняется: " & key
    TestRecord InStr(description, key) > 0 And InStr(description, inputCell.Address(False, False)) > 0 And _
        InStr(description, "rngCoupledMaterialSettings") > 0, "Ошибка указывает ключ и фактическое место ввода: " & description
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет число проверок и отдельную причину каждого отрицательного исхода.
Private Sub TestRecord(ByVal condition As Boolean, ByVal description As String)
    If condition Then
        mPassedTest = mPassedTest + 1
    Else
        mFailedTest = mFailedTest + 1
        mReportTest = mReportTest & "FAIL: " & description & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет фактический движок Validation после открытия сохранённой книги, без нормализации разделителей.
Private Sub TestLawValidation(ByVal table As Object)
    Dim cell As Object: Set cell = TestFindValueCell(table, "Coupled.ConcreteLaw")
    Dim saved As Variant: saved = cell.Formula
    TestRecord cell.Validation.Type = 3, "Выбор пространственного закона является списком"
    TestRecord cell.Validation.ShowError And cell.Validation.AlertStyle = 1, "Недопустимые значения останавливаются проверкой данных"
    Dim value As Variant, expected As Boolean
    For Each value In Array("CDP", "ASD", "CDP,ASD", "CDP;ASD", "UnknownLaw")
        cell.Value2 = CStr(value)
        expected = (value = "CDP" Or value = "ASD")
        TestRecord cell.Validation.Value = expected, "Нативная проверка варианта закона: " & CStr(value)
    Next value
    cell.Formula = saved
End Sub
