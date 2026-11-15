Attribute VB_Name = "modResultFreshness"
Option Explicit

' ==========================================================================
' Актуальность показанного расчёта на Results
' ==========================================================================
' Подпись относится к использованному вводу и хранится с RunID снимка.
' Строка объявления описывает показанный лист и ход последнего запуска.

Public Const RESULT_INPUT_STAMP As String = "Calculation.InputStamp"
Public Const RESULT_INCOMPLETE_COUNT As String = "Calculation.IncompleteCount"
Private Const INPUT_VERSION As String = "1"
Private Const NOTICE_CANCELLED As String = "Расчёт отменён. Показаны ранее сохранённые результаты."
Private Const NOTICE_ATTEMPT_FAILED As String = "Новый расчёт не выполнен. Показаны ранее сохранённые результаты."
Private Const NOTICE_DISPLAY_FAILED As String = "Вывод прерван. Данные листа могут быть неполными. Обновите результаты из файла или выполните расчёт."
Private Const NOTICE_SOURCE_FAILED As String = "Не удалось прочитать выбранный снимок. На листе показаны ранее сохранённые данные."
Private Const NOTICE_SOURCE_CHANGED As String = "Выбран другой источник. Нажмите «Обновить схему и результаты»."

Private Type TResultFreshnessMetadata
    RunID As Double ' Относится к фактически видимому или проверенному файловому снимку.
    InputStamp As String
    SourceType As String
    RecipeType As String
    IncompleteCount As Long
End Type

' Снимает подпись до запуска решателя. Ошибка необязательного сравнения
' не подменяет адресную диагностику ввода расчётчика. Отмена сохраняет номер.
Public Function CaptureCalculationInputStamp(ByVal workbook As Object, ByVal sourceType As String, ByVal recipeType As String, _
        ByVal hasSpatial As Boolean, ByVal hasStability As Boolean) As String
    On Error GoTo Failed
    Dim digest As New CFileChecksum
    digest.BeginDigest

    Dim data As Variant, row As Long, key As String
    data = InputTable(workbook, "rngSystemSettings")
    DigestCell digest, "rngSystemSettings"
    For row = 1 To UBound(data, 1)
        key = LCase$(FreshnessText(data(row, 1)))
        If IsCalculationSettingKey(key, hasStability) Then
            DigestCell digest, key
            DigestCell digest, data(row, 2)
        End If
    Next row

    DigestColumns digest, workbook, "rngUnitSettings", Array(1, 2, 4)
    DigestColumns digest, workbook, "rngSignConventionSettings", Array(1, 2)
    DigestColumns digest, workbook, "rngConcreteMaterialParameters", Array(1, 2, 3), False, "Concrete.Rb.mc2"
    DigestColumns digest, workbook, "rngSteelMaterialParameters", Array(1, 2, 3)
    DigestLoadRows digest, workbook, "rngLoadCombinations"
    DigestProfiles digest, workbook
    If hasStability Then
        DigestLoadRows digest, workbook, "rngStabilityDurationLoads"
        DigestColumns digest, workbook, "rngSP35Table721", Array(1, 2, 3, 4, 5, 6)
    End If
    If hasSpatial Then
        DigestColumns digest, workbook, "rngCoupledMaterialSettings", Array(1, 2)
        DigestColumns digest, workbook, "rngCoupledSteelClasses", Array(1, 2, 3, 4, 5, 6, 7), True
        DigestTransverse digest, workbook, sourceType, recipeType
    End If
    CaptureCalculationInputStamp = INPUT_VERSION & "|" & CStr(Abs(hasSpatial)) & "|" & CStr(Abs(hasStability)) & "|" & digest.EndDigest
    Exit Function
Failed:
    If Err.Number = 18 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Считает сочетания с недоступной проверкой, даже если их общий статус FAIL.
Public Function IncompleteCalculationCount(ByVal batch As CBatchSectionCalculator) As Long
    Dim row As Long, result As CCombinationResult
    For row = 1 To batch.Count
        Set result = batch.ResultAt(row)
        If Not result.InputErrorMeta Is Nothing Or Not result.InternalErrorMeta Is Nothing Or _
                Not result.NumericalFailureMeta Is Nothing Then IncompleteCalculationCount = IncompleteCalculationCount + 1
    Next row
End Function

' Обновляет только надпись уже показанного листа, без открытия или замены .ndm.
Public Sub RefreshVisibleResultFreshness(ByVal workbook As Object, Optional ByVal sourceResolved As Boolean = False)
    Dim summary As New CBatchResultWriter, notice As String
    notice = summary.FreshnessNotice(workbook)
    Select Case notice
        Case NOTICE_CANCELLED, NOTICE_ATTEMPT_FAILED, NOTICE_DISPLAY_FAILED: Exit Sub
        Case NOTICE_SOURCE_FAILED, NOTICE_SOURCE_CHANGED
            If Not sourceResolved Then Exit Sub
    End Select
    If Not summary.HasDisplayedCombinations(workbook) Then Exit Sub
    Dim metadata As TResultFreshnessMetadata
    ReadFreshnessMetadata ReadPhysicalResultTable(workbook, "rngNDMSectionProperties"), metadata
    WriteFreshnessFromMetadata workbook, metadata
End Sub

' Сохраняет причину незавершённого запуска при обновлении того же снимка.
Public Function PriorAttemptNoticeForSnapshot(ByVal workbook As Object, ByVal snapshot As Object) As String
    Dim summary As New CBatchResultWriter, notice As String
    notice = summary.FreshnessNotice(workbook)
    If notice <> NOTICE_CANCELLED And notice <> NOTICE_ATTEMPT_FAILED Then Exit Function
    Dim previous As TResultFreshnessMetadata, selected As TResultFreshnessMetadata
    ReadFreshnessMetadata ReadPhysicalResultTable(workbook, "rngNDMSectionProperties"), previous
    ReadFreshnessMetadata snapshot.Table("rngNDMSectionProperties"), selected
    If previous.RunID = selected.RunID And previous.InputStamp = selected.InputStamp Then PriorAttemptNoticeForSnapshot = notice
End Function

' Рисует актуальность из одной опубликованной совокупности, а не из A5 файла.
Public Sub RefreshPublishedResultFreshness(ByVal workbook As Object, ByVal snapshot As Object, _
        Optional ByVal priorAttempt As String = vbNullString)
    Dim summary As New CBatchResultWriter
    If snapshot.IsGeometryPreview Or Len(priorAttempt) > 0 Then
        summary.SetFreshnessNotice workbook, priorAttempt
        Exit Sub
    End If
    Dim metadata As TResultFreshnessMetadata
    ReadFreshnessMetadata snapshot.Table("rngNDMSectionProperties"), metadata
    WriteFreshnessFromMetadata workbook, metadata
End Sub

' Не даёт выдать архивный результат за новый или читать незаконченный Excel-вывод.
Public Sub MarkResultAttemptStopped(ByVal workbook As Object, ByVal cancelled As Boolean, ByVal displayStarted As Boolean)
    Dim summary As New CBatchResultWriter
    If displayStarted Then
        summary.SetFreshnessNotice workbook, NOTICE_DISPLAY_FAILED
    ElseIf summary.FreshnessNotice(workbook) <> NOTICE_DISPLAY_FAILED And summary.HasDisplayedCombinations(workbook) Then
        summary.SetFreshnessNotice workbook, IIf(cancelled, NOTICE_CANCELLED, NOTICE_ATTEMPT_FAILED)
    End If
End Sub

' Отмечает сбой чтения File, сохраняя предыдущий вывод и исходное исключение.
Public Sub MarkResultSourceReadFailed(ByVal workbook As Object)
    Dim summary As New CBatchResultWriter
    If summary.FreshnessNotice(workbook) = NOTICE_DISPLAY_FAILED Then Exit Sub
    If summary.HasDisplayedCombinations(workbook) Then summary.SetFreshnessNotice workbook, NOTICE_SOURCE_FAILED
End Sub

' Через этот барьер проходят только схема и экспорт из физического Excel.
Public Sub RequireCompleteExcelResultDisplay(ByVal workbook As Object)
    Dim summary As New CBatchResultWriter
    If summary.FreshnessNotice(workbook) = NOTICE_DISPLAY_FAILED Then _
        Err.Raise vbObjectError + 4982, "RequireCompleteExcelResultDisplay", NOTICE_DISPLAY_FAILED
End Sub

' Фильтрует события конкретных листов. Смена файла требует штатного обновления.
Public Sub CheckResultFreshnessAfterEdit(ByVal workbook As Object, ByVal sheet As Object, ByVal target As Range)
    If sheet.Name <> "Config" And sheet.Name <> "Сечения" Then Exit Sub
    If sheet.Name = "Config" Then
        Dim settings As Range: Set settings = workbook.Names.Item("rngSystemSettings").RefersToRange
        Dim edited As Range: Set edited = workbook.Application.Intersect(settings.Columns(2), target)
        If Not edited Is Nothing Then
            Dim cell As Range, key As String, summary As New CBatchResultWriter
            For Each cell In edited.Cells
                key = FreshnessText(cell.Offset(0, -1).Value2)
                If key = "Results.SnapshotStorage" Or key = "Results.SnapshotName" Then
                    Select Case summary.FreshnessNotice(workbook)
                        Case NOTICE_CANCELLED, NOTICE_ATTEMPT_FAILED, NOTICE_DISPLAY_FAILED: Exit Sub
                    End Select
                    If summary.HasDisplayedCombinations(workbook) Then summary.SetFreshnessNotice workbook, NOTICE_SOURCE_CHANGED
                    Exit Sub
                End If
            Next cell
        End If
    End If
    RefreshVisibleResultFreshness workbook
End Sub

' Сопоставляет физический ввод с версионированной подписью RunID.
Private Sub WriteFreshnessFromMetadata(ByVal workbook As Object, ByRef metadata As TResultFreshnessMetadata)
    Dim notice As String, parts As Variant
    parts = Split(metadata.InputStamp, "|")
    If UBound(parts) = 3 Then
        If parts(0) = INPUT_VERSION And Len(parts(3)) = 64 Then
            Dim current As String
            current = CaptureCalculationInputStamp(workbook, metadata.SourceType, metadata.RecipeType, parts(1) = "1", parts(2) = "1")
            If current = metadata.InputStamp Then
                If metadata.IncompleteCount > 0 Then notice = "Часть проверок не завершена. Сочетаний с недоступными проверками: " & CStr(metadata.IncompleteCount) & "."
            Else
                notice = "Исходные данные изменены или недоступны. Показаны ранее сохранённые результаты."
            End If
        Else
            notice = "Показан сохранённый расчёт. Актуальность исходных данных не подтверждена."
        End If
    Else
        notice = "Показан сохранённый расчёт. Актуальность исходных данных не подтверждена."
        If metadata.IncompleteCount > 0 Then notice = notice & " Часть проверок не завершена."
    End If
    Dim summary As New CBatchResultWriter
    summary.SetFreshnessNotice workbook, notice
End Sub

' Читает метаданные только из общих свойств одного RunID.
Private Sub ReadFreshnessMetadata(ByVal data As Variant, ByRef metadata As TResultFreshnessMetadata)
    If Not IsArray(data) Then Exit Sub
    If UBound(data, 2) < 4 Then Exit Sub
    Dim row As Long, rowRunID As Double, runValue As Variant
    For row = 2 To UBound(data, 1)
        If FreshnessText(data(row, 2)) = "ALL" Then
            runValue = data(row, 1)
            If VarType(runValue) = vbDate Or IsNumeric(runValue) Then
                rowRunID = CDbl(runValue) ' Подготовленный снимок передаёт Date, лист и файл передают Double.
                If metadata.RunID = 0# Then metadata.RunID = rowRunID
                If rowRunID = metadata.RunID Then
                    Select Case FreshnessText(data(row, 3))
                        Case RESULT_INPUT_STAMP: metadata.InputStamp = FreshnessText(data(row, 4))
                        Case RESULT_INCOMPLETE_COUNT
                            If IsNumeric(data(row, 4)) Then metadata.IncompleteCount = CLng(data(row, 4))
                        Case "Geometry.Source": metadata.SourceType = FreshnessText(data(row, 4))
                        Case "Geometry.RecipeType": metadata.RecipeType = FreshnessText(data(row, 4))
                    End Select
                End If
            End If
        End If
    Next row
End Sub

' Перечисляет расчётные ключи, исключая выбор источника и оформление схемы.
Private Function IsCalculationSettingKey(ByVal key As String, ByVal hasStability As Boolean) As Boolean
    Select Case key
        Case "general.plateauslopefactor", "general.diagramextension", "calculation.zeromomentperdepth"
            IsCalculationSettingKey = True
        Case Else
            IsCalculationSettingKey = Left$(key, 5) = "load." Or Left$(key, 7) = "solver." Or _
                Left$(key, 9) = "capacity." Or Left$(key, 4) = "sls." Or (hasStability And Left$(key, 10) = "stability.")
    End Select
End Function

' Читает один именованный блок без обращения к текущему листу или книге.
Private Function InputTable(ByVal workbook As Object, ByVal rangeName As String) As Variant
    InputTable = workbook.Names.Item(rangeName).RefersToRange.Value2
End Function

' Сворачивает только указанные смысловые столбцы: единицы и комментарии вне ввода пропускаются.
Private Sub DigestColumns(ByVal digest As CFileChecksum, ByVal workbook As Object, ByVal table As String, ByVal columns As Variant, _
        Optional ByVal skipBlankRows As Boolean = False, Optional ByVal withoutThird As String = vbNullString)
    Dim data As Variant, row As Long, column As Variant
    data = InputTable(workbook, table)
    DigestCell digest, table
    For row = 1 To UBound(data, 1)
        If row = 1 Or Not skipBlankRows Or Len(FreshnessText(data(row, 1))) > 0 Or IsError(data(row, 1)) Then
            DigestCell digest, row
            For Each column In columns
                If Len(withoutThird) = 0 Or column <> 3 Or StrComp(FreshnessText(data(row, 1)), withoutThird, vbTextCompare) <> 0 Then _
                    DigestCell digest, data(row, column)
            Next column
        End If
    Next row
End Sub

' Пустой ID выключает всю строку нагрузки. Положение активного ID сохраняется в сводке.
Private Sub DigestLoadRows(ByVal digest As CFileChecksum, ByVal workbook As Object, ByVal rangeName As String)
    Dim data As Variant, row As Long, column As Long
    data = InputTable(workbook, rangeName)
    DigestCell digest, rangeName
    For row = 1 To UBound(data, 1)
        If row = 1 Or Len(FreshnessText(data(row, 1))) > 0 Or IsError(data(row, 1)) Then
            DigestCell digest, row
            For column = 1 To UBound(data, 2): DigestCell digest, data(row, column): Next column
        End If
    Next row
End Sub

' Действующие данные профилей хранятся отдельно от живых выборов Visualization.*.
Private Sub DigestProfiles(ByVal digest As CFileChecksum, ByVal workbook As Object)
    Dim data As Variant, row As Long, column As Long, key As String
    data = InputTable(workbook, "rngCalculationProfiles")
    DigestCell digest, "rngCalculationProfiles"
    For row = 1 To UBound(data, 1)
        key = LCase$(FreshnessText(data(row, 2)))
        If key = "key" Or key = "profileid" Or Left$(key, 8) = "profile." Or Left$(key, 12) = "calculation." Or Left$(key, 14) = "materialmodel." Then
            DigestCell digest, key
            For column = 3 To UBound(data, 2) - 1: DigestCell digest, data(row, column): Next column
        End If
    Next row
End Sub

' Читает только включённые параметры Transverse сохранённого типа сечения.
Private Sub DigestTransverse(ByVal digest As CFileChecksum, ByVal workbook As Object, ByVal sourceType As String, ByVal recipeType As String)
    Const block As String = "Transverse"
    Dim tableNames As Variant, tables As New CSectionTypeRegistry
    If sourceType = "AutoCADImport" Then
        tableNames = Array("rngAutoCADTransverseRebar")
    ElseIf Len(recipeType) > 0 Then
        tableNames = tables.RecipeTableNames(recipeType)
    Else
        Exit Sub
    End If
    Dim capture As New CSectionInputSnapshotReader, context As CSectionRecipeContext, units As New CUnitSystem
    Set context = capture.Capture(workbook, tableNames, units) ' Сравниваются значения ячеек, без перевода длин.
    Dim tableName As Variant, group As Variant, reader As CSectionInputReader, mode As String, key As Variant
    For Each tableName In tableNames
        Set reader = context.Reader(CStr(tableName))
        DigestCell digest, CStr(tableName)
        For Each group In reader.RowKeys(block)
            mode = LCase$(FreshnessText(reader.CellValue(block, CStr(group), "kind")))
            If Len(mode) > 0 Then
                DigestCell digest, CStr(group)
                DigestCell digest, mode
                If mode = "контур" Or mode = "ветвь" Then
                    For Each key In Array("d", "spacing", "steel")
                        DigestCell digest, reader.CellValue(block, CStr(group), CStr(key))
                    Next key
                    If mode = "контур" Then
                        DigestCell digest, reader.CellValue(block, CStr(group), "cover")
                        DigestCell digest, reader.CellValue(block, CStr(group), "loop")
                    Else
                        DigestCell digest, reader.CellValue(block, CStr(group), "path")
                    End If
                End If
            End If
        Next group
    Next tableName
End Sub

' Хеширует тип и точные биты числа либо длину и UTF-16 текста без локали.
Private Sub DigestCell(ByVal digest As CFileChecksum, ByVal value As Variant)
    Dim tag As Long, count As Long, number As Double, valueText As String
    tag = VarType(value)
    Select Case tag
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, 14
            tag = vbDouble: number = CDbl(value)
        Case vbString, vbError: valueText = CStr(value)
        Case vbBoolean: number = Abs(CBool(value))
    End Select
    digest.DigestMemory VarPtr(tag), 4
    Select Case tag
        Case vbDouble, vbBoolean
            digest.DigestMemory VarPtr(number), 8
        Case vbString, vbError
            count = LenB(valueText)
            digest.DigestMemory VarPtr(count), 4
            digest.DigestMemory StrPtr(valueText), count
    End Select
End Sub

' Возвращает текст ключа без попытки преобразовать ошибку Excel в идентификатор.
Private Function FreshnessText(ByVal value As Variant) As String
    If Not IsError(value) And Not IsNull(value) Then FreshnessText = Trim$(CStr(value))
End Function
