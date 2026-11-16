Attribute VB_Name = "modPlanarTorsionResults"
Option Explicit

' ===========================================================================
' Вывод предварительной проверки кручения на лист Results
' ===========================================================================
' Выводит готовые сочетания и напряжения продольных стержней двумя массивами.
' Именованные якоря задают положение заголовков. Отдельные имена запоминают
' только строки данных этого модуля, в том числе после сохранения книги.

Private Const SUMMARY_ANCHOR As String = "rngPlanarTorsionSummary"
Private Const STEEL_ANCHOR As String = "rngPlanarTorsionSteel"
Private Const SUMMARY_TITLE As String = "Кручение и нормальные напряжения"
Private Const STEEL_TITLE As String = "Продольная арматура при кручении"
Private Const SUMMARY_WIDTH As Long = 9
Private Const STEEL_WIDTH As Long = 5
Private Const COMBINATION_HEADER As String = "Сочетание"
Private Const OWNED_SUFFIX As String = "Owned"

' Публикует только готовые результаты сочетаний с кручением. Обе области
' проверяются до записи. Момент и напряжения переводятся в OUTPUT-единицы.
Public Sub WritePlanarTorsionResults(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem)
    If batch Is Nothing Or units Is Nothing Then TorsionOutputError "Для вывода кручения нужны готовый расчёт сочетаний и единицы вывода."
    Dim summary As Object, steel As Object, caseCount As Long
    caseCount = batch.Count
    If Not ReadAnchors(workbook, summary, steel) Then
        Dim i As Long, result As CCombinationResult
        For i = 1 To caseCount
            Set result = batch.ResultAt(i)
            If Not result.PlanarTorsion Is Nothing Then TorsionOutputError "На листе Results нужны оба якоря " & SUMMARY_ANCHOR & " и " & STEEL_ANCHOR & "."
        Next i
        Exit Sub
    End If
    RequireWritableSheet summary.Worksheet

    Dim oldSummary As Long, oldSteel As Long
    oldSummary = OwnedRowCount(workbook, summary, SUMMARY_ANCHOR, SUMMARY_TITLE, SUMMARY_WIDTH)
    oldSteel = OwnedRowCount(workbook, steel, STEEL_ANCHOR, STEEL_TITLE, STEEL_WIDTH)

    Dim summaries As Variant, details As Variant, summaryRows As Long, detailRows As Long
    Dim summaryHeaders As Variant, steelHeaders As Variant
    summaryHeaders = Array(COMBINATION_HEADER, "T, " & units.OutputMomentUnit, "q, Н/мм", _
        "Напряжение хомута, " & units.OutputStressUnit, "Запас хомута", "Статус хомута", _
        "Использование N/M + T", "Статус N/M + T", "ResultComment")
    steelHeaders = Array(COMBINATION_HEADER, "Стержень", "Начальное напряжение от T, " & units.OutputStressUnit, _
        "Начальная деформация от T", "Итоговое напряжение, " & units.OutputStressUnit)
    BuildTables batch, units, summaryHeaders, steelHeaders, AvailableBodyRows(summary), AvailableBodyRows(steel), _
        summaries, details, summaryRows, detailRows

    ValidateWriteArea workbook, summary, SUMMARY_ANCHOR, SUMMARY_TITLE, summaryHeaders, oldSummary, summaryRows
    ValidateWriteArea workbook, steel, STEEL_ANCHOR, STEEL_TITLE, steelHeaders, oldSteel, detailRows
    WriteTable workbook, summary, SUMMARY_ANCHOR, SUMMARY_TITLE, SUMMARY_WIDTH, oldSummary, summaryRows, summaries
    WriteTable workbook, steel, STEEL_ANCHOR, STEEL_TITLE, STEEL_WIDTH, oldSteel, detailRows, details
End Sub

' Очищает только зарегистрированные строки данных. Заголовки и области
' соседних таблиц сохраняются. В книге без обоих якорей блок отсутствует.
Public Sub ClearPlanarTorsionResults(ByVal workbook As Object)
    Dim summary As Object, steel As Object
    If Not ReadAnchors(workbook, summary, steel) Then Exit Sub
    Dim summaryRows As Long, steelRows As Long
    summaryRows = OwnedRowCount(workbook, summary, SUMMARY_ANCHOR, SUMMARY_TITLE, SUMMARY_WIDTH)
    steelRows = OwnedRowCount(workbook, steel, STEEL_ANCHOR, STEEL_TITLE, STEEL_WIDTH)
    If summaryRows = 0 And steelRows = 0 Then Exit Sub
    RequireWritableSheet summary.Worksheet
    If summaryRows > 0 Then RequireUnmerged summary.Offset(2, 0).Resize(summaryRows, SUMMARY_WIDTH), SUMMARY_ANCHOR
    If steelRows > 0 Then RequireUnmerged steel.Offset(2, 0).Resize(steelRows, STEEL_WIDTH), STEEL_ANCHOR
    ClearOwnedBody workbook, summary, SUMMARY_ANCHOR, SUMMARY_WIDTH, summaryRows
    ClearOwnedBody workbook, steel, STEEL_ANCHOR, STEEL_WIDTH, steelRows
End Sub

' Собирает компактные матрицы с одной строкой заголовков. Сочетания без
' результата кручения пропускаются. Размеры проверяются до выделения памяти.
Private Sub BuildTables(ByVal batch As CBatchSectionCalculator, ByVal units As CUnitSystem, _
        ByRef summaryHeaders As Variant, ByRef steelHeaders As Variant, ByVal summaryLimit As Long, ByVal detailLimit As Long, _
        ByRef summaries As Variant, ByRef details As Variant, ByRef summaryRows As Long, ByRef detailRows As Long)
    Dim i As Long, rowCount As Long, result As CCombinationResult
    Dim outcome() As CTorsionPlanarResult, steelTable() As Variant, rowsByCase() As Long
    If batch.Count > 0 Then
        ReDim outcome(1 To batch.Count): ReDim steelTable(1 To batch.Count): ReDim rowsByCase(1 To batch.Count)
    End If
    For i = 1 To batch.Count
        Set result = batch.ResultAt(i): Set outcome(i) = result.PlanarTorsion
        If Not outcome(i) Is Nothing Then
            summaryRows = summaryRows + 1
            If summaryRows > summaryLimit Then TorsionOutputError "Сводка кручения не помещается под " & SUMMARY_ANCHOR & "."
            steelTable(i) = outcome(i).SteelTable: rowCount = SteelRows(steelTable(i))
            If rowCount > detailLimit - detailRows Then TorsionOutputError "Строки стержней не помещаются под " & STEEL_ANCHOR & "."
            rowsByCase(i) = rowCount: detailRows = detailRows + rowCount
        End If
    Next i

    ReDim summaries(1 To summaryRows + 1, 1 To SUMMARY_WIDTH)
    ReDim details(1 To detailRows + 1, 1 To STEEL_WIDTH)
    PutHeaders summaries, summaryHeaders
    PutHeaders details, steelHeaders
    Dim summaryIndex As Long, detailIndex As Long, combination As String, policy As New CResultStatusPolicy
    summaryIndex = 1: detailIndex = 1
    For i = 1 To batch.Count
        If Not outcome(i) Is Nothing Then
            combination = batch.CombinationID(i)
            summaryIndex = summaryIndex + 1
            PutSummary summaries, summaryIndex, combination, outcome(i), policy, units
            If rowsByCase(i) > 0 Then PutSteel details, detailIndex, combination, steelTable(i), rowsByCase(i), units
        End If
    Next i
End Sub

' Возвращает число строк сохранённых стержней и проверяет их контракт.
Private Function SteelRows(ByRef table As Variant) As Long
    If IsEmpty(table) Then Exit Function
    If Not IsArray(table) Then TorsionOutputError "Результат продольных стержней должен быть таблицей для " & STEEL_ANCHOR & "."
    On Error GoTo InvalidShape
    If UBound(table, 2) - LBound(table, 2) + 1 <> 4 Then GoTo InvalidShape
    SteelRows = UBound(table, 1) - LBound(table, 1) + 1
    Exit Function
InvalidShape:
    On Error GoTo 0
    TorsionOutputError "Для " & STEEL_ANCHOR & " нужны четыре поля результата каждого стержня."
End Function

' Заполняет первую строку массива видимыми заголовками столбцов.
Private Sub PutHeaders(ByRef data As Variant, ByRef headers As Variant)
    Dim column As Long
    For column = LBound(headers) To UBound(headers)
        data(1, column - LBound(headers) + 1) = headers(column)
    Next column
End Sub

' Переносит одну сводку. Числа недоступной проверки остаются пустыми,
' её внешний статус и причина берутся из общих правил результатов.
Private Sub PutSummary(ByRef data As Variant, ByVal rowIndex As Long, ByVal combination As String, _
        ByVal outcome As CTorsionPlanarResult, ByVal policy As CResultStatusPolicy, ByVal units As CUnitSystem)
    Dim hoop As CResultMeta, normal As CResultMeta, comment As String, normalComment As String
    Set hoop = outcome.HoopMeta: Set normal = outcome.NormalMeta
    data(rowIndex, 1) = combination
    data(rowIndex, 2) = units.InternalTorqueToOutput(outcome.Torque)
    If hoop.Calculated Then
        data(rowIndex, 3) = outcome.ShearFlow ' Поток показан в Н/мм независимо от выбранной пары OUTPUT-единиц.
        data(rowIndex, 4) = units.InternalStressToOutput(outcome.HoopStress)
        data(rowIndex, 5) = outcome.HoopReserve
    End If
    data(rowIndex, 6) = policy.ExternalStatus(hoop)
    If outcome.HasNormalState Then data(rowIndex, 7) = outcome.NormalUtilization
    data(rowIndex, 8) = policy.ExternalStatus(normal)
    If hoop.InternalStatus <> rsSuccess Then comment = Trim$(hoop.ResultComment)
    normalComment = Trim$(normal.ResultComment)
    If Len(normalComment) > 0 And StrComp(comment, normalComment, vbTextCompare) <> 0 Then
        If Len(comment) > 0 Then comment = comment & " "
        comment = comment & normalComment
    End If
    data(rowIndex, 9) = comment
End Sub

' Переносит сечение арматуры, начальные напряжение и деформацию от T и
' итоговое напряжение из принятого решения. Деформация безразмерна.
Private Sub PutSteel(ByRef data As Variant, ByRef rowIndex As Long, ByVal combination As String, _
        ByRef table As Variant, ByVal rowCount As Long, ByVal units As CUnitSystem)
    Dim i As Long, inputRow As Long, firstColumn As Long
    inputRow = LBound(table, 1): firstColumn = LBound(table, 2)
    For i = 1 To rowCount
        rowIndex = rowIndex + 1
        data(rowIndex, 1) = combination
        data(rowIndex, 2) = CStr(table(inputRow, firstColumn))
        data(rowIndex, 3) = units.InternalStressToOutput(CDbl(table(inputRow, firstColumn + 1)))
        data(rowIndex, 4) = CDbl(table(inputRow, firstColumn + 2))
        data(rowIndex, 5) = units.InternalStressToOutput(CDbl(table(inputRow, firstColumn + 3)))
        inputRow = inputRow + 1
    Next i
End Sub

' Сверяет наличие двух одноклеточных якорей на Results и разнесение колонок.
' Отсутствие обеих таблиц допустимо для книги без результатов кручения.
Private Function ReadAnchors(ByVal workbook As Object, ByRef summary As Object, ByRef steel As Object) As Boolean
    If workbook Is Nothing Then TorsionOutputError "Книга Excel не передана."
    Set summary = NamedRangeIfPresent(workbook, SUMMARY_ANCHOR)
    Set steel = NamedRangeIfPresent(workbook, STEEL_ANCHOR)
    If summary Is Nothing And steel Is Nothing Then
        If Not FindName(workbook, SUMMARY_ANCHOR & OWNED_SUFFIX) Is Nothing Or _
                Not FindName(workbook, STEEL_ANCHOR & OWNED_SUFFIX) Is Nothing Then _
            TorsionOutputError "Не найдены якоря для зарегистрированных таблиц кручения на Results."
        Exit Function
    End If
    If summary Is Nothing Or steel Is Nothing Then TorsionOutputError "Восстановите оба якоря " & SUMMARY_ANCHOR & " и " & STEEL_ANCHOR & "."
    ValidateAnchor workbook, summary, SUMMARY_ANCHOR, SUMMARY_WIDTH
    ValidateAnchor workbook, steel, STEEL_ANCHOR, STEEL_WIDTH
    If summary.Column + SUMMARY_WIDTH - 1 >= steel.Column And steel.Column + STEEL_WIDTH - 1 >= summary.Column Then _
        TorsionOutputError "Колонки областей " & SUMMARY_ANCHOR & " и " & STEEL_ANCHOR & " пересекаются."
    ReadAnchors = True
End Function

' Проверяет доступность заголовка и требуемых столбцов на листе Results.
Private Sub ValidateAnchor(ByVal workbook As Object, ByVal anchor As Object, ByVal anchorName As String, ByVal width As Long)
    If anchor.Areas.Count <> 1 Or anchor.Cells.CountLarge <> 1 Then TorsionOutputError "Имя " & anchorName & " должно указывать на одну ячейку."
    If Not anchor.Worksheet.Parent Is workbook Then TorsionOutputError "Имя " & anchorName & " должно указывать на текущую книгу."
    If StrComp(anchor.Worksheet.Name, "Results", vbTextCompare) <> 0 Then TorsionOutputError "Якорь " & anchorName & " должен находиться на листе Results."
    If anchor.Column + width - 1 > anchor.Worksheet.Columns.Count Or AvailableBodyRows(anchor) < 0 Then _
        TorsionOutputError "Заголовки таблицы " & anchorName & " выходят за пределы Results."
    If anchor.MergeCells Then
        Dim merged As Object: Set merged = anchor.MergeArea
        If merged.Rows.Count <> 1 Or merged.Column <> anchor.Column Or merged.Columns.Count > width Then _
            TorsionOutputError "Объединение заголовка " & anchorName & " должно находиться внутри одной строки таблицы."
    End If
End Sub

' Считает доступные строки под названием и строкой названий столбцов.
Private Function AvailableBodyRows(ByVal anchor As Object) As Long
    AvailableBodyRows = anchor.Worksheet.Rows.Count - anchor.Row - 1
End Function

' Возвращает длину сохранённого тела таблицы только при совпадении
' именованной области, колонки ID и её заголовка с текущим якорем.
Private Function OwnedRowCount(ByVal workbook As Object, ByVal anchor As Object, ByVal anchorName As String, _
        ByVal title As String, ByVal width As Long) As Long
    Dim stored As Object: Set stored = NamedRangeIfPresent(workbook, anchorName & OWNED_SUFFIX)
    If stored Is Nothing Then Exit Function
    If Not stored.Worksheet Is anchor.Worksheet Then TorsionOutputError "Область " & anchorName & OWNED_SUFFIX & " находится на другом листе."
    If stored.Areas.Count <> 1 Or stored.Row <> anchor.Row + 2 Or stored.Column <> anchor.Column Or stored.Columns.Count <> width Then _
        TorsionOutputError "Сохранённая область " & anchorName & OWNED_SUFFIX & " не совпадает с таблицей."
    If Not ExactText(anchor.Value2, title) Or Not ExactText(anchor.Offset(1, 0).Value2, COMBINATION_HEADER) Then _
        TorsionOutputError "У сохранённой таблицы " & anchorName & " изменены опознавательные заголовки."
    OwnedRowCount = stored.Rows.Count
End Function

' Проверяет заголовок, прежние строки и свободное продолжение таблицы.
' При смене OUTPUT-единиц допускается заменить старые обозначения единиц.
Private Sub ValidateWriteArea(ByVal workbook As Object, ByVal anchor As Object, ByVal anchorName As String, ByVal title As String, _
        ByRef headers As Variant, ByVal oldRows As Long, ByVal newRows As Long)
    Dim titleValue As Variant: titleValue = anchor.Value2
    If Not IsEmpty(titleValue) And Not ExactText(titleValue, title) Then TorsionOutputError "В якоре " & anchorName & " находятся посторонние данные."
    Dim column As Long, headerValue As Variant, headerRange As Object, width As Long
    width = UBound(headers) - LBound(headers) + 1
    Set headerRange = anchor.Offset(1, 0).Resize(1, width)
    RequireUnmerged headerRange, anchorName
    If AnyFormula(headerRange) Then TorsionOutputError "В строке заголовков " & anchorName & " находятся формулы."
    headerValue = headerRange.Value2
    For column = 1 To width
        If Not CompatibleHeader(headerValue(1, column), CStr(headers(LBound(headers) + column - 1))) Then _
            TorsionOutputError "Под якорем " & anchorName & " находятся посторонние данные в заголовке столбца " & CStr(column) & "."
    Next column
    If oldRows > 0 Then RequireUnmerged anchor.Offset(2, 0).Resize(oldRows, width), anchorName
    If newRows > oldRows Then
        Dim extension As Object: Set extension = anchor.Offset(oldRows + 2, 0).Resize(newRows - oldRows, width)
        RequireUnmerged extension, anchorName
        If workbook.Application.WorksheetFunction.CountA(extension) <> 0 Then _
            TorsionOutputError "Расширение таблицы " & anchorName & " пересекло занятые ячейки Results."
    End If
End Sub

' Сверяет название столбца, разрешая прежнее обозначение единицы у
' размерных величин. Постоянная единица потока всегда остаётся Н/мм.
Private Function CompatibleHeader(ByVal value As Variant, ByVal expected As String) As Boolean
    If IsEmpty(value) Or ExactText(value, expected) Then
        CompatibleHeader = True
        Exit Function
    End If
    If VarType(value) <> vbString Then Exit Function
    If StrComp(expected, "q, Н/мм", vbBinaryCompare) = 0 Then Exit Function
    Dim index As Long: index = InStr(1, expected, ", ", vbBinaryCompare)
    If index <= 0 Or Len(CStr(value)) <= index + 1 Then Exit Function
    CompatibleHeader = (StrComp(Left$(CStr(value), index + 1), Left$(expected, index + 1), vbBinaryCompare) = 0)
End Function

' Опознаёт текстовое значение без преобразования формул с ошибкой,
' пустых ячеек и чисел к строке.
Private Function ExactText(ByVal value As Variant, ByVal expected As String) As Boolean
    If VarType(value) = vbString Then ExactText = (StrComp(CStr(value), expected, vbBinaryCompare) = 0)
End Function

' Находит формулы в одиночной или смешанной области заголовка.
Private Function AnyFormula(ByVal target As Object) As Boolean
    Dim value As Variant: value = target.HasFormula
    If IsNull(value) Then
        AnyFormula = True
    Else
        AnyFormula = CBool(value)
    End If
End Function

' Проверяет, что блочная запись и очистка физически затрагивают только
' свои клетки, без пересечения объединённых диапазонов.
Private Sub RequireUnmerged(ByVal target As Object, ByVal anchorName As String)
    Dim value As Variant: value = target.MergeCells
    If IsNull(value) Then TorsionOutputError "В таблице " & anchorName & " есть объединённые ячейки."
    If CBool(value) Then TorsionOutputError "В таблице " & anchorName & " есть объединённые ячейки."
End Sub

' Уважает режим защиты пользователя, включая разрешение на запись макросам.
Private Sub RequireWritableSheet(ByVal sheet As Object)
    If sheet.ProtectContents And Not sheet.ProtectionMode Then _
        TorsionOutputError "Лист Results защищён от записи. Проверьте доступность " & SUMMARY_ANCHOR & " и " & STEEL_ANCHOR & "."
End Sub

' Публикует готовую матрицу и убирает оставшийся хвост прежней таблицы.
' При расширении заранее отмечает только проверенные пустые строки, чтобы
' следующая очистка смогла удалить и частично записанный результат.
Private Sub WriteTable(ByVal workbook As Object, ByVal anchor As Object, ByVal anchorName As String, ByVal title As String, _
        ByVal width As Long, ByVal oldRows As Long, ByVal newRows As Long, ByRef data As Variant)
    If IsEmpty(anchor.Value2) Then anchor.Value2 = title
    If IsEmpty(anchor.Offset(1, 0).Value2) Then anchor.Offset(1, 0).Value2 = COMBINATION_HEADER ' Опознавательный заголовок нужен уже при регистрации строки данных.
    If newRows > oldRows Then RememberOwnedBody workbook, anchor, anchorName, width, newRows
    If newRows > 0 Then
        Dim idColumns As Long: idColumns = 1
        If anchorName = STEEL_ANCHOR Then idColumns = 2
        anchor.Offset(2, 0).Resize(newRows, idColumns).NumberFormat = "@" ' Идентификаторы с ведущими нулями остаются текстом.
    End If
    anchor.Offset(1, 0).Resize(newRows + 1, width).Value2 = data
    If oldRows > newRows Then
        anchor.Offset(newRows + 2, 0).Resize(oldRows - newRows, width).ClearContents
        RememberOwnedBody workbook, anchor, anchorName, width, newRows
    End If
End Sub

' Снимает сохранённую запись об области после очистки ровно прежних строк.
Private Sub ClearOwnedBody(ByVal workbook As Object, ByVal anchor As Object, ByVal anchorName As String, _
        ByVal width As Long, ByVal oldRows As Long)
    If oldRows = 0 Then Exit Sub
    anchor.Offset(2, 0).Resize(oldRows, width).ClearContents
    RememberOwnedBody workbook, anchor, anchorName, width, 0
End Sub

' Сохраняет точный прямоугольник данных отдельно от постоянного якоря.
Private Sub RememberOwnedBody(ByVal workbook As Object, ByVal anchor As Object, ByVal anchorName As String, _
        ByVal width As Long, ByVal rowCount As Long)
    Dim ownedName As String, entry As Object
    ownedName = anchorName & OWNED_SUFFIX: Set entry = FindName(workbook, ownedName)
    If rowCount = 0 Then
        If Not entry Is Nothing Then entry.Delete
        Exit Sub
    End If
    Dim area As Object, reference As String
    Set area = anchor.Offset(2, 0).Resize(rowCount, width)
    reference = "='" & Replace$(area.Worksheet.Name, "'", "''") & "'!" & area.Address
    If entry Is Nothing Then
        workbook.Names.Add Name:=ownedName, RefersTo:=reference
    Else
        entry.RefersTo = reference
    End If
End Sub

' Ищет существующее имя на уровне книги без создания нового диапазона.
Private Function FindName(ByVal workbook As Object, ByVal name As String) As Object
    On Error Resume Next
    Set FindName = workbook.Names.Item(name)
    On Error GoTo 0
End Function

' Возвращает область сохранённого имени. Повреждённая ссылка требует
' исправления, чтобы следующая запись не затронула соседние ячейки.
Private Function NamedRangeIfPresent(ByVal workbook As Object, ByVal name As String) As Object
    Dim entry As Object: Set entry = FindName(workbook, name)
    If entry Is Nothing Then Exit Function
    On Error Resume Next
    Set NamedRangeIfPresent = entry.RefersToRange
    On Error GoTo 0
    If NamedRangeIfPresent Is Nothing Then TorsionOutputError "Имя " & name & " должно ссылаться на действительную область Results."
End Function

' Передаёт причину остановки записи с названием соответствующей области.
Private Sub TorsionOutputError(ByVal description As String)
    Err.Raise vbObjectError + 5680, "modPlanarTorsionResults", description
End Sub
