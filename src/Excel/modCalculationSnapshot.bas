Attribute VB_Name = "modCalculationSnapshot"
Option Explicit

' ==========================================================================
' Общая загрузка и восстановление результатов без повторного расчёта
' ==========================================================================
' После расчёта вывод использует готовые данные в памяти. Кнопка в File
' читает текущий выбранный файл расчёта либо геометрии; Excel читает только Results.
' Кратковременный контекст чтения принадлежит одной операции и одной книге.
Private mReadView As CResultReadSource ' Адресный источник существует только внутри одной операции.
Private mReadViewResolved As Boolean
Private mReadSnapshot As CCalculationSnapshot ' Неизменяемый набор только на время одной операции чтения.
Private mReadWorkbook As Object ' Книга, которой принадлежит текущий контекст.
Private mReadDepth As Long ' Вложенные чтения схемы и экспорта не уничтожают внешний контекст.
Private Const PLOT_NOT_CURRENT As String = "Схема требует обновления для нового расчёта."
Private Const OUTPUT_INDEX_NAME As String = "rngNDMElementResultIndex"

' Открывает короткую область чтения; между пользовательскими действиями кэша нет.
Public Sub BeginCalculationSnapshotRead(ByVal workbook As Object, Optional ByVal snapshot As CCalculationSnapshot = Nothing)
    If mReadDepth = 0 Then
        Set mReadView = Nothing: mReadViewResolved = False
        If Not snapshot Is Nothing Then snapshot.Validate
        Set mReadWorkbook = workbook: Set mReadSnapshot = snapshot
    ElseIf Not workbook Is mReadWorkbook Then
        Err.Raise 5, "modCalculationSnapshot", "Вложенное чтение относится к другой книге."
    ElseIf Not snapshot Is Nothing Then
        snapshot.Validate
        Set mReadSnapshot = snapshot: Set mReadView = Nothing: mReadViewResolved = False
    End If
    mReadDepth = mReadDepth + 1
End Sub

' Освобождает большие массивы и ссылку на книгу после успеха или ошибки.
Public Sub EndCalculationSnapshotRead()
    If mReadDepth > 0 Then mReadDepth = mReadDepth - 1
    If mReadDepth = 0 Then
        Set mReadSnapshot = Nothing: Set mReadWorkbook = Nothing
        Set mReadView = Nothing: mReadViewResolved = False
    End If
End Sub

' Выбирает источник только по текущему Excel/File. File сразу читает указанное имя,
' Excel не обращается к файлу независимо от прежней привязки или смены имени.
Public Function CurrentCalculationSnapshot(ByVal workbook As Object) As CCalculationSnapshot
    Dim source As CResultReadSource: Set source = CurrentResultSource(workbook)
    If Not source Is Nothing Then Set CurrentCalculationSnapshot = source.LoadAll
End Function

' Возвращает адресный File либо готовый снимок текущего вывода; Excel без prepared читает лист.
Public Function CurrentResultSource(ByVal workbook As Object) As CResultReadSource
    If workbook Is mReadWorkbook Then
        If mReadViewResolved Then Set CurrentResultSource = mReadView: Exit Function
    End If
    Dim source As CResultReadSource, store As New CCalculationSnapshotStore
    Dim readingFile As Boolean
    On Error GoTo Failed
    If workbook Is mReadWorkbook And Not mReadSnapshot Is Nothing Then
        Set source = New CResultReadSource: source.UsePrepared mReadSnapshot
    ElseIf store.StorageMode(workbook) = "File" Then
        readingFile = True
        Set source = New CResultReadSource: source.OpenWorkbook workbook, store
    End If
    If mReadDepth > 0 And workbook Is mReadWorkbook Then Set mReadView = source: mReadViewResolved = True
    Set CurrentResultSource = source
    Exit Function
Failed:
    Dim number As Long, reason As String, origin As String
    number = Err.Number: reason = Err.Description: origin = Err.Source
    On Error Resume Next
    If readingFile And number <> 18 Then MarkResultSourceReadFailed workbook
    On Error GoTo 0
    Err.Raise number, origin, reason
End Function

' Выводит готовый расчёт из памяти; ручной File читает общие блоки и выбранное НДС.
' Ручное обновление Excel читает схему с листа и не восстанавливает верхние сводки.
Public Sub UpdateSectionResultsForWorkbook(ByVal workbook As Object, Optional ByVal prepared As CCalculationSnapshot = Nothing, _
        Optional ByRef plotWarning As String = "", Optional ByRef plotHasWarning As Boolean = False, _
        Optional ByRef publicationStarted As Boolean = False, Optional ByRef publicationRolledBack As Boolean = False)
    plotHasWarning = False: plotWarning = vbNullString
    Dim guard As CExcelAppStateGuard: Set guard = New CExcelAppStateGuard
    Dim entered As Boolean, sheetStarted As Boolean, sheetComplete As Boolean, indexWasMissing As Boolean
    Dim backup As Workbook, targets As New Collection, copies As New Collection
    Dim formulaTargets As New Collection, formulas As New Collection
    Dim copyObjects As Boolean, copySettingSaved As Boolean
    On Error GoTo Failed
    guard.Enter workbook.Application, True, (Not prepared Is Nothing And workbook.Application.EnableCancelKey = xlDisabled) ' Сохраняет непрерываемость связанной записи файла и таблиц.
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim mode As String: mode = store.StorageMode(workbook)
    Dim snapshot As CCalculationSnapshot: Set snapshot = prepared
    If snapshot Is Nothing Then
        If mode = "File" Then
            UpdateIndexedResultsForWorkbook workbook, plotWarning, plotHasWarning
            guard.Restore True
            CompleteWorkbookDrawing workbook
            Exit Sub
        Else
            store.ClearReference workbook ' Источник Excel — только лист; служебная файловая привязка не участвует в чтении.
            RequireCompleteExcelResultDisplay workbook
            UpdateSectionPlotForWorkbook workbook, True, plotWarning, True
            RefreshVisibleResultFreshness workbook, True
            plotHasWarning = (Len(plotWarning) > 0)
            guard.Restore True
    CompleteWorkbookDrawing workbook
            Exit Sub
        End If
    End If
    snapshot.Validate
    ValidateRestoredSnapshotLayout workbook, snapshot, mode = "Excel"
    RequireSnapshotPublicationReady workbook, mode
    Dim writer As CNDMResultsWriter: Set writer = New CNDMResultsWriter
    Dim summary As CBatchResultWriter: Set summary = New CBatchResultWriter
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook workbook
    Dim fullFormatting As Boolean: fullFormatting = settings.GetRequiredBoolean("General.ReformatResultsEnabled")
    If mode = "Excel" Then
        indexWasMissing = Not HasWorksheetResultIndex(workbook)
        writer.AddExcelPublicationRanges workbook, snapshot, targets
        summary.AddPublicationRanges workbook, snapshot, fullFormatting, targets
        copyObjects = workbook.Application.CopyObjectsWithCells: copySettingSaved = True
        workbook.Application.CopyObjectsWithCells = False ' В журнал ячеек не входят привязанные к ним фигуры пользователя.
        RememberResultsRanges workbook, targets, backup, copies, formulaTargets, formulas
    End If
    BeginCalculationSnapshotRead workbook, snapshot: entered = True
    sheetStarted = True
    publicationStarted = True ' До этого места проверены снимок, раскладка и параметры вывода.
    writer.RestoreSnapshotBase workbook, snapshot, fullFormatting, mode = "Excel"
    If snapshot.IsGeometryPreview Then
        summary.ClearSummary workbook, fullFormatting
    Else
        summary.RestoreSummary workbook, snapshot, fullFormatting
    End If
    If mode = "Excel" Then
        store.ClearReference workbook
    Else
        store.LinkFile workbook
    End If
    sheetComplete = True
    If Not backup Is Nothing Then backup.Close False: Set backup = Nothing
    If copySettingSaved Then workbook.Application.CopyObjectsWithCells = copyObjects: copySettingSaved = False
    If snapshot.IsGeometryPreview Then
        UpdateSectionGeometryPreviewForWorkbook workbook, plotWarning
        plotHasWarning = (Len(plotWarning) > 0)
    ElseIf prepared Is Nothing Then
        UpdateSectionPlotForWorkbook workbook, False, plotWarning
        plotHasWarning = (Len(plotWarning) > 0)
    Else
        plotWarning = TryUpdateCalculatedSnapshotPlot(workbook, snapshot, plotHasWarning)
    End If
    RefreshPublishedResultFreshness workbook, snapshot
    EndCalculationSnapshotRead
    entered = False
    guard.Restore True
    CompleteWorkbookDrawing workbook
    Exit Sub
Failed:
    Dim errorNumber As Long, reason As String, source As String
    errorNumber = Err.Number: reason = ResultOutputErrorDescription(Err.Description): source = Err.source
    If entered Then EndCalculationSnapshotRead: entered = False
    If sheetStarted And Not sheetComplete Then
        workbook.Application.EnableCancelKey = xlDisabled ' Возвращение согласованного вида нельзя отменить второй раз.
        If mode = "Excel" Then
            Dim recoveryReason As String
            publicationRolledBack = TryRestoreResultsRanges(targets, copies, formulaTargets, formulas, recoveryReason)
            If Len(recoveryReason) > 0 Then reason = WorkbookMessageParagraph(reason, recoveryReason)
        ElseIf TryRepairPublishedFileView(workbook, plotWarning, reason) Then
            plotHasWarning = True
            guard.Restore True
            CompleteWorkbookDrawing workbook
            Exit Sub
        End If
    End If
    On Error Resume Next
    If Not backup Is Nothing Then backup.Close False
    If copySettingSaved Then workbook.Application.CopyObjectsWithCells = copyObjects
    If indexWasMissing And (Not sheetStarted Or publicationRolledBack) Then
        If HasWorksheetResultIndex(workbook) Then workbook.Names.Item(OUTPUT_INDEX_NAME).Delete
    End If
    If publicationRolledBack Then MarkResultAttemptStopped workbook, errorNumber = 18, False
    On Error GoTo 0
    guard.Restore
    Err.Raise errorNumber, source, reason
End Sub

' Проверяет существование сгенерированного имени. До первого вывода
' индекс может отсутствовать, а при откате появившееся имя удаляется.
Private Function HasWorksheetResultIndex(ByVal workbook As Object) As Boolean
    Dim indexName As Object
    On Error Resume Next
    Set indexName = workbook.Names.Item(OUTPUT_INDEX_NAME)
    On Error GoTo 0
    HasWorksheetResultIndex = Not indexName Is Nothing
End Function

' Сохраняет точные ячейки будущей записи. Объединения и оформление переносит
' Excel; формулы возвращаются отдельно из исходной книги по своим адресам.
Private Sub RememberResultsRanges(ByVal workbook As Object, ByVal targets As Collection, ByRef backup As Workbook, _
        ByVal copies As Collection, ByVal formulaTargets As Collection, ByVal formulas As Collection)
    Dim current As Object: Set current = workbook.Application.ActiveWindow
    Set backup = workbook.Application.Workbooks.Add(xlWBATWorksheet)
    backup.Windows.Item(1).Visible = False
    If Not current Is Nothing Then current.Activate
    Dim sheets As Object: Set sheets = CreateObject("Scripting.Dictionary")
    Dim target As Range, shadow As Worksheet, saved As Range
    Dim containsFormulas As Variant, formulaCells As Range, area As Range
    For Each target In targets
        If sheets.Exists(target.Worksheet.Name) Then
            Set shadow = sheets.Item(target.Worksheet.Name)
        Else
            If sheets.Count = 0 Then Set shadow = backup.Worksheets.Item(1) Else Set shadow = backup.Worksheets.Add
            sheets.Add target.Worksheet.Name, shadow
        End If
        Set saved = shadow.Range(target.Address)
        saved.UnMerge ' Смежные области могут включать одну общую строку заголовка.
        target.Copy saved
        copies.Add saved
        containsFormulas = target.HasFormula
        If IsNull(containsFormulas) Or containsFormulas Then
            If target.CountLarge = 1 Then Set formulaCells = target Else Set formulaCells = target.SpecialCells(xlCellTypeFormulas)
            For Each area In formulaCells.Areas
                formulaTargets.Add area: formulas.Add area.Formula
                shadow.Range(area.Address).ClearContents ' Обратное копирование не переносит внешние формульные ссылки.
            Next area
        End If
    Next target
End Sub

' Возвращает только сохранённые области. При отказе Excel барьер чтения остаётся.
Private Function TryRestoreResultsRanges(ByVal targets As Collection, ByVal copies As Collection, _
        ByVal formulaTargets As Collection, ByVal formulas As Collection, ByRef failure As String) As Boolean
    On Error GoTo Failed
    Dim i As Long, target As Range, saved As Range
    If targets.Count <> copies.Count Then Exit Function
    For i = 1 To targets.Count
        Set target = targets.Item(i): Set saved = copies.Item(i)
        target.UnMerge
        saved.Copy target
    Next i
    For i = 1 To formulaTargets.Count
        Set target = formulaTargets.Item(i)
        target.Formula = formulas.Item(i)
    Next i
    TryRestoreResultsRanges = True
Failed:
    If Not TryRestoreResultsRanges Then
        Dim cause As String: cause = Err.Description
        failure = "Не удалось полностью восстановить прежний вывод"
        If Not target Is Nothing Then failure = failure & " в " & target.Address(External:=True)
        failure = failure & ". " & cause
    End If
End Function

' В File уже записан авторитетный файл. Восстанавливает его имя и вывод
' штатным адресным путём, без повторного вызова расчёта или записи файла.
Private Function TryRepairPublishedFileView(ByVal workbook As Object, ByRef plotWarning As String, ByVal originalError As String) As Boolean
    On Error GoTo Failed
    Dim warning As String, hasWarning As Boolean, sheetComplete As Boolean
    UpdateIndexedResultsForWorkbook workbook, warning, hasWarning, sheetComplete, True
    GoTo Repaired
Failed:
    If Not sheetComplete Then Exit Function
    warning = Err.Description ' Таблицы согласованы даже при отказе схемы.
Repaired:
    plotWarning = WorkbookMessageParagraph("После ошибки записи вывод восстановлен из сохранённого файла результатов.", originalError)
    plotWarning = WorkbookMessageParagraph(plotWarning, warning)
    TryRepairPublishedFileView = True
End Function

' Проверяет вместимость готовых таблиц до очистки; число LC не берётся из текущего Config.
Private Sub ValidateRestoredSnapshotLayout(ByVal workbook As Object, ByVal snapshot As Object, ByVal includeElements As Boolean)
    Dim key As Variant, values As Variant, anchor As Object, rowCount As Long
    Set anchor = workbook.Names.Item("rngNDMElementGeometry").RefersToRange ' В старом payload её может не быть; текущая книга всё равно выводит шапку до любых очисток.
    If Not ResultRowsFitSheet(anchor.Row, 1, anchor.Worksheet.Rows.Count) Then Err.Raise 5, "ValidateRestoredSnapshotLayout", "Не помещается шапка rngNDMElementGeometry."
    For Each key In snapshot.BaseKeys()
        Set anchor = workbook.Names.Item(CStr(key)).RefersToRange
        rowCount = snapshot.TableRowCount(CStr(key))
        If Not includeElements Then
            If CStr(key) = "rngNDMElementResults" Or CStr(key) = "rngNDMElementGeometry" Or CStr(key) = "rngNDMSectionAnnotations" Then rowCount = 1
        End If
        If includeElements And Not snapshot.IsGeometryPreview Then
            If CStr(key) = "rngNDMSectionProperties" Then rowCount = rowCount + snapshot.TableRowCount("LoadCase.Properties") - 1
            If CStr(key) = "rngNDMSectionAnnotations" Then rowCount = rowCount + snapshot.TableRowCount("LoadCase.Annotations") - 1
        End If
        If Not ResultRowsFitSheet(anchor.Row, rowCount, anchor.Worksheet.Rows.Count) Then
            Dim reason As String: reason = ResultRowLimitDescription(CStr(key), anchor.Row, rowCount, anchor.Worksheet.Rows.Count)
            If CStr(key) = "rngNDMElementResults" Or CStr(key) = "rngNDMElementGeometry" Then
                If RestoredSnapshotMandatoryRowsFit(workbook, snapshot) Then reason = reason & vbCrLf & "Что сделать:" & vbCrLf & "1) " & ExternalSnapshotStorageAdvice(workbook) & " Повторите «Обновить схему и результаты»."
            End If
            Err.Raise 5, "ValidateRestoredSnapshotLayout", reason
        End If
    Next key
    If snapshot.HasTable(COUPLED_FIELD_TABLE) Then
        Set anchor = NamedCoupledFieldAnchor(workbook)
        If anchor Is Nothing Then Err.Raise 5, "ValidateRestoredSnapshotLayout", _
            "В книге отсутствует именованный диапазон " & COUPLED_FIELD_TABLE & "."
        rowCount = 1
        If includeElements Then rowCount = snapshot.TableRowCount(COUPLED_FIELD_TABLE)
        If Not ResultRowsFitSheet(anchor.Row, rowCount, anchor.Worksheet.Rows.Count) Then Err.Raise 5, _
            "ValidateRestoredSnapshotLayout", ResultRowLimitDescription(COUPLED_FIELD_TABLE, anchor.Row, rowCount, anchor.Worksheet.Rows.Count)
    End If
    If snapshot.IsGeometryPreview Then Exit Sub
    Dim names As Variant, keys As Variant, nextNames As Variant, headers As Variant, i As Long
    names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    keys = Array("Summary", "Strength.Data", "Cracks.Data", "Stability.Data")
    nextNames = Array("rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor", "rngNDMSectionGeometry")
    headers = Array(4, 4, 5, 2)
    For i = 0 To 3
        Set anchor = workbook.Names.Item(CStr(names(i))).RefersToRange: values = snapshot.Table(CStr(keys(i)))
        Dim nextAnchor As Object: Set nextAnchor = workbook.Names.Item(CStr(nextNames(i))).RefersToRange
        If anchor.Row + UBound(values, 1) + CLng(headers(i)) + 1 >= nextAnchor.Row Then Err.Raise 5, , _
            "В снимке больше строк, чем в текущей раскладке Results: " & CStr(names(i)) & " пересекает " & CStr(nextNames(i)) & ". Увеличьте расстояние между именованными диапазонами."
    Next i
End Sub

' Проверяет таблицы, остающиеся на листе при File, без копирования матриц.
' Используется только при отказе вместимости таблицы состояний: File нельзя
' предлагать, если общая геометрия или готовые сводки тоже не помещаются.
Private Function RestoredSnapshotMandatoryRowsFit(ByVal workbook As Object, ByVal snapshot As Object) As Boolean
    Dim key As Variant, anchor As Object
    For Each key In snapshot.BaseKeys()
        If CStr(key) <> "rngNDMElementResults" And CStr(key) <> "rngNDMElementGeometry" And CStr(key) <> "rngNDMSectionAnnotations" Then
            Set anchor = workbook.Names.Item(CStr(key)).RefersToRange
            If Not ResultRowsFitSheet(anchor.Row, snapshot.TableRowCount(CStr(key)), anchor.Worksheet.Rows.Count) Then Exit Function
        End If
    Next key
    If snapshot.IsGeometryPreview Then RestoredSnapshotMandatoryRowsFit = True: Exit Function
    Dim names As Variant, keys As Variant, headerRows As Variant, i As Long
    names = Array("rngBatchSummary", "rngStrengthSummaryAnchor", "rngCrackSummaryAnchor", "rngStabilitySummaryAnchor")
    keys = Array("Summary", "Strength.Data", "Cracks.Data", "Stability.Data")
    headerRows = Array(0, 4, 4, 5)
    For i = 0 To 3
        Set anchor = workbook.Names.Item(CStr(names(i))).RefersToRange
        If Not ResultRowsFitSheet(anchor.Row - CLng(headerRows(i)), CDbl(snapshot.TableRowCount(CStr(keys(i)))) + CLng(headerRows(i)), anchor.Worksheet.Rows.Count) Then Exit Function
    Next i
    RestoredSnapshotMandatoryRowsFit = True
End Function

' Проверяет фактическую вместимость листа с учётом первой строки, без переполнения Long.
' Предел передаётся из Worksheet.Rows.Count: .xls и .xlsm имеют разную высоту листа.
Public Function ResultRowsFitSheet(ByVal firstRow As Long, ByVal rowCount As Double, ByVal rowLimit As Long) As Boolean
    If firstRow < 1 Or rowCount < 1# Or rowLimit < firstRow Then Exit Function
    ResultRowsFitSheet = rowCount <= CDbl(rowLimit) - CDbl(firstRow) + 1#
End Function

' Называет потребное число строк и фактический предел текущего листа, включая смещение якоря.
Public Function ResultRowLimitDescription(ByVal tableName As String, ByVal firstRow As Long, ByVal rowCount As Double, ByVal rowLimit As Long) As String
    ResultRowLimitDescription = tableName & " требует " & Format$(rowCount, "0") & " строк от строки " & CStr(firstRow) & _
        " (последняя строка " & Format$(CDbl(firstRow) + rowCount - 1#, "0") & "). " & _
        "Этот лист Excel поддерживает " & CStr(rowLimit) & " строк; ниже начала диапазона доступно " & CStr(rowLimit - firstRow + 1) & "."
End Function

' Предлагает File с реальным адресом настройки; выбор пользователя автоматически не изменяет.
Public Function ExternalSnapshotStorageAdvice(ByVal workbook As Object) As String
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader
    settings.LoadFromRange workbook.Names.Item("rngSystemSettings").RefersToRange
    ExternalSnapshotStorageAdvice = "На Config установите Results.SnapshotStorage = File (" & settings.InputLocation("Results.SnapshotStorage") & _
        "), чтобы сохранять полный снимок во внешнем файле .ndm: для данных снимка на листе Excel недостаточно строк."
End Function

' Объясняет подтверждённый отказ Excel из-за объединённых ячеек, сохраняя прочие причины без подмены.
' Одного кода 1004 недостаточно: он используется Excel для множества разных ошибок.
Public Function ResultOutputErrorDescription(ByVal original As String) As String
    ResultOutputErrorDescription = original
    Dim normalized As String: normalized = LCase$(Replace$(original, "ё", "е"))
    If (InStr(normalized, "объединен") > 0 And InStr(normalized, "ячей") > 0) Or InStr(normalized, "merged cell") > 0 Then
        ResultOutputErrorDescription = "Не удалось вывести результаты на лист Results: область вывода попадает в объединённые ячейки. " & _
            "Снимите объединение в месте вывода или перенесите именованный диапазон в свободную необъединённую область, затем повторите действие."
    End If
End Function

' Сохраняет полный снимок только в File; оба режима выводят результаты из памяти.
' Excel оставляет обычные читаемые таблицы без второго архива подробных сводок.
Public Function SaveAndDisplayCalculationSnapshot(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, _
        Optional ByRef plotHasWarning As Boolean = False, Optional ByRef publicationStarted As Boolean = False, _
        Optional ByRef publicationRolledBack As Boolean = False) As String
    plotHasWarning = False
    snapshot.Validate
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim mode As String: mode = store.StorageMode(workbook)
    ValidateRestoredSnapshotLayout workbook, snapshot, mode = "Excel"
    RequireSnapshotPublicationReady workbook, mode
    If mode = "File" Then
        store.SaveFile workbook, snapshot
        publicationStarted = True ' Теперь выбранный файл уже содержит новый сохранённый результат.
    End If
    Dim warning As String
    UpdateSectionResultsForWorkbook workbook, snapshot, warning, plotHasWarning, publicationStarted, publicationRolledBack
    SaveAndDisplayCalculationSnapshot = warning
End Function

' Проверяет защиту перед первой записью или заменой файла. Excel запрещает
' обратное копирование форматов на лист с защитой даже при UserInterfaceOnly.
Private Sub RequireSnapshotPublicationReady(ByVal workbook As Object, ByVal mode As String)
    PrepareIndexedResultsWrite workbook
    If mode <> "Excel" Then Exit Sub
    Dim sheet As Object: Set sheet = workbook.Names.Item("rngNDMElementResults").RefersToRange.Worksheet
    If Not sheet.ProtectContents Then Exit Sub
    Dim settings As New CSystemSettingsReader
    settings.LoadFromRange workbook.Names.Item("rngSystemSettings").RefersToRange
    Err.Raise 5, "RequireSnapshotPublicationReady", _
        "Защита листа «" & sheet.Name & "» не позволяет Excel вернуть оформление при прерывании записи. Прежний вывод сохранён." & _
        vbCrLf & vbCrLf & "Снимите защиту для вывода в Excel либо выберите File (" & settings.InputLocation("Results.SnapshotStorage") & ")."
End Sub

' Обновляет историю нагружения с таблицами Results. Схема НДС зависит от
' настроек и наличия результатов элементов. Ошибка вывода оставляет расчёт сохранённым.
Private Function TryUpdateCalculatedSnapshotPlot(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, _
        ByRef plotHasWarning As Boolean) As String
    plotHasWarning = False
    On Error GoTo Failed
    Dim settings As CSystemSettingsReader: Set settings = New CSystemSettingsReader: settings.LoadFromWorkbook workbook
    Dim updateScheme As Boolean: updateScheme = settings.GetRequiredBoolean("Plot.Enabled")
    If updateScheme Then updateScheme = settings.GetRequiredBoolean("Plot.AutoUpdateAfterCalculation")
    Dim containerNotice As String
    If updateScheme And snapshot.TableRowCount("rngNDMElementResults") > 1 Then
        UpdateSectionPlotForWorkbook workbook, False, containerNotice
    Else
        Dim traceNotice As String: traceNotice = TryUpdateCoupledLoadChart(workbook, settings)
        Dim plotter As New CSectionPlotter
        If updateScheme Then
            plotter.ClearExisting workbook, "Схема не обновлена: нет доступных расчётных состояний LC.", containerNotice
        Else
            plotter.MarkResultsNotDisplayed workbook, PLOT_NOT_CURRENT
        End If
        containerNotice = WorkbookMessageParagraph(containerNotice, traceNotice)
    End If
    plotHasWarning = (Len(containerNotice) > 0)
    TryUpdateCalculatedSnapshotPlot = containerNotice
    Exit Function
Failed:
    Dim reason As String: reason = Err.Description
    On Error Resume Next
    plotter.MarkResultsNotDisplayed workbook, PLOT_NOT_CURRENT
    On Error GoTo 0
    plotHasWarning = True
    TryUpdateCalculatedSnapshotPlot = "Предупреждение:" & vbCrLf & "Схема не обновлена. Расчётные результаты сохранены." & vbCrLf & vbCrLf & _
        reason & vbCrLf & "Исправьте настройки оформления и нажмите «Обновить схему и результаты». Повторный расчёт не требуется."
End Function

' Выбирает предупреждение по числу строк и фактической вместимости листа.
' mandatoryRows содержит записи (имя диапазона, число строк, первая строка)
' всех таблиц, остающихся на Results при File. RowLimit запрещает предложение
' File: переполнение обязательной геометрии режим хранения не устраняет.
Public Function CalculationSnapshotOutputPolicy(ByVal storageMode As String, ByVal estimatedRows As Double, _
        ByRef mandatoryRows As Variant, ByVal rowLimit As Long, Optional ByVal estimatedBoundaryRows As Double = 0#) As String
    Dim i As Long, entry As Variant, largeMandatory As Boolean
    If IsArray(mandatoryRows) Then
        For i = LBound(mandatoryRows) To UBound(mandatoryRows)
            entry = mandatoryRows(i)
            If Not ResultRowsFitSheet(CLng(entry(2)), CDbl(entry(1)), rowLimit) Then
                CalculationSnapshotOutputPolicy = "RowLimit"
                Exit Function
            End If
            If CDbl(entry(1)) > 60000# Then largeMandatory = True
        Next i
    End If
    If storageMode = "Excel" And (estimatedRows > 60000# Or estimatedBoundaryRows > 60000#) Then CalculationSnapshotOutputPolicy = "File"
    If largeMandatory Then
        If Len(CalculationSnapshotOutputPolicy) > 0 Then
            CalculationSnapshotOutputPolicy = "FileAndMandatory"
        Else
            CalculationSnapshotOutputPolicy = "Mandatory"
        End If
    End If
End Function

' Предупреждает о крупных обязательных таблицах в Excel и File, независимо
' от способа формирования сетки. Переключение в File предлагается только
' для больших состояний или собственных границ и при вместимости остальных таблиц Results.
Public Sub OfferExternalCalculationSnapshot(ByVal workbook As Object, ByVal estimatedRows As Double, ByVal interactive As Boolean, _
        Optional ByVal mandatoryRows As Variant, Optional ByVal estimatedBoundaryRows As Double = 0#)
    If Not interactive Then Exit Sub
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    Dim policy As String: policy = CalculationSnapshotOutputPolicy(store.StorageMode(workbook), estimatedRows, _
        mandatoryRows, workbook.Worksheets.Item("Results").Rows.Count, estimatedBoundaryRows)
    If Len(policy) = 0 Or policy = "RowLimit" Then Exit Sub ' Отказ по пределу листа сообщает общая предварительная проверка до расчёта.
    Dim message As String, i As Long, entry As Variant, table As Object
    If policy = "Mandatory" Or policy = "FileAndMandatory" Then
        message = "На Results предстоит вывести большие обязательные таблицы."
        For i = LBound(mandatoryRows) To UBound(mandatoryRows)
            entry = mandatoryRows(i)
            If CDbl(entry(1)) > 60000# Then
                Set table = workbook.Names.Item(CStr(entry(0))).RefersToRange
                message = message & vbCrLf & CStr(entry(0)) & " (" & table.Parent.Name & "!" & _
                    table.Address(False, False) & "): около " & Format$(CDbl(entry(1)), "0") & " строк."
            End If
        Next i
        message = message & vbCrLf & vbCrLf & "Эти таблицы сохраняются на листе и при Results.SnapshotStorage = File. " & _
            "Переключение снимка в файл не уменьшает число строк общей геометрии, общих контуров, характеристик и сводок."
    End If
    If policy = "Mandatory" Then
        ShowWorkbookMessage message, vbExclamation, "Большой объём Results"
        Exit Sub
    End If
    If Len(message) > 0 Then message = message & vbCrLf & vbCrLf
    message = message & "Предварительная оценка подробных таблиц:" & vbCrLf & _
        "Состояния элементов: около " & Format$(estimatedRows, "0") & " строк." & vbCrLf & _
        "Фактические границы элементов: около " & Format$(estimatedBoundaryRows, "0") & " строк." & vbCrLf & vbCrLf & _
        "Внешний файл .ndm сохранит полные состояния и фактические границы элементов без их вывода на лист. " & _
        "Общая геометрия, общие контуры, использованные диаграммы и подробные сводки останутся на Results." & vbCrLf & vbCrLf & _
        "Переключить Results.SnapshotStorage на File?" & vbCrLf & _
        "Да — переключить и продолжить. Нет — продолжить с хранением в Excel."
    Dim answer As VbMsgBoxResult: answer = MsgBox(message, vbQuestion Or vbYesNo, "Хранение большого снимка")
    If answer <> vbYes Then Exit Sub
    Dim values As Variant
    Set table = workbook.Names.Item("rngSystemSettings").RefersToRange: values = table.Value2
    For i = 1 To UBound(values, 1)
        If CStr(values(i, 1)) = "Results.SnapshotStorage" Then table.Cells(i, 2).Value2 = "File": Exit Sub
    Next i
    Err.Raise 5, , "В rngSystemSettings отсутствует Results.SnapshotStorage."
End Sub

' Удаляет привязку готового расчёта перед обычной записью preview или тестового writer.
Public Sub ClearPublishedCalculationSnapshot(ByVal workbook As Object)
    InvalidateElementResultIndex workbook
    PrepareIndexedResultsWrite workbook
    Dim store As CCalculationSnapshotStore: Set store = New CCalculationSnapshotStore
    store.ClearReference workbook
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Доказывает освобождение большого снимка и книги после законченной операции.
Public Function TestCalculationSnapshotReadReleased() As Boolean
    TestCalculationSnapshotReadReleased = (mReadDepth = 0 And mReadSnapshot Is Nothing And mReadWorkbook Is Nothing And mReadView Is Nothing)
End Function

' Ручное обновление File восстанавливает общие таблицы и выбранную схему без всех НДС.
' Источник сначала проверяет обязательные общие блоки, затем допускает изменение Results.
Private Sub UpdateIndexedResultsForWorkbook(ByVal workbook As Object, ByRef plotWarning As String, ByRef plotHasWarning As Boolean, _
        Optional ByRef displayFinished As Boolean = False, Optional ByVal repairing As Boolean = False)
    Dim entered As Boolean, sourceVerified As Boolean, displayStarted As Boolean
    On Error GoTo Failed
    BeginCalculationSnapshotRead workbook: entered = True
    Dim source As CResultReadSource: Set source = CurrentResultSource(workbook)
    source.ValidateForDisplay
    sourceVerified = True
    ValidateRestoredSnapshotLayout workbook, source, False
    Dim writer As New CNDMResultsWriter, summary As New CBatchResultWriter, store As New CCalculationSnapshotStore
    Dim settings As New CSystemSettingsReader: settings.LoadFromWorkbook workbook
    Dim fullFormatting As Boolean: fullFormatting = settings.GetRequiredBoolean("General.ReformatResultsEnabled")
    Dim preview As Boolean: preview = source.IsGeometryPreview
    Dim worstCriterion As String, priorAttempt As String
    If Not preview Then
        worstCriterion = summary.SavedWorstCriterion(settings)
        If Not repairing Then priorAttempt = PriorAttemptNoticeForSnapshot(workbook, source)
    End If
    displayStarted = True
    writer.RestoreIndexedBase workbook, source, fullFormatting
    If preview Then
        summary.ClearSummary workbook, fullFormatting
    Else
        summary.RestoreIndexedSummary workbook, source, fullFormatting, worstCriterion
    End If
    store.LinkFile workbook
    RefreshPublishedResultFreshness workbook, source, priorAttempt
    displayFinished = True
    Dim updateScheme As Boolean: updateScheme = True
    If repairing And Not preview Then
        Dim plotter As New CSectionPlotter
        plotter.MarkResultsNotDisplayed workbook, PLOT_NOT_CURRENT
        updateScheme = settings.GetRequiredBoolean("Plot.Enabled")
        If updateScheme Then updateScheme = settings.GetRequiredBoolean("Plot.AutoUpdateAfterCalculation")
    End If
    If preview Then
        UpdateSectionGeometryPreviewForWorkbook workbook, plotWarning
    ElseIf updateScheme Then
        UpdateSectionPlotForWorkbook workbook, False, plotWarning
    Else
        plotWarning = TryUpdateCoupledLoadChart(workbook, settings)
    End If
    plotHasWarning = (Len(plotWarning) > 0)
    EndCalculationSnapshotRead
    entered = False
    Exit Sub
Failed:
    Dim number As Long, reason As String, origin As String
    number = Err.Number: reason = Err.Description: origin = Err.Source
    If entered Then EndCalculationSnapshotRead
    On Error Resume Next
    If displayStarted And Not displayFinished Then MarkResultAttemptStopped workbook, False, True
    If Not sourceVerified And number <> 18 Then MarkResultSourceReadFailed workbook
    On Error GoTo 0
    Err.Raise number, origin, reason
End Sub
