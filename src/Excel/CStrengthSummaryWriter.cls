VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CStrengthSummaryWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подробная таблица НДС по прочности и несущей способности на Results
' ==========================================================================
' Записывает от rngStrengthSummaryAnchor сохранённые состояния StrengthState
' и CapacityState с переводом в пользовательские единицы.
' Для отсутствующей физической зоны оставляет пустую ячейку: например,
' если всё сечение растянуто, сжатые напряжения и деформации не подменяются
' значениями другого знака.

Private mCapturedSnapshot As CCalculationSnapshot ' Приёмник готовых значений без записи и оформления листа.
Private mMinimalFormatting As Boolean ' Сохраняет готовое оформление; обновляет данные, шапки при изменении и цвета статусов.

Private Const HEADER_ROWS As Long = 4 ' Строки шапки над первой строкой LC.
Private Const TOTAL_COLUMNS As Long = STRENGTH_SNAPSHOT_COLUMNS

Private Const COL_LEFT_FIRST As Long = 1 ' Столбец таблицы: первый столбец общих сведений о сочетании.
Private Const COL_LEFT_LAST As Long = 9 ' Столбец таблицы: последний столбец общих сведений о сочетании.
Private Const COL_SEPARATOR1 As Long = 10 ' Столбец таблицы: пустой столбец перед текущим НДС.
Private Const COL_STATE_FIRST As Long = 11 ' Столбец таблицы: первый столбец текущего НДС.
Private Const COL_STATE_LAST As Long = 30 ' Столбец таблицы: последний столбец текущего НДС.
Private Const COL_SEPARATOR2 As Long = 31 ' Столбец таблицы: пустой столбец перед несущей способностью.
Private Const COL_CAPACITY_FIRST As Long = 32 ' Столбец таблицы: первый столбец несущей способности.
Private Const COL_CAPACITY_LAST As Long = 52 ' Столбец таблицы: последний столбец несущей способности.

Private Const COL_ID As Long = 1 ' Столбец таблицы: номер сочетания.
Private Const COL_RESULT_COMMENT As Long = 2 ' Столбец таблицы: пояснение результата сочетания.
Private Const COL_STATUS As Long = 3 ' Столбец таблицы: общий статус сочетания.
Private Const COL_REF_N As Long = 4 ' Продольная сила для расчёта в центре приведённого сечения.
Private Const COL_REF_MX As Long = 5 ' Момент для расчёта относительно центральной оси X.
Private Const COL_REF_MY As Long = 6 ' Момент для расчёта относительно центральной оси Y.
Private Const COL_LOAD_N As Long = 7 ' Столбец таблицы: расчётная продольная сила.
Private Const COL_LOAD_MX As Long = 8 ' Столбец таблицы: расчётный момент Mx.
Private Const COL_LOAD_MY As Long = 9 ' Столбец таблицы: расчётный момент My.

Private Const COL_LIM_EBC As Long = 11 ' Столбец таблицы: предельная деформация сжатого бетона.
Private Const COL_LIM_EBT As Long = 12 ' Столбец таблицы: предельная деформация растянутого бетона.
Private Const COL_LIM_ESC As Long = 13 ' Столбец таблицы: предельная деформация сжатой арматуры.
Private Const COL_LIM_EST As Long = 14 ' Столбец таблицы: предельная деформация растянутой арматуры.
Private Const COL_STATE_EPS0 As Long = 15 ' Столбец таблицы: текущее НДС: деформация epsilon0.
Private Const COL_STATE_KX As Long = 16 ' Столбец таблицы: текущее НДС: кривизна kappax.
Private Const COL_STATE_KY As Long = 17 ' Столбец таблицы: текущее НДС: кривизна kappay.
Private Const COL_STATE_EBC As Long = 18 ' Столбец таблицы: текущее НДС: наибольшая деформация сжатия бетона.
Private Const COL_STATE_EBT As Long = 19 ' Столбец таблицы: текущее НДС: наибольшая деформация растяжения бетона.
Private Const COL_STATE_ESC As Long = 20 ' Столбец таблицы: текущее НДС: наибольшая деформация сжатия арматуры.
Private Const COL_STATE_EST As Long = 21 ' Столбец таблицы: текущее НДС: наибольшая деформация растяжения арматуры.
Private Const COL_STATE_SBC As Long = 22 ' Столбец таблицы: текущее НДС: наибольшее напряжение сжатия бетона.
Private Const COL_STATE_SBT As Long = 23 ' Столбец таблицы: текущее НДС: наибольшее напряжение растяжения бетона.
Private Const COL_STATE_SSC As Long = 24 ' Столбец таблицы: текущее НДС: наибольшее напряжение сжатия арматуры.
Private Const COL_STATE_SST As Long = 25 ' Столбец таблицы: текущее НДС: наибольшее напряжение растяжения арматуры.
Private Const COL_STATE_H As Long = 26 ' Столбец таблицы: текущее НДС: высота сечения.
Private Const COL_STATE_X As Long = 27 ' Столбец таблицы: текущее НДС: высота сжатой зоны.
Private Const COL_STATE_DIAGEXT As Long = 28 ' Столбец таблицы: текущее НДС: использование продолжения диаграмм.
Private Const COL_STATE_RESERVE As Long = 29 ' Столбец таблицы: текущее НДС: коэффициент запаса.
Private Const COL_STATE_STATUS As Long = 30 ' Столбец таблицы: текущее НДС: статус расчёта.

Private Const COL_CAP_EPS0 As Long = 32 ' Столбец таблицы: предельное НДС: деформация epsilon0.
Private Const COL_CAP_KX As Long = 33 ' Столбец таблицы: предельное НДС: кривизна kappax.
Private Const COL_CAP_KY As Long = 34 ' Столбец таблицы: предельное НДС: кривизна kappay.
Private Const COL_CAP_LIMIT_STATE As Long = 35 ' Столбец таблицы: предельное НДС: определяющий предел деформаций.
Private Const COL_CAP_LOAD_PATH As Long = 36 ' Столбец таблицы: предельное НДС: траектория изменения нагрузки.
Private Const COL_CAP_METHOD As Long = 37 ' Столбец таблицы: предельное НДС: способ поиска предела.
Private Const COL_CAP_LAMBDA As Long = 38 ' Столбец таблицы: предельное НДС: множитель предельной нагрузки.
Private Const COL_CAP_N As Long = 39 ' Столбец таблицы: предельное НДС: предельная сила N.
Private Const COL_CAP_MX As Long = 40 ' Столбец таблицы: предельное НДС: предельный момент Mx.
Private Const COL_CAP_MY As Long = 41 ' Столбец таблицы: предельное НДС: предельный момент My.
Private Const COL_CAP_T As Long = 42 ' Крутящий момент в предельной точке.
Private Const COL_CAP_QX As Long = 43 ' Поперечная сила X в предельной точке.
Private Const COL_CAP_QY As Long = 44 ' Поперечная сила Y в предельной точке.
Private Const COL_CAP_EBC As Long = 45 ' Столбец таблицы: предельное НДС: наибольшая деформация сжатия бетона.
Private Const COL_CAP_EBT As Long = 46 ' Столбец таблицы: предельное НДС: наибольшая деформация растяжения бетона.
Private Const COL_CAP_ESC As Long = 47 ' Столбец таблицы: предельное НДС: наибольшая деформация сжатия арматуры.
Private Const COL_CAP_EST As Long = 48 ' Столбец таблицы: предельное НДС: наибольшая деформация растяжения арматуры.
Private Const COL_CAP_H As Long = 49 ' Столбец таблицы: предельное НДС: высота сечения.
Private Const COL_CAP_X As Long = 50 ' Столбец таблицы: предельное НДС: высота сжатой зоны.
Private Const COL_CAP_RESERVE As Long = 51 ' Столбец таблицы: предельное НДС: коэффициент запаса.
Private Const COL_CAP_STATUS As Long = 52 ' Столбец таблицы: предельное НДС: статус расчёта.
Private Const COL_SEPARATOR3 As Long = 53 ' Пустой столбец перед поперечной арматурой.
Private Const COL_TRANSVERSE_POINT As Long = 54 ' Контрольная точка поперечной арматуры при заданных усилиях.
Private Const COL_TRANSVERSE_STRESS As Long = 55 ' Знаковое напряжение в той же точке.
Private Const COL_TRANSVERSE_STRAIN As Long = 56 ' Знаковая деформация при заданных усилиях.
Private Const COL_TRANSVERSE_STRAIN_LIMIT As Long = 57 ' Предел деформации класса стали и её знака.
Private Const COL_TRANSVERSE_UTIL As Long = 58 ' Отношение деформации к пределу своего знака.
Private Const COL_TRANSVERSE_STATUS As Long = 59 ' Статус критерия деформации поперечной арматуры.
Private Const COL_TRANSVERSE_RESERVE As Long = 60 ' Коэффициент запаса, когда поперечная арматура определяет предел.
Private Const COL_TRANSVERSE_LIMIT As Long = 61 ' Определяющая точка и причина арматурного предела.
Private Const COL_TRANSVERSE_LIMIT_STRESS As Long = 62 ' Напряжение в сохранённом предельном НДС хомута.
Private Const COL_TRANSVERSE_LIMIT_STRAIN As Long = 63 ' Деформация в сохранённом предельном НДС хомута.
Private Const COL_TRANSVERSE_LIMIT_STRAIN_LIMIT As Long = 64 ' Знаковый предел деформации класса стали в предельном НДС.
Private Const COL_TRANSVERSE_LIMIT_STATUS As Long = 65 ' Статус хомута при определении предельного множителя.

Private Const HEADER_COLOR As Long = &HD9D9D9 ' Серый фон шапки, как в остальных подробных блоках.
Private Const SMALL_VALUE As Double = 0.000000001 ' Порог для определения отсутствующего численного поля при выводе таблицы.
Private Const HEADER_NOTES_ROW As Long = 3 ' Строка пояснений в шапке таблицы.
Private Const HEADER_LABEL_ROW As Long = 4 ' Строка обозначений и единиц в шапке таблицы.

' Возвращает число строк шапки над rngStrengthSummaryAnchor.
Public Function HeaderRowsAboveAnchor() As Long
    HeaderRowsAboveAnchor = HEADER_ROWS
End Function

' Возвращает полную высоту блока: шапка плюс строки сочетаний.
Public Function RequiredRows(ByVal combinationCount As Long) As Long
    If combinationCount < 1 Then combinationCount = 1
    RequiredRows = HEADER_ROWS + combinationCount
End Function

' Возвращает число строк данных по высоте rngLoadCombinations, чтобы пустые
' все доступные строки получили полные рамки.
Public Function RequiredDataRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredDataRowsForWorkbook = combinationCount
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1

    On Error GoTo SafeExit
    If workbook Is Nothing Then GoTo SafeExit
    Dim loads As Object
    Set loads = workbook.Names.Item("rngLoadCombinations").RefersToRange
    If loads.Rows.Count > 1 Then
        If loads.Rows.Count - 1 > RequiredDataRowsForWorkbook Then _
            RequiredDataRowsForWorkbook = loads.Rows.Count - 1
    End If
SafeExit:
    If RequiredDataRowsForWorkbook < 1 Then RequiredDataRowsForWorkbook = 1
End Function

' Возвращает полную высоту блока для фактической книги.
Public Function RequiredRowsForWorkbook(ByVal workbook As Object, ByVal combinationCount As Long) As Long
    RequiredRowsForWorkbook = RequiredRows(RequiredDataRowsForWorkbook(workbook, combinationCount))
End Function

' Возвращает ширину блока вместе с пустыми разделителями.
Public Function RequiredColumns() As Long
    RequiredColumns = TOTAL_COLUMNS
End Function

' Записывает подробную таблицу НДС по прочности и предельной несущей способности.
Public Sub WriteSummary(ByVal workbook As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByVal units As CUnitSystem = Nothing, Optional ByVal section As CSectionModel = Nothing, Optional ByVal fullFormatting As Boolean = True)
    mMinimalFormatting = Not fullFormatting
    On Error GoTo Failed
    If workbook Is Nothing Then Err.Raise vbObjectError + 3996, "CStrengthSummaryWriter", "Книга Excel не передана."
    If batch Is Nothing Then Err.Raise vbObjectError + 3997, "CStrengthSummaryWriter", "Пакетный расчетчик не передан."

    Dim target As Object
    Set target = TargetRange(workbook, RequiredDataRowsForWorkbook(workbook, batch.Count))
    If mMinimalFormatting Then
        mMinimalFormatting = (CStr(target.Cells(target.Rows.Count, 1).NumberFormat) = "@")
    End If
    If mCapturedSnapshot Is Nothing Then PrepareTarget target
    WriteHeaders target, units
    If mCapturedSnapshot Is Nothing Then WriteBlockTitle target, "Расчет НДС и несущей способности по модели прочности"
    WriteData target, batch, units, section
    FormatTable target, batch
    Exit Sub

Failed:
    Err.Raise Err.Number, "CStrengthSummaryWriter.WriteSummary", Err.Description
End Sub

' Очищает данные подробной таблицы прочности. При fullFormatting=True
' удаляет также шапку, объединения и форматы; иначе сохраняет готовое оформление.
Public Sub ClearSummary(ByVal workbook As Object, Optional ByVal fullFormatting As Boolean = True)
    mMinimalFormatting = Not fullFormatting
    If workbook Is Nothing Then Err.Raise vbObjectError + 3998, "CStrengthSummaryWriter", "Книга Excel не передана."
    PrepareTarget ClearTargetRange(workbook)
End Sub

' Возвращает диапазон текущей таблицы вместе с шапкой и строками сочетаний.
' Возвращает область, которую меняет публикация этого блока, включая его заголовок.
Friend Function PublicationRange(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, ByVal fullFormatting As Boolean) As Object
    mMinimalFormatting = Not fullFormatting
    Dim target As Object
    If snapshot.IsGeometryPreview Then
        Set target = ClearTargetRange(workbook)
    Else
        Set target = TargetRange(workbook, snapshot.TableRowCount("Strength.Data"))
    End If
    If target.Row > 1 Then Set target = target.Offset(-1, 0).Resize(target.Rows.Count + 1, target.Columns.Count)
    Set PublicationRange = target
End Function

' Возвращает область шапки и данных по фактическому якорю Results.
Private Function TargetRange(ByVal workbook As Object, Optional ByVal dataRows As Long = 1) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3999, "CStrengthSummaryWriter", _
        "Диапазон rngStrengthSummaryAnchor привязан к " & anchor.Worksheet.Name & "!" & anchor.Address(False, False) & _
        ". Для шапки таблицы нужны " & CStr(HEADER_ROWS) & " строк выше этой ячейки. Перенесите именованный диапазон ниже."
    Set TargetRange = anchor.Offset(-HEADER_ROWS, 0).Resize(RequiredRows(dataRows), TOTAL_COLUMNS)
End Function

' Возвращает диапазон очистки этой таблицы до начала следующего подробного раздела Results.
Private Function ClearTargetRange(ByVal workbook As Object) As Object
    Dim anchor As Object
    Set anchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    If anchor.Row <= HEADER_ROWS Then Err.Raise vbObjectError + 3999, "CStrengthSummaryWriter", _
        "Диапазон rngStrengthSummaryAnchor привязан к " & anchor.Worksheet.Name & "!" & anchor.Address(False, False) & _
        ". Для шапки таблицы нужны " & CStr(HEADER_ROWS) & " строк выше этой ячейки. Перенесите именованный диапазон ниже."

    Dim topRow As Long
    topRow = anchor.Row - HEADER_ROWS

    Dim rowsToClear As Long
    rowsToClear = RequiredRows(1)
    On Error GoTo UseDefaultHeight
    Dim crackWriter As CCrackSummaryWriter
    Set crackWriter = New CCrackSummaryWriter
    Dim crackAnchor As Object
    Set crackAnchor = workbook.Names.Item("rngCrackSummaryAnchor").RefersToRange
    Dim crackTopRow As Long
    crackTopRow = crackAnchor.Row - crackWriter.HeaderRowsAboveAnchor
    If crackTopRow > topRow Then
        rowsToClear = crackTopRow - topRow
        If mMinimalFormatting Then rowsToClear = rowsToClear - 1
    End If

UseDefaultHeight:
    If topRow + rowsToClear - 1 > anchor.Worksheet.Rows.Count Then rowsToClear = anchor.Worksheet.Rows.Count - topRow + 1
    Set ClearTargetRange = anchor.Worksheet.Cells(topRow, anchor.Column).Resize(rowsToClear, TOTAL_COLUMNS)
End Function

' Готовит область к записи: при полном оформлении очищает её целиком.
' В минимальном режиме оставляет шапку и форматы, удаляя данные и заливку рассчитанных статусов.
Private Sub PrepareTarget(ByVal target As Object)
    If mMinimalFormatting Then
        If target.Rows.Count > HEADER_ROWS Then
            With target.Offset(HEADER_ROWS, 0).Resize(target.Rows.Count - HEADER_ROWS, TOTAL_COLUMNS)
                .ClearContents
                .Interior.ColorIndex = -4142
            End With
        End If
        Exit Sub
    End If
    ClearBlockTitleRow target
    target.UnMerge
    target.Clear
End Sub

' Записывает отдельную строку-название над подробным блоком, не вмешиваясь в многоуровневую шапку таблицы.
Private Sub WriteBlockTitle(ByVal target As Object, ByVal titleText As String)
    If target.Row <= 1 Then Exit Sub
    If mMinimalFormatting Then
        target.Worksheet.Cells(target.Row - 1, target.Column).Value2 = titleText
        Exit Sub
    End If

    Dim titleRow As Long
    titleRow = target.Row - 1
    target.Worksheet.Cells(titleRow, target.Column).Resize(1, TOTAL_COLUMNS).Interior.Color = RGB(189, 215, 238)
    With target.Worksheet.Cells(titleRow, target.Column)
        .Value2 = titleText
        .Font.Name = "Arial"
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = RGB(255, 0, 0)
        .HorizontalAlignment = -4131
        .VerticalAlignment = -4108
        .WrapText = False
    End With
End Sub

' Очищает название и заливку только в ширине блока прочности. Ячейки справа
' принадлежат другим данным листа и не должны изменяться при повторной записи.
Private Sub ClearBlockTitleRow(ByVal target As Object)
    If target.Row <= 1 Then Exit Sub
    target.Worksheet.Cells(target.Row - 1, target.Column).Resize(1, TOTAL_COLUMNS).Clear
End Sub

' Собирает многоуровневую шапку с актуальными единицами. В минимальном
' режиме сохраняет совпадающую шапку; при отличии возвращается к полному оформлению.
Private Sub WriteHeaders(ByVal target As Object, ByVal units As CUnitSystem)
    Dim headers() As Variant
    ReDim headers(1 To HEADER_ROWS, 1 To TOTAL_COLUMNS)

    headers(1, COL_ID) = "Combination ID"
    headers(1, COL_RESULT_COMMENT) = "ResultComment"
    headers(1, COL_STATUS) = "общий статус"
    headers(1, COL_REF_N) = "относительно центральных осей приведенного сечения"
    headers(1, COL_LOAD_N) = "в точке приложения нагрузок"
    headers(1, COL_LIM_EBC) = "Общие данные"
    headers(1, COL_STATE_EPS0) = "НДС от заданного сочетания нагрузок (по модели прочности)"
    headers(1, COL_CAP_EPS0) = "Расчет предельной несущей способности по выбранной " & ChrW$(&H3BB) & "-траектории (по модели прочности)"

    headers(2, COL_STATE_EPS0) = "плоскость деформации"
    headers(2, COL_STATE_EBC) = "фактические макс. деформации"
    headers(2, COL_STATE_SBC) = "фактические напряжения"
    headers(2, COL_STATE_H) = "параметры сечения и статус"
    headers(2, COL_CAP_EPS0) = "плоскость деформации"
    headers(2, COL_CAP_LIMIT_STATE) = "параметры расчета"
    headers(2, COL_CAP_LAMBDA) = "предельные компоненты"
    headers(2, COL_CAP_EBC) = "фактические макс. деформации"
    headers(2, COL_CAP_H) = "параметры сечения при " & ChrW$(&H3BB) & "ult и статус"
    PutHeaderNotes headers
    PutColumnLabels headers, units
    PutCapacitySpatialHeaders headers, OutputMomentUnit(units), OutputForceUnit(units)
    PutTransverseHeaders headers, ChrW$(&H3C3) & ", " & OutputStressUnit(units)
    If Not mCapturedSnapshot Is Nothing Then
        mCapturedSnapshot.AddTable "Strength.Headers", headers
        Exit Sub
    End If
    If mMinimalFormatting Then
        If ResultHeaderMatches(target, headers, merges:=HeaderMergeBlocks()) Then Exit Sub
        mMinimalFormatting = False
        target.Resize(HEADER_ROWS, TOTAL_COLUMNS).UnMerge
    End If
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).Value2 = headers
    MergeHeaders target
End Sub

' Задаёт подписи и единицы пространственных усилий в точке приложения нагрузки.
Private Sub PutCapacitySpatialHeaders(ByRef headers() As Variant, ByVal momentUnit As String, ByVal forceUnit As String)
    headers(HEADER_NOTES_ROW, COL_CAP_T) = "полный T при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_QX) = "полная Qx при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_QY) = "полная Qy при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_LABEL_ROW, COL_CAP_T) = "Tult, " & momentUnit
    headers(HEADER_LABEL_ROW, COL_CAP_QX) = "Qx,ult, " & forceUnit
    headers(HEADER_LABEL_ROW, COL_CAP_QY) = "Qy,ult, " & forceUnit
End Sub

' Задаёт две группы столбцов хомута. Подпись напряжения содержит единицу
' исходного расчёта, в том числе при восстановлении снимка.
Private Sub PutTransverseHeaders(ByRef headers() As Variant, ByVal stressLabel As String)
    Dim epsilon As String: epsilon = ChrW$(&H3B5)
    headers(1, COL_TRANSVERSE_POINT) = "Хомуты и поперечные ветви"
    headers(2, COL_TRANSVERSE_POINT) = "При заданных усилиях"
    headers(2, COL_TRANSVERSE_RESERVE) = "На пределе поперечной арматуры"
    headers(HEADER_NOTES_ROW, COL_TRANSVERSE_POINT) = "Точка наибольшего использования"
    headers(HEADER_NOTES_ROW, COL_TRANSVERSE_UTIL) = "|" & epsilon & " / " & epsilon & "lim|"
    headers(HEADER_NOTES_ROW, COL_TRANSVERSE_LIMIT) = "Предельная деформация хомута или ветви"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_POINT) = "Группа, класс и точка"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_STRESS) = stressLabel
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_STRAIN) = epsilon
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_STRAIN_LIMIT) = epsilon & "lim"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_UTIL) = "Использование"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_STATUS) = "Статус НДС"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_RESERVE) = ChrW$(&H3BB) & " арматуры"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_LIMIT) = "Точка и причина предела"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_LIMIT_STRESS) = stressLabel
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_LIMIT_STRAIN) = epsilon
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_LIMIT_STRAIN_LIMIT) = epsilon & "lim"
    headers(HEADER_LABEL_ROW, COL_TRANSVERSE_LIMIT_STATUS) = "Статус предела"
End Sub

' Заполняет строку обычных пояснений над нижним ярусом обозначений.
Private Sub PutHeaderNotes(ByRef headers() As Variant)
    headers(HEADER_NOTES_ROW, COL_RESULT_COMMENT) = "пояснение расчетного статуса блока прочности"
    headers(HEADER_NOTES_ROW, COL_REF_N) = "продольная сила N, переданная в расчёт прочности, в центре приведённого сечения"
    headers(HEADER_NOTES_ROW, COL_REF_MX) = "момент Mx, переданный в расчёт прочности, относительно оси X через центр приведённого сечения"
    headers(HEADER_NOTES_ROW, COL_REF_MY) = "момент My, переданный в расчёт прочности, относительно оси Y через центр приведённого сечения"
    headers(HEADER_NOTES_ROW, COL_LOAD_N) = "продольная сила N, переданная в расчёт прочности, в пользовательской точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_LOAD_MX) = "момент Mx, переданный в расчёт прочности, в пользовательской точке приложения нагрузки. Продольный изгиб учтён, если он включён для сочетания."
    headers(HEADER_NOTES_ROW, COL_LOAD_MY) = "момент My, переданный в расчёт прочности, в пользовательской точке приложения нагрузки. Продольный изгиб учтён, если он включён для сочетания."
    headers(HEADER_NOTES_ROW, COL_LIM_EBC) = "предельная деформация бетона при сжатии"
    headers(HEADER_NOTES_ROW, COL_LIM_EBT) = "предельная деформация бетона при растяжении; пусто, если растянутый бетон не работает"
    headers(HEADER_NOTES_ROW, COL_LIM_ESC) = "предельная деформация продольной арматуры при сжатии"
    headers(HEADER_NOTES_ROW, COL_LIM_EST) = "предельная деформация продольной арматуры при растяжении"
    headers(HEADER_NOTES_ROW, COL_STATE_EPS0) = ChrW$(&H3B5) & "0 в точке (0;0) внутренней системы координат модели; от этой точки заданы координаты всех элементов сечения"
    headers(HEADER_NOTES_ROW, COL_STATE_KX) = "кривизна kx при координате y во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_STATE_KY) = "кривизна ky при координате x во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_STATE_EBC) = "максимальная сжимающая деформация бетона, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_EBT) = "максимальная растягивающая деформация бетона, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_ESC) = "максимальная сжимающая деформация продольной арматуры, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_EST) = "максимальная растягивающая деформация продольной арматуры, если такая зона есть"
    headers(HEADER_NOTES_ROW, COL_STATE_SBC) = "максимальное сжатие бетона"
    headers(HEADER_NOTES_ROW, COL_STATE_SBT) = "максимальное растяжение бетона; пусто, если диаграмма не несет растяжение"
    headers(HEADER_NOTES_ROW, COL_STATE_SSC) = "максимальное сжатие продольной арматуры"
    headers(HEADER_NOTES_ROW, COL_STATE_SST) = "максимальное растяжение продольной арматуры"
    headers(HEADER_NOTES_ROW, COL_STATE_H) = "полная высота сечения по нормали к нейтральной линии"
    headers(HEADER_NOTES_ROW, COL_STATE_X) = "высота сжатой зоны по нормали к нейтральной линии"
    headers(HEADER_NOTES_ROW, COL_STATE_DIAGEXT) = "yes - равновесие найдено через продолжение диаграммы; no - продолжение было разрешено, но не понадобилось; - - продолжение выключено"
    headers(HEADER_NOTES_ROW, COL_STATE_RESERVE) = "минимальный запас по предельным деформациям"
    headers(HEADER_NOTES_ROW, COL_STATE_STATUS) = "статус фактического НДС по модели прочности"
    headers(HEADER_NOTES_ROW, COL_CAP_EPS0) = ChrW$(&H3B5) & "0 в точке (0;0) внутренней системы координат модели; от этой точки заданы координаты всех элементов сечения"
    headers(HEADER_NOTES_ROW, COL_CAP_KX) = "кривизна kx при координате y во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_CAP_KY) = "кривизна ky при координате x во внутренней системе координат модели"
    headers(HEADER_NOTES_ROW, COL_CAP_LIMIT_STATE) = "причина найденного предельного состояния"
    headers(HEADER_NOTES_ROW, COL_CAP_LOAD_PATH) = "какие компоненты масштабировались"
    headers(HEADER_NOTES_ROW, COL_CAP_METHOD) = "фактический метод поиска"
    headers(HEADER_NOTES_ROW, COL_CAP_LAMBDA) = "предельный множитель выбранной траектории"
    headers(HEADER_NOTES_ROW, COL_CAP_N) = "полная N при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_MX) = "полный Mx при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_MY) = "полный My при " & ChrW$(&H3BB) & "ult в точке приложения нагрузки"
    headers(HEADER_NOTES_ROW, COL_CAP_EBC) = "сжатая деформация бетона в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_EBT) = "растянутая деформация бетона в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_ESC) = "сжатая деформация продольной арматуры в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_EST) = "растянутая деформация продольной арматуры в CapacityState, если зона есть"
    headers(HEADER_NOTES_ROW, COL_CAP_H) = "полная высота по нормали к нейтральной линии CapacityState"
    headers(HEADER_NOTES_ROW, COL_CAP_X) = "высота сжатой зоны CapacityState"
    headers(HEADER_NOTES_ROW, COL_CAP_RESERVE) = "запас по найденному множителю " & ChrW$(&H3BB)
    headers(HEADER_NOTES_ROW, COL_CAP_STATUS) = "статус расчета несущей способности"
End Sub

' Заполняет нижний ярус обозначений с единицами пользователя.
Private Sub PutColumnLabels(ByRef headers() As Variant, ByVal units As CUnitSystem)
    headers(HEADER_LABEL_ROW, COL_REF_N) = "N, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_REF_MX) = "Mx, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_REF_MY) = "My, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_N) = "N, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_MX) = "Mx, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_LOAD_MY) = "My, " & OutputMomentUnit(units)

    headers(HEADER_LABEL_ROW, COL_LIM_EBC) = ChrW$(&H3B5) & "bc,ult"
    headers(HEADER_LABEL_ROW, COL_LIM_EBT) = ChrW$(&H3B5) & "bt,ult"
    headers(HEADER_LABEL_ROW, COL_LIM_ESC) = ChrW$(&H3B5) & "sc,ult"
    headers(HEADER_LABEL_ROW, COL_LIM_EST) = ChrW$(&H3B5) & "st,ult"
    headers(HEADER_LABEL_ROW, COL_STATE_EPS0) = ChrW$(&H3B5) & "0"
    headers(HEADER_LABEL_ROW, COL_STATE_KX) = ChrW$(&H3BA) & "x, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_KY) = ChrW$(&H3BA) & "y, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_STATE_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_STATE_ESC) = ChrW$(&H3B5) & "sc"
    headers(HEADER_LABEL_ROW, COL_STATE_EST) = ChrW$(&H3B5) & "st"
    headers(HEADER_LABEL_ROW, COL_STATE_SBC) = ChrW$(&H3C3) & "bc, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_SBT) = ChrW$(&H3C3) & "bt, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_SSC) = ChrW$(&H3C3) & "sc, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_SST) = ChrW$(&H3C3) & "st, " & OutputStressUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_H) = "h, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_X) = "x, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_STATE_DIAGEXT) = "ExtUsed"
    headers(HEADER_LABEL_ROW, COL_STATE_RESERVE) = "запас (" & ChrW$(&H3B5) & ")"
    headers(HEADER_LABEL_ROW, COL_STATE_STATUS) = "статус"

    headers(HEADER_LABEL_ROW, COL_CAP_EPS0) = ChrW$(&H3B5) & "0"
    headers(HEADER_LABEL_ROW, COL_CAP_KX) = ChrW$(&H3BA) & "x, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_KY) = ChrW$(&H3BA) & "y, " & OutputCurvatureUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_LIMIT_STATE) = "LimitState"
    headers(HEADER_LABEL_ROW, COL_CAP_LOAD_PATH) = "LoadPath"
    headers(HEADER_LABEL_ROW, COL_CAP_METHOD) = "Method"
    headers(HEADER_LABEL_ROW, COL_CAP_LAMBDA) = ChrW$(&H3BB) & "ult"
    headers(HEADER_LABEL_ROW, COL_CAP_N) = "Nult, " & OutputForceUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_MX) = "Mx,ult, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_MY) = "My,ult, " & OutputMomentUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_EBC) = ChrW$(&H3B5) & "bc"
    headers(HEADER_LABEL_ROW, COL_CAP_EBT) = ChrW$(&H3B5) & "bt"
    headers(HEADER_LABEL_ROW, COL_CAP_ESC) = ChrW$(&H3B5) & "sc"
    headers(HEADER_LABEL_ROW, COL_CAP_EST) = ChrW$(&H3B5) & "st"
    headers(HEADER_LABEL_ROW, COL_CAP_H) = "h, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_X) = "x, " & OutputLengthUnit(units)
    headers(HEADER_LABEL_ROW, COL_CAP_RESERVE) = "запас (" & ChrW$(&H3BB) & ")"
    headers(HEADER_LABEL_ROW, COL_CAP_STATUS) = "статус"
End Sub

' Объединяет прямоугольные участки шапки.
Private Sub MergeHeaders(ByVal target As Object)
    Dim block As Variant
    For Each block In HeaderMergeBlocks()
        MergeRange target, CLng(block(0)), CLng(block(2)), CLng(block(1)), CLng(block(3))
    Next block
End Sub

' Использует одно описание объединений для проверки и построения шапки.
Private Function HeaderMergeBlocks() As Collection
    Dim blocks As Collection: Set blocks = New Collection
    blocks.Add Array(1, HEADER_LABEL_ROW, COL_ID, COL_ID)
    blocks.Add Array(1, HEADER_LABEL_ROW, COL_RESULT_COMMENT, COL_RESULT_COMMENT)
    blocks.Add Array(1, HEADER_LABEL_ROW, COL_STATUS, COL_STATUS)
    blocks.Add Array(1, 2, COL_REF_N, COL_REF_MY)
    blocks.Add Array(1, 2, COL_LOAD_N, COL_LOAD_MY)
    blocks.Add Array(1, 2, COL_LIM_EBC, COL_LIM_EST)
    blocks.Add Array(1, 1, COL_STATE_EPS0, COL_STATE_STATUS)
    blocks.Add Array(2, 2, COL_STATE_EPS0, COL_STATE_KY)
    blocks.Add Array(2, 2, COL_STATE_EBC, COL_STATE_EST)
    blocks.Add Array(2, 2, COL_STATE_SBC, COL_STATE_SST)
    blocks.Add Array(2, 2, COL_STATE_H, COL_STATE_STATUS)
    blocks.Add Array(1, 1, COL_CAP_EPS0, COL_CAP_STATUS)
    blocks.Add Array(2, 2, COL_CAP_EPS0, COL_CAP_KY)
    blocks.Add Array(2, 2, COL_CAP_LIMIT_STATE, COL_CAP_METHOD)
    blocks.Add Array(2, 2, COL_CAP_LAMBDA, COL_CAP_QY)
    blocks.Add Array(2, 2, COL_CAP_EBC, COL_CAP_EST)
    blocks.Add Array(2, 2, COL_CAP_H, COL_CAP_STATUS)
    blocks.Add Array(1, 1, COL_TRANSVERSE_POINT, COL_TRANSVERSE_LIMIT_STATUS)
    blocks.Add Array(2, 2, COL_TRANSVERSE_POINT, COL_TRANSVERSE_STATUS)
    blocks.Add Array(2, 2, COL_TRANSVERSE_RESERVE, COL_TRANSVERSE_LIMIT_STATUS)
    Set HeaderMergeBlocks = blocks
End Function

' Записывает строки сочетаний одним массивом.
Private Sub WriteData(ByVal target As Object, ByVal batch As CBatchSectionCalculator, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel)
    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows <= 0 Then Exit Sub

    Dim data() As Variant
    ReDim data(1 To dataRows, 1 To TOTAL_COLUMNS)

    Dim i As Long, workbook As Object: Set workbook = target.Worksheet.Parent
    For i = 1 To batch.Count
        PutCombination data, OutputDataOffset(batch, i, dataRows), batch, i, units, section, workbook
    Next i
    If Not mCapturedSnapshot Is Nothing Then
        mCapturedSnapshot.AddTable "Strength.Data", data
        Exit Sub
    End If
    target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).Value2 = data
End Sub

' Заполняет одну строку подробного блока прочности.
Private Sub PutCombination(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel, ByVal workbook As Object)
    Dim policy As CResultStatusPolicy
    Set policy = StatusPolicy()

    data(rowIndex, COL_ID) = batch.CombinationID(index)
    data(rowIndex, COL_RESULT_COMMENT) = CalculationResultMessage(batch, index, batch.ResultAt(index).StrengthMeta, workbook)
    data(rowIndex, COL_STATUS) = policy.ExternalStatus(batch.ResultAt(index).StrengthMeta)
    data(rowIndex, COL_STATE_STATUS) = policy.NotApplicable
    data(rowIndex, COL_CAP_STATUS) = policy.NotApplicable

    If Not batch.StrengthDirectStateRequested(index) And Not batch.StrengthCapacityRequested(index) Then Exit Sub

    data(rowIndex, COL_REF_N) = OutputForce(batch.StrengthReferenceN(index), units)
    data(rowIndex, COL_REF_MX) = OutputMx(batch.StrengthReferenceMx(index), units)
    data(rowIndex, COL_REF_MY) = OutputMy(batch.StrengthReferenceMy(index), units)
    data(rowIndex, COL_LOAD_N) = OutputForce(batch.StrengthLoadPointN(index), units)
    data(rowIndex, COL_LOAD_MX) = OutputMx(batch.StrengthLoadPointMx(index), units)
    data(rowIndex, COL_LOAD_MY) = OutputMy(batch.StrengthLoadPointMy(index), units)

    PutStrengthLimits data, rowIndex, batch, index, units
    PutStrengthState data, rowIndex, batch, index, units, section, policy
    PutCapacityState data, rowIndex, batch, index, units, section, policy
    If batch.HasSpatialLoads(index) Then PutTransverseState data, rowIndex, batch, index, units, policy
End Sub

' Отдельно показывает максимальное использование хомута при заданных усилиях
' и множитель достижения предельной деформации поперечной арматуры.
' Если подготовка не завершилась, показывает ошибку запрошенного этапа.
Private Sub PutTransverseState(ByRef data() As Variant, ByVal row As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem, _
        ByVal policy As CResultStatusPolicy)
    data(row, COL_TRANSVERSE_STATUS) = policy.NotApplicable
    data(row, COL_TRANSVERSE_LIMIT_STATUS) = policy.NotApplicable
    Dim layout As CTransverseRebarLayout: Set layout = batch.CoupledTransverseLayout
    If Not layout Is Nothing Then
        If layout.Count = 0 Then Exit Sub ' Подтверждено, что поперечные стержни отсутствуют.
    End If

    Dim result As CCombinationResult: Set result = batch.ResultAt(index)
    data(row, COL_TRANSVERSE_STATUS) = TransverseUnavailableStatus(result.DirectStateMeta, policy)
    data(row, COL_TRANSVERSE_LIMIT_STATUS) = TransverseUnavailableStatus(result.CapacityMeta, policy)
    If layout Is Nothing Then Exit Sub ' Ошибка подготовки помешала получить проверенную раскладку.
    Dim state As CSectionStateResult: Set state = result.StateRepository.FindState(sstStrengthState)
    Dim description As String, utilization As Double, stress As Double, strain As Double, strainLimit As Double
    If Not state Is Nothing Then
        If state.CopyCoupledTransverseControl(description, utilization, stress, strain, strainLimit) Then
            data(row, COL_TRANSVERSE_POINT) = description
            data(row, COL_TRANSVERSE_STRESS) = OutputStress(stress, units)
            data(row, COL_TRANSVERSE_STRAIN) = strain
            data(row, COL_TRANSVERSE_STRAIN_LIMIT) = strainLimit
            data(row, COL_TRANSVERSE_UTIL) = utilization
            data(row, COL_TRANSVERSE_STATUS) = state.CoupledTransverseStatus
        End If
    End If
    Dim capacity As CCapacityResult: Set capacity = result.StrengthResult.Capacity
    If Not capacity.IsTransverseLimit Then Exit Sub
    data(row, COL_TRANSVERSE_LIMIT) = capacity.LimitState
    data(row, COL_TRANSVERSE_LIMIT_STATUS) = capacity.TransverseLimitStatus
    If Not capacity.HasSteelLimitPoint Then Exit Sub ' При отказе на постоянной части множителя переменной нагрузки нет.
    data(row, COL_TRANSVERSE_RESERVE) = capacity.CoupledLimitLambda
    Set state = capacity.SteelLimitStateResult
    If state Is Nothing Then Exit Sub
    If state.CopyCoupledTransverseControl(description, utilization, stress, strain, strainLimit) Then
        data(row, COL_TRANSVERSE_LIMIT_STRESS) = OutputStress(stress, units)
        data(row, COL_TRANSVERSE_LIMIT_STRAIN) = strain
        data(row, COL_TRANSVERSE_LIMIT_STRAIN_LIMIT) = strainLimit
    End If
End Sub

' Ошибки общего этапа мешают проверить хомут. Успех или физический отказ
' по другому механизму сам по себе не задаёт отдельный результат для хомута.
Private Function TransverseUnavailableStatus(ByVal meta As CResultMeta, ByVal policy As CResultStatusPolicy) As String
    Select Case meta.InternalStatus
        Case rsNumericalFailure, rsInvalidInput, rsInvalidConfiguration, rsInternalError
            TransverseUnavailableStatus = policy.ExternalStatus(meta)
        Case Else
            TransverseUnavailableStatus = policy.NotApplicable
    End Select
End Function

' Возвращает позицию LC внутри data-части подробной таблицы.
' Основной поток читает offset из rngLoadCombinations; прямые тесты без offset продолжают выводиться компактно.
Private Function OutputDataOffset(ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal dataRows As Long) As Long
    OutputDataOffset = batch.SourceDataOffset(index)
    If OutputDataOffset < 1 Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = index
    If OutputDataOffset > dataRows Then OutputDataOffset = dataRows
End Function

' Записывает физические пределы деформаций фактически использованной модели прочности.
' Растянутый предел бетона показывает только при включённой растянутой ветви.
Private Sub PutStrengthLimits(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, ByVal units As CUnitSystem)
    data(rowIndex, COL_LIM_EBC) = batch.StrengthConcreteCompressionLimitForCombination(index)
    If batch.StrengthConcreteTensionLimitEnabledForCombination(index) Then _
        data(rowIndex, COL_LIM_EBT) = batch.StrengthConcreteTensionLimitForCombination(index)
    data(rowIndex, COL_LIM_ESC) = batch.StrengthSteelCompressionLimitForCombination(index)
    data(rowIndex, COL_LIM_EST) = batch.StrengthSteelTensionLimitForCombination(index)
End Sub

' Пишет фактический StrengthState, если он действительно был рассчитан.
Private Sub PutStrengthState(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel, ByVal policy As CResultStatusPolicy)
    If Not batch.StrengthDirectStateRequested(index) Then
        data(rowIndex, COL_STATE_STATUS) = policy.NotApplicable
        Exit Sub
    End If

    Dim stateResult As CSectionStateResult
    Set stateResult = batch.ResultAt(index).StateRepository.FindState(sstStrengthState)
    If stateResult Is Nothing Then
        data(rowIndex, COL_STATE_STATUS) = DirectStatusText(batch, index, policy)
        Exit Sub
    End If

    PutStatePlane data, rowIndex, stateResult, units, COL_STATE_EPS0, COL_STATE_KX, COL_STATE_KY
    PutStateExtremes data, rowIndex, stateResult, units, _
        COL_STATE_EBC, COL_STATE_EBT, COL_STATE_ESC, COL_STATE_EST, _
        COL_STATE_SBC, COL_STATE_SBT, COL_STATE_SSC, COL_STATE_SST
    PutStateGeometry data, rowIndex, stateResult, units, section, COL_STATE_H, COL_STATE_X
    data(rowIndex, COL_STATE_DIAGEXT) = ExtensionUsedText(batch, index, stateResult)
    If batch.ResultAt(index).StrengthResult.DirectState.StrainReserve > SMALL_VALUE Then _
        data(rowIndex, COL_STATE_RESERVE) = batch.ResultAt(index).StrengthResult.DirectState.StrainReserve
    data(rowIndex, COL_STATE_STATUS) = policy.ExternalStatus(stateResult.ResultMeta)
End Sub

' Показывает использование продолжения диаграммы только тогда, когда само
' продолжение разрешено настройкой General.DiagramExtension.
' При выключенной настройке выводится "-", потому что extension не мог
' применяться и обычная пара yes/no здесь не несет полезного смысла.
Private Function ExtensionUsedText(ByVal batch As CBatchSectionCalculator, _
        ByVal index As Long, ByVal stateResult As CSectionStateResult) As String
    If Not batch.ResultAt(index).DiagramExtensionEnabled Then
        ExtensionUsedText = "-"
    ElseIf stateResult.ExtensionUsed Then
        ExtensionUsedText = "yes"
    Else
        ExtensionUsedText = "no"
    End If
End Function

' Записывает путь нагрузки, метод, статус и подтверждённое предельное НДС.
' Если конечного состояния нет, последние пробные усилия или деформации
' вместо него не подставляются.
Private Sub PutCapacityState(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal units As CUnitSystem, ByVal section As CSectionModel, ByVal policy As CResultStatusPolicy)
    If Not batch.StrengthCapacityRequested(index) Then
        data(rowIndex, COL_CAP_STATUS) = policy.NotApplicable
        Exit Sub
    End If

    data(rowIndex, COL_CAP_LIMIT_STATE) = batch.ResultAt(index).StrengthResult.Capacity.LimitState
    data(rowIndex, COL_CAP_LOAD_PATH) = batch.ResultAt(index).CapacityPathDisplayName
    data(rowIndex, COL_CAP_METHOD) = batch.ResultAt(index).StrengthResult.Capacity.SolutionMethod
    If batch.ResultAt(index).StrengthResult.Capacity.LambdaCapacity > SMALL_VALUE Then data(rowIndex, COL_CAP_LAMBDA) = batch.ResultAt(index).StrengthResult.Capacity.LambdaCapacity

    Dim stateResult As CSectionStateResult
    Dim capacity As CCapacityResult: Set capacity = batch.ResultAt(index).StrengthResult.Capacity
    Set stateResult = batch.ResultAt(index).StateRepository.FindState(sstCapacityState)
    If capacity.HasUltimatePoint And (Not stateResult Is Nothing Or capacity.ReserveFactor > 0#) Then _
        PutCapacityLoadPointComponents data, rowIndex, capacity, units
    If Not stateResult Is Nothing And Not (capacity.HasSteelLimitPoint And Not capacity.HasUltimatePoint) Then
        PutStatePlane data, rowIndex, stateResult, units, COL_CAP_EPS0, COL_CAP_KX, COL_CAP_KY
        PutStateExtremes data, rowIndex, stateResult, units, _
            COL_CAP_EBC, COL_CAP_EBT, COL_CAP_ESC, COL_CAP_EST, _
            0, 0, 0, 0
        PutStateGeometry data, rowIndex, stateResult, units, section, COL_CAP_H, COL_CAP_X
    End If
    If batch.ResultAt(index).StrengthResult.Capacity.ReserveFactor > SMALL_VALUE Then _
        data(rowIndex, COL_CAP_RESERVE) = batch.ResultAt(index).StrengthResult.Capacity.ReserveFactor
    data(rowIndex, COL_CAP_STATUS) = CapacityStatusText(batch, index, policy)
End Sub

' Выводит шесть подтверждённых предельных усилий в пользовательской точке приложения.
Private Sub PutCapacityLoadPointComponents(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal capacity As CCapacityResult, ByVal units As CUnitSystem)
    data(rowIndex, COL_CAP_N) = OutputForce(capacity.NUltimate, units)
    data(rowIndex, COL_CAP_MX) = OutputMx(capacity.MxUltimate, units)
    data(rowIndex, COL_CAP_MY) = OutputMy(capacity.MyUltimate, units)
    If units Is Nothing Then
        data(rowIndex, COL_CAP_T) = capacity.TUltimate
        data(rowIndex, COL_CAP_QX) = capacity.QxUltimate
        data(rowIndex, COL_CAP_QY) = capacity.QyUltimate
    Else
        data(rowIndex, COL_CAP_T) = units.InternalTorqueToOutput(capacity.TUltimate)
        data(rowIndex, COL_CAP_QX) = units.InternalShearToOutput(capacity.QxUltimate)
        data(rowIndex, COL_CAP_QY) = units.InternalShearToOutput(capacity.QyUltimate)
    End If
End Sub

' Записывает сохранённое поле деформаций: epsilon0 в начале координат модели
' и кривизны в выбранных выходных единицах.
Private Sub PutStatePlane(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal units As CUnitSystem, _
        ByVal colEps0 As Long, ByVal colKx As Long, ByVal colKy As Long)
    data(rowIndex, colEps0) = stateResult.Epsilon0
    data(rowIndex, colKx) = OutputCurvature(stateResult.KappaX, units)
    data(rowIndex, colKy) = OutputCurvature(stateResult.KappaY, units)
End Sub

' Пишет экстремумы состояния по флагам наличия физической зоны.
Private Sub PutStateExtremes(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal units As CUnitSystem, _
        ByVal colEbc As Long, ByVal colEbt As Long, ByVal colEsc As Long, ByVal colEst As Long, _
        ByVal colSbc As Long, ByVal colSbt As Long, ByVal colSsc As Long, ByVal colSst As Long)
    If stateResult.HasConcreteCompression Then data(rowIndex, colEbc) = stateResult.MinConcreteStrain
    If stateResult.HasConcreteTension Then data(rowIndex, colEbt) = stateResult.MaxConcreteStrain
    If stateResult.HasSteelCompression Then data(rowIndex, colEsc) = stateResult.MinSteelStrain
    If stateResult.HasSteelTension Then data(rowIndex, colEst) = stateResult.MaxSteelStrain

    If colSbc > 0 And stateResult.HasConcreteCompression Then data(rowIndex, colSbc) = OutputStress(stateResult.MinConcreteStress, units)
    If colSbt > 0 And stateResult.HasConcreteTension Then data(rowIndex, colSbt) = OutputStress(stateResult.MaxConcreteStress, units)
    If colSsc > 0 And stateResult.HasSteelCompression Then data(rowIndex, colSsc) = OutputStress(stateResult.MinSteelStress, units)
    If colSst > 0 And stateResult.HasSteelTension Then data(rowIndex, colSst) = OutputStress(stateResult.MaxSteelStress, units)
End Sub

' Записывает высоту сечения h и высоту сжатой зоны x для сохранённого поля деформаций.
Private Sub PutStateGeometry(ByRef data() As Variant, ByVal rowIndex As Long, _
        ByVal stateResult As CSectionStateResult, ByVal units As CUnitSystem, _
        ByVal section As CSectionModel, ByVal colH As Long, ByVal colX As Long)
    Dim depthH As Double, compressedDepth As Double, hasH As Boolean, hasX As Boolean
    stateResult.GetGeometryDepths section, depthH, compressedDepth, hasH, hasX
    If hasH Then data(rowIndex, colH) = OutputLength(depthH, units)
    If hasX Then data(rowIndex, colX) = OutputLength(compressedDepth, units)
End Sub

' Оформляет подробную таблицу прочности: шапку, рамки, числа и статусы.
' В минимальном режиме обновляет только подсветку статусов.
Private Sub FormatTable(ByVal target As Object, ByVal batch As CBatchSectionCalculator, _
        Optional ByRef preparedValues As Variant)
    If Not mCapturedSnapshot Is Nothing Then Exit Sub
    If mMinimalFormatting Then
        ApplyStatusCells target, batch, target.Rows.Count - HEADER_ROWS, preparedValues
        Exit Sub
    End If
    target.Font.Name = "Arial"
    target.Font.Size = 9
    target.HorizontalAlignment = -4108
    target.VerticalAlignment = -4108
    target.WrapText = False

    FormatBlock target, COL_LEFT_FIRST, COL_LEFT_LAST
    FormatBlock target, COL_STATE_FIRST, COL_STATE_LAST
    FormatBlock target, COL_CAPACITY_FIRST, COL_CAPACITY_LAST
    FormatBlock target, COL_TRANSVERSE_POINT, COL_TRANSVERSE_LIMIT_STATUS
    target.Columns(COL_SEPARATOR1).Clear
    target.Columns(COL_SEPARATOR2).Clear
    target.Columns(COL_SEPARATOR3).Clear

    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).HorizontalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).VerticalAlignment = -4108
    target.Resize(HEADER_ROWS, TOTAL_COLUMNS).WrapText = True
    target.Rows(HEADER_NOTES_ROW).Font.Bold = False
    target.Rows(HEADER_NOTES_ROW).HorizontalAlignment = -4131
    target.Rows(HEADER_NOTES_ROW).WrapText = False
    target.Rows(HEADER_LABEL_ROW).WrapText = False

    Dim dataRows As Long
    dataRows = target.Rows.Count - HEADER_ROWS
    If dataRows > 0 Then
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).HorizontalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).VerticalAlignment = -4108
        target.Offset(HEADER_ROWS, 0).Resize(dataRows, TOTAL_COLUMNS).NumberFormat = "0.000"
        target.Offset(HEADER_ROWS, COL_ID - 1).Resize(dataRows, 3).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_RESULT_COMMENT - 1).Resize(dataRows, 1).HorizontalAlignment = -4131
        target.Offset(HEADER_ROWS, COL_STATE_DIAGEXT - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_STATE_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_CAP_LIMIT_STATE - 1).Resize(dataRows, 3).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_CAP_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_TRANSVERSE_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        target.Offset(HEADER_ROWS, COL_TRANSVERSE_LIMIT_STATUS - 1).Resize(dataRows, 1).NumberFormat = "@"
        With target.Offset(HEADER_ROWS, COL_TRANSVERSE_POINT - 1).Resize(dataRows, 1)
            .HorizontalAlignment = -4131
            .ShrinkToFit = True ' Длинное имя группы и координаты остаются в одной строке.
        End With
        With target.Offset(HEADER_ROWS, COL_TRANSVERSE_LIMIT - 1).Resize(dataRows, 1)
            .HorizontalAlignment = -4131
            .ShrinkToFit = True
        End With
        ApplyStrainNumberFormat target, dataRows
        ApplyStatusCells target, batch, dataRows, preparedValues
    End If
    ApplyHeaderSubscripts target
End Sub

' Форматирует относительные деформации и кривизны плоскости НДС с нужной точностью вывода.
Private Sub ApplyStrainNumberFormat(ByVal target As Object, ByVal dataRows As Long)
    target.Offset(HEADER_ROWS, COL_LIM_EBC - 1).Resize(dataRows, COL_LIM_EST - COL_LIM_EBC + 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_STATE_EPS0 - 1).Resize(dataRows, 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_STATE_KX - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
    target.Offset(HEADER_ROWS, COL_STATE_EBC - 1).Resize(dataRows, COL_STATE_EST - COL_STATE_EBC + 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_CAP_EPS0 - 1).Resize(dataRows, 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_CAP_KX - 1).Resize(dataRows, 2).NumberFormat = "0.000000"
    target.Offset(HEADER_ROWS, COL_CAP_EBC - 1).Resize(dataRows, COL_CAP_EST - COL_CAP_EBC + 1).NumberFormat = "0.0000"
    target.Offset(HEADER_ROWS, COL_TRANSVERSE_STRAIN - 1).Resize(dataRows, 2).NumberFormat = "0.000E+00"
    target.Offset(HEADER_ROWS, COL_TRANSVERSE_LIMIT_STRAIN - 1).Resize(dataRows, 2).NumberFormat = "0.000E+00"
    target.Offset(HEADER_ROWS, COL_TRANSVERSE_RESERVE - 1).Resize(dataRows, 1).NumberFormat = "0.000000"
End Sub

' Оформляет непрерывный блок таблицы, не затрагивая пустые разделители.
Private Sub FormatBlock(ByVal target As Object, ByVal firstCol As Long, ByVal lastCol As Long)
    Dim block As Object
    Set block = target.Offset(0, firstCol - 1).Resize(target.Rows.Count, lastCol - firstCol + 1)
    block.Borders.LineStyle = 1
    block.Borders.Weight = 2
    block.Rows(1).Resize(HEADER_ROWS).Font.Bold = True
    target.Offset(0, firstCol - 1).Resize(HEADER_ROWS, lastCol - firstCol + 1).Interior.Color = HEADER_COLOR
End Sub

' Красит строковые блоки подробной таблицы по фактическому статусу каждой проверки.
Private Sub ApplyStatusCells(ByVal target As Object, ByVal batch As CBatchSectionCalculator, ByVal dataRows As Long, _
        Optional ByRef preparedValues As Variant)
    Dim rows() As Long, i As Long
    If batch Is Nothing Then
        Dim storedRows As Variant: storedRows = ResultSnapshotStatusRows(target, HEADER_ROWS, preparedValues)
        If IsEmpty(storedRows) Then Exit Sub
        rows = storedRows
    Else
    If batch.Count <= 0 Then Exit Sub
    ReDim rows(1 To batch.Count)
    For i = 1 To batch.Count
        rows(i) = HEADER_ROWS + OutputDataOffset(batch, i, dataRows)
    Next i
    End If
    ApplyTableStatusFills target, rows, Array( _
        COL_STATUS, COL_STATE_STATUS, COL_CAP_STATUS, COL_TRANSVERSE_STATUS, COL_TRANSVERSE_LIMIT_STATUS), _
        False, 0, preparedValues, HEADER_ROWS
End Sub


' Проставляет нижние индексы в обозначениях нижнего яруса шапки.
Private Sub ApplyHeaderSubscripts(ByVal target As Object)
    Dim labels As Object
    Set labels = target.Rows(HEADER_LABEL_ROW)
    Dim headers As Variant: headers = labels.Value2

    Dim cols As Variant
    cols = Array(COL_LIM_EBC, COL_LIM_EBT, COL_LIM_ESC, COL_LIM_EST, _
        COL_STATE_EPS0, COL_STATE_KX, COL_STATE_KY, _
        COL_STATE_EBC, COL_STATE_EBT, COL_STATE_ESC, COL_STATE_EST, _
        COL_STATE_SBC, COL_STATE_SBT, COL_STATE_SSC, COL_STATE_SST, _
        COL_CAP_LAMBDA, COL_CAP_N, COL_CAP_MX, COL_CAP_MY, COL_CAP_T, COL_CAP_QX, COL_CAP_QY, _
        COL_CAP_EPS0, COL_CAP_KX, COL_CAP_KY, _
        COL_CAP_EBC, COL_CAP_EBT, COL_CAP_ESC, COL_CAP_EST)

    Dim i As Long
    For i = LBound(cols) To UBound(cols)
        ApplySubscriptToken labels, headers, CLng(cols(i)), "bc"
        ApplySubscriptToken labels, headers, CLng(cols(i)), "bt"
        ApplySubscriptToken labels, headers, CLng(cols(i)), "sc"
        ApplySubscriptToken labels, headers, CLng(cols(i)), "st"
        ApplySubscriptToken labels, headers, CLng(cols(i)), "ult"
    Next i
    ApplySubscriptToken labels, headers, COL_CAP_LAMBDA, "ult"
    ApplySubscriptToken labels, headers, COL_STATE_EPS0, "0"
    ApplySubscriptToken labels, headers, COL_STATE_KX, "x"
    ApplySubscriptToken labels, headers, COL_STATE_KY, "y"
    ApplySubscriptToken labels, headers, COL_CAP_EPS0, "0"
    ApplySubscriptToken labels, headers, COL_CAP_KX, "x"
    ApplySubscriptToken labels, headers, COL_CAP_KY, "y"
    ApplySubscriptToken labels, headers, COL_TRANSVERSE_STRAIN_LIMIT, "lim"
    ApplySubscriptToken labels, headers, COL_TRANSVERSE_LIMIT_STRAIN_LIMIT, "lim"
End Sub

' Ставит нижний индекс найденному текстовому токену внутри одной ячейки.
Private Sub ApplySubscriptToken(ByVal labels As Object, ByRef headers As Variant, ByVal columnIndex As Long, ByVal token As String)
    Dim textValue As String
    textValue = CStr(headers(1, columnIndex))
    Dim startPos As Long
    startPos = InStr(1, textValue, token, vbTextCompare)
    If startPos = 0 Then Exit Sub
    Dim cell As Object: Set cell = labels.Cells(1, columnIndex)
    Do While startPos > 0
        cell.Characters(startPos, Len(token)).Font.Subscript = True
        startPos = InStr(startPos + Len(token), textValue, token, vbTextCompare)
    Loop
End Sub

' Возвращает внешний статус НДС при заданной нагрузке по сохранённому результату и общим
' правилам.
Private Function DirectStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CResultStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.ResultAt(index).DirectStateMeta
    DirectStatusText = policy.ExternalStatus(meta)
End Function

' Возвращает внешний статус несущей способности по ResultMeta.
Private Function CapacityStatusText(ByVal batch As CBatchSectionCalculator, ByVal index As Long, _
        ByVal policy As CResultStatusPolicy) As String
    Dim meta As CResultMeta
    Set meta = batch.ResultAt(index).CapacityMeta
    CapacityStatusText = policy.ExternalStatus(meta)
End Function


' Объединяет прямоугольный участок шапки.
Private Sub MergeRange(ByVal target As Object, ByVal firstRow As Long, ByVal firstCol As Long, _
        ByVal lastRow As Long, ByVal lastCol As Long)
    If lastRow < firstRow Or lastCol < firstCol Then Exit Sub
    target.Worksheet.Range(target.Cells(firstRow, firstCol), target.Cells(lastRow, lastCol)).Merge
End Sub



' Переводит длину из внутренних мм в выходную единицу пользователя.
' Если CUnitSystem не передан, возвращает исходное значение в мм.
Private Function OutputLength(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputLength = value Else OutputLength = units.InternalLengthToOutput(value)
End Function

' Возвращает подпись выходной единицы длины для заголовков таблицы прочности.
Private Function OutputLengthUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputLengthUnit = "mm" Else OutputLengthUnit = units.OutputLengthUnit
End Function

' Переводит напряжение из внутренних МПа в выбранную пользователем единицу.
Private Function OutputStress(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputStress = value Else OutputStress = units.InternalStressToOutput(value)
End Function

' Возвращает ту же единицу напряжений, в которую переводятся значения строк НДС.
Private Function OutputStressUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputStressUnit = "MPa" Else OutputStressUnit = units.OutputStressUnit
End Function

' Переводит кривизну плоскости деформаций из внутренних 1/мм
' в пользовательскую единицу вывода.
Private Function OutputCurvature(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputCurvature = value Else OutputCurvature = units.InternalCurvatureToOutput(value)
End Function

' Возвращает подпись выходной единицы кривизны для kx/ky.
Private Function OutputCurvatureUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputCurvatureUnit = "1/mm" Else OutputCurvatureUnit = units.OutputCurvatureUnit
End Function

' Переводит продольную силу из внутренних Н в пользовательскую систему вывода.
Private Function OutputForce(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputForce = value Else OutputForce = units.InternalForceToOutput(value)
End Function

' Возвращает подпись выходной единицы силы для заголовков таблицы прочности.
Private Function OutputForceUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputForceUnit = "N" Else OutputForceUnit = units.OutputForceUnit
End Function

' Переводит Mx с учетом пользовательской системы знаков.
Private Function OutputMx(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMx = value Else OutputMx = units.InternalMomentMxToOutput(value)
End Function

' Переводит My с учетом пользовательской системы знаков.
Private Function OutputMy(ByVal value As Double, ByVal units As CUnitSystem) As Double
    If units Is Nothing Then OutputMy = value Else OutputMy = units.InternalMomentMyToOutput(value)
End Function

' Возвращает подпись выходной единицы момента для заголовков таблицы прочности.
Private Function OutputMomentUnit(ByVal units As CUnitSystem) As String
    If units Is Nothing Then OutputMomentUnit = "N*mm" Else OutputMomentUnit = units.OutputMomentUnit
End Function

' Возвращает единый словарь пользовательских статусов для подробного блока прочности.
Private Function StatusPolicy() As CResultStatusPolicy
    Static policy As CResultStatusPolicy
    If policy Is Nothing Then Set policy = New CResultStatusPolicy
    Set StatusPolicy = policy
End Function





' Передаёт отдельный приёмник готовых массивов; при Nothing сохраняется обычный вывод.
Friend Property Set CaptureSnapshot(ByVal value As CCalculationSnapshot)
    Set mCapturedSnapshot = value
End Property

' Восстанавливает готовую подробную таблицу и её сохранённое оформление без решения НДС.
Public Sub RestoreSummary(ByVal workbook As Object, ByVal snapshot As CCalculationSnapshot, Optional ByVal fullFormatting As Boolean = True)
    snapshot.Validate
    RestoreSummaryTables workbook, snapshot, fullFormatting
End Sub

' Восстанавливает только готовые общие сводки из проверенного адресного источника.
Public Sub RestoreIndexedSummary(ByVal workbook As Object, ByVal source As CResultReadSource, ByVal fullFormatting As Boolean)
    source.ValidateForDisplay
    RestoreSummaryTables workbook, source, fullFormatting
End Sub

' Записывает те же массивы и оформление для двух явно проверенных входных контрактов.
Private Sub RestoreSummaryTables(ByVal workbook As Object, ByVal snapshot As Object, ByVal fullFormatting As Boolean)
    Set mCapturedSnapshot = Nothing
    mMinimalFormatting = Not fullFormatting
    Dim data As Variant: data = snapshot.Table("Strength.Data")
    Dim target As Object: Set target = TargetRange(workbook, UBound(data, 1))
    Dim headers As Variant: headers = snapshot.Table("Strength.Headers")
    If UBound(headers, 2) <> UBound(data, 2) Then Err.Raise 5, "CStrengthSummaryWriter", _
        "В снимке не совпадает число столбцов подписей и данных таблицы прочности."
    If UBound(headers, 2) <> TOTAL_COLUMNS Then ExpandStrengthSnapshot headers, data
    If mMinimalFormatting Then mMinimalFormatting = CStr(target.Cells(target.Rows.Count, 1).NumberFormat) = "@" And ResultHeaderMatches(target, headers, merges:=HeaderMergeBlocks())
    PrepareTarget target
    If Not mMinimalFormatting Then
        target.Resize(HEADER_ROWS, TOTAL_COLUMNS).Value2 = headers
        MergeHeaders target
    End If
    WriteBlockTitle target, "Расчет НДС и несущей способности по модели прочности"
    target.Offset(HEADER_ROWS, 0).Resize(UBound(data, 1), TOTAL_COLUMNS).Value2 = data
    FormatTable target, Nothing, data
End Sub

' Восстанавливает шапку и значения форматов 49, 57 и 62 столбцов. Три
' добавленных предельных усилия имеют единицы снимка и пустые значения.
Private Sub ExpandStrengthSnapshot(ByRef headers As Variant, ByRef data As Variant)
    Dim newHeaders() As Variant, expanded() As Variant, sourceColumns As Long
    sourceColumns = UBound(headers, 2)
    CopyStrengthSnapshotMatrix headers, newHeaders
    PutCapacitySpatialHeaders newHeaders, StrengthSnapshotUnit(headers(HEADER_LABEL_ROW, COL_CAP_MX)), _
        StrengthSnapshotUnit(headers(HEADER_LABEL_ROW, COL_CAP_N))
    If sourceColumns <> LEGACY_HOOP_STATES_STRENGTH_SNAPSHOT_COLUMNS Then
        Dim stressLabel As String: stressLabel = ChrW$(&H3C3)
        If sourceColumns = LEGACY_HOOP_STRENGTH_SNAPSHOT_COLUMNS Then _
            stressLabel = CStr(headers(HEADER_LABEL_ROW, 52)) ' Единица напряжения хомута в формате 57 столбцов.
        PutTransverseHeaders newHeaders, stressLabel
    End If
    headers = newHeaders
    CopyStrengthSnapshotMatrix data, expanded
    If sourceColumns = LEGACY_HOOP_STRENGTH_SNAPSHOT_COLUMNS Then
        Dim destination As Variant, row As Long, column As Long
        destination = Array(COL_TRANSVERSE_POINT, COL_TRANSVERSE_STRESS, COL_TRANSVERSE_UTIL, COL_TRANSVERSE_STATUS, _
            COL_TRANSVERSE_RESERVE, COL_TRANSVERSE_LIMIT, COL_TRANSVERSE_LIMIT_STATUS)
        For row = 1 To UBound(data, 1)
            For column = LBound(destination) To UBound(destination)
                expanded(row, CLng(destination(column))) = data(row, 51 + column)
            Next column
        Next row
    End If
    data = expanded
End Sub

' Переносит общие старые столбцы, оставляя место для T, Qx и Qy. В формате
' 62 столбцов состояние хомута уже полное и переносится тем же смещением.
Private Sub CopyStrengthSnapshotMatrix(ByRef values As Variant, ByRef expanded() As Variant)
    Dim row As Long, column As Long, endColumn As Long, offset As Long
    endColumn = LEGACY_STRENGTH_SNAPSHOT_COLUMNS
    If UBound(values, 2) = LEGACY_HOOP_STATES_STRENGTH_SNAPSHOT_COLUMNS Then endColumn = UBound(values, 2)
    offset = COL_CAP_QY - COL_CAP_MY
    ReDim expanded(1 To UBound(values, 1), 1 To TOTAL_COLUMNS)
    For row = 1 To UBound(values, 1)
        For column = 1 To COL_CAP_MY
            expanded(row, column) = values(row, column)
        Next column
        For column = COL_CAP_T To endColumn
            expanded(row, column + offset) = values(row, column)
        Next column
    Next row
End Sub

' Читает единицу после последней запятой в сохранённой подписи усилия.
Private Function StrengthSnapshotUnit(ByVal label As String) As String
    Dim separator As Long: separator = InStrRev(label, ",")
    If separator > 0 Then StrengthSnapshotUnit = Trim$(Mid$(label, separator + 1))
End Function
