Attribute VB_Name = "modCoupledLoadChart"
Option Explicit

' ==========================================================================
' График принятых шагов связанного расчёта на Results
' ==========================================================================
' Читает только сохранённые свойства выбранного сочетания. Обе ветви
' независимы, а точка численной остановки не считается пределом.

Public Const COUPLED_TRACE_DIRECT As String = "Coupled.Trace.Direct."
Public Const COUPLED_TRACE_CAPACITY As String = "Coupled.Trace.Capacity."
Private Const TRACE_PREFIX As String = "Coupled.Trace."
Private Const TRACE_CHART As String = "chtNDMCoupledLoadTrace"

' Обновляет диаграмму при наличии сохранённых точек. Плоское сочетание,
' предварительная геометрия и снимки без истории освобождают место.
Public Function TryUpdateCoupledLoadChart(ByVal workbook As Object, ByVal settings As CSystemSettingsReader) As String
    On Error GoTo Failed
    Dim properties As Variant, source As CResultReadSource
    Set source = CurrentResultSource(workbook)
    If source Is Nothing Then
        properties = ReadPhysicalResultTable(workbook, "rngNDMSectionProperties")
    ElseIf Not source.IsGeometryPreview Then
        If source.HasTable("LoadCase.Properties") Then properties = source.Table("LoadCase.Properties")
    End If
    If Not HasResultTableRows(properties) Then ClearCoupledLoadChart workbook: Exit Function
    Dim combination As String, values As Object, row As Long, key As String
    For row = 2 To UBound(properties, 1)
        If Left$(CStr(properties(row, 3)), Len(TRACE_PREFIX)) = TRACE_PREFIX Then Exit For
    Next row
    If row > UBound(properties, 1) Then ClearCoupledLoadChart workbook: Exit Function
    combination = ReadOutputCombinationID(workbook, settings)
    Set values = CreateObject("Scripting.Dictionary")
    For row = 2 To UBound(properties, 1)
        If StrComp(CStr(properties(row, 2)), combination, vbTextCompare) = 0 Then
            key = CStr(properties(row, 3))
            If Left$(key, Len(TRACE_PREFIX)) = TRACE_PREFIX Or Left$(key, 17) = "Coupled.Capacity." Then _
                values.Add key, properties(row, 4)
        End If
    Next row
    If Not values.Exists(COUPLED_TRACE_DIRECT & "1.Lambda") And Not values.Exists(COUPLED_TRACE_CAPACITY & "1.Lambda") Then
        ClearCoupledLoadChart workbook
        Exit Function
    End If
    Dim chart As Object: Set chart = PrepareCoupledLoadChart(workbook, combination)
    Dim directCount As Long, capacityCount As Long, maxPoints As Long
    directCount = DrawLoadBranch(chart, values, COUPLED_TRACE_DIRECT, "Заданные усилия", RGB(56, 96, 135))
    capacityCount = DrawLoadBranch(chart, values, COUPLED_TRACE_CAPACITY, "Поиск предела", RGB(152, 85, 44))
    maxPoints = directCount
    If capacityCount > maxPoints Then maxPoints = capacityCount
    chart.HasLegend = True: chart.Legend.Position = -4107 ' xlLegendPositionBottom.
    chart.Axes(1).HasTitle = True: chart.Axes(1).AxisTitle.Text = "Номер принятой точки ветви"
    chart.Axes(1).MinimumScale = 1#: chart.Axes(1).MajorUnit = (maxPoints + 7) \ 8
    chart.Axes(1).TickLabels.NumberFormat = "0"
    chart.Axes(2).HasTitle = True: chart.Axes(2).AxisTitle.Text = ChrW$(&H3BB) & " или доля постоянной части"
    chart.Axes(2).TickLabels.NumberFormat = "0.00"
    Exit Function
Failed:
    Dim message As String: message = Err.Description
    ClearCoupledLoadChart workbook
    TryUpdateCoupledLoadChart = "График нагружения на Results не обновлён. " & message
End Function

' Создаёт одно окно справа от подробной таблицы прочности. Положение
' следует её именованному якорю и действующим размерам таблицы.
Private Function PrepareCoupledLoadChart(ByVal workbook As Object, ByVal combination As String) As Object
    Dim writer As New CStrengthSummaryWriter, anchor As Object, container As Object, chart As Object
    Set anchor = workbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    Set anchor = anchor.Offset(-writer.HeaderRowsAboveAnchor, writer.RequiredColumns + 1)
    Set container = CoupledChartObject(workbook)
    If container Is Nothing Then
        Set container = anchor.Worksheet.ChartObjects.Add(anchor.Left, anchor.Top, 470#, 240#)
        container.Name = TRACE_CHART
        container.Placement = 3 ' xlFreeFloating.
    ElseIf container.Left < anchor.Left Then
        container.Left = anchor.Left ' После расширения таблицы сохранённый график может перекрывать новые столбцы.
    End If
    Set chart = container.Chart
    Do While chart.SeriesCollection.Count > 0: chart.SeriesCollection(1).Delete: Loop
    chart.ChartType = 74 ' xlXYScatterLines: соединяем только принятые точки.
    chart.HasTitle = True: chart.ChartTitle.Text = "График нагружения — " & combination
    chart.ChartTitle.Font.Size = 12
    Set PrepareCoupledLoadChart = chart
End Function

' Показывает постоянную часть и изменяемый множитель раздельно.
' Маркеры ставятся в подтверждённых точках соответствующей ветви.
Private Function DrawLoadBranch(ByVal chart As Object, ByVal values As Object, ByVal prefix As String, _
        ByVal name As String, ByVal color As Long) As Long
    Dim stepIndex As Long, beforeCount As Long, afterCount As Long, count As Long
    stepIndex = 1
    Do While values.Exists(prefix & CStr(stepIndex) & ".Lambda"): stepIndex = stepIndex + 1: Loop
    count = stepIndex - 1
    If count = 0 Then Exit Function
    DrawLoadBranch = count
    Dim xOffset() As Double, yOffset() As Double, xVariable() As Double, yVariable() As Double
    ReDim xOffset(1 To count): ReDim yOffset(1 To count): ReDim xVariable(1 To count): ReDim yVariable(1 To count)
    Dim offset As Double, lambda As Double, key As String
    For stepIndex = 1 To count
        key = prefix & CStr(stepIndex)
        offset = CDbl(values(key & ".Offset")): lambda = CDbl(values(key & ".Lambda"))
        If lambda = 0# Then
            beforeCount = beforeCount + 1
            xOffset(beforeCount) = stepIndex: yOffset(beforeCount) = offset
        End If
        If offset = 1# Then
            afterCount = afterCount + 1
            xVariable(afterCount) = stepIndex: yVariable(afterCount) = lambda
        End If
    Next stepIndex
    If beforeCount > 0 Then
        ReDim Preserve xOffset(1 To beforeCount): ReDim Preserve yOffset(1 To beforeCount)
        AddLoadSeries chart, name & " · постоянная часть", xOffset, yOffset, color, True
    End If
    If afterCount = 0 Then Exit Function
    ReDim Preserve xVariable(1 To afterCount): ReDim Preserve yVariable(1 To afterCount)
    AddLoadSeries chart, name, xVariable, yVariable, color
    Dim label As String
    If prefix = COUPLED_TRACE_DIRECT Then
        If lambda = 1# Then label = "Заданная точка"
    Else
        label = "Последняя точка поиска"
        If values.Exists("Coupled.Capacity.PeakLambda") Then
            If CDbl(values("Coupled.Capacity.PeakLambda")) = lambda Then label = "Подтверждённый максимум"
        ElseIf values.Exists("Coupled.Capacity.SteelLambda") Then
            If CDbl(values("Coupled.Capacity.SteelLambda")) = lambda Then label = "Предельная деформация арматуры"
        End If
    End If
    If Len(label) > 0 Then AddLoadSeries chart, label, Array(xVariable(afterCount)), Array(lambda), color, False, True
End Function

' Передаёт рассчитанные координаты в нативный ряд без дополнительных
' ячеек и скрытых формул. Линия никогда не сглаживается.
Private Sub AddLoadSeries(ByVal chart As Object, ByVal name As String, ByVal x As Variant, ByVal y As Variant, _
        ByVal color As Long, Optional ByVal dashed As Boolean = False, Optional ByVal markerOnly As Boolean = False)
    Dim series As Object: Set series = chart.SeriesCollection.NewSeries
    series.Name = name: series.XValues = x: series.Values = y
    series.MarkerStyle = 8: series.MarkerSize = 4
    series.MarkerForegroundColor = color: series.MarkerBackgroundColor = color
    series.Format.Line.ForeColor.RGB = color
    If dashed Then series.Format.Line.DashStyle = 4 ' msoLineDash.
    If markerOnly Then
        series.MarkerStyle = 2: series.MarkerSize = 8
        series.Format.Line.Visible = 0
    End If
    series.Smooth = False
End Sub

' Удаляет только собственную диаграмму подробных результатов.
Public Sub ClearCoupledLoadChart(ByVal workbook As Object)
    Dim chart As Object: Set chart = CoupledChartObject(workbook)
    If Not chart Is Nothing Then chart.Delete
End Sub

' Пустая книга и пользовательские диаграммы не требуют изменений.
Private Function CoupledChartObject(ByVal workbook As Object) As Object
    On Error Resume Next
    Set CoupledChartObject = workbook.Worksheets("Results").ChartObjects(TRACE_CHART)
End Function
