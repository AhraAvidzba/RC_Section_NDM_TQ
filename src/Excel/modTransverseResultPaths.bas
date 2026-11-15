Attribute VB_Name = "modTransverseResultPaths"
Option Explicit

' ============================================================================
' Сохранённые оси поперечной арматуры связанного НДС
' ============================================================================
' Аннотации принадлежат одному RunID, сочетанию, профилю и виду НДС.
' Сегменты хранятся в координатах модели с указанной выходной единицей.
' Не участвуют в построении границ бетона и восстановлении исходного сечения.

Private Const PATH_PREFIX As String = "TRANSVERSE_"
Private Const PATH_LINE As String = "TRANSVERSE_LINE"
Private Const PATH_ARC As String = "TRANSVERSE_ARC"
Private Const PATH_CIRCLE As String = "TRANSVERSE_CIRCLE"

' Отделяет геометрию расчётных состояний от общих аннотаций сечения.
Public Function IsCalculatedSectionAnnotation(ByVal annotationType As String) As Boolean
    IsCalculatedSectionAnnotation = (Left$(annotationType, 6) = "CRACK_" Or Left$(annotationType, Len(PATH_PREFIX)) = PATH_PREFIX)
End Function

' Записывает оси хомутов и ветвей, ранее подготовленные для равновесия.
' Идентификатор запуска записывается числом и доступен сразу в снимке в памяти.
Public Sub PutTransversePathRows(ByRef data() As Variant, ByRef row As Long, ByVal runID As Double, _
        ByVal combinationID As String, ByVal profileID As String, ByVal stateType As String, _
        ByVal layout As CTransverseRebarLayout, ByVal units As CUnitSystem)
    Dim i As Long, j As Long, path As CGeometryPath, edge As TRegionEdge, extra As String
    Dim lengthUnit As String: lengthUnit = "mm"
    If Not units Is Nothing Then lengthUnit = units.OutputLengthUnit
    For i = 1 To layout.Count
        Set path = layout.Path(i)
        extra = SectionOptionalJson("ProfileId", profileID, "StateType", stateType, "Closed", CStr(Abs(CLng(path.IsClosed))), _
            "DiameterMM", SectionOptionalNumber(layout.Diameter(i)), "SpacingMM", SectionOptionalNumber(layout.Spacing(i)), _
            "SteelClass", layout.SteelClass(i), "InputReference", layout.SourceReference(i))
        For j = 1 To path.SegmentCount
            path.CopySegment j, edge
            data(row, 1) = runID: data(row, 2) = PATH_LINE
            data(row, 3) = PATH_PREFIX & combinationID & "_" & stateType & "_" & path.ID & "_" & CStr(j)
            PutTransverseCoordinates data, row, edge, units
            If edge.Sweep <> 0# Then
                If Abs(edge.Sweep) = 2# * GEOM_PI Then
                    data(row, 2) = PATH_CIRCLE
                    data(row, 10) = CStr(Sgn(edge.Sweep)) ' Полный оборот хранит направление без округления угла.
                Else
                    data(row, 2) = PATH_ARC
                    data(row, 10) = SectionOptionalNumber(edge.Sweep)
                End If
            End If
            data(row, 13) = "Ось поперечной арматуры связанного сечения."
            data(row, 14) = combinationID: data(row, 16) = path.ID: data(row, 17) = j
            data(row, 23) = lengthUnit: data(row, 24) = extra
            row = row + 1
        Next j
    Next i
End Sub

' Переводит координаты сегмента из мм в единицу, указанную в аннотации.
Private Sub PutTransverseCoordinates(ByRef data() As Variant, ByVal row As Long, ByRef edge As TRegionEdge, ByVal units As CUnitSystem)
    Dim coordinates As Variant, i As Long, value As Double
    coordinates = Array(edge.X1, edge.Y1, edge.X2, edge.Y2, edge.CenterX, edge.CenterY)
    For i = 0 To UBound(coordinates)
        value = coordinates(i)
        If Not units Is Nothing Then value = units.InternalLengthToOutput(value)
        data(row, 4 + i) = value
    Next i
End Sub

' Читает только пути выбранного состояния из таблицы текущего снимка.
' Отсутствие записей допустимо для плоского НДС и явно заданного «Нет».
Public Function ReadSavedTransversePaths(ByVal workbook As Object, ByVal runID As Double, ByVal combinationID As String, _
        ByVal profileID As String, ByVal stateType As String, Optional ByRef data As Variant) As Collection
    If IsMissing(data) Then data = ReadAnchoredResultTable(workbook, "rngNDMSectionAnnotations")
    If Not HasResultTableRows(data) Then Exit Function
    Dim anchor As Object: Set anchor = workbook.Names.Item("rngNDMSectionAnnotations").RefersToRange
    Dim found As Object: Set found = CreateObject("Scripting.Dictionary"): found.CompareMode = vbTextCompare
    Dim result As New Collection, path As CGeometryPath, edge As TRegionEdge
    Dim row As Long, kind As String, info As Object, pathID As String, closed As String, lengthUnit As String
    Dim closures As Object: Set closures = CreateObject("Scripting.Dictionary"): closures.CompareMode = vbTextCompare
    For row = 2 To UBound(data, 1)
        kind = ReadSavedAnnotationText(data, anchor, row, 2)
        If Left$(kind, Len(PATH_PREFIX)) = PATH_PREFIX And SafeText(data(row, 14)) = combinationID Then
            Set info = ReadTransverseRecord(data, anchor, row)
            If SectionOptionalText(info, "ProfileId") = profileID And SectionOptionalText(info, "StateType") = stateType Then
                If ReadGeometryNumber(data, anchor, row, 1) <> runID Then TransverseRowError anchor, row, "Ось относится к другому запуску."
                pathID = ReadSavedAnnotationText(data, anchor, row, 16)
                closed = SectionOptionalText(info, "Closed")
                If Len(pathID) = 0 Or (closed <> "0" And closed <> "1") Then _
                    TransverseRowError anchor, row, "Не задана группа или признак замкнутости оси."
                If Not found.Exists(pathID) Then
                    Set path = New CGeometryPath: path.Initialize pathID
                    found.Add pathID, path: closures.Add pathID, closed
                    result.Add path
                Else
                    Set path = found.Item(pathID)
                    If closures.Item(pathID) <> closed Then TransverseRowError anchor, row, "Не совпадает признак замкнутости оси."
                End If
                If ReadGeometryNumber(data, anchor, row, 17) <> path.SegmentCount + 1 Then _
                    TransverseRowError anchor, row, "Нарушен порядок сегментов оси."
                lengthUnit = ReadSavedAnnotationText(data, anchor, row, 23)
                ReadTransverseEdge data, anchor, row, kind, lengthUnit, edge
                path.Append CStr(path.SegmentCount + 1), edge
            End If
        End If
    Next row
    For Each path In result
        path.Seal closures.Item(path.ID) = "1"
    Next path
    If result.Count > 0 Then Set ReadSavedTransversePaths = result
End Function

' Читает ключи состояния, запрашивая адрес Excel только при ошибке.
Private Function ReadTransverseRecord(ByRef data As Variant, ByVal anchor As Object, ByVal row As Long) As Object
    On Error GoTo Failed
    Set ReadTransverseRecord = ReadSectionOptionalJson(ReadSavedAnnotationText(data, anchor, row, 24))
    If Len(SectionOptionalText(ReadTransverseRecord, "ProfileId")) = 0 Or _
            Len(SectionOptionalText(ReadTransverseRecord, "StateType")) = 0 Then _
        Err.Raise 5, , "Не указан профиль или вид НДС."
    Exit Function
Failed:
    TransverseRowError anchor, row, Err.Description
End Function

' Восстанавливает прямую или точную дугу в мм из общей схемы аннотаций.
Private Sub ReadTransverseEdge(ByRef data As Variant, ByVal anchor As Object, ByVal row As Long, ByVal kind As String, _
        ByVal lengthUnit As String, ByRef edge As TRegionEdge)
    If lengthUnit <> "mm" And lengthUnit <> "cm" And lengthUnit <> "m" Then TransverseRowError anchor, row, "Неверная единица длины оси."
    edge.X1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 4), lengthUnit)
    edge.Y1 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 5), lengthUnit)
    edge.X2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 6), lengthUnit)
    edge.Y2 = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 7), lengthUnit)
    edge.Sweep = 0#: edge.CenterX = 0#: edge.CenterY = 0#
    Select Case kind
        Case PATH_ARC, PATH_CIRCLE
            edge.CenterX = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 8), lengthUnit)
            edge.CenterY = OutputLengthToInternalByUnit(ReadGeometryNumber(data, anchor, row, 9), lengthUnit)
            If kind = PATH_CIRCLE Then
                Dim direction As String: direction = ReadSavedAnnotationText(data, anchor, row, 10)
                If direction <> "1" And direction <> "-1" Then TransverseRowError anchor, row, "Неверное направление обхода окружности."
                edge.Sweep = CDbl(direction) * 2# * GEOM_PI
            Else
                edge.Sweep = ReadSavedContourArcSweep(data, anchor, row, 10)
            End If
        Case PATH_LINE
        Case Else
            TransverseRowError anchor, row, "Неверный тип сегмента оси."
    End Select
End Sub

' Формирует ссылку на строку в текущей раскладке Results.
Private Function TransverseRowAddress(ByVal anchor As Object, ByVal row As Long) As String
    TransverseRowAddress = anchor.Parent.Name & "!" & anchor.Offset(row - 1, 0).Resize(1, 24).Address(False, False)
End Function

' Формирует диагностируемый отказ до отрисовки новой схемы.
Private Sub TransverseRowError(ByVal anchor As Object, ByVal row As Long, ByVal reason As String)
    Err.Raise vbObjectError + 4376, "ReadSavedTransversePaths", reason & " " & TransverseRowAddress(anchor, row) & ". Повторите расчёт."
End Sub
