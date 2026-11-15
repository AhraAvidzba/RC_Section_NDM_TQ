Attribute VB_Name = "modCoupledFieldResults"
Option Explicit

' ==========================================================================
' Пространственные поля сохранённого связанного состояния
' ==========================================================================
' Хранит дополнительные компоненты отдельно от девяти столбцов основной
' таблицы НДС. Схема и AutoCAD используют выбранный источник Excel/File.

Public Const COUPLED_FIELD_TABLE As String = "rngNDMCoupledFields"
Public Const COUPLED_FIELD_COLUMNS As Long = 10

' Возвращает позицию в существующей палитре по знаку пространственного
' значения. Эти ключи не описывают продольное состояние материала.
Public Function CoupledFieldPaletteState(ByVal value As Double) As String
    Select Case Sgn(value)
        Case -1: CoupledFieldPaletteState = "Compression"
        Case 1: CoupledFieldPaletteState = "Tension"
        Case Else: CoupledFieldPaletteState = "NearZero"
    End Select
End Function

' Создаёт шапку и место для строк подтверждённых состояний.
' Касательные напряжения хранятся в МПа, продольное перемещение в мм.
Public Function NewCoupledFieldTable(ByVal rows As Long) As Variant
    Dim data() As Variant, headers As Variant, i As Long
    headers = Array("RunID", "LoadCase", "ProfileId", "StateType", "ElementID", _
        "TauXZ, MPa", "TauYZ, MPa", "GammaXZ", "GammaYZ", "Warping, mm")
    ReDim data(1 To rows + 1, 1 To COUPLED_FIELD_COLUMNS)
    For i = 1 To COUPLED_FIELD_COLUMNS
        data(1, i) = headers(i - 1)
    Next i
    NewCoupledFieldTable = data
End Function

' Читает величину для одного запуска, сочетания, профиля и состояния.
' Результат индексируется по ElementID; ошибки указывают источник данных.
Public Function ReadCoupledFieldValues(ByVal workbook As Object, ByVal quantity As EVisualizationQuantity, _
        ByVal runID As Double, ByVal combinationID As String, ByVal profileID As String, ByVal stateType As String) As Object
    Dim name As String, firstColumn As String, secondColumn As String
    name = VisualizationQuantityToText(quantity)
    Select Case quantity
        Case vqTau
            firstColumn = "TauXZ": secondColumn = "TauYZ"
        Case vqGamma
            firstColumn = "GammaXZ": secondColumn = "GammaYZ"
        Case Else
            firstColumn = name
    End Select
    Dim source As CResultReadSource, anchor As Object, data As Variant, origin As String
    Set source = CurrentResultSource(workbook)
    If source Is Nothing Then
        Set anchor = NamedCoupledFieldAnchor(workbook)
        If Not anchor Is Nothing Then
            origin = anchor.Worksheet.Name & "!" & anchor.Address(False, False)
            data = ReadPhysicalResultTable(workbook, COUPLED_FIELD_TABLE)
        End If
    Else
        origin = "Results.SnapshotName, " & COUPLED_FIELD_TABLE
        If source.HasTable(COUPLED_FIELD_TABLE) Then data = source.Table(COUPLED_FIELD_TABLE)
    End If
    Dim description As String
    description = combinationID & "/" & profileID & "/" & stateType & " (" & name & ")"
    If Not HasResultTableRows(data) Then CoupledFieldUnavailable origin, description
    Dim columns As Variant: columns = Array("RunID", "LoadCase", "ProfileId", "StateType", "ElementID", firstColumn)
    Dim indices(0 To 5) As Long, i As Long, row As Long, colOther As Long
    For i = 0 To 5
        indices(i) = ResultColumn(data, CStr(columns(i)))
    Next i
    If Len(secondColumn) > 0 Then colOther = ResultColumn(data, secondColumn)

    ValidateSnapshotTableShape COUPLED_FIELD_TABLE, data
    Dim values As Object, elementID As String, number As Double, other As Double, storedRun As Double
    Set values = CreateObject("Scripting.Dictionary")
    values.CompareMode = vbTextCompare
    For row = 2 To UBound(data, 1)
        If StrComp(SafeText(data(row, indices(1))), combinationID, vbTextCompare) = 0 And _
                StrComp(SafeText(data(row, indices(2))), profileID, vbTextCompare) = 0 And _
                StrComp(SafeText(data(row, indices(3))), stateType, vbTextCompare) = 0 Then
            elementID = SafeText(data(row, indices(4)))
            If Len(elementID) = 0 Or values.Exists(elementID) Then _
                InvalidCoupledFieldValue origin, anchor, row, indices(4), "Пустой или повторный ElementID."
            storedRun = ReadCoupledFieldNumber(data, row, indices(0), origin, anchor)
            If storedRun <> runID Then _
                InvalidCoupledFieldValue origin, anchor, row, indices(0), "Поле относится к другому запуску расчёта."
            number = ReadCoupledFieldNumber(data, row, indices(5), origin, anchor)
            If colOther > 0 Then
                other = ReadCoupledFieldNumber(data, row, colOther, origin, anchor)
                number = Sqr(number * number + other * other)
            End If
            values.Add elementID, number
        End If
    Next row
    If values.Count = 0 Then CoupledFieldUnavailable origin, description
    Set ReadCoupledFieldValues = values
End Function

' Ищет необязательный диапазон. В книгах без пространственной таблицы
' ошибка выдаётся только при запросе её поля через профиль.
Public Function NamedCoupledFieldAnchor(ByVal workbook As Object) As Object
    On Error Resume Next
    Set NamedCoupledFieldAnchor = workbook.Names.Item(COUPLED_FIELD_TABLE).RefersToRange
    On Error GoTo 0
End Function

' Возвращает размерность канонических полей дополнительной таблицы.
Public Function CoupledFieldUnit(ByVal quantity As EVisualizationQuantity) As String
    If VisualizationQuantityIsStress(quantity) Then
        CoupledFieldUnit = "MPa"
    ElseIf quantity = vqWarping Then
        CoupledFieldUnit = "mm"
    Else
        CoupledFieldUnit = "-"
    End If
End Function

' Читает действительное сохранённое число; пустота не считается нулём.
Private Function ReadCoupledFieldNumber(ByRef data As Variant, ByVal row As Long, ByVal col As Long, _
        ByVal origin As String, ByVal anchor As Object) As Double
    Dim value As Double
    If Not TryReadStrictDouble(data(row, col), value) Then _
        InvalidCoupledFieldValue origin, anchor, row, col, "В пространственном поле требуется числовое значение."
    ReadCoupledFieldNumber = value
End Function

' Сообщает фактический адрес неподходящей строки Excel или номер записи файла.
Private Sub InvalidCoupledFieldValue(ByVal origin As String, ByVal anchor As Object, ByVal row As Long, _
        ByVal col As Long, ByVal reason As String)
    Dim address As String
    If anchor Is Nothing Then
        address = origin & ", строка " & CStr(row) & ", столбец " & CStr(col)
    Else
        address = anchor.Worksheet.Name & "!" & anchor.Offset(row - 1, col - 1).Address(False, False)
    End If
    Err.Raise vbObjectError + 4389, "modCoupledFieldResults", address & " (" & COUPLED_FIELD_TABLE & "). " & reason
End Sub

' Сообщает об отсутствии сохранённого поля для точной комбинации ключей.
Private Sub CoupledFieldUnavailable(ByVal origin As String, ByVal description As String)
    If Len(origin) = 0 Then origin = COUPLED_FIELD_TABLE
    Err.Raise vbObjectError + 4388, "modCoupledFieldResults", "В " & origin & " нет поля для " & description & _
        ". Выберите состояние связанного расчёта или продольное поле Stress либо Strain."
End Sub

' Возвращает значение элемента; неполный набор поля не считается нулевым.
Public Function CoupledFieldValue(ByVal values As Object, ByVal elementID As String) As Double
    If Not values.Exists(elementID) Then Err.Raise vbObjectError + 4389, "modCoupledFieldResults", _
        "В " & COUPLED_FIELD_TABLE & " отсутствует сохранённое поле элемента " & elementID & ". Повторите расчёт."
    CoupledFieldValue = CDbl(values.Item(elementID))
End Function
