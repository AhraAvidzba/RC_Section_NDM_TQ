Attribute VB_Name = "modCoupledWorkbookInput"
Option Explicit

' ==========================================================================
' Ввод связанного расчёта из книги
' ==========================================================================
' Дополняет настройки и читает поперечную арматуру по типу сохранённого
' сечения. Ошибка этой подготовки относится к сочетаниям с T, Qx или Qy.
' Нагрузки и текущие настройки читаются заново, бетон и продольные стержни
' берутся из сформированного сечения без повторной генерации или импорта.

' Подключает чистый решатель после успешной адресной проверки ввода.
' Настройки энергии и длины заданы в фиксированных Н/мм и мм.
Public Sub ConfigureWorkbookCoupledAnalysis(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal settings As CSystemSettingsReader, ByVal units As CUnitSystem, _
        ByVal provider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator)
    On Error GoTo Failed
    settings.LoadCoupledSettings workbook
    Dim poisson As Double, gt As Double, gc As Double, length As Double, yieldRatio As Double, peak As Double
    poisson = settings.GetRequiredDouble("Coupled.Poisson")
    If poisson < 0# Or poisson >= 0.5 Then CalibrationError settings, "Coupled.Poisson", "Введите число от 0 включительно до 0,5."
    gt = PositiveCalibrationValue(settings, "Coupled.Gt")
    gc = PositiveCalibrationValue(settings, "Coupled.Gc")
    length = PositiveCalibrationValue(settings, "Coupled.Localization")
    yieldRatio = PositiveCalibrationValue(settings, "Coupled.YieldRatio")
    If yieldRatio >= 1# Then CalibrationError settings, "Coupled.YieldRatio", "Введите отношение меньше 1."
    peak = PositiveCalibrationValue(settings, "Coupled.PeakStrain")

    Dim steelClass As String, layout As CTransverseRebarLayout, supplementary As Object
    steelClass = settings.GetString("Coupled.SteelClass")
    If Len(steelClass) = 0 And section.SourceType = "AutoCADImport" Then
        Dim bar As Long
        For bar = 1 To section.RebarCount
            If Len(section.RebarSteelClassForMaterial(bar, steelClass)) = 0 Then _
                CalibrationError settings, "Coupled.SteelClass", "Укажите основной класс продольной арматуры, импортированной из AutoCAD."
        Next bar
    End If
    Set layout = ReadSavedSectionTransverse(workbook, section, units, steelClass)
    ValidateConcreteMaterial section
    Set supplementary = ReadSupplementarySteel(workbook, section, layout, steelClass, provider.PlateauSlopeFactor)

    Dim analysis As New CCoupledSectionAnalysis
    analysis.Initialize section, layout, provider, poisson, gt, gc, length, yieldRatio, peak, steelClass, _
        settings.GetRequiredLong("Solver.MaxIterations"), supplementary
    batch.ConfigureCoupledAnalysis analysis
    Exit Sub
Failed:
    If Err.Number = 18 Then Err.Raise Err.Number, Err.Source, Err.Description ' Отмена прерывает запуск, ошибка ввода относится к сочетанию.
    Dim reason As String: reason = Err.Description
    batch.ConfigureCoupledAnalysis Nothing, reason
End Sub

' Находит Transverse по источнику уже сохранённого сечения. Для AutoCAD
' координаты ветвей читаются в осях импортированной модели (МСК).
Private Function ReadSavedSectionTransverse(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal units As CUnitSystem, ByVal steelClass As String) As CTransverseRebarLayout
    Const AUTOCAD_TABLE As String = "rngAutoCADTransverseRebar"
    Dim tables As Variant, imported As Boolean
    imported = (StrComp(section.SourceType, "AutoCADImport", vbTextCompare) = 0)
    If imported Then
        tables = Array(AUTOCAD_TABLE)
    Else
        If Len(section.RecipeType) = 0 Then _
            Err.Raise vbObjectError + 4194, "ReadSavedSectionTransverse", _
                "В сохранённой геометрии отсутствует тип рецепта (rngNDMSectionProperties). Нажмите «Сформировать сечение»."
        Dim registry As New CSectionTypeRegistry
        tables = registry.RecipeTableNames(section.RecipeType)
    End If
    Dim reader As New CSectionInputSnapshotReader, context As CSectionRecipeContext
    Set context = reader.Capture(workbook, tables, units)
    Dim sectionInput As CSectionInputReader, layout As CTransverseRebarLayout, name As Variant
    For Each name In tables
        Set sectionInput = context.Reader(CStr(name))
        Set layout = sectionInput.ReadTransverseLayout(section, steelClass)
        If layout.IsSpecified Then Set ReadSavedSectionTransverse = layout: Exit Function
    Next name
    Dim address As String: address = Join(tables, ", ")
    If imported Then address = AUTOCAD_TABLE & ", " & sectionInput.SourceReference("Transverse", "Group1", "kind")
    Err.Raise vbObjectError + 4194, "ReadSavedSectionTransverse", _
        "В блоке Transverse на листе «Сечения» укажите поперечную арматуру или выберите «Нет» (" & address & ")."
End Function

' Сверяет идентификаторы бетона, пока для всего сечения задана одна калибровка.
Private Sub ValidateConcreteMaterial(ByVal section As CSectionModel)
    Dim i As Long, firstConcrete As Long
    firstConcrete = section.ConcreteMaterialID(1)
    For i = 2 To section.ConcreteCount
        If section.ConcreteMaterialID(i) <> firstConcrete Then _
            Err.Raise vbObjectError + 4194, "ValidateConcreteMaterial", _
                "Для связанного расчёта задана одна калибровка бетона, а в сохранённом сечении несколько ID материалов (rngNDMSectionGeometry, элемент " & section.ConcreteID(i) & ")."
    Next i
End Sub

' Читает дополнительные классы, фактически указанные в геометрии.
' Строка хранится как готовая модель или адресная ошибка её набора ULS/SLS.
Private Function ReadSupplementarySteel(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal layout As CTransverseRebarLayout, ByVal primaryClass As String, ByVal slopeFactor As Double) As Object
    Dim needed As Object: Set needed = CreateObject("Scripting.Dictionary"): needed.CompareMode = vbTextCompare
    Dim i As Long, grade As String
    For i = 1 To section.RebarCount
        grade = section.RebarSteelClassForMaterial(i, primaryClass)
        If StrComp(grade, primaryClass, vbTextCompare) <> 0 Then needed(grade) = True
    Next i
    For i = 1 To layout.Count
        grade = layout.SteelClass(i)
        If StrComp(grade, primaryClass, vbTextCompare) <> 0 Then needed(grade) = True
    Next i
    If needed.Count = 0 Then Exit Function

    Dim table As Object
    On Error Resume Next
    Set table = workbook.Names.Item("rngCoupledSteelClasses").RefersToRange
    Dim wasCancelled As Boolean: wasCancelled = (Err.Number = 18)
    On Error GoTo 0
    If wasCancelled Then Err.Raise 18, "ReadSupplementarySteel"
    If table Is Nothing Then Err.Raise vbObjectError + 4194, "ReadSupplementarySteel", _
        "На Config отсутствует таблица дополнительных классов арматуры rngCoupledSteelClasses."
    If table.Areas.Count <> 1 Or table.Columns.Count <> 7 Or table.Rows.Count < 2 Then _
        SteelInputError table, 1, 1, "Восстановите таблицу классов из семи столбцов."
    Dim values As Variant: values = table.Value2
    Dim epsilon As String: epsilon = ChrW$(&H3B5) ' Греческая буква отсутствует в штатной кодовой странице исходников VBA.
    Dim headers As Variant: headers = Array("Класс стали", "Набор", "Es, МПа", "Rs, МПа", "Rsc, МПа", epsilon & "s2", epsilon & "sc2")
    For i = 1 To 7
        If IsError(values(1, i)) Then SteelInputError table, 1, i, "Исправьте заголовок столбца."
        If CStr(values(1, i)) <> headers(i - 1) Then SteelInputError table, 1, i, "Требуется заголовок «" & headers(i - 1) & "»."
    Next i

    Dim result As Object: Set result = CreateObject("Scripting.Dictionary")
    Dim row As Long, key As String, models As Object, classModel As CCoupledSteelModel, inputError As String
    For row = 2 To UBound(values, 1)
        If IsError(values(row, 1)) Then SteelInputError table, row, 1, "Исправьте обозначение класса."
        grade = Trim$(CStr(values(row, 1)))
        If needed.Exists(grade) And Len(grade) > 0 Then
            key = ReadSteelValueSet(values(row, 2), table, row)
            If Not result.Exists(key) Then
                Set models = CreateObject("Scripting.Dictionary"): models.CompareMode = vbTextCompare
                result.Add key, models
            Else
                Set models = result(key)
            End If
            If models.Exists(grade) Then
                models(grade) = SteelInputMessage(table, row, 1, "Для класса «" & grade & "» и набора " & key & " задано несколько строк.")
            Else
                Set classModel = ReadSteelModel(values, table, row, slopeFactor, inputError)
                If classModel Is Nothing Then
                    models.Add grade, inputError
                Else
                    models.Add grade, classModel
                End If
            End If
        End If
    Next row
    Set ReadSupplementarySteel = result
End Function

' Распознаёт обозначения наборов, используемые в существующих профилях.
Private Function ReadSteelValueSet(ByVal value As Variant, ByVal table As Object, ByVal row As Long) As String
    If Not IsError(value) Then
        Select Case UCase$(Trim$(CStr(value)))
            Case "ULS(I)": ReadSteelValueSet = "ULS(I)": Exit Function
            Case "SLS(II)": ReadSteelValueSet = "SLS(II)": Exit Function
        End Select
    End If
    SteelInputError table, row, 2, "Введите ULS(I) или SLS(II), как в расчётном профиле."
End Function

' Создаёт двухлинейный закон с модулем и сопротивлениями в МПа.
' Числовая ошибка сохраняется только для набора, к которому относится строка.
Private Function ReadSteelModel(ByRef data As Variant, ByVal table As Object, ByVal row As Long, _
        ByVal slopeFactor As Double, ByRef inputError As String) As CCoupledSteelModel
    Dim values(3 To 7) As Double, column As Long
    On Error GoTo Failed
    For column = 3 To 7
        If Not TryReadStrictDouble(data(row, column), values(column)) Or values(column) <= 0# Then _
            SteelInputError table, row, column, "Введите положительное число."
    Next column
    If values(6) <= values(4) / values(3) Then SteelInputError table, row, 6, "Предельная деформация растяжения должна превышать Rs/Es."
    If values(7) <= values(5) / values(3) Then SteelInputError table, row, 7, "Предельная деформация сжатия должна превышать Rsc/Es."
    Dim model As New CCoupledSteelModel, slope As Double
    slope = slopeFactor * values(3)
    model.Initialize values(3), values(4), values(5), slope, slope, values(6), values(7)
    inputError = vbNullString
    Set ReadSteelModel = model
    Exit Function
Failed:
    If Err.Number = 18 Then Err.Raise Err.Number, Err.Source, Err.Description
    inputError = Err.Description
End Function

' Возвращает адрес фактической ячейки строки класса для диагностики.
Private Function SteelInputMessage(ByVal table As Object, ByVal row As Long, ByVal column As Long, ByVal reason As String) As String
    SteelInputMessage = reason & " " & table.Worksheet.Name & "!" & table.Cells(row, column).Address(False, False) & " (rngCoupledSteelClasses)."
End Function

' Возвращает ошибку ключа, структуры или числового значения каталога.
Private Sub SteelInputError(ByVal table As Object, ByVal row As Long, ByVal column As Long, ByVal reason As String)
    Err.Raise vbObjectError + 4194, "ReadSupplementarySteel", SteelInputMessage(table, row, column, reason)
End Sub

' Проверяет положительное значение с диагностикой по имени и адресу ячейки.
Private Function PositiveCalibrationValue(ByVal settings As CSystemSettingsReader, ByVal key As String) As Double
    PositiveCalibrationValue = settings.GetRequiredDouble(key)
    If PositiveCalibrationValue <= 0# Then CalibrationError settings, key, "Введите положительное число."
End Function

' Возвращает ошибку инженерного диапазона через общий словарь адресов Config.
Private Sub CalibrationError(ByVal settings As CSystemSettingsReader, ByVal key As String, ByVal correction As String)
    Err.Raise vbObjectError + 4194, "modCoupledWorkbookInput", _
        settings.InputErrorMessage(key, "Значение вне допустимого диапазона.", correction)
End Sub
