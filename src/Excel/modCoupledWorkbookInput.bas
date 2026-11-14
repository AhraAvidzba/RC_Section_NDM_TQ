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

    Dim steelClass As String, layout As CTransverseRebarLayout
    If section.RebarCount > 0 Then steelClass = settings.GetRequiredString("Coupled.SteelClass")
    Set layout = ReadSavedRecipeTransverse(workbook, section, units, steelClass)
    If layout.Count > 0 And Len(steelClass) = 0 Then steelClass = settings.GetRequiredString("Coupled.SteelClass")
    ValidateSavedMaterialClasses section, layout, steelClass, settings.InputLocation("Coupled.SteelClass")

    Dim analysis As New CCoupledSectionAnalysis
    analysis.Initialize section, layout, provider, poisson, gt, gc, length, yieldRatio, peak, steelClass, _
        settings.GetRequiredLong("Solver.MaxIterations")
    batch.ConfigureCoupledAnalysis analysis
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Description
    batch.ConfigureCoupledAnalysis Nothing, reason
End Sub

' Находит блок Transverse для фактически сохранённого типа сечения.
' Старый снимок без метки нужно сформировать средствами текущей книги.
Private Function ReadSavedRecipeTransverse(ByVal workbook As Object, ByVal section As CSectionModel, _
        ByVal units As CUnitSystem, ByVal steelClass As String) As CTransverseRebarLayout
    If StrComp(section.SourceType, "AutoCADImport", vbTextCompare) = 0 Then _
        Err.Raise vbObjectError + 4194, "ReadSavedRecipeTransverse", _
            "Для сечения из AutoCAD пока нет источника данных поперечного армирования. Проверьте предпосылки связанного расчёта."
    If Len(section.RecipeType) = 0 Then _
        Err.Raise vbObjectError + 4194, "ReadSavedRecipeTransverse", _
            "В сохранённой геометрии отсутствует тип рецепта (rngNDMSectionProperties). Нажмите «Сформировать сечение»."

    Dim registry As New CSectionTypeRegistry, tables As Variant, name As Variant
    tables = registry.RecipeTableNames(section.RecipeType)
    Dim reader As New CSectionInputSnapshotReader, context As CSectionRecipeContext
    Set context = reader.Capture(workbook, tables, units)
    Dim sectionInput As CSectionInputReader, layout As CTransverseRebarLayout
    For Each name In tables
        Set sectionInput = context.Reader(CStr(name))
        Set layout = sectionInput.ReadTransverseLayout(section, steelClass)
        If layout.IsSpecified Then Set ReadSavedRecipeTransverse = layout: Exit Function
    Next name
    Err.Raise vbObjectError + 4194, "ReadSavedRecipeTransverse", _
        "В блоке Transverse на листе «Сечения» укажите поперечную арматуру или выберите «Нет» (" & Join(tables, ", ") & ")."
End Function

' Сверяет материалы сохранённой геометрии с фактическими калибровками.
' Один блок Config задаёт один бетон и один класс двухлинейной стали.
Private Sub ValidateSavedMaterialClasses(ByVal section As CSectionModel, ByVal layout As CTransverseRebarLayout, _
        ByVal steelClass As String, ByVal classLocation As String)
    Dim i As Long, firstConcrete As Long, material As String
    firstConcrete = section.ConcreteMaterialID(1)
    For i = 2 To section.ConcreteCount
        If section.ConcreteMaterialID(i) <> firstConcrete Then _
            Err.Raise vbObjectError + 4194, "ValidateSavedMaterialClasses", _
                "Для связанного расчёта задана одна калибровка бетона, а в сохранённом сечении несколько ID материалов (rngNDMSectionGeometry, элемент " & section.ConcreteID(i) & ")."
    Next i
    For i = 1 To section.RebarCount
        material = section.RebarSteelClass(i)
        If StrComp(material, steelClass, vbTextCompare) <> 0 Then _
            Err.Raise vbObjectError + 4194, "ValidateSavedMaterialClasses", _
                "Класс продольного стержня «" & section.RebarID(i) & "» (rngNDMSectionGeometry) — «" & material & _
                "». Для него нужны согласованные параметры стали (" & classLocation & ")."
    Next i
    For i = 1 To layout.Count
        material = layout.SteelClass(i)
        If StrComp(material, steelClass, vbTextCompare) <> 0 Then _
            Err.Raise vbObjectError + 4194, "ValidateSavedMaterialClasses", _
                "Для поперечной арматуры класса «" & material & "» (" & layout.SourceReference(i) & _
                ") нужны согласованные параметры стали (" & classLocation & ")."
    Next i
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
