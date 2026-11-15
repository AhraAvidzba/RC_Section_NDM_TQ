Attribute VB_Name = "modTestCoupledRecipeWorkbook"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Полный маршрут Form и Run с собственным блоком Transverse рецепта.
' Запускается только в отдельной книге, использует малые нагрузки и обычный ввод.

Private Const CHECK_ID As String = "SpatialRecipe"
Private Const CHECK_PROFILE As String = "PR1"
Private Const OFFSET_X As Double = 3#, OFFSET_Y As Double = -4#
Private Const MAX_EQUILIBRIUM_MPA As Double = 0.000002

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет оба обязательных для этого сценария рецепта либо выбранный рецепт отдельно.
Public Function RunTestCoupledRecipeWorkbook(ByVal storage As String, Optional ByVal selection As String = "Required") As String
    Dim passed As Long, report As String
    On Error GoTo Failed
    Dim book As Workbook: Set book = ThisWorkbook
    If InStr(1, book.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestCoupledRecipeWorkbook", "Запустите проверку из изолированной копии книги."
    If storage <> "Excel" And storage <> "File" Then _
        Err.Raise 5, "RunTestCoupledRecipeWorkbook", "Выберите Excel или File."
    Dim recipes As Variant
    Select Case selection
        Case "Required": recipes = Array("PointedSection", "HollowRectangle")
        Case "All": recipes = Array("PointedSection", "HollowRectangle", "Circle")
        Case "PointedSection", "HollowRectangle", "Circle": recipes = Array(selection)
        Case Else: Err.Raise 5, "RunTestCoupledRecipeWorkbook", "Неизвестный набор рецептов."
    End Select
    Dim recipe As Variant
    For Each recipe In recipes
        TestRunSpatialRecipe book, CStr(recipe), storage, passed, report
    Next recipe
    RunTestCoupledRecipeWorkbook = report & "TOTAL_COUPLED_RECIPE_WORKBOOK: passed=" & CStr(passed) & "; failed=0"
    Exit Function
Failed:
    RunTestCoupledRecipeWorkbook = report & "FAIL: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_COUPLED_RECIPE_WORKBOOK: passed=" & CStr(passed) & "; failed=1"
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает реальные именованные входы и выполняет обычные команды книги.
Private Sub TestRunSpatialRecipe(ByVal book As Workbook, ByVal recipe As String, ByVal storage As String, _
        ByRef passed As Long, ByRef report As String)
    Dim label As String: label = recipe & " | " & storage & " | "
    modTestPerformance.TestConfigurePerformanceFixture "NDS", 1, "Newton", recipe, "Normal", "No"
    Dim settings As Range, profiles As Range, calibration As Range, grades As Range, loads As Range
    Set settings = book.Names("rngSystemSettings").RefersToRange
    Set profiles = book.Names("rngCalculationProfiles").RefersToRange
    Set calibration = book.Names("rngCoupledMaterialSettings").RefersToRange
    Set grades = book.Names("rngCoupledSteelClasses").RefersToRange
    Set loads = book.Names("rngLoadCombinations").RefersToRange
    TestSpatialRecipeCell(settings, "Results.SnapshotStorage").Value2 = storage
    TestSpatialRecipeCell(settings, "Results.SnapshotName").Value2 = "SpatialRecipe_" & recipe & "_" & storage
    TestSpatialRecipeCell(settings, "General.ExecutionReportEnabled").Value2 = "No"
    TestSpatialRecipeCell(settings, "Mesh.CellGeometry").Value2 = "Contour"
    TestSpatialRecipeCell(settings, "Mesh.StepMode").Value2 = "Auto 0.2" ' Крупный допустимый шаг для проверки маршрута и точных границ.
    TestSpatialRecipeCell(settings, "Plot.Enabled").Value2 = "Yes"
    TestSpatialRecipeCell(settings, "Plot.AutoUpdateAfterCalculation").Value2 = "Yes"
    TestSpatialRecipeCell(settings, "Output.CombinationMode").Value2 = "Worst"
    TestSpatialRecipeCell(profiles, "MaterialModel.Strength.ValueSet", 2, 3).Value2 = "ULS(I)"
    TestSpatialRecipeCell(profiles, "MaterialModel.Strength.ConcreteTension", 2, 3).Value2 = "UseDiagram"
    TestSpatialRecipeCell(profiles, "MaterialModel.Strength.SteelDiagram", 2, 3).Value2 = "TwoLine"
    TestSpatialRecipeCell(profiles, "Visualization.State", 2, 3).Value2 = "StrengthState"
    TestSpatialRecipeCell(profiles, "Visualization.Quantity", 2, 3).Value2 = "TauXZ"
    TestSpatialRecipeCell(calibration, "Coupled.Poisson").Value2 = 0.2
    TestSpatialRecipeCell(calibration, "Coupled.Gt").Value2 = 0.1
    TestSpatialRecipeCell(calibration, "Coupled.Gc").Value2 = 10#
    TestSpatialRecipeCell(calibration, "Coupled.Localization").Value2 = 100#
    TestSpatialRecipeCell(calibration, "Coupled.YieldRatio").Value2 = 0.6
    TestSpatialRecipeCell(calibration, "Coupled.PeakStrain").Value2 = 0.002

    Dim unitSettings As New CSystemSettingsReader, units As New CUnitSystem
    unitSettings.LoadFromWorkbook book: units.LoadFromSettings unitSettings
    Dim lengthInput As Double: lengthInput = units.InputLengthToInternal(1#)
    TestSpatialRecipeCell(settings, "Load.ReferenceOffsetX").Value2 = OFFSET_X / lengthInput
    TestSpatialRecipeCell(settings, "Load.ReferenceOffsetY").Value2 = OFFSET_Y / lengthInput
    Dim target As Variant: target = Array(-2500#, 30000#, -44000#, 120000#, 180#, -360#) ' Н, Н*мм; поперечные силы и перенос не коллинеарны, поэтому добавка к T ненулевая.
    loads.Cells(2, 1).Value2 = CHECK_ID
    loads.Cells(2, 2).Value2 = target(0) / units.InputForceToInternal(1#)
    loads.Cells(2, 3).Value2 = target(1) / units.InputMomentMxToInternal(1#)
    loads.Cells(2, 4).Value2 = target(2) / units.InputMomentMyToInternal(1#)
    loads.Cells(2, 5).Value2 = CHECK_PROFILE
    loads.Cells(2, 6).Value2 = "Auto"
    loads.Cells(2, 8).Value2 = target(3) / units.InputTorqueToInternal(1#)
    loads.Cells(2, 9).Value2 = target(4) / units.InputShearToInternal(1#)
    loads.Cells(2, 10).Value2 = target(5) / units.InputShearToInternal(1#)

    Dim geometry As Range, scheme As Range
    Set geometry = book.Names("rng" & recipe & "Geometry").RefersToRange
    TestPrepareSpatialRecipeGeometry geometry, recipe, lengthInput
    Set scheme = TestSpatialRecipeCell(geometry, "Group1")
    scheme.Resize(3, 8).ClearContents ' Машинные ключи строк остаются на месте.
    Dim started As Double: started = Timer
    Dim attention As Boolean, message As String
    message = FormSectionForWorkbook(book, requiresAttention:=attention)
    report = report & "TIMING_FORM_SECONDS: " & label & TestSpatialRecipeTime(started) & vbCrLf
    If attention Then Err.Raise 5, "FormSectionForWorkbook", label & Replace(message, vbCrLf, " ")
    Dim section As CSectionModel: Set section = ReadFormedSectionFromResults(book, includeMaterialDetails:=True)
    TestRequireSpatialRecipe section.RecipeType = recipe, label & "тип восстановлен из сформированного снимка", passed, report
    TestRequireSpatialRecipe section.RebarCount = 4, label & "сформировано четыре продольных стержня", passed, report
    TestSpatialRecipeContour section, recipe, label, passed, report
    Dim primaryClass As String: primaryClass = "A400"
    If section.RebarCount > 0 Then primaryClass = section.RebarSteelClassForMaterial(1, primaryClass)
    TestSpatialRecipeCell(calibration, "Coupled.SteelClass").Value2 = primaryClass
    Dim grade As String: grade = "Check_" & recipe
    grades.Cells(2, 1).Value2 = grade: grades.Cells(2, 2).Value2 = "ULS(I)"
    grades.Cells(2, 3).Value2 = 200000#: grades.Cells(2, 4).Value2 = 240#: grades.Cells(2, 5).Value2 = 240#
    grades.Cells(2, 6).Value2 = 0.02: grades.Cells(2, 7).Value2 = 0.02
    Dim expectedLength As Double
    expectedLength = TestSetSpatialRecipeArmature(section, geometry, scheme, recipe, lengthInput, grade)

    started = Timer
    message = RunSectionCalculationForWorkbook(book)
    report = report & "TIMING_CALC_SECONDS: " & label & TestSpatialRecipeTime(started) & vbCrLf
    TestRequireSpatialRecipe ReadLoadCasePropertyText(book, CHECK_ID, "DirectStateStatus") = "OK", _
        label & "получено прямое 6D равновесие " & Replace(message, vbCrLf, " "), passed, report
    TestRequireSpatialRecipe ReadLoadCasePropertyText(book, "ALL", "Geometry.RecipeType") = recipe And _
        ReadLoadCasePropertyText(book, CHECK_ID, "Coupled.StateSolver") = "SplineNewton", _
        label & "сохранены рецепт и фактический решатель", passed, report
    TestRequireSpatialRecipe ReadLoadCasePropertyText(book, CHECK_ID, "Coupled.Transverse.Status") = "OK", _
        label & "хомут имеет собственный статус", passed, report
    TestRequireSpatialRecipe InStr(1, ReadLoadCasePropertyText(book, CHECK_ID, "Coupled.Transverse.Control"), grade, vbBinaryCompare) > 0, _
        label & "в расчёте применён класс из блока выбранного рецепта", passed, report
    TestSpatialRecipeWrench book, target, units, label, passed, report

    Dim rows As Variant: rows = ReadElementResultTable(book, CHECK_ID, "StrengthState")
    TestRequireSpatialRecipe HasResultTableRows(rows), label & "активный носитель отдаёт результаты элементов", passed, report
    If storage = "File" Then
        TestRequireSpatialRecipe Not HasResultTableRows(ReadPhysicalResultTable(book, "rngNDMElementResults")), _
            label & "подробности читаются из .ndm", passed, report
    End If
    Dim runID As Double: runID = CDbl(rows(2, ResultColumn(rows, "RunID")))
    Dim paths As Collection: Set paths = ReadSavedTransversePaths(book, runID, CHECK_ID, CHECK_PROFILE, "StrengthState")
    TestRequireSpatialRecipe Not paths Is Nothing, label & "найден сохранённый путь своего Transverse", passed, report
    TestRequireSpatialRecipe paths.Count = 1, label & "найдена только одна принятая группа", passed, report
    Dim path As CGeometryPath: Set path = paths(1)
    TestRequireSpatialRecipe Abs(path.Length - expectedLength) <= 0.000001 And path.IsClosed = (recipe = "Circle"), _
        label & "сохранены длина и замкнутость", passed, report
    Dim values As Object: Set values = ReadCoupledFieldValues(book, vqTauXZ, runID, CHECK_ID, CHECK_PROFILE, "StrengthState")
    TestRequireSpatialRecipe values.Count = section.ConcreteCount + section.RebarCount, _
        label & "доступно полное сохранённое касательное поле", passed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Задаёт размеры в миллиметрах и четыре стержня одного ряда через видимые входные таблицы.
Private Sub TestPrepareSpatialRecipeGeometry(ByVal geometry As Range, ByVal recipe As String, ByVal lengthInput As Double)
    Dim keys As Variant, sizes As Variant, faces As Variant, primary As String, i As Long
    Select Case recipe
        Case "PointedSection"
            keys = Array("B", "H", "Left.W", "Left.R1", "Left.R2", "Right.W", "Right.R1", "Right.R2")
            sizes = Array(800#, 250#, 0#, 100#, 0#, 150#, 40#, 50#)
            faces = Array("Top", "Bottom", "Left", "Right")
        Case "HollowRectangle"
            keys = Array("B", "H", "R", "Bo", "Ho", "Ro", "Xo", "Yo")
            sizes = Array(500#, 800#, 180#, 200#, 500#, 30#, 0#, 0#)
            faces = Array("Top", "Bottom", "Left", "Right", "OpeningTop", "OpeningBottom", "OpeningLeft", "OpeningRight")
        Case "Circle"
            keys = Array("D"): sizes = Array(300#): faces = Array("Circle")
    End Select
    For i = 0 To UBound(keys)
        TestSpatialRecipeCell(geometry, CStr(keys(i)), 1, 3).Value2 = sizes(i) / lengthInput
    Next i
    If recipe <> "Circle" Then TestSpatialRecipeCell(geometry, "CornerMode", 1, 3).Value2 = "Скругление"

    Dim main As Range, extra As Range, inputRow As Range
    Set main = TestSpatialRecipeBlock(geometry, "Main")
    Set extra = TestSpatialRecipeBlock(geometry, "Extra")
    primary = CStr(faces(0))
    For i = 0 To UBound(faces)
        Set inputRow = TestSpatialRecipeCell(main, CStr(faces(i)), 1, 3)
        inputRow.Value2 = 40# / lengthInput ' Расстояние до оси всех сторон участвует в построении общего опорного пути.
        inputRow.Offset(0, 1).Value2 = 0# ' Диаметр ноль отключает сторону, в том числе проецируемый ряд отверстия.
        If i < 4 Then inputRow.Offset(0, 2).Value2 = 0 ' Для сторон отверстия n не применяется.
    Next i
    Set inputRow = TestSpatialRecipeCell(main, primary, 1, 3)
    inputRow.Offset(0, 1).Value2 = 12# / lengthInput
    inputRow.Offset(0, 2).Value2 = 4
    inputRow.Offset(0, 3).Value2 = 0#: inputRow.Offset(0, 4).Value2 = 0#
    If recipe = "Circle" Then
        inputRow.Offset(0, 5).Value2 = "IncludeStart" ' Четыре различные точки на замкнутом пути.
    Else
        inputRow.Offset(0, 5).Value2 = "ExcludeBoth"
    End If
    Set inputRow = TestSpatialRecipeCell(extra, primary, 1, 3)
    inputRow.ClearContents: inputRow.Offset(0, 4).ClearContents ' d2 и d3 отключают дополнительные ряды.
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выделяет данные блока между машинными ключами, пропуская заголовки и строку единиц.
Private Function TestSpatialRecipeBlock(ByVal geometry As Range, ByVal key As String) As Range
    Dim header As Range, ending As Range
    Set header = TestSpatialRecipeCell(geometry, key, 1, 1)
    If key = "Main" Then
        Set ending = TestSpatialRecipeCell(geometry, "Extra", 1, 1)
    Else
        Set ending = TestSpatialRecipeCell(geometry, "Transverse", 1, 1)
    End If
    Set TestSpatialRecipeBlock = header.Resize(ending.Row - header.Row, geometry.Columns.Count)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет реальные границы рецепта и формулы площадей круга и пустотелого прямоугольника.
Private Sub TestSpatialRecipeContour(ByVal section As CSectionModel, ByVal recipe As String, _
        ByVal label As String, ByRef passed As Long, ByRef report As String)
    Dim count As Long, hasCurve As Boolean, opening As Boolean, i As Long
    count = section.Contours.Count
    For i = 1 To count
        hasCurve = hasCurve Or (section.Contours.SegmentType(i) <> "Line")
        opening = opening Or (section.Contours.LoopRole(i) = "Opening")
    Next i
    TestRequireSpatialRecipe count > 0 And hasCurve And opening = (recipe = "HollowRectangle"), _
        label & "сохранены реальные дуги и принадлежность отверстия", passed, report
    Dim area As Double, cx As Double, cy As Double, expectedArea As Double
    section.ConcreteMeshCentroid cx, cy, area
    Select Case recipe
        Case "Circle"
            expectedArea = GEOM_PI * 150# ^ 2
        Case "HollowRectangle"
            expectedArea = 500# * 800# - 200# * 500# - (4# - GEOM_PI) * (180# ^ 2 - 30# ^ 2) ' Из каждой площадки четырёх углов вычитается четверть круга.
        Case Else
            TestRequireSpatialRecipe area > 0#, label & "площадь сформированной сетки положительна", passed, report
            Exit Sub
    End Select
    TestRequireSpatialRecipe Abs(area - expectedArea) <= 0.000002 * expectedArea, _
        label & "площадь по независимой геометрической формуле", passed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Круг получает хомут по контуру, остальные рецепты короткую ветвь внутри бетона.
Private Function TestSetSpatialRecipeArmature(ByVal section As CSectionModel, ByVal geometry As Range, ByVal scheme As Range, _
        ByVal recipe As String, ByVal lengthInput As Double, ByVal grade As String) As Double
    scheme.Offset(0, 1).Value2 = 8# / lengthInput
    scheme.Offset(0, 2).Value2 = 120# / lengthInput
    scheme.Offset(0, 4).Value2 = grade
    If recipe = "Circle" Then
        scheme.Value2 = "Контур": scheme.Offset(0, 3).Value2 = 10# / lengthInput
        TestSetSpatialRecipeArmature = GEOM_PI * (TestSpatialRecipeGeometryValue(geometry, "D") * lengthInput - 28#)
        Exit Function
    End If
    Dim query As New CSectionGeometryQuery, region As CConcreteRegion, i As Long, x As Double, y As Double
    query.Initialize section: Set region = query.ConcreteDomain
    For i = 1 To section.ConcreteCount
        x = section.ConcreteX(i): y = section.ConcreteY(i)
        If query.BoundaryDistance(region, x, y) > 20# Then Exit For
    Next i
    If i > section.ConcreteCount Then Err.Raise 5, "TestSetSpatialRecipeArmature", "Нет внутренней точки для направленной ветви 20 мм."
    scheme.Value2 = "Ветвь"
    scheme.Offset(0, 6).Value2 = CStr((x - 10#) / lengthInput) & ";" & CStr(y / lengthInput) & ";" & _
        CStr((x + 10#) / lengthInput) & ";" & CStr(y / lengthInput)
    TestSetSpatialRecipeArmature = 20#
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет восстановленные решателем шесть интегралов с независимо перенесённым заданием.
Private Sub TestSpatialRecipeWrench(ByVal book As Workbook, ByVal target As Variant, ByVal units As CUnitSystem, _
        ByVal label As String, ByRef passed As Long, ByRef report As String)
    Dim expected As Variant, names As Variant, stem As String
    names = Array("N", "Mx", "My", "T", "Qx", "Qy")
    expected = Array(units.InternalForceToOutput(target(0)), units.InternalMomentMxToOutput(target(1) + OFFSET_Y * target(0)), _
        units.InternalMomentMyToOutput(target(2) + OFFSET_X * target(0)), _
        units.InternalTorqueToOutput(target(3) + OFFSET_X * target(5) - OFFSET_Y * target(4)), _
        units.InternalShearToOutput(target(4)), units.InternalShearToOutput(target(5)))
    stem = "State.StrengthState.Equilibrium."
    Dim area As Double, lever As Double, forceScale As Double, momentScale As Double
    area = units.OutputAreaToInternal(CDbl(ReadLoadCasePropertyText(book, CHECK_ID, stem & "ResidualArea")))
    lever = units.OutputLengthToInternal(CDbl(ReadLoadCasePropertyText(book, CHECK_ID, stem & "ResidualLength")))
    TestRequireSpatialRecipe area > 0# And lever > 0# And ReadLoadCasePropertyText(book, CHECK_ID, stem & "Reference") = "ConcreteMeshCentroid", _
        label & "однозначно указана точка приведения", passed, report
    forceScale = units.InternalShearToOutput(1#) * area
    momentScale = units.InternalTorqueToOutput(1#) * area * lever
    Dim i As Long, actual As Double, requested As Double, normalizer As Double, delta As Double
    For i = 0 To 5
        actual = CDbl(ReadLoadCasePropertyText(book, CHECK_ID, stem & "Restored." & names(i)))
        requested = CDbl(ReadLoadCasePropertyText(book, CHECK_ID, stem & "Target." & names(i)))
        normalizer = forceScale
        If i >= 1 And i <= 3 Then normalizer = momentScale
        delta = Abs(actual - expected(i)) / normalizer
        TestRequireSpatialRecipe delta <= MAX_EQUILIBRIUM_MPA And Abs(requested - expected(i)) <= 0.000000001 * GeomMax(1#, Abs(expected(i))), _
            label & CStr(names(i)) & " | невязка в МПа " & CStr(delta), passed, report
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает размер из столбца val блока General после пользовательской подписи.
Private Function TestSpatialRecipeGeometryValue(ByVal geometry As Range, ByVal key As String) As Double
    TestSpatialRecipeGeometryValue = CDbl(TestSpatialRecipeCell(geometry, key, 1, 3).Value2)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает значение по смысловому ключу, не фиксируя координаты таблицы.
Private Function TestSpatialRecipeCell(ByVal table As Range, ByVal key As String, Optional ByVal keyColumn As Long = 1, _
        Optional ByVal valueColumn As Long = 2) As Range
    Dim row As Long
    For row = 1 To table.Rows.Count
        If StrComp(CStr(table.Cells(row, keyColumn).Value2), key, vbBinaryCompare) = 0 Then
            Set TestSpatialRecipeCell = table.Cells(row, valueColumn)
            Exit Function
        End If
    Next row
    Err.Raise 5, "TestSpatialRecipeCell", "Не найден ключ " & key & " в " & table.Address(External:=True) & "."
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает одну проверку и оставляет точную причину первой ошибки.
Private Sub TestRequireSpatialRecipe(ByVal agreed As Boolean, ByVal text As String, ByRef passed As Long, ByRef report As String)
    If Not agreed Then Err.Raise 5, "TestRequireSpatialRecipe", text
    passed = passed + 1
    report = report & "PASS: " & text & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает полночь в локальных независимых замерах Form и Calc.
Private Function TestSpatialRecipeTime(ByVal started As Double) As String
    Dim seconds As Double: seconds = Timer - started
    If seconds < 0# Then seconds = seconds + 86400#
    TestSpatialRecipeTime = CStr(seconds)
End Function
