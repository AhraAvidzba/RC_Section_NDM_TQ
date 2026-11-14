Attribute VB_Name = "modTestCoupledTransverseRebar"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: геометрия и взвешивание поперечной арматуры
' ==========================================================================
' Проверяет полный путь относительно отверстий и наружной границы,
' сохранение фактической длины, шага и направлений в квадратуре.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает интегралы прямого стержня и замкнутого кругового хомута
' с аналитическими значениями, проверяет диагностические отказы.
Public Function RunCoupledTransverseRebarTests() As String
    On Error GoTo Failed
    Dim report As String, passed As Long, failed As Long
    Dim regionQuery As New CSectionGeometryQuery
    Dim region As CConcreteRegion, inner As CConcreteRegion
    Set region = regionQuery.RectangleRegion(0#, 0#, 200#, 160#)
    Set inner = regionQuery.CircleRegion(0#, 0#, 22#)
    Set region = regionQuery.SubtractContainedRegion(region, inner)
    Dim geometry As New CGeometryRegion, mesh As New CFiberMeshBuilder
    geometry.Initialize region
    mesh.BuildMesh geometry, 12#, 12#, cellGeometry:="Contour"
    Dim builder As New CSectionModelBuilder, section As CSectionModel
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    geometry.BuildContours section.Contours

    Dim quadrature As New CCoupledSectionQuadrature, bars As New CTransverseRebarQuadrature
    quadrature.Initialize section, 4, 3
    Dim layout As New CTransverseRebarLayout
    TestExpect Not layout.IsSpecified, "empty input remains unknown", passed, failed, report
    TestExpect InStr(TestFailure(bars, layout, quadrature), "подтвердите") > 0, _
        "unknown does not become absence", passed, failed, report
    layout.DeclareNone
    bars.Prepare layout, quadrature
    TestExpect layout.IsSpecified And bars.Count = 0, "explicit absence", passed, failed, report

    Set layout = New CTransverseRebarLayout
    Dim rectangleLoop As New CSectionDefinitionBuilder, path As CGeometryPath
    rectangleLoop.Initialize "TestTransverse"
    Set path = rectangleLoop.CirclePath("H1", 0#, 0#, 56#)
    layout.AddGroup path, 8#, 85#, "A400", "rngTransverseRebar H1"
    Set path = rectangleLoop.LinePath("B1", -80#, 60#, 80#, 60#)
    layout.AddGroup path, 10#, 140#, "A400", "rngTransverseRebar B1"
    bars.Prepare layout, quadrature

    Dim value(1 To 3, 1 To 2) As Double
    Dim i As Long, group As Long, x As Double, y As Double, area As Double, tx As Double, ty As Double, txy As Double
    For i = 1 To bars.Count
        bars.Point i, x, y, area, tx, ty, txy, group
        value(1, group) = value(1, group) + area * tx
        value(2, group) = value(2, group) + area * ty
        value(3, group) = value(3, group) + area * txy
    Next i
    Dim expected As Double
    expected = GEOM_PI * 16# * (2# * GEOM_PI * 56#) / 85# / 2#
    TestExpect Abs(value(1, 1) / expected - 1#) < 1E-4 And Abs(value(2, 1) / expected - 1#) < 1E-4, _
        "circle integrated directional weights", passed, failed, report
    TestExpect Abs(value(3, 1)) < 1E-4 * expected, "circle mixed term", passed, failed, report
    expected = GEOM_PI * 25# * 160# / 140#
    TestExpect Abs(value(1, 2) / expected - 1#) < 1E-12 And value(2, 2) = 0# And value(3, 2) = 0#, _
        "branch equivalent volume per length", passed, failed, report

    Set layout = New CTransverseRebarLayout
    Set path = rectangleLoop.LinePath("B2", -80#, 0#, 80#, 0#)
    layout.AddGroup path, 8#, 125#, "A400", "rngTransverseRebar B2"
    TestExpect InStr(TestFailure(bars, layout, quadrature), "rngTransverseRebar B2") > 0 And bars.Count = 0, _
        "opening crossing fails with source row", passed, failed, report

    Set layout = New CTransverseRebarLayout
    Set path = rectangleLoop.LinePath("B3", -80#, 24#, 80#, 24#)
    layout.AddGroup path, 8#, 125#, "A400"
    TestExpect Len(TestFailure(bars, layout, quadrature)) > 0, _
        "bar diameter intersects hole although axis stays in concrete", passed, failed, report

    Set layout = New CTransverseRebarLayout
    Set path = rectangleLoop.CirclePath("H2", 0#, 0#, 24#)
    layout.AddGroup path, 8#, 125#, "A400"
    TestExpect Len(TestFailure(bars, layout, quadrature)) > 0, _
        "concentric arc too close to hole", passed, failed, report
    Set layout = New CTransverseRebarLayout
    layout.AddContourGroup section, "OuterCover", "", 8#, 85#, 12#, "A400"
    Set path = layout.Path(1)
    TestExpect path.IsClosed And Abs(path.Length - 592#) < 1E-8, _
        "outer hoop follows concrete with cover plus radius", passed, failed, report
    Dim openingID As String
    For i = 1 To section.Contours.Count
        If section.Contours.LoopRole(i) = "Opening" Then
            openingID = section.Contours.LoopID(i)
            Exit For
        End If
    Next i
    layout.AddContourGroup section, "OpeningCover", openingID, 8#, 120#, 13#, "A400"
    Set path = layout.Path(2)
    TestExpect path.IsClosed And Abs(path.Length - 78# * GEOM_PI) < 1E-8, _
        "opening hoop offsets into the concrete", passed, failed, report
    bars.Prepare layout, quadrature
    TestExpect bars.Count > 0, "both automatic hoop paths admit their full diameter", passed, failed, report

    RunCoupledTransverseRebarTests = report & "TOTAL_COUPLED_TRANSVERSE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunCoupledTransverseRebarTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает структурированную причину отказа при подготовке арматуры.
Private Function TestFailure(ByVal bars As CTransverseRebarQuadrature, ByVal layout As CTransverseRebarLayout, _
        ByVal quadrature As CCoupledSectionQuadrature) As String
    On Error GoTo Failed
    bars.Prepare layout, quadrature
    Exit Function
Failed:
    TestFailure = Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет в отчёт независимую проверку и обновляет счётчики.
Private Sub TestExpect(ByVal condition As Boolean, ByVal name As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If condition Then
        passed = passed + 1
        report = report & "OK: " & name & vbCrLf
    Else
        failed = failed + 1
        report = report & "FAIL: " & name & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает четвёртый блок тем же ключевым читателем, что используют рецепты.
' Ячейки с ошибками в неприменимых полях специально не читаются.
Public Function RunCoupledTransverseInputTests() As String
    On Error GoTo Failed
    Dim values(1 To 8, 1 To 11) As Variant, headers As Variant, i As Long
    Dim section As CSectionModel, geometry As CGeometryRegion, mesh As New CFiberMeshBuilder
    Set geometry = TestFixtureDisk(200#)
    mesh.BuildMesh geometry, 20#, 20#, cellGeometry:="Contour"
    Dim builder As New CSectionModelBuilder
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    geometry.BuildContours section.Contours
    values(1, 1) = "General": values(3, 1) = "key": values(3, 2) = "val"
    values(4, 1) = "Transverse"
    headers = Array("key", "kind", "d", "spacing", "cover", "steel", "loop", "path", "", "", "note")
    For i = 0 To UBound(headers)
        values(6, i + 1) = headers(i)
    Next i
    values(7, 1) = "Group1": values(8, 1) = "Group2"
    values(7, 3) = CVErr(2015): values(8, 8) = CVErr(2015)

    Dim passed As Long, failed As Long, report As String, result As CTransverseRebarLayout
    Set result = TestReadTransverse(values, section)
    TestExpect Not result.IsSpecified, "unfilled fourth block remains unknown", passed, failed, report
    values(7, 2) = "Нет"
    Set result = TestReadTransverse(values, section)
    TestExpect result.IsSpecified And result.Count = 0, "explicit no steel ignores unused numeric fields", passed, failed, report
    values(7, 2) = "Контур": values(7, 3) = 8#: values(7, 4) = 100#: values(7, 5) = 15#
    values(7, 8) = CVErr(2015)
    Set result = TestReadTransverse(values, section)
    TestExpect result.Count = 1 And result.Path(1).IsClosed And result.SteelClass(1) = "A400", _
        "typical contour uses configured default steel", passed, failed, report
    values(8, 2) = "Ветвь": values(8, 3) = 10#: values(8, 4) = 140#: values(8, 5) = CVErr(2015)
    values(8, 8) = "-80;25;50;25"
    Set result = TestReadTransverse(values, section)
    TestExpect result.Count = 2 And Not result.Path(2).IsClosed And Abs(result.Path(2).Length - 130#) < 1E-9, _
        "one extra branch and unused cover", passed, failed, report
    values(8, 8) = "-80;25;text;25"
    TestExpect InStr(TestReadFailure(values, section), "R8C8") > 0, _
        "bad branch coordinate includes actual source", passed, failed, report
    RunCoupledTransverseInputTests = report & "TOTAL_COUPLED_TRANSVERSE_INPUT: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunCoupledTransverseInputTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт снимок с явными адресами, которые можно проверить независимо от Excel.
Private Function TestReadTransverse(ByRef values As Variant, ByVal section As CSectionModel) As CTransverseRebarLayout
    Dim addresses(1 To 8, 1 To 11) As Variant, row As Long, col As Long
    For row = 1 To 8
        For col = 1 To 11
            addresses(row, col) = "R" & row & "C" & col
        Next col
    Next row
    Dim table As New CSectionInputTable, reader As New CSectionInputReader
    table.Initialize "TestTransverse", values, addresses
    Dim units As New CUnitSystem
    units.InitializeDefaults
    reader.Initialize table, units
    Set TestReadTransverse = reader.ReadTransverseLayout(section, "A400")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает причину неверного ввода для сверки адреса ячейки.
Private Function TestReadFailure(ByRef values As Variant, ByVal section As CSectionModel) As String
    On Error GoTo Failed
    Dim result As CTransverseRebarLayout
    Set result = TestReadTransverse(values, section)
    Exit Function
Failed:
    TestReadFailure = Err.Description
End Function
