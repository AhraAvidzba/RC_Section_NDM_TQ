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
