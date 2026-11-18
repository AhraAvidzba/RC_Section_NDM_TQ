Attribute VB_Name = "modTestCoupledQuadrature"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: рабочая квадратура Гаусса и дискретизация пространственного поля
' ==========================================================================
' Независимые интегралы полиномов проверяются аналитически. Отверстие
' проверяет сохранение точной площади, центра и пригодности всех узлов.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет 100 двумерных полиномиальных интегралов, область с отверстием и ошибки настроек.
Public Function RunTestCoupledQuadrature() As String
    On Error GoTo Failed
    Dim passed As Long, failed As Long, report As String
    TestRectanglePolynomials passed, failed, report
    TestOpeningQuadrature passed, failed, report
    TestDiscretizationErrors passed, failed, report
    RunTestCoupledQuadrature = report & "TOTAL_COUPLED_QUADRATURE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledQuadrature = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит контрольную модель тем же общим сборщиком сетки и достоверных контуров.
Private Function TestBuildQuadratureSection(ByVal geometry As CGeometryRegion) As CSectionModel
    Dim mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder
    mesh.BuildMesh geometry, geometry.Width / 20#, geometry.Height / 20#, cellGeometry:="Contour"
    Set TestBuildQuadratureSection = builder.BuildFromGenerated(mesh, Nothing)
    geometry.BuildContours TestBuildQuadratureSection.Contours
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает все произведения степеней 0..9 с точными средними по прямоугольнику.
Private Sub TestRectanglePolynomials(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim section As CSectionModel: Set section = TestBuildQuadratureSection(TestFixtureBox(180#, 120#))
    Dim quadrature As New CCoupledSectionQuadrature: quadrature.Initialize section, 2, 2
    Dim x() As Double, y() As Double, area() As Double, indices() As Long, shape() As Double, dx() As Double, dy() As Double
    quadrature.Prepare 5, x, y, area, indices, shape, dx, dy, True
    TestRecord UBound(area) = 100, "four full spans contain 100 Gauss nodes", passed, failed, report
    Dim i As Long, px As Long, py As Long, integral As Double, expected As Double
    For px = 0 To 9
        For py = 0 To 9
            integral = 0#: expected = 0#
            If px Mod 2 = 0 And py Mod 2 = 0 Then expected = 1# / ((px + 1) * (py + 1))
            For i = 1 To UBound(area)
                integral = integral + area(i) / 21600# * (x(i) / 90#) ^ px * (y(i) / 60#) ^ py
            Next i
            TestRecord Abs(integral - expected) < 0.000000000001, "polynomial x^" & px & " y^" & py, passed, failed, report
        Next py
    Next px
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет смешение полных интервалов и контурных ячеек без узлов внутри отверстия.
Private Sub TestOpeningQuadrature(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim query As New CSectionGeometryQuery, region As CConcreteRegion, geometry As CGeometryRegion
    Set region = query.SubtractContainedRegion(query.RectangleRegion(0#, 0#, 180#, 120#), query.RectangleRegion(13#, -20#, 4#, 4#))
    Set geometry = TestFixtureGeometry(region)
    Dim section As CSectionModel: Set section = TestBuildQuadratureSection(geometry)
    Dim quadrature As New CCoupledSectionQuadrature: quadrature.Initialize section, 2, 2
    Dim x() As Double, y() As Double, area() As Double, indices() As Long, shape() As Double, dx() As Double, dy() As Double
    quadrature.Prepare 5, x, y, area, indices, shape, dx, dy, True
    Dim i As Long, a As Double, firstX As Double, firstY As Double, inside As Boolean, worldX As Double, worldY As Double
    inside = True
    For i = 1 To UBound(area)
        worldX = x(i) + quadrature.Frame.OriginX: worldY = y(i) + quadrature.Frame.OriginY
        inside = inside And geometry.ContainsPoint(worldX, worldY) And area(i) > 0#
        a = a + area(i): firstX = firstX + worldX * area(i): firstY = firstY + worldY * area(i)
    Next i
    TestRecord inside, "all mixed quadrature nodes remain inside actual concrete", passed, failed, report
    TestRecord Abs(a - 21584#) < 0.0000001, "opening exact area", passed, failed, report
    TestRecord Abs(firstX + 208#) < 0.000001 And Abs(firstY - 320#) < 0.000001, _
        "opening exact first moments", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет непротиворечивые параметры и неизменяемость кэширующихся настроек.
Private Sub TestDiscretizationErrors(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim parameters As New CCoupledDiscretization, number As Long
    On Error Resume Next
    parameters.Initialize 2, 2, 4, True: number = Err.Number: Err.Clear
    On Error GoTo 0
    TestRecord number = 5 And Not parameters.IsReady, "invalid Gauss order is rejected", passed, failed, report
    parameters.Initialize 2, 2, 5, True
    TestRecord parameters.IsReady And parameters.SpansX = 2 And parameters.SpansY = 2 And _
        parameters.PointsPerSpan = 5 And parameters.UseFullSpanGauss5, "prepared settings preserve all inputs", passed, failed, report
    On Error Resume Next
    parameters.Initialize 1, 1, 5, False: number = Err.Number: Err.Clear
    On Error GoTo 0
    TestRecord number = 5 And parameters.SpansX = 2 And parameters.UseFullSpanGauss5, _
        "prepared settings cannot change under operator cache", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет именованные результаты без изменения ожиданий после неудачи.
Private Sub TestRecord(ByVal ok As Boolean, ByVal name As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    report = report & IIf(ok, "PASS ", "FAIL ") & name & vbCrLf
End Sub
