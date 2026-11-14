Attribute VB_Name = "modTestCoupledElasticState"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: аналитические контрольные случаи упругой задачи
' ==========================================================================
' Интегралы сравниваются отдельно от контроля решения дискретных уравнений.
' Формулы для круга и кольца не используются рабочим расчётным классом.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет шесть усилий на заданной нагрузке, закручивание Сен-Венана
' и поле от перерезывающей силы в круге, затем повторяет кручение для кольца.
Public Function RunCoupledElasticStateTests() As String
    On Error GoTo Failed
    Dim solver As CCoupledElasticState
    Set solver = New CCoupledElasticState
    Dim geometry As CGeometryRegion
    Dim section As CSectionModel
    Dim deformation() As Double, resultants() As Double, stress() As Double
    Dim passed As Long, failed As Long, report As String
    Dim ok As Boolean
    Const RADIUS As Double = 100#
    Const E_MODULUS As Double = 32000#
    Const SHEAR_MODULUS As Double = 16000#
    Const MAX_RELATIVE_ERROR As Double = 0.015
    Const PURE_SHEAR_FORCE As Double = 10000#
    Const TORQUE_VALUE As Double = 2500000#
    Dim area As Double, inertia As Double, torsion As Double

    Set geometry = TestFixtureDisk(2# * RADIUS)
    Set section = TestBuildSection(geometry)
    area = GEOM_PI * RADIUS ^ 2
    inertia = GEOM_PI * RADIUS ^ 4 / 4#
    torsion = 2# * inertia
    solver.PrepareModel section, E_MODULUS, 0#, 2, 2, 9
    ok = solver.Solve(10000#, 2000000#, -1600000#, TORQUE_VALUE, 0#, 0#, deformation, resultants, stress)
    TestRecordSolve ok, "circle axial bending torsion", solver, passed, failed, report
    If ok Then
        TestExpectNear deformation(1), 10000# / (E_MODULUS * area), 1E-8, "circle axial strain", passed, failed, report
        TestExpectNear deformation(2), 2000000# / (E_MODULUS * inertia), MAX_RELATIVE_ERROR, _
            "circle x curvature", passed, failed, report
        TestExpectNear deformation(3), -1600000# / (E_MODULUS * inertia), MAX_RELATIVE_ERROR, _
            "circle y curvature", passed, failed, report
        TestExpectNear deformation(4), TORQUE_VALUE / (SHEAR_MODULUS * torsion), MAX_RELATIVE_ERROR, _
            "circle twist", passed, failed, report
    End If

    ok = solver.Solve(0#, 0#, 0#, 0#, PURE_SHEAR_FORCE, 0#, deformation, resultants, stress)
    TestRecordSolve ok, "circle shear", solver, passed, failed, report
    If ok Then
        Dim i As Long, x As Double, y As Double, weight As Double
        Dim referenceX As Double, referenceY As Double, errorNorm As Double, fieldNorm As Double
        For i = 1 To UBound(stress, 2)
            solver.Point i, x, y, weight
            referenceX = PURE_SHEAR_FORCE * (3# * RADIUS ^ 2 - 3# * x ^ 2 - y ^ 2) / (8# * inertia)
            referenceY = -PURE_SHEAR_FORCE * x * y / (4# * inertia)
            errorNorm = errorNorm + weight * ((stress(5, i) - referenceX) ^ 2 + (stress(6, i) - referenceY) ^ 2)
            fieldNorm = fieldNorm + weight * (referenceX ^ 2 + referenceY ^ 2)
        Next i
        TestRecord Sqr(errorNorm / fieldNorm) < 0.035, "circle shear field relative L2=" & _
            CStr(Sqr(errorNorm / fieldNorm)), passed, failed, report
        TestExpectNear resultants(5), PURE_SHEAR_FORCE, 1E-8, "circle shear integral", passed, failed, report
    End If

    Dim query As New CSectionGeometryQuery
    Dim annulus As CConcreteRegion
    Set annulus = query.SubtractContainedRegion(query.CircleRegion(0#, 0#, RADIUS), query.CircleRegion(0#, 0#, 55#))
    geometry.Initialize annulus
    Set section = TestBuildSection(geometry)
    solver.PrepareModel section, E_MODULUS, 0#, 2, 2, 9
    ok = solver.Solve(0#, 0#, 0#, TORQUE_VALUE, 0#, 0#, deformation, resultants, stress)
    TestRecordSolve ok, "annulus torsion", solver, passed, failed, report
    If ok Then
        torsion = GEOM_PI * (RADIUS ^ 4 - 55# ^ 4) / 2#
        TestExpectNear deformation(4), TORQUE_VALUE / (SHEAR_MODULUS * torsion), MAX_RELATIVE_ERROR, _
            "annulus twist", passed, failed, report
    End If
    TestAsymmetricSixComponents passed, failed, report
    TestReinforcedElasticReplacement passed, failed, report
    RunCoupledElasticStateTests = report & "TOTAL_COUPLED_ELASTIC: passed=" & CStr(passed) & "; failed=" & CStr(failed)
    Exit Function
Failed:
    TestAsymmetricSixComponents passed, failed, report
    TestReinforcedElasticReplacement passed, failed, report
    RunCoupledElasticStateTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет совместную нагрузку на невыпуклом контуре с эксцентричным
' отверстием. Независимые упругие N, Mx, My получены из точных A, I и Ixy.
' Упругая суперпозиция дополнительно проверяется во всех материальных точках.
Private Sub TestAsymmetricSixComponents(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const E_MODULUS As Double = 32000#
    Dim query As New CSectionGeometryQuery, region As CConcreteRegion
    Dim outline As CGeometryPath
    Set outline = query.AdjacentRectangles("AsymmetricSix", Array( _
        Array(-100#, -80#, 200#, 100#), Array(-100#, 20#, 135#, 80#)))
    Set region = query.SubtractContainedRegion(query.RegionFromPath(outline), query.CircleRegion(-40#, -24#, 21#))
    Dim geometry As New CGeometryRegion, section As CSectionModel
    geometry.Initialize region
    Set section = TestBuildSection(geometry)
    Dim solver As New CCoupledElasticState
    solver.PrepareModel section, E_MODULUS, 0#, 3, 3, 9

    Dim combined() As Double, resultants() As Double, stress() As Double
    Dim ok As Boolean, force As Double, mx As Double, my As Double, torque As Double, qx As Double, qy As Double
    force = -52000#: mx = 3100000#: my = -1900000#: torque = 2200000#: qx = 4400#: qy = -3700#
    ok = solver.Solve(force, mx, my, torque, qx, qy, combined, resultants, stress)
    TestRecordSolve ok, "asymmetric perforated all six", solver, passed, failed, report
    If Not ok Then Exit Sub
    Dim properties As TGeometryProperties, originX As Double, originY As Double, meshArea As Double
    query.RegionProperties region, properties
    section.ConcreteMeshCentroid originX, originY, meshArea
    Dim dx As Double, dy As Double, centralMx As Double, centralMy As Double
    dx = properties.CentroidX - originX: dy = properties.CentroidY - originY
    centralMx = mx - force * dy: centralMy = my - force * dx
    Dim determinant As Double, kx As Double, ky As Double, eps0 As Double
    determinant = properties.Ixc * properties.Iyc - properties.Ixyc ^ 2
    kx = (centralMx * properties.Iyc - centralMy * properties.Ixyc) / (E_MODULUS * determinant)
    ky = (centralMy * properties.Ixc - centralMx * properties.Ixyc) / (E_MODULUS * determinant)
    eps0 = force / (E_MODULUS * properties.Area) - kx * dy - ky * dx
    TestExpectNear combined(1), eps0, 0.015, "asymmetric axial strain", passed, failed, report
    TestExpectNear combined(2), kx, 0.015, "asymmetric x curvature", passed, failed, report
    TestExpectNear combined(3), ky, 0.015, "asymmetric y curvature", passed, failed, report

    Dim torquePart() As Double, shearPart() As Double, normalPart() As Double
    Dim plane() As Double
    ok = solver.Solve(0#, 0#, 0#, torque, 0#, 0#, plane, resultants, torquePart)
    TestRecordSolve ok, "asymmetric pure torque", solver, passed, failed, report
    If Not ok Then Exit Sub
    ok = solver.Solve(0#, 0#, 0#, 0#, qx, qy, plane, resultants, shearPart)
    TestRecordSolve ok, "asymmetric biaxial shear", solver, passed, failed, report
    If Not ok Then Exit Sub
    ok = solver.Solve(force, mx, my, 0#, 0#, 0#, plane, resultants, normalPart)
    TestRecordSolve ok, "asymmetric normal forces", solver, passed, failed, report
    If Not ok Then Exit Sub
    Dim k As Long, p As Long, fieldError As Double, maximumStress As Double
    For p = 1 To UBound(stress, 2)
        For k = 1 To UBound(stress, 1)
            maximumStress = GeomMax(maximumStress, Abs(stress(k, p)))
            fieldError = GeomMax(fieldError, Abs(stress(k, p) - torquePart(k, p) - shearPart(k, p) - normalPart(k, p)))
        Next k
    Next p
    TestRecord fieldError <= 1E-10 * maximumStress, "asymmetric component superposition error=" & CStr(fieldError), _
        passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При ν=0 связь осевого растяжения с поперечными деформациями отсутствует.
' Продольный стержень замещает бетон своей площадью, поперечная ветвь
' исключает свой объём бетона и не передаёт продольное усилие.
Private Sub TestReinforcedElasticReplacement(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const E_CONCRETE As Double = 32000#
    Const E_STEEL As Double = 200000#
    Const force As Double = 15000#
    Const barArea As Double = 100#
    Dim geometry As CGeometryRegion, section As CSectionModel
    Set geometry = TestFixtureBox(200#, 160#)
    Set section = TestBuildSection(geometry)
    section.AddRebarElement 0#, 0#, Sqr(4# * barArea / GEOM_PI), barArea, "A400"
    Dim layout As New CTransverseRebarLayout
    Dim geometryBuilder As New CSectionDefinitionBuilder
    Dim branch As CGeometryPath
    geometryBuilder.Initialize "TestElasticBranch"
    Set branch = geometryBuilder.LinePath("TopBranch", -55#, 30#, 55#, 30#)
    layout.AddGroup branch, 8#, 90#, "A400"

    Dim solver As New CCoupledElasticState
    solver.PrepareModel section, E_CONCRETE, 0#, 2, 2, 9, layout, E_STEEL
    Dim deformation() As Double, resultants() As Double, stress() As Double
    Dim ok As Boolean
    ok = solver.Solve(force, 0#, 0#, 0#, 0#, 0#, deformation, resultants, stress)
    TestRecordSolve ok, "reinforced axial equilibrium", solver, passed, failed, report
    If Not ok Then Exit Sub
    Dim removed As Double, centroidY As Double, axial As Double, bending As Double, cross As Double
    removed = GEOM_PI * 16# * 110# / 90#
    centroidY = 30#
    axial = E_CONCRETE * (200# * 160# - removed) + (E_STEEL - E_CONCRETE) * barArea
    bending = E_CONCRETE * (200# * 160# ^ 3 / 12# - removed * centroidY ^ 2)
    cross = -E_CONCRETE * removed * centroidY
    TestExpectNear deformation(1), force * bending / (axial * bending - cross ^ 2), 1E-6, _
        "reinforced net axial area", passed, failed, report
    TestExpectNear deformation(2), -force * cross / (axial * bending - cross ^ 2), 0.01, _
        "transverse replacement moment", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт достоверные границы диска и кольца через рабочую модель сечения.
Private Function TestBuildSection(ByVal geometry As CGeometryRegion) As CSectionModel
    Dim mesh As New CFiberMeshBuilder
    mesh.BuildMesh geometry, geometry.Width / 20#, geometry.Height / 20#, cellGeometry:="Contour"
    Dim builder As New CSectionModelBuilder
    Set TestBuildSection = builder.BuildFromGenerated(mesh, Nothing)
    geometry.BuildContours TestBuildSection.Contours
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Приводит ошибку вычисления вместе с человеком читаемой причиной.
Private Sub TestRecordSolve(ByVal ok As Boolean, ByVal name As String, ByVal solver As CCoupledElasticState, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If Not ok Then name = name & " " & solver.ResultMeta.ResultComment
    TestRecord ok, name, passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет величину с заранее назначенным относительным допуском.
Private Sub TestExpectNear(ByVal value As Double, ByVal reference As Double, ByVal tolerance As Double, _
        ByVal name As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    TestRecord Abs(value / reference - 1#) < tolerance, _
        name & " actual=" & CStr(value) & " expected=" & CStr(reference), passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет невязку и поправку, собранные для текущего состояния, с прямым
' линейным решением. Проверяет в том числе продольный градиент и все усилия.
Public Function RunCoupledEquationAssemblyTests() As String
    On Error GoTo Failed
    Const E_MODULUS As Double = 32000#
    Dim solver As New CCoupledElasticState
    Dim section As CSectionModel
    Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    solver.PrepareModel section, E_MODULUS, 0#, 1, 1, 8

    Dim external() As Double, state() As Double, residual() As Double, linearized() As Double
    Dim i As Long, j As Long, point As Long, passed As Long, failed As Long, report As String
    Dim stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double, strain(1 To 6) As Double
    Dim forces(1 To 6) As Double, trialForces(1 To 6) As Double
    For i = 1 To 3
        tangent(i, i) = E_MODULUS
        tangent(i + 3, i + 3) = E_MODULUS / 2#
    Next i
    forces(1) = 30000#: forces(2) = 1900000#: forces(3) = -1100000#
    forces(4) = 1400000#: forces(5) = 3200#: forces(6) = -1800#
    solver.LoadEquations forces(1), forces(2), forces(3), forces(4), forces(5), forces(6), external
    ReDim state(1 To solver.DegreeCount)
    For i = 1 To solver.DegreeCount
        state(i) = Sin(i) * 1E-7
    Next i
    solver.StartEquations state, external, residual, linearized, True
    For point = 1 To solver.PointCount
        solver.PointStrain point, state, strain
        For j = 1 To 6
            stress(j) = tangent(j, j) * strain(j)
        Next j
        solver.AddPointEquilibrium point, state, stress, tangent, residual, linearized, True, trialForces
    Next point
    Dim system As New CCoupledLinearSystem, ok As Boolean
    For i = 1 To UBound(residual)
        residual(i) = -residual(i)
    Next i
    ok = system.Solve(linearized, residual, solver.DegreeCount)
    TestRecord ok, "linearization solves mixed equations", passed, failed, report
    If Not ok Then GoTo ReportResult
    For i = 1 To UBound(state)
        state(i) = state(i) + residual(i)
    Next i
    solver.StartEquations state, external, residual, linearized, False
    Erase trialForces
    For point = 1 To solver.PointCount
        solver.PointStrain point, state, strain
        For j = 1 To 6
            stress(j) = tangent(j, j) * strain(j)
        Next j
        solver.AddPointEquilibrium point, state, stress, tangent, residual, linearized, False, trialForces
    Next point
    Dim norm As Double
    For i = 1 To UBound(residual)
        norm = GeomMax(norm, Abs(residual(i)))
    Next i
    TestRecord norm < 1E-10, "internal residual after one correction=" & CStr(norm), passed, failed, report
    For i = 1 To 6
        TestExpectNear trialForces(i), forces(i), 1E-8, _
            "assembled component " & CStr(i), passed, failed, report
    Next i
ReportResult:
    RunCoupledEquationAssemblyTests = report & "TOTAL_COUPLED_EQUATIONS: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunCoupledEquationAssemblyTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' Сохраняет итог отдельного сравнения и увеличивает соответствующий счетчик.
Private Sub TestRecord(ByVal ok As Boolean, ByVal name As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If ok Then
        passed = passed + 1
        report = report & "OK: " & name & vbCrLf
    Else
        failed = failed + 1
        report = report & "FAIL: " & name & vbCrLf
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет две правые части одной неспряжённой системы с заданными решениями.
' Нулевой главный диагональный член требует смены ведущей строки.
Public Function RunCoupledFactorizationTests() As String
    On Error GoTo Failed
    Dim system As New CCoupledLinearSystem
    Dim matrix(1 To 3, 1 To 3) As Double, rhs(1 To 3) As Double
    Dim passed As Long, failed As Long, report As String, ok As Boolean
    matrix(2, 1) = 2#: matrix(3, 1) = 1#
    matrix(1, 2) = 1#: matrix(2, 2) = 1#
    matrix(1, 3) = 3#: matrix(3, 3) = 1#
    ok = system.Factor(matrix, 3)
    TestRecord ok, "factor pivoted matrix", passed, failed, report
    If ok Then
        rhs(1) = 7#: rhs(2) = 3#: rhs(3) = 6#
        ok = system.SolveFactored(rhs)
        TestRecord ok And Abs(rhs(1) - 1#) + Abs(rhs(2) - 2#) + Abs(rhs(3) - 3#) < 1E-12, _
            "factor solve first vector", passed, failed, report
        rhs(1) = 2#: rhs(2) = 3#: rhs(3) = 6#
        ok = system.SolveFactored(rhs)
        TestRecord ok And Abs(rhs(1) - 2#) + Abs(rhs(2) - 1#) + Abs(rhs(3)) < 1E-12, _
            "factor solve second vector", passed, failed, report
        TestRecord matrix(1, 1) = 0# And matrix(2, 1) = 2# And matrix(3, 3) = 1#, _
            "original system retained", passed, failed, report
    End If
    Erase matrix
    ok = system.Factor(matrix, 3)
    TestRecord Not ok And system.FailureCode = sfcSingularTangent, "singular result", passed, failed, report
    ok = system.SolveFactored(rhs)
    TestRecord Not ok, "failed refactorization discards previous factors", passed, failed, report
    RunCoupledFactorizationTests = report & "TOTAL_COUPLED_FACTORIZATION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunCoupledFactorizationTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет депланацию при чистом изгибе и кручении однородных сечений.
' Нормировка в кручении использует только угол закрутки и габарит.
Public Function RunCoupledWarpingTests() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const E_MODULUS As Double = 32000#
    Const TORQUE As Double = 2500000#
    Dim solver As New CCoupledElasticState, section As CSectionModel
    Dim deformation() As Double, resultants() As Double, stress() As Double, solutionField() As Double, solutionArray As Variant
    Dim torsionScale As Double, w As Double, maximum As Double, x As Long, y As Long, ok As Boolean

    Set section = TestBuildSection(TestFixtureBox(200#, 120#))
    solver.PrepareModel section, E_MODULUS, 0#, 3, 3, 9
    ok = solver.Solve(0#, 1800000#, 0#, 0#, 0#, 0#, deformation, resultants, stress, solutionArray)
    TestRecordSolve ok, "rectangle pure bending for warping", solver, passed, failed, report
    If ok Then
        solutionField = solutionArray
        For x = -80 To 80 Step 40
            For y = -50 To 50 Step 25
                w = Abs(solver.WarpingAt(x, y, solutionField))
                If w > maximum Then maximum = w
            Next y
        Next x
        TestRecord maximum < 1E-11, "pure bending has no internal w, max mm=" & CStr(maximum), passed, failed, report
    End If

    ok = solver.Solve(0#, 0#, 0#, TORQUE, 0#, 0#, deformation, resultants, stress, solutionArray)
    TestRecordSolve ok, "rectangle torsion for warping", solver, passed, failed, report
    If ok Then
        solutionField = solutionArray
        torsionScale = Abs(deformation(4)) * 100# ^ 2
        w = solver.WarpingAt(75#, 45#, solutionField)
        TestRecord Abs(w) > 0.05 * torsionScale, "rectangle off-axis warping mm=" & CStr(w), passed, failed, report
        TestRecord Abs(solver.WarpingAt(-75#, 45#, solutionField) + w) < 1E-8 * torsionScale And _
            Abs(solver.WarpingAt(75#, -45#, solutionField) + w) < 1E-8 * torsionScale, "warping is odd across each axis", passed, failed, report
        TestRecord Abs(solver.WarpingAt(0#, 0#, solutionField)) < 1E-8 * torsionScale, "zero warping at center", passed, failed, report
    End If

    Set section = TestBuildSection(TestFixtureDisk(200#))
    solver.PrepareModel section, E_MODULUS, 0#, 3, 3, 9
    ok = solver.Solve(0#, 0#, 0#, TORQUE, 0#, 0#, deformation, resultants, stress, solutionArray)
    TestRecordSolve ok, "circle torsion for warping", solver, passed, failed, report
    If ok Then
        solutionField = solutionArray
        torsionScale = Abs(deformation(4)) * 100# ^ 2
        maximum = 0#
        For x = -60 To 60 Step 30
            For y = -60 To 60 Step 30
                w = Abs(solver.WarpingAt(x, y, solutionField))
                If w > maximum Then maximum = w
            Next y
        Next x
        TestRecord maximum < 0.015 * torsionScale, "circle warping relative discrete error=" & CStr(maximum / torsionScale), passed, failed, report
    End If
    RunCoupledWarpingTests = report & "TOTAL_COUPLED_WARPING: passed=" & CStr(passed) & "; failed=" & CStr(failed)
    Exit Function
Failed:
    RunCoupledWarpingTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает нелинейное продольное равновесие с заданным кубическим законом.
' Моменты и сдвиги сверяет по шести независимым интегралам напряжений.
Public Function RunCoupledNewtonTests() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel, operators As New CCoupledElasticState
    Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    operators.PrepareModel section, 30000#, 0#, 1, 1, 8
    Dim material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    material.TestInitialize operators.PointCount
    solver.Initialize operators, material, 30000#
    Const AREA As Double = 21600#
    Const STRAIN As Double = 0.001
    Const FORCE As Double = AREA * (30000# * STRAIN + 1E+10 * STRAIN ^ 3)
    Dim field() As Double, resultants() As Double, acceptedField() As Double, ok As Boolean
    ok = solver.TryStep(FORCE, 0#, 0#, 0#, 0#, 0#, 1E-10, 1E-9, 12)
    TestRecord ok, "nonlinear axial equilibrium " & solver.ResultMeta.ResultComment, passed, failed, report
    If ok Then
        solver.CopyTrial field, resultants
        TestExpectNear field(1), STRAIN, 1E-9, "nonlinear axial strain", passed, failed, report
        TestExpectNear resultants(1), FORCE, 1E-9, "nonlinear axial resultant", passed, failed, report
        TestRecord material.TestAcceptedPeak = 0#, "successful trial is uncommitted", passed, failed, report
        solver.AcceptStep
        TestRecord Abs(material.TestAcceptedPeak - STRAIN) < 1E-12 And material.TestAcceptedSteps = 1, _
            "successful step publishes material history once", passed, failed, report
    End If

    material.TestFail = True
    ok = solver.TryStep(1.1 * FORCE, 0#, 0#, 0#, 0#, 0#, 1E-10, 1E-9, 12)
    TestRecord Not ok And solver.ResultMeta.InternalStatus = rsNumericalFailure, _
        "local update failure has a numerical status", passed, failed, report
    TestRecord InStr(solver.ResultMeta.DiagnosticDetails, " 2") > 0, _
        "point failure includes point index", passed, failed, report
    solver.CopyAccepted acceptedField
    TestRecord Abs(acceptedField(1) - STRAIN) < 1E-12 And material.TestAcceptedSteps = 1, _
        "material failure leaves accepted field and history", passed, failed, report
    material.TestFail = False
    ok = solver.TryStep(0.8 * FORCE, 200000#, 100000#, 300000#, 2000#, -1000#, 1E-9, 1E-8, 25)
    TestRecord ok, "six component nonlinear equilibrium " & solver.ResultMeta.ResultComment, passed, failed, report
    If ok Then
        solver.CopyTrial field, resultants
        Dim expected(1 To 6) As Double, i As Long
        expected(1) = 0.8 * FORCE: expected(2) = 200000#: expected(3) = 100000#
        expected(4) = 300000#: expected(5) = 2000#: expected(6) = -1000#
        For i = 1 To 6
            TestExpectNear resultants(i), expected(i), 0.00001, "recovered component " & CStr(i), passed, failed, report
        Next i
        solver.DiscardStep
    End If
    TestRecord material.TestAcceptedSteps = 1, "discard leaves history", passed, failed, report
    solver.ResetHistory
    solver.CopyAccepted acceptedField
    TestRecord acceptedField(1) = 0# And material.TestAcceptedPeak = 0#, _
        "independent combination starts without history", passed, failed, report
    RunCoupledNewtonTests = report & "TOTAL_COUPLED_NEWTON: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunCoupledNewtonTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
