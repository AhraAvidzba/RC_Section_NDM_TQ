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
    Dim section As CSectionModel: Set section = TestAsymmetricCoupledSection(region)
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
' Готовит единую невыпуклую область с эксцентричным круглым отверстием.
Private Function TestAsymmetricCoupledSection(ByRef region As CConcreteRegion, Optional ByVal shiftX As Double = 0#, _
        Optional ByVal shiftY As Double = 0#, Optional ByVal rotation As Double = 0#, _
        Optional ByVal reflectX As Boolean = False) As CSectionModel
    Dim query As New CSectionGeometryQuery, outline As CGeometryPath, geometry As New CGeometryRegion
    Set outline = query.AdjacentRectangles("AsymmetricSix", Array( _
        Array(-100#, -80#, 200#, 100#), Array(-100#, 20#, 135#, 80#)))
    Dim x As Double, y As Double: x = -40#: y = -24#
    query.TransformCoordinates x, y, shiftX, shiftY, rotation, reflectX
    Set outline = outline.Transformed(shiftX, shiftY, rotation, reflectX)
    Set region = query.SubtractContainedRegion(query.RegionFromPath(outline), query.CircleRegion(x, y, 21#))
    geometry.Initialize region
    Set TestAsymmetricCoupledSection = TestBuildSection(geometry)
End Function

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
    TestRecord solver.EquilibriumEquationDescription(4) = "равновесие по T" And _
        solver.EquilibriumEquationDescription(7) = "внутреннее равновесие по депланации w", _
        "equation diagnostics distinguish torsion and local warping", passed, failed, report
    TestRecord solver.EquilibriumEquationDescription(solver.DegreeCount - 7) = "продольный градиент N" And _
        solver.EquilibriumEquationDescription(solver.DegreeCount - 5) = "продольный градиент My" And _
        Len(solver.EquilibriumEquationDescription(solver.DegreeCount - 3)) = 0, _
        "equation diagnostics identify gradients and exclude gauge constraints", passed, failed, report
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
    Const FORCE As Double = AREA * (30000# * STRAIN + 1E+10# * STRAIN ^ 3)
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
    TestRecord InStr(solver.ResultMeta.DiagnosticDetails, "последняя точка отказа: 2") > 0, _
        "point failure includes point index", passed, failed, report
    solver.CopyAccepted acceptedField
    TestRecord Abs(acceptedField(1) - STRAIN) < 1E-12 And material.TestAcceptedSteps = 1, _
        "material failure leaves accepted field and history", passed, failed, report
    material.TestFail = False
    material.TestMaximumIncrement = 1E-14
    ok = solver.TryStep(1.1 * FORCE, 0#, 0#, 0#, 0#, 0#, 1E-10, 1E-9, 12)
    TestRecord Not ok And InStr(solver.ResultMeta.DiagnosticDetails, "ни одна уменьшенная поправка") > 0, _
        "reports whether every reduced correction was rejected by the material", passed, failed, report
    solver.CopyAccepted acceptedField
    TestRecord Abs(acceptedField(1) - STRAIN) < 1E-12 And material.TestAcceptedSteps = 1, _
        "rejected damping probes leave committed history unchanged", passed, failed, report
    material.TestMaximumIncrement = 0#
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


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет принятие постоянной части и уменьшение трудных шагов по lambda.
' Аналитический материал ограничивает приращение относительно принятой истории.
Public Function RunTestCoupledContinuation() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#
    Const FORCE0 As Double = 0.6 * 30000# * 0.001 * AREA
    Const FORCE1 As Double = 1000000#
    Dim section As CSectionModel, operators As New CCoupledElasticState
    Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    operators.PrepareModel section, 30000#, 0#, 1, 1, 8
    Dim material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    material.TestInitialize operators.PointCount
    material.TestMaximumIncrement = 0.00022
    solver.Initialize operators, material, 30000#
    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    path.Initialize FORCE0, FORCE1 - FORCE0, 0#, 200000#, 0#, 100000#, 0#, 300000#, 0#, 2000#, 0#, -1000#
    branch.Initialize solver, path, 1, 100, 20, 1E-8, 1E-8, True
    Dim field() As Double, resultants() As Double, expected() As Double, ok As Boolean
    ok = branch.Advance(0#)
    TestRecord ok, "constant part reached " & branch.ResultMeta.DiagnosticDetails, passed, failed, report
    If Not ok Then GoTo Done
    branch.CopyAccepted field, resultants
    TestExpectNear resultants(1), FORCE0, 1E-7, "preload resultant", passed, failed, report
    TestRecord branch.AcceptedOffsetFraction = 1# And branch.AcceptedLambda = 0# And branch.RejectedAttempts > 0, _
        "preload reduced an unsafe increment and reached its exact target", passed, failed, report
    Dim trace As CCoupledLoadTrace, traceCount As Long, fraction As Double, lambda As Double
    Set trace = branch.SavedTrace
    trace.ReadPoint 1, fraction, lambda
    TestRecord fraction > 0# And fraction < 1# And lambda = 0#, "trace records only accepted offset fraction", passed, failed, report

    ok = branch.Advance(1#)
    TestRecord ok, "direction reached " & branch.ResultMeta.DiagnosticDetails, passed, failed, report
    If Not ok Then GoTo Done
    branch.CopyAccepted field, resultants
    path.TargetSix 1#, expected
    Dim i As Long
    For i = 1 To 6
        TestExpectNear resultants(i), expected(i), 0.00002, "path component " & CStr(i), passed, failed, report
    Next i
    TestRecord branch.AcceptedLambda = 1# And branch.RejectedAttempts > 0, _
        "direction reduced an unsafe increment and reached exact lambda", passed, failed, report

    Dim acceptedBefore As Long, firstDegree As Double
    acceptedBefore = material.TestAcceptedSteps: firstDegree = field(1)
    traceCount = trace.Count
    trace.ReadPoint traceCount, fraction, lambda
    TestRecord traceCount = branch.SavedFieldCount And fraction = 1# And lambda = 1#, _
        "trace and material history end at the same target", passed, failed, report
    material.TestFail = True
    ok = branch.Advance(1.1)
    TestRecord Not ok And branch.ResultMeta.InternalStatus = rsNumericalFailure And branch.AcceptedLambda = 1#, _
        "minimum step failure stays numerical at last accepted lambda", passed, failed, report
    branch.CopyAccepted field, resultants
    TestRecord firstDegree = field(1) And acceptedBefore = material.TestAcceptedSteps And trace.Count = traceCount, _
        "failed increments leave accepted section, history and trace unchanged", passed, failed, report
    material.TestFail = False
    TestRecord branch.Advance(1.1), "recovered path can continue", passed, failed, report
    ok = branch.Advance(1#)
    TestRecord Not ok And branch.ResultMeta.ResultCode = rcInvalidInput, "history cannot be reversed by moving lambda back", passed, failed, report
    path.Initialize 0#, 0#, 0#, 0#, 0#, 0#
    branch.Initialize solver, path, 1, 100, 20, 1E-8, 1E-8
    TestRecord branch.Advance(0#), "zero load returns a confirmed state", passed, failed, report
    branch.CopyAccepted field, resultants
    TestRecord field(1) = 0# And resultants(1) = 0# And material.TestAcceptedPeak = 0#, _
        "independent zero load starts with a virgin history", passed, failed, report
    TestExactSegmentEndpoints solver, passed, failed, report
    TestLoadSecantRetry passed, failed, report
Done:
    RunTestCoupledContinuation = report & "TOTAL_COUPLED_CONTINUATION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledContinuation = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Воспроизводит округление суммы текущего множителя и номинального шага.
' Проверяет отказ и принятие ровно на границе 1/64 до исчерпания бюджета.
Public Function RunTestCoupledMinimumLoadStep() As String
    Const START As Double = 18.55811, FINISH As Double = 20.337165
    Const FORCE_SCALE As Double = 10000#, AREA As Double = 21600#
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim operators As New CCoupledElasticState, material As New CTestCoupledPointResponse
    Dim solver As New CCoupledNewtonSolver, branch As New CCoupledLoadContinuation, path As New CLoadPathVector
    operators.PrepareModel TestBuildSection(TestFixtureBox(180#, 120#)), 30000#, 0#, 1, 1, 5
    material.TestInitialize operators.PointCount
    solver.Initialize operators, material, 30000#
    path.Initialize 0#, FORCE_SCALE, 0#, 0#, 0#, 0#
    branch.Initialize solver, path, 2, 7, 20, 1E-8, 1E-8, True
    If Not branch.Advance(START) Then Err.Raise 5, , "Не достигнуто исходное состояние контрольного шага."

    Dim trace As CCoupledLoadTrace, traceCount As Long, acceptedCount As Long, solved As Boolean
    Set trace = branch.SavedTrace
    traceCount = trace.Count: acceptedCount = material.TestAcceptedSteps
    material.TestFail = True
    solved = branch.Advance(FINISH)
    TestRecord Not solved And branch.ResultMeta.InternalStatus = rsNumericalFailure, _
        "отказ на граничном подшаге остаётся численным", passed, failed, report
    TestRecord branch.RejectedAttempts = 7, "проверены все семь проб вплоть до 1/64", passed, failed, report
    TestRecord branch.AcceptedLambda = START And trace.Count = traceCount And material.TestAcceptedSteps = acceptedCount, _
        "численный отказ сохраняет принятую нагрузку и историю", passed, failed, report

    material.TestFail = False
    branch.Initialize solver, path, 2, 7, 20, 1E-8, 1E-8, True
    If Not branch.Advance(START) Then Err.Raise 5, , "Не восстановлено исходное состояние контрольного шага."
    Set trace = branch.SavedTrace
    traceCount = trace.Count: acceptedCount = material.TestAcceptedSteps
    Dim field() As Double, resultants() As Double, expected() As Double, minimum As Double
    branch.CopyAccepted field, resultants
    minimum = ((FINISH - START) / 2#) / 64#
    material.TestMaximumIncrement = 1.5 * minimum * FORCE_SCALE / (AREA * (30000# + 3E+10 * field(1) ^ 2))
    solved = branch.Advance(FINISH)
    TestRecord Not solved And branch.ResultMeta.InternalStatus = rsNumericalFailure, _
        "бюджет заканчивается после принятия, целевое усилие остаётся открытым", passed, failed, report
    TestRecord branch.AcceptedLambda = START + minimum, "граничный подшаг принят при прежнем нижнем пределе", passed, failed, report
    TestRecord branch.RejectedAttempts = 6 And trace.Count = traceCount + 1 And material.TestAcceptedSteps = acceptedCount + 1, _
        "шесть отказов сохраняют историю, граничная проба пишет ровно один шаг", passed, failed, report
    branch.CopyAccepted field, resultants
    path.TargetSix START + minimum, expected
    TestRecord operators.ResultantResidualNorm(resultants, expected) < 1E-6, _
        "в принятой граничной точке выполнено равновесие шести усилий", passed, failed, report
    RunTestCoupledMinimumLoadStep = report & "TOTAL_COUPLED_MINIMUM_LOAD_STEP: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledMinimumLoadStep = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет остаток порядка округления Double в четвёртом переменном
' и десятом постоянном шаге. Целевое усилие решается на штатном шаге.
Private Sub TestExactSegmentEndpoints(ByVal solver As CCoupledNewtonSolver, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation, trace As CCoupledLoadTrace
    path.Initialize 0#, 100000#, 0#, 200000#, 0#, 100000#
    branch.Initialize solver, path, 4, 100, 20, 1E-8, 1E-8, True
    TestRecord branch.Advance(0.5), "exact endpoint fixture first segment", passed, failed, report
    Set trace = branch.SavedTrace
    Dim before As Long: before = trace.Count
    Dim fraction As Double, lambda As Double, solved As Boolean
    solved = branch.Advance(0.65)
    trace.ReadPoint trace.Count, fraction, lambda
    TestRecord solved And lambda = 0.65 And trace.Count = before + 4, _
        "binary rounding does not cause a fifth variable substep", passed, failed, report

    path.Initialize 100000#, 0#, 200000#, 0#, 100000#, 0#
    branch.Initialize solver, path, 10, 10, 20, 1E-8, 1E-8, True
    solved = branch.Advance(0#)
    Set trace = branch.SavedTrace
    trace.ReadPoint trace.Count, fraction, lambda
    TestRecord solved And fraction = 1# And lambda = 0# And trace.Count = 10, _
        "rounded offset accumulation reaches the target within ten probes", passed, failed, report
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет повтор целевой нагрузки и ограничения на длину прогноза и число проб.
Private Sub TestLoadSecantRetry(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim operators As New CCoupledElasticState
    operators.PrepareModel TestBuildSection(TestFixtureBox(180#, 120#)), 30000#, 0#, 1, 1, 5
    Dim path As New CLoadPathVector
    path.Initialize 0#, 648000#, 0#, 200000#, 0#, -100000#, 0#, 300000#, 0#, 2000#, 0#, -1000#

    TestLoadSecantRetryCase operators, path, 0.75, 2, True, "bounded secant", passed, failed, report
    TestLoadSecantRetryCase operators, path, 1#, 2, False, "excessive secant span", passed, failed, report
    TestLoadSecantRetryCase operators, path, 0.75, 1, False, "exhausted trial budget", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Последнее принятое поле не проходит порог, а прогноз на штатном шаге проходит.
' Все варианты начинают с одной пары физических состояний и одной истории.
Private Sub TestLoadSecantRetryCase(ByVal operators As CCoupledElasticState, ByVal path As CLoadPathVector, _
        ByVal target As Double, ByVal attempts As Long, ByVal expectSuccess As Boolean, ByVal name As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    Dim branch As New CCoupledLoadContinuation
    material.TestInitialize operators.PointCount
    solver.Initialize operators, material, 30000#
    branch.Initialize solver, path, 1, attempts, 20, 1E-8, 1E-8, True
    Dim ok As Boolean: ok = branch.Advance(0.25)
    If ok Then ok = branch.Advance(0.5)
    TestRecord ok, name & " establishes both predecessor fields", passed, failed, report
    If Not ok Then Exit Sub

    Dim trace As CCoupledLoadTrace: Set trace = branch.SavedTrace
    Dim field() As Double, resultants() As Double, expected() As Double
    Dim strainBefore As Double, peakBefore As Double, fraction As Double, lambda As Double
    branch.CopyAccepted field, resultants
    strainBefore = field(1): peakBefore = material.TestAcceptedPeak
    material.TestMinimumAxialStrain = 0.0006 ' Разделяет исходное поле и прогноз без изменения закона напряжений.
    ok = branch.Advance(target)
    Dim meta As CResultMeta: Set meta = branch.ResultMeta
    branch.CopyAccepted field, resultants
    trace.ReadPoint trace.Count, fraction, lambda
    TestRecord ok = expectSuccess, name & " returns expected outcome", passed, failed, report
    TestRecord fraction = 1# And trace.Count = branch.SavedFieldCount, _
        name & " preserves the accepted load trace", passed, failed, report
    If expectSuccess Then
        path.TargetSix target, expected
        TestRecord operators.ResultantResidualNorm(resultants, expected) <= 0.0000005, _
            name & " balances all six requested forces", passed, failed, report
        TestRecord lambda = target And branch.AcceptedLambda = target And branch.PreviousAcceptedLambda = 0.5, _
            name & " accepts the exact target after its original predecessor", passed, failed, report
        TestRecord trace.Count = 3 And material.TestAcceptedSteps = 3 And material.TestAcceptedPeak > peakBefore, _
            name & " commits one material state and one trace entry", passed, failed, report
        TestRecord InStr(1, meta.DiagnosticDetails, "секущим прогнозом", vbTextCompare) > 0, _
            name & " records the recovery mechanism", passed, failed, report
    Else
        TestRecord meta.InternalStatus = rsNumericalFailure And lambda = 0.5 And branch.AcceptedLambda = 0.5, _
            name & " retains numerical status and the last accepted load", passed, failed, report
        TestRecord trace.Count = 2 And material.TestAcceptedSteps = 2 And field(1) = strainBefore And _
            material.TestAcceptedPeak = peakBefore, name & " does not commit a rejected trial", passed, failed, report
    End If
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет предел деформации центрального стержня с явной формулой N.
' Контрольное кручение не влияет на продольное равновесие этого материала.
Public Function RunTestCoupledSteelLimit() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#
    Const BAR_AREA As Double = 1000#
    Const STEEL_LIMIT As Double = 0.001
    Const LIMIT_FORCE As Double = (AREA - BAR_AREA) * (30000# * STEEL_LIMIT + 1E+10# * STEEL_LIMIT ^ 3) + _
        BAR_AREA * (100# + 2000# * (STEEL_LIMIT - 100# / 200000#))
    Dim section As CSectionModel, operators As New CCoupledElasticState
    Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, 12#, BAR_AREA, "TestGrade"
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModulus:=200000#
    Dim material As New CTestCoupledPointResponse, steel As New CCoupledSteelModel, models As Object
    material.TestInitialize operators.PointCount
    steel.Initialize 200000#, 100#, 100#, 2000#, 2000#, STEEL_LIMIT, STEEL_LIMIT
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "TestGrade", steel
    Dim response As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    response.Initialize operators, material, models
    solver.Initialize operators, response, 30000#
    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    Dim low As Double, high As Double, onOffset As Boolean, expected As Double, steps As Variant
    Dim trial As Variant
    For Each trial In Array(0.98, 1.02)
        expected = CDbl(trial)
        path.Initialize 0#, LIMIT_FORCE / expected, 0#, 0#, 0#, 0#, 0#, 300000#, 0#, 0#, 0#, 0#
        For Each steps In Array(1, 4, 16)
            branch.Initialize solver, path, CLng(steps), 100, 25, 1E-8, 1E-8
            branch.TraceSteelLimit response, 1#, 1.25, 0.003
            TestRecord branch.CopySteelLimit(low, high, onOffset), _
                "steel limit is bracketed, steps=" & CStr(steps) & " " & branch.ResultMeta.DiagnosticDetails, passed, failed, report
            TestRecord low <= expected And high >= expected And high - low <= 0.0001 And Not onOffset, _
                "analytic limit lies within the refined near-one bracket", passed, failed, report
            TestRecord branch.AcceptedLambda <= low And response.AcceptedSteelUtilization <= 1#, _
                "over-limit probes did not modify accepted history", passed, failed, report
        Next steps
    Next trial

    path.Initialize LIMIT_FORCE / 0.8, 1000#, 0#, 0#, 0#, 0#
    branch.Initialize solver, path, 4, 100, 25, 1E-8, 1E-8
    branch.TraceSteelLimit response, 1#, 1.5, 0.0001
    TestRecord branch.CopySteelLimit(low, high, onOffset) And onOffset And low <= 0.8 And high >= 0.8, _
        "temporary fixed-load failure is classified by a confirmed steel limit", passed, failed, report
    TestRecord Not branch.HasAcceptedVariableState, "incomplete preload cannot confirm a variable-load lower bound", passed, failed, report

    path.Initialize 0#, LIMIT_FORCE, 0#, 0#, 0#, 0#
    branch.Initialize solver, path, 4, 100, 25, 1E-8, 1E-8
    branch.TraceSteelLimit response, 1#, 0.9, 0.0001
    TestRecord Not branch.CopySteelLimit(low, high, onOffset) And branch.ResultMeta.InternalStatus = rsNumericalFailure, _
        "search bound is not a material limit", passed, failed, report
    TestRecord branch.HasAcceptedVariableState And branch.AcceptedLambda = 0.9, _
        "search bound retains last admissible step", passed, failed, report

    material.TestFail = True
    path.Initialize LIMIT_FORCE / 4#, LIMIT_FORCE, 0#, 0#, 0#, 0#
    branch.Initialize solver, path, 4, 100, 25, 1E-8, 1E-8
    branch.TraceSteelLimit response, 1#, 1.5, 0.0001
    TestRecord Not branch.CopySteelLimit(low, high, onOffset) And Not branch.HasAcceptedVariableState And _
        branch.ResultMeta.InternalStatus = rsNumericalFailure, "unresolved preload is never a physical failure", passed, failed, report
    RunTestCoupledSteelLimit = report & "TOTAL_COUPLED_STEEL_LIMIT: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledSteelLimit = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает полную производную собранной невязки с независимыми
' разностями. Кубический тестовый закон имеет гладкую неупругую касательную.
Public Function RunTestCoupledGradientLinearization() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim operators As New CCoupledElasticState, material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5
    material.TestInitialize operators.PointCount
    solver.Initialize operators, material, 30000#

    Dim field() As Double, residual() As Double, frozen() As Double, corrected() As Double
    Dim plus() As Double, minus() As Double, scratch() As Double, column As Long, row As Long, dof As Long
    Dim frozenError As Double, correctedError As Double, saved As Double, expected As Double
    Const STEP_SIZE As Double = 0.00000008
    dof = operators.DegreeCount
    ReDim field(1 To dof)
    field(1) = 0.0004: field(2) = 0.00015: field(3) = 0.0001: field(4) = 0.0002
    field(dof - 7) = 0.001: field(dof - 6) = 0.002: field(dof - 5) = -0.001
    field(dof - 4) = 0.00025  ' Последние четыре степени свободы задают только условия однозначности.
    If Not solver.TestEquilibriumLinearization(field, True, False, residual, frozen) Then Err.Raise 5, , "Frozen assembly failed."
    If Not solver.TestEquilibriumLinearization(field, True, True, residual, corrected) Then Err.Raise 5, , "Corrected assembly failed."
    For column = 1 To dof
        saved = field(column)
        field(column) = saved + STEP_SIZE
        If Not solver.TestEquilibriumLinearization(field, False, False, plus, scratch) Then Err.Raise 5, , "Positive field perturbation failed."
        field(column) = saved - STEP_SIZE
        If Not solver.TestEquilibriumLinearization(field, False, False, minus, scratch) Then Err.Raise 5, , "Negative field perturbation failed."
        field(column) = saved
        For row = 1 To dof
            expected = (plus(row) - minus(row)) / (2# * STEP_SIZE)
            If Abs(frozen(column, row) - expected) > frozenError Then frozenError = Abs(frozen(column, row) - expected)
            If Abs(corrected(column, row) - expected) > correctedError Then correctedError = Abs(corrected(column, row) - expected)
        Next row
    Next column
    TestRecord frozenError > 1#, "frozen tangent omits a measured coupling " & CStr(frozenError), passed, failed, report
    TestRecord correctedError < 0.001, "corrected assembly agrees with full residual differences " & CStr(correctedError), passed, failed, report
    For column = dof - 7 To dof - 4: field(column) = 0#: Next column
    If Not solver.TestEquilibriumLinearization(field, True, False, residual, frozen) Then Err.Raise 5, , "Zero-gradient assembly failed."
    If Not solver.TestEquilibriumLinearization(field, True, True, residual, corrected) Then Err.Raise 5, , "Zero-gradient corrected assembly failed."
    correctedError = 0#
    For column = 1 To dof
        For row = 1 To dof
            correctedError = correctedError + Abs(corrected(column, row) - frozen(column, row))
        Next row
    Next column
    TestRecord correctedError = 0#, "zero gradient requires no correction", passed, failed, report
    RunTestCoupledGradientLinearization = report & "TOTAL_COUPLED_GRADIENT_LINEARIZATION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledGradientLinearization = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет деформационный контроль на ветви с известным максимумом
' N = A*(E*eps - c*eps^3). Момент кручения масштабируется с N.
Public Function RunTestCoupledDeformationControl() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim operators As New CCoupledElasticState, material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5
    material.TestInitialize operators.PointCount
    material.TestUseDescendingAxialLaw
    solver.Initialize operators, material, 30000#
    Dim baseLoad(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    baseLoad(1) = 4# * AREA: baseLoad(4) = 30000#
    direction(1) = 20# * AREA: direction(4) = 200000#
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    If Not solver.TryStep(baseLoad(1), 0#, 0#, baseLoad(4), 0#, 0#, 1E-8, 1E-8, 30) Then _
        Err.Raise 5, , "Independent preload failed."
    solver.AcceptStep

    Dim field() As Double, resultants() As Double, strain As Variant, lambda As Double
    Dim previousLambda As Double, expected As Double, eps As Double
    Dim history As New Collection
    solver.CopyAccepted field: history.Add field
    For Each strain In Array(0.0002, 0.0006, 0.00095, 0.00105, 0.0012, 0.0015)
        eps = CDbl(strain)
        If Not solver.TryDeformationStep(baseLoad, direction, control, eps, previousLambda, lambda, 1E-8, 1E-8, 30) Then _
            Err.Raise 5, , "Deformation step " & CStr(eps) & " failed. " & solver.ResultMeta.ResultComment & " " & solver.ResultMeta.DiagnosticDetails
        solver.CopyTrial field, resultants
        expected = (AREA * (30000# * eps - 1E+10# * eps ^ 3) - baseLoad(1)) / direction(1)
        TestExpectNear lambda, expected, 0.00000001, "strain-controlled multiplier at " & CStr(eps), passed, failed, report
        TestExpectNear field(1), eps, 0.00000001, "strain control equation", passed, failed, report
        TestExpectNear resultants(4), baseLoad(4) + lambda * direction(4), 0.000001, _
            "scaled torque in the same solution", passed, failed, report
        If eps >= 0.00105 Then TestRecord lambda < previousLambda, "branch descends after peak", passed, failed, report
        previousLambda = lambda
        solver.AcceptStep
        history.Add field
    Next strain
    TestRecord solver.ReplayAcceptedFields(history, 3), "replay restores the prepeak material history", passed, failed, report
    solver.CopyAccepted field
    TestExpectNear field(1), 0.0006, 0.00000001, "restored field before peak", passed, failed, report
    TestExpectNear material.TestAcceptedPeak, 0.0006, 0.00000001, "replay restores history instead of retaining accepted postpeak strain", passed, failed, report
    Dim prefix As New Collection, i As Long
    For i = 1 To 3: prefix.Add history(i): Next i
    For i = 1 To 6: resultants(i) = baseLoad(i) + 0.592 * direction(i): Next i
    Dim branch As New CCoupledDeformationContinuation, peak As Double, accuracy As Double, designVerified As Boolean
    branch.Trace solver, Nothing, baseLoad, direction, prefix, 0.096, 0.592, resultants, 1.5, 0.001, 30, 1E-8, 1E-8
    TestRecord branch.CopyPeak(peak, accuracy, designVerified), _
        "confirms first peak across two independently replayed increments " & branch.ResultMeta.ResultComment & _
        " " & branch.ResultMeta.DiagnosticDetails, passed, failed, report
    TestExpectNear peak, 0.8, 0.001, "first peak agrees with independent closed form", passed, failed, report
    TestRecord accuracy > 0# And accuracy <= 0.001 And Not designVerified, _
        "reports resolution of peak below design", passed, failed, report
    TestRecord Abs(peak - 0.8) <= accuracy, "numerical resolution covers the independent peak", passed, failed, report
    Dim atPeak As Collection
    branch.CopyLimitState atPeak, resultants
    If atPeak Is Nothing Then
        TestRecord False, "independently confirmed peak has no saved field", passed, failed, report
    Else
        field = atPeak.Item(atPeak.Count)
        expected = (AREA * (30000# * field(1) - 1E+10# * field(1) ^ 3) - baseLoad(1)) / direction(1)
        TestExpectNear expected, peak, 0.00000001, "saved strain gives the confirmed peak multiplier on the independent cubic law", passed, failed, report
        TestExpectNear resultants(4), baseLoad(4) + peak * direction(4), 0.000001, "actual resultants of the saved peak", passed, failed, report
        TestRecord solver.ReplayAcceptedFields(atPeak, atPeak.Count), "saved peak history replays", passed, failed, report
        TestExpectNear material.TestAcceptedPeak, field(1), 0.00000001, "no descending shoulder in the saved history", passed, failed, report
    End If
    Dim capacity As New CCapacityResult, vector(1 To 6) As Double
    If peak > 0# Then
        capacity.InitializeCoupledPeak peak, accuracy, designVerified, vector, "Auto", vbNullString
        TestRecord capacity.Status = "FAIL" And capacity.LambdaCapacity = peak, "peak below load gives physical FAIL", passed, failed, report
    End If
    capacity.InitializeCoupledPeak 0.99998, 0.00004, False, vector, "Auto", vbNullString
    TestRecord capacity.Status = "NumFail" And capacity.LambdaCapacity = 0# And capacity.HasUltimatePoint, _
        "uncertain peak near design remains explicit without definitive reserve", passed, failed, report
    capacity.InitializeCoupledPeak 1.12, 0.001, True, vector, "Auto", vbNullString
    TestRecord capacity.Status = "OK" And capacity.LambdaCapacity = 1.12, "peak above verified design load gives reserve", passed, failed, report
    RunTestCoupledDeformationControl = report & "TOTAL_COUPLED_DEFORMATION_CONTROL: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledDeformationControl = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет поправку множителя закрепления и выбор шага по механическому полю.
Public Function RunTestCoupledKinematicGauge() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim operators As New CCoupledElasticState, solver As New CCoupledNewtonSolver
    Dim center As New CTestCoupledPointResponse, plusSide As New CTestCoupledPointResponse, minusSide As New CTestCoupledPointResponse
    operators.PrepareModel TestBuildSection(TestFixtureBox(180#, 120#)), 30000#, 0#, 1, 1, 5
    center.TestInitialize operators.PointCount
    plusSide.TestInitialize operators.PointCount
    minusSide.TestInitialize operators.PointCount
    solver.Initialize operators, center, 30000#, plusSide, minusSide
    TestRecord operators.DegreeCount - solver.PhysicalDegreeCount = 4, "четыре множителя отделены от механики", passed, failed, report
    TestMechanicalSecant solver.PhysicalDegreeCount, operators.DegreeCount, passed, failed, report
    TestWarpGaugeCorrection operators, solver, passed, failed, report
    RunTestCoupledKinematicGauge = report & "TOTAL_COUPLED_KINEMATIC_GAUGE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledKinematicGauge = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает точную механическую секущую, отклонённое поле и чистое приращение множителя связи.
Private Sub TestMechanicalSecant(ByVal mechanicalCount As Long, ByVal degreeCount As Long, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim branch As New CCoupledDeformationContinuation
    Dim left() As Double, current() As Double, trial() As Double
    ReDim left(1 To degreeCount): ReDim current(1 To degreeCount): ReDim trial(1 To degreeCount)
    current(1) = 1E-6: trial(1) = 1.5E-6
    TestRecord branch.FitsMechanicalSecant(left, current, trial, mechanicalCount, 0.5), _
        "точная механическая секущая", passed, failed, report
    trial(mechanicalCount + 3) = -5.66313495971864E-7
    TestRecord branch.FitsMechanicalSecant(left, current, trial, mechanicalCount, 0.5), _
        "поправка множителя не меняет секущую", passed, failed, report
    trial(1) = 1.52E-6
    TestRecord Not branch.FitsMechanicalSecant(left, current, trial, mechanicalCount, 0.5), _
        "четыре процента отклонения механики превышают допуск", passed, failed, report
    current(1) = 0#: trial(1) = 0#
    TestRecord Not branch.FitsMechanicalSecant(left, current, trial, mechanicalCount, 0.5), _
        "чистый шаг множителя не продолжает механическую ветвь", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет линейную поправку, фактическую пересборку и штатное принятие нулевого поля НДС.
Private Sub TestWarpGaugeCorrection(ByVal operators As CCoupledElasticState, ByVal solver As CCoupledNewtonSolver, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const TOLERANCE As Double = 1E-7
    Dim field() As Double, external() As Double, residual() As Double, tangent() As Double, control() As Double
    Dim baseLoad(1 To 6) As Double, direction(1 To 6) As Double
    Dim i As Long, count As Long, warpGauge As Long, gauge As Long, first As Long, second As Long
    Dim weight As Double, maximumWeight As Double, initialValue As Double, change As Double, norm As Double
    count = operators.DegreeCount: warpGauge = solver.PhysicalDegreeCount + 3
    ReDim field(1 To count): ReDim external(1 To count): ReDim control(1 To count)
    operators.StartEquations field, external, residual, tangent, True
    For i = 1 To solver.PhysicalDegreeCount
        weight = tangent(i, warpGauge)
        If weight > 0# Then
            If first = 0 Then first = i Else second = i
            maximumWeight = GeomMax(maximumWeight, weight)
        End If
    Next i
    If second = 0 Then Err.Raise 5, , "Для проверки нужны две ненулевые связи продольного перемещения."
    initialValue = 1.5 * TOLERANCE / maximumWeight
    field(warpGauge) = initialValue
    operators.StartEquations field, external, residual, tangent, False
    If Not operators.TryWarpGaugeCorrection(residual, TOLERANCE, gauge, change) Then _
        Err.Raise 5, , "Допустимый интервал поправки не найден."
    TestRecord gauge = warpGauge And change < 0#, "поправка адресована продольному множителю", passed, failed, report
    field(gauge) = field(gauge) + change
    operators.StartEquations field, external, residual, tangent, False
    For i = 1 To count: norm = GeomMax(norm, Abs(residual(i))): Next i
    TestRecord norm <= TOLERANCE, "пересборка проходит силовые уравнения и связи", passed, failed, report

    residual(1) = 1.1 * TOLERANCE
    TestRecord Not operators.TryWarpGaugeCorrection(residual, TOLERANCE, gauge, change), _
        "множитель не исправляет независимое уравнение", passed, failed, report
    residual(1) = 0#: residual(first) = 1.5 * TOLERANCE: residual(second) = -1.5 * TOLERANCE
    TestRecord Not operators.TryWarpGaugeCorrection(residual, TOLERANCE, gauge, change), _
        "несовместимые поправки к одному множителю отклоняются", passed, failed, report

    Dim history As New Collection, trial() As Double, forces() As Double, stored() As Double, factor As Double
    field(warpGauge) = initialValue: history.Add field
    If Not solver.ReplayAcceptedFields(history, 1) Then Err.Raise 5, , "Не воспроизведено поле для пробы решателя."
    direction(1) = 21600#: control(1) = 1#
    If Not solver.TryDeformationStep(baseLoad, direction, control, 0#, 0#, factor, TOLERANCE, 1E-8, 1, consistentControl:=True) Then _
        Err.Raise 5, , "Штатный решатель не принял поправку. " & solver.ResultMeta.DiagnosticDetails
    TestRecord InStr(1, solver.ResultMeta.DiagnosticDetails, "Итерации: 0", vbBinaryCompare) > 0, _
        "поправка принята до шага Ньютона", passed, failed, report
    solver.CopyTrial trial, forces
    Dim changed As Long
    For i = 1 To count
        If trial(i) <> field(i) Then changed = changed + 1
    Next i
    TestRecord changed = 1 And trial(warpGauge) <> initialValue, "изменён только продольный множитель", passed, failed, report
    norm = 0#
    For i = 1 To 6: norm = GeomMax(norm, Abs(forces(i))): Next i
    TestRecord norm = 0# And factor = 0#, "шесть усилий и множитель нагрузки сохранены", passed, failed, report
    solver.CopyAccepted stored
    TestRecord stored(warpGauge) = initialValue, "проба сохраняет принятую историю", passed, failed, report
    solver.AcceptStep
    solver.CopyAccepted stored
    TestRecord stored(warpGauge) = trial(warpGauge), "принятое значение совпадает с проверенной пробой", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет достижение предельной деформации арматуры до и после пика N.
' Поздний предел завершает допустимую ветвь, а результат относится к уточнённому большему максимуму.
Public Function RunTestCoupledDeformationSteelCrossing() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    TestCoupledDeformationSteelCrossing 0.00125, 410000#, 396328.125, True, False, passed, failed, report
    TestCoupledDeformationSteelCrossing 0.0008, 450000#, 415920#, False, False, passed, failed, report
    TestCoupledDeformationSteelCrossing 0.0008, 415920#, 415920#, False, True, passed, failed, report
    RunTestCoupledDeformationSteelCrossing = report & "TOTAL_COUPLED_DEFORMATION_STEEL_CROSSING: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledDeformationSteelCrossing = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет определяющий максимум закрытой арматурным событием ветви выше, ниже и около λ=1.
' Контрольное кручение имеет постоянную и масштабируемую части; численный отказ до события не закрывает ветвь.
Public Function RunTestCoupledClosedBranchPeak() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    TestCoupledDeformationSteelCrossing 0.00125, 410000#, 396328.125, True, False, passed, failed, report, torque:=200000#
    TestCoupledDeformationSteelCrossing 0.00125, 450000#, 396328.125, True, False, passed, failed, report, torque:=200000#
    TestCoupledDeformationSteelCrossing 0.00125, 440000#, 396328.125, True, False, passed, failed, report, torque:=200000#
    TestCoupledDeformationSteelCrossing 0.00125, 410000#, 396328.125, True, False, passed, failed, report, freezeAt:=0.00115, torque:=200000#
    RunTestCoupledClosedBranchPeak = report & "TOTAL_COUPLED_CLOSED_BRANCH_PEAK: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledClosedBranchPeak = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для прямоугольника 180×120 мм и центрального стержня 100 мм²
' N = 21500·(30000·ε − 10¹⁰·ε³) + 100·min(200000·ε, 100) в Н.
' После начала текучести постоянное усилие в стержне не сдвигает
' аналитический максимум N = 440000 Н при ε = 0,001.
Private Sub TestCoupledDeformationSteelCrossing(ByVal steelStrainLimit As Double, ByVal designForce As Double, _
        ByVal crossingForce As Double, ByVal afterPeak As Boolean, ByVal atDesignLoad As Boolean, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String, Optional ByVal freezeAt As Double = 0#, _
        Optional ByVal torque As Double = 0#)
    Const STEEL_AREA As Double = 100#
    Const NET_CONCRETE_AREA As Double = 21500#
    Const STEEL_MODULUS As Double = 200000#
    Const STEEL_YIELD As Double = 100#
    Const LAMBDA_TOLERANCE As Double = 0.0001
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, Sqr(4# * STEEL_AREA / GEOM_PI), STEEL_AREA, "A400"
    Dim steel As New CCoupledSteelModel, models As Object
    steel.Initialize STEEL_MODULUS, STEEL_YIELD, STEEL_YIELD, 0#, 0#, steelStrainLimit, steelStrainLimit
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "A400", steel
    Dim operators As New CCoupledElasticState, concrete As New CTestCoupledPointResponse
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModels:=models
    concrete.TestInitialize operators.PointCount
    concrete.TestUseDescendingAxialLaw
    concrete.TestFreezeAxialTailAt = freezeAt
    Dim material As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    material.Initialize operators, concrete, models
    solver.Initialize operators, material, 30000#

    Dim baseLoad(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    Dim history As New Collection, trace As New CCoupledLoadTrace, field() As Double, resultants() As Double
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#: direction(1) = designForce
    direction(4) = torque
    If torque <> 0# Then baseLoad(4) = 30000#
    Dim strain As Variant, eps As Double, lambda As Double, previousLambda As Double, initialLambda As Double
    Dim stress As Double, expected As Double
    For Each strain In Array(0.0004, 0.0006)
        eps = CDbl(strain)
        If Not solver.TryDeformationStep(baseLoad, direction, control, eps, lambda, initialLambda, 1E-8, 1E-8, 30) Then _
            Err.Raise 5, , "Could not seed the analytical branch. " & solver.ResultMeta.DiagnosticDetails
        solver.CopyTrial field, resultants
        solver.AcceptStep
        history.Add field
        previousLambda = lambda: lambda = initialLambda
        trace.Add 1#, lambda
        stress = STEEL_MODULUS * eps
        If stress > STEEL_YIELD Then stress = STEEL_YIELD
        expected = (NET_CONCRETE_AREA * (30000# * eps - 1E+10# * eps ^ 3) + STEEL_AREA * stress) / designForce
        TestExpectNear lambda, expected, 1E-8, "analytical seed at strain " & CStr(eps), passed, failed, report
    Next strain

    Dim branch As New CCoupledDeformationContinuation, meta As CResultMeta
    branch.Trace solver, material, baseLoad, direction, history, previousLambda, lambda, resultants, _
        1.45, LAMBDA_TOLERANCE, 30, 1E-8, 1E-8, trace
    Set meta = branch.ResultMeta
    Dim low As Double, high As Double, point As Long, limitFields As Collection
    Dim peak As Double, accuracy As Double, designVerified As Boolean
    If freezeAt > 0# Then
        TestRecord meta.InternalStatus = rsNumericalFailure And branch.SteelEventPoint = 0, _
            "numerical tail failure supplies no physical closing event " & meta.DiagnosticDetails, passed, failed, report
        TestRecord Not branch.CopyPeak(peak, accuracy, designVerified) And Not branch.CopySteelLimit(low, high, point), _
            "an open reinforced branch supplies no governing capacity", passed, failed, report
        TestRecord InStr(1, meta.ResultComment, "локальный максимум", vbTextCompare) > 0, _
            "a verified local maximum remains diagnostic before physical closure", passed, failed, report
        TestRecord material.AcceptedSteelUtilization < 1#, "numerical failure is below the steel strain limit", passed, failed, report
        Exit Sub
    End If
    expected = crossingForce / designForce
    TestRecord branch.SteelEventPoint = operators.ConcretePointCount + 1, "central bar controls the localized event", passed, failed, report
    If afterPeak Then
        TestRecord meta.InternalStatus = rsSuccess And branch.CopyPeak(peak, accuracy, designVerified), _
            "physical closure certifies the refined preceding peak " & meta.DiagnosticDetails, passed, failed, report
        If peak <= 0# Then Exit Sub
        TestRecord accuracy > 0# And accuracy <= LAMBDA_TOLERANCE And Abs(peak - 440000# / designForce) <= accuracy, _
            "peak resolution covers the independent maximum at the requested tolerance", passed, failed, report
        TestRecord Abs(branch.BestAdmissibleLambda - peak) <= accuracy, _
            "no later admissible equilibrium exceeds the certified maximum", passed, failed, report
        TestRecord Not branch.CopySteelLimit(low, high, point) And branch.CopyLimitState(limitFields, resultants), _
            "late steel limit returns the earlier governing maximum state", passed, failed, report
        TestRecord InStr(1, meta.ResultComment, "завершённой предельной деформацией арматуры", vbTextCompare) > 0, _
            "comment identifies physical closure and its separate maximum", passed, failed, report
        TestRecord InStr(1, meta.DiagnosticDetails, "Арматурное событие: lambda до = ", vbTextCompare) > 0 And _
            InStr(1, meta.DiagnosticDetails, "Координата деформации", vbTextCompare) > 0, _
            "diagnostics resolve the late steel event in deformation coordinates", passed, failed, report
        Dim capacity As New CCapacityResult, vector(1 To 6) As Double, expectedStatus As String
        If designForce < 440000# Then
            expectedStatus = "OK"
        ElseIf designForce > 440000# Then
            expectedStatus = "FAIL"
        Else
            expectedStatus = "NumFail"
        End If
        capacity.InitializeCoupledPeak peak, accuracy, designVerified, vector, "Auto", meta.DiagnosticDetails
        TestRecord capacity.Status = expectedStatus, "capacity classifies the known maximum as " & expectedStatus, passed, failed, report
        TestRecord designForce >= 440000# Or designVerified, "a safe design load is verified before the governing maximum", passed, failed, report
        TestRecord designForce <= 440000# Or Not designVerified, "a load above the known maximum is not falsely verified", passed, failed, report
        If Not limitFields Is Nothing Then
            field = limitFields.Item(limitFields.Count)
            TestRecord Abs(field(1) - 0.001) <= 0.00001 And field(1) < steelStrainLimit, _
                "saved maximum strain precedes the closing steel rupture", passed, failed, report
            TestRecord Abs(resultants(4) - baseLoad(4) - peak * direction(4)) <= 0.00001, _
                "the same peak state balances the fixed and scaled torque", passed, failed, report
            TestRecord solver.ReplayAcceptedFields(limitFields, limitFields.Count), "governing peak prefix replays", passed, failed, report
            TestRecord material.AcceptedSteelUtilization <= 1# And concrete.TestAcceptedPeak < steelStrainLimit, _
                "replayed governing history excludes the descending closing tail", passed, failed, report
        Else
            TestRecord False, "closed physical branch must retain its maximum state", passed, failed, report
        End If
        expected = 440000# / designForce
    ElseIf atDesignLoad Then
        TestRecord meta.InternalStatus = rsNumericalFailure And meta.ResultCode = rcNumericalFailure, _
            "event interval containing exact design load stays unresolved", passed, failed, report
        TestRecord Not branch.CopySteelLimit(low, high, point) And Not branch.CopyLimitState(limitFields, resultants), _
            "uncertainty at design load supplies no governing capacity state", passed, failed, report
        TestRecord InStr(1, meta.ResultComment, "охватывает расчётную нагрузку", vbTextCompare) > 0, _
            "comment identifies the event near the design load: " & meta.ResultComment, passed, failed, report
    Else
        TestRecord meta.InternalStatus = rsCheckFailed And meta.ResultCode = rcPhysicalLimitExceeded, _
            "steel limit before the peak gives a physical failure", passed, failed, report
        TestRecord branch.CopySteelLimit(low, high, point), "rising steel limit has a bracket", passed, failed, report
        TestRecord low <= expected + 1E-8 And high >= expected - 1E-8 And high - low <= LAMBDA_TOLERANCE, _
            "rising steel bracket contains the independent analytical load", passed, failed, report
        TestRecord point = operators.ConcretePointCount + 1, "central bar controls the rising steel limit", passed, failed, report
        If branch.CopyLimitState(limitFields, resultants) Then
            field = limitFields(limitFields.Count)
            TestExpectNear resultants(1), low * designForce, 0.01, "saved rising-limit equilibrium", passed, failed, report
            TestRecord field(1) <= steelStrainLimit + 1E-10 And field(1) >= steelStrainLimit - 0.000001, _
                "saved field is within the admissible side of the strain interval", passed, failed, report
        Else
            TestRecord False, "rising steel limit has no saved state", passed, failed, report
        End If
    End If
    Dim saved As CCoupledLoadTrace, offsetFraction As Double, tracedLambda As Double
    Set saved = branch.SavedTrace
    If saved Is Nothing Then
        TestRecord False, "traced branch is missing", passed, failed, report
    ElseIf saved.Count <= 2 Then
        TestRecord False, "traced branch contains only the seed fields", passed, failed, report
    Else
        saved.ReadPoint saved.Count, offsetFraction, tracedLambda
        TestExpectNear tracedLambda, expected, LAMBDA_TOLERANCE + 1E-8, "trace ends at the governing limit", passed, failed, report
        TestRecord offsetFraction = 1# And (afterPeak And Abs(tracedLambda - expected) <= accuracy Or _
            Not afterPeak And tracedLambda <= expected + 1E-8), "trace retains the admissible side of the crossing", passed, failed, report
    End If
    TestRecord material.AcceptedSteelUtilization <= 1# + 1E-10, "no inadmissible material state committed", passed, failed, report
    ' Сохраняемый след и поля должны описывать все действительно принятые подшаги локализации, включая численно неопределённый исход.
    Dim retained As Collection, acceptedField() As Double, sameField As Boolean, coordinate As Long
    Set retained = branch.AcceptedFields
    If retained Is Nothing Or saved Is Nothing Then
        TestRecord False, "accepted steel-refinement state history is available", passed, failed, report
    Else
        TestRecord retained.Count = saved.Count, "accepted fields align with the complete load trace", passed, failed, report
        field = retained(retained.Count): solver.CopyAccepted acceptedField
        sameField = (UBound(field) = UBound(acceptedField))
        If sameField Then
            For coordinate = 1 To UBound(field)
                If field(coordinate) <> acceptedField(coordinate) Then sameField = False: Exit For
            Next coordinate
        End If
        TestRecord sameField, "last saved field exactly matches the committed solver state", passed, failed, report
        stress = STEEL_MODULUS * field(1)
        If stress > STEEL_YIELD Then stress = STEEL_YIELD
        expected = NET_CONCRETE_AREA * (30000# * field(1) - 1E+10# * field(1) ^ 3) + STEEL_AREA * stress
        saved.ReadPoint saved.Count, offsetFraction, tracedLambda
        TestRecord Abs(expected - tracedLambda * designForce) <= 0.01, _
            "last retained field independently equilibrates its saved load actual=" & CStr(expected) & _
            " expected=" & CStr(tracedLambda * designForce), passed, failed, report
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет направленное замещение бетона, выбор разных классов и откат стали.
' Бетон здесь аналитический, чтобы независимо вычислить поправки.
Public Function RunTestCoupledReinforcedResponse() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, 12#, GEOM_PI * 36#, "A400"
    Dim builder As New CSectionDefinitionBuilder, transverse As New CTransverseRebarLayout
    builder.Initialize "TestCoupledReinforcedResponse"
    transverse.AddGroup builder.LinePath("TestBranch", -40#, 25#, 50#, 25#), 8#, 80#, "A240"
    Dim operators As New CCoupledElasticState
    operators.PrepareModel section, 30000#, 0#, 1, 1, 8, transverse, 200000#
    Dim longSteel As New CCoupledSteelModel, transverseSteel As New CCoupledSteelModel, models As Object
    longSteel.Initialize 200000#, 400#, 350#, 1500#, 1000#, 0.025, 0.018
    transverseSteel.Initialize 200000#, 240#, 240#, 2000#, 0#, 0.02, 0.02
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "A400", longSteel: models.Add "A240", transverseSteel
    Dim concrete As New CTestCoupledPointResponse
    concrete.TestInitialize operators.PointCount
    Dim material As New CCoupledReinforcedResponse, response As ICoupledPointResponse
    material.Initialize operators, concrete, models
    Set response = material
    Dim strain(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim firstRebar As Long, firstHoop As Long, point As Long, resultCode As ESolverFailureCode
    firstRebar = operators.ConcretePointCount + 1: firstHoop = firstRebar + 1
    TestRecord operators.SteelClassAt(firstRebar) = "A400" And operators.SteelClassAt(firstHoop) = "A240", _
        "distinct material classes follow geometry", passed, failed, report

    strain(1) = 0.003: strain(3) = 0.004: strain(5) = 0.001
    For point = 1 To operators.PointCount
        resultCode = response.EvaluatePoint(point, strain, stress, tangent)
        If resultCode <> sfcNone Then GoTo Done
        Select Case point
            Case 1
                TestExpectNear stress(3), 760#, 1E-10, "concrete integrated once", passed, failed, report
            Case firstRebar
                TestExpectNear stress(3), 403# - 760#, 1E-10, "longitudinal steel minus concrete", passed, failed, report
                TestExpectNear stress(5), -15#, 1E-10, "longitudinal steel adds no transverse shear stiffness", passed, failed, report
                TestExpectNear tangent(3, 3), 1500# - 510000#, 1E-10, "tangent of replacement", passed, failed, report
            Case firstHoop
                TestExpectNear stress(1), 243.6 - 90#, 1E-10, "transverse projected strain and separate steel", passed, failed, report
                TestExpectNear stress(3), -760#, 1E-10, "transverse bar replaces concrete without axial steel", passed, failed, report
        End Select
    Next point
    TestExpectNear material.TrialSteelUtilization, 0.004 / 0.025, 1E-10, "maximum steel trial utilization", passed, failed, report
    TestRecord material.AcceptedSteelUtilization = 0#, "trial leaves accepted utilization untouched", passed, failed, report
    strain(1) = 0.03 ' Возмущение с превышением поперечного предела не записывается в базовое состояние.
    resultCode = response.EvaluatePoint(firstHoop, strain, stress, tangent, recordTrial:=False)
    TestRecord resultCode = sfcNone And material.TrialSteelCriticalPoint = firstRebar And _
        material.TrialSteelUtilization = 0.004 / 0.025, "directional probe leaves the base steel limit intact", passed, failed, report
    strain(1) = 0.003
    response.AcceptTrial
    TestExpectNear material.AcceptedSteelUtilization, 0.004 / 0.025, 1E-10, "accepted utilization", passed, failed, report
    TestExpectNear material.AcceptedSteelStress(firstHoop), 243.6, 1E-10, "accepted hoop stress belongs to the base trial", passed, failed, report

    strain(3) = 0#: resultCode = response.EvaluatePoint(firstRebar, strain, stress, tangent)
    TestRecord resultCode = sfcNone And Abs(stress(3)) > 300#, "unloading remembers accepted plastic history", passed, failed, report
    strain(3) = 0.0035: resultCode = response.EvaluatePoint(firstRebar, strain, stress, tangent)
    TestExpectNear stress(3), 303# - 533.75, 1E-10, "rejecting reverse trial preserves accepted steel", passed, failed, report
    response.ClearHistory
    strain(3) = 0.001: resultCode = response.EvaluatePoint(firstRebar, strain, stress, tangent)
    TestExpectNear stress(3), 160#, 1E-10, "new combination resets steel plastic history", passed, failed, report
Done:
    TestRecord resultCode = sfcNone, "point evaluations converged", passed, failed, report
    RunTestCoupledReinforcedResponse = report & "TOTAL_COUPLED_REINFORCED_RESPONSE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledReinforcedResponse = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет главные напряжения, границу текучести и градиент потенциала.
' Поворот, невырожденное смешанное состояние и гидростатика заданы независимо.
Public Function RunTestCoupledConcreteTensor() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const E As Double = 30000#
    Const NU As Double = 0.2
    Dim tensor(1 To 6, 1 To 6) As Double, stress(1 To 6) As Double, flow(1 To 6) As Double
    Dim trace As Double, q As Double, high As Double, mid As Double, low As Double
    CoupledElasticTensor E, NU, tensor
    TestExpectNear tensor(1, 1), 33333.3333333333, 1E-12, "isotropic normal stiffness", passed, failed, report
    TestExpectNear tensor(1, 2), 8333.33333333333, 1E-12, "isotropic Poisson stiffness", passed, failed, report
    TestExpectNear tensor(6, 6), 12500#, 1E-12, "engineering shear stiffness", passed, failed, report
    stress(1) = -20#: stress(2) = -20#: stress(3) = -20#
    CoupledStressInvariants stress, trace, q, high, mid, low
    TestRecord trace = -60# And q = 0# And high = -20# And mid = -20# And low = -20#, _
        "hydrostatic values exact", passed, failed, report
    stress(1) = 0#: stress(2) = 0#: stress(3) = 0#: stress(6) = 9#
    CoupledStressInvariants stress, trace, q, high, mid, low
    TestExpectNear q, Sqr(3#) * 9#, 1E-12, "physical shear contribution to invariant", passed, failed, report
    TestRecord Abs(high - 9#) < 1E-12 And Abs(mid) < 1E-12 And Abs(low + 9#) < 1E-12, _
        "pure shear has opposite principal stresses", passed, failed, report
    stress(1) = -4#: stress(2) = -4#: stress(3) = -2#: stress(4) = 6#: stress(6) = 0#
    CoupledStressInvariants stress, trace, q, high, mid, low
    TestRecord Abs(high - 2#) < 1E-12 And Abs(mid + 2#) < 1E-12 And Abs(low + 10#) < 1E-12, _
        "rotated principal directions preserve spectrum", passed, failed, report
    TestRecord Abs(CoupledBarcelonaYield(-30#, 30#, 0#, 30#, 3#, 0.12, 3#)) < 1E-12, _
        "uniaxial compression at specified cohesion", passed, failed, report
    TestRecord Abs(CoupledBarcelonaYield(3#, 3#, 3#, 30#, 3#, 0.12, 3#)) < 1E-12, _
        "uniaxial tension at specified cohesion", passed, failed, report

    Const DELTA As Double = 0.00001
    Dim otherFlow(1 To 6) As Double, original As Double, plus As Double, minus As Double, value As Double, i As Long
    stress(5) = 2.5: stress(6) = -1.25
    CoupledStressInvariants stress, trace, q, high, mid, low
    value = CoupledConcretePotential(stress, trace, q, 0.32, 3#, 0.1, flow)
    For i = 1 To 6
        original = stress(i)
        stress(i) = original + DELTA
        CoupledStressInvariants stress, trace, q, high, mid, low
        plus = CoupledConcretePotential(stress, trace, q, 0.32, 3#, 0.1, otherFlow)
        stress(i) = original - DELTA
        CoupledStressInvariants stress, trace, q, high, mid, low
        minus = CoupledConcretePotential(stress, trace, q, 0.32, 3#, 0.1, otherFlow)
        stress(i) = original
        TestRecord Abs(flow(i) - (plus - minus) / (2# * DELTA)) < 1E-8, _
            "potential derivative " & CStr(i), passed, failed, report
    Next i
    RunTestCoupledConcreteTensor = report & "TOTAL_COUPLED_CONCRETE_TENSOR: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConcreteTensor = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Контролирует обратный Эйлер при фиксированных параметрах поверхности.
' Чистый сдвиг и гидростатическая вершина имеют независимый явный ответ.
Public Function RunTestCoupledConcreteReturn() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const E As Double = 30000#
    Const NU As Double = 0.2
    Const FC As Double = 30#
    Const FT As Double = 3#
    Const ALPHA As Double = 0.12
    Const GAMMA As Double = 3#
    Dim update As New CCoupledConcreteReturn
    update.Initialize E, NU, ALPHA, GAMMA, 0.1
    Dim elastic(1 To 6) As Double, corrected(1 To 6) As Double, increment(1 To 6) As Double
    Dim tensor(1 To 6, 1 To 6) As Double, flow(1 To 6) As Double
    Dim multiplier As Double, multiplierDivisor As Double, value As Double, trace As Double, q As Double, high As Double, mid As Double, low As Double
    Dim i As Long, j As Long, caseNo As Long, dilation As Double
    Dim beta As Double: beta = FC / FT * (1# - ALPHA) - (1# + ALPHA)
    Dim ok As Boolean
    elastic(3) = 1#
    ok = TestConcreteReturn(update, elastic, FC, FT, 0#, corrected, increment, multiplier, multiplierDivisor)
    TestRecord ok And corrected(3) = 1# And multiplier = 0# And multiplierDivisor = 1# And increment(3) = 0#, _
        "elastic trial stays unchanged", passed, failed, report
    CoupledElasticTensor E, NU, tensor
    For caseNo = 1 To 6
        For i = 1 To 6: elastic(i) = 0#: Next i
        dilation = 0.3
        Select Case caseNo
            Case 1: elastic(3) = -50#
            Case 2: elastic(3) = 12#
            Case 3: elastic(6) = 10#: dilation = 0#
            Case 4: elastic(1) = 9#: elastic(2) = 9#: elastic(3) = 9#
            Case 5: elastic(1) = -48#: elastic(2) = 5#: elastic(3) = -13#: elastic(4) = 8#: elastic(5) = -11#
            Case 6
                elastic(1) = -28#: elastic(2) = -28#: elastic(3) = -28#: elastic(6) = 35#: dilation = 0#
        End Select
        ok = TestConcreteReturn(update, elastic, FC, FT, dilation, corrected, increment, multiplier, multiplierDivisor)
        TestRecord ok And multiplier > 0# And multiplierDivisor = 1#, "return " & CStr(caseNo), passed, failed, report
        If ok Then
            CoupledStressInvariants corrected, trace, q, high, mid, low
            value = CoupledBarcelonaYield(trace, q, high, FC, FT, ALPHA, GAMMA)
            TestRecord Abs(value) < 4E-9, "yield consistency " & CStr(caseNo), passed, failed, report
            value = CoupledConcretePotential(corrected, trace, q, dilation, FT, 0.1, flow)
            Dim residual As Double, maxBackward As Double, maxFlow As Double
            maxBackward = 0#: maxFlow = 0#
            For i = 1 To 6
                residual = corrected(i) - elastic(i)
                For j = 1 To 6: residual = residual + tensor(i, j) * increment(j): Next j
                If Abs(residual) > maxBackward Then maxBackward = Abs(residual)
                residual = Abs(increment(i) - multiplier * flow(i))
                If residual > maxFlow Then maxFlow = residual
            Next i
            TestRecord maxBackward < 1E-10 And maxFlow < 1E-12, _
                "backward Euler equations " & CStr(caseNo), passed, failed, report
            If caseNo = 3 Then TestExpectNear corrected(6), FC * (1# - ALPHA) / (Sqr(3#) + beta), _
                1E-9, "pure shear analytical return", passed, failed, report
            If caseNo = 4 Then TestExpectNear corrected(1), FC * (1# - ALPHA) / (3# * ALPHA + beta), _
                1E-9, "hydrostatic analytical return", passed, failed, report
            If caseNo = 6 Then TestExpectNear corrected(6), (FC * (1# - ALPHA) + 28# * (3# * ALPHA + GAMMA)) / _
                (Sqr(3#) + GAMMA), 1E-9, "closed-form return crosses to compressive principal branch", passed, failed, report
        End If
    Next caseNo
    For i = 1 To 3: elastic(i) = 9#: elastic(i + 3) = 0#: Next i
    ok = TestConcreteReturn(update, elastic, FC, FT, 0#, corrected, increment, multiplier, multiplierDivisor)
    TestRecord Not ok, "hydrostatic tension without volumetric flow reports inability to return", passed, failed, report
    For i = 1 To 3: elastic(i) = 4.5: Next i
    elastic(6) = 0.025
    Dim smallDilation As Variant, apexStress As Double, apexVolume As Double
    apexStress = FC * (1# - ALPHA) / (3# * ALPHA + beta)
    apexVolume = (elastic(1) - apexStress) / (E / (1# - 2# * NU))
    For Each smallDilation In Array(0.000001, 1E-170)
        dilation = CDbl(smallDilation)
        ok = TestConcreteReturn(update, elastic, FC, FT, dilation, corrected, increment, multiplier, multiplierDivisor)
        TestRecord ok And multiplier > 0# And multiplierDivisor = 1#, "near-apex return for dilation " & CStr(dilation), passed, failed, report
        If ok Then
            TestExpectNear corrected(1), apexStress, 1E-9, "near-apex analytical pressure", passed, failed, report
            TestExpectNear (dilation / multiplierDivisor) * multiplier, apexVolume, 1E-12, "near-apex volumetric flow", passed, failed, report
            TestExpectNear tensor(6, 6) * increment(6) + corrected(6), elastic(6), 1E-12, _
                "near-apex engineering shear balance", passed, failed, report
        End If
    Next smallDilation
    ok = TestConcreteReturn(update, elastic, FC, FT, 0#, corrected, increment, multiplier, multiplierDivisor)
    TestRecord Not ok, "near-apex deviator does not enable volumetric flow at zero dilation", passed, failed, report
    dilation = 1E-300
    dilation = dilation * 3.822077E-18 ' Субнормальное положительное число строится при выполнении, без литерала ниже нормального диапазона.
    TestRecord dilation > 0# And dilation < 2.22507385850721E-308, "subnormal dilation remains positive", passed, failed, report
    Dim trialShear As Variant, apexTag As String
    Dim maxPressureError As Double, maxVolumeError As Double, maxShearStress As Double, maxShearIncrementError As Double
    For Each trialShear In Array(0.025, 0#)
        elastic(6) = CDbl(trialShear)
        apexTag = "hydrostatic apex"
        If elastic(6) <> 0# Then apexTag = "scaled deviator apex"
        CoupledStressInvariants elastic, trace, q, high, mid, low
        TestRecord Abs(q - Sqr(3#) * elastic(6)) < 1E-14 And (q > 0#) = (elastic(6) > 0#), _
            apexTag & " starts in the intended branch", passed, failed, report
        ok = TestConcreteReturn(update, elastic, FC, FT, dilation, corrected, increment, multiplier, multiplierDivisor)
        TestRecord ok And multiplier > 0# And multiplierDivisor > 0# And multiplierDivisor < 1#, _
            apexTag & " uses a scaled multiplier", passed, failed, report
        If ok And multiplierDivisor > 0# Then
            TestExpectNear (dilation / multiplierDivisor) * multiplier, apexVolume, 1E-12, _
                apexTag & " analytical volumetric flow", passed, failed, report
            maxPressureError = 0#: maxVolumeError = 0#: maxShearStress = 0#: maxShearIncrementError = 0#
            For i = 1 To 3
                maxPressureError = GeomMax(maxPressureError, Abs(corrected(i) - apexStress))
                maxVolumeError = GeomMax(maxVolumeError, Abs(increment(i) - apexVolume))
                maxShearStress = GeomMax(maxShearStress, Abs(corrected(i + 3)))
                maxShearIncrementError = GeomMax(maxShearIncrementError, _
                    Abs(tensor(i + 3, i + 3) * increment(i + 3) + corrected(i + 3) - elastic(i + 3))) ' Сдвиговая компонента деформации инженерная.
            Next i
            TestRecord maxPressureError < 1E-9 And maxVolumeError < 1E-12, _
                apexTag & " analytical pressure and normal increments", passed, failed, report
            TestRecord maxShearStress < 1E-12 And maxShearIncrementError < 1E-12, _
                apexTag & " constitutive engineering shears", passed, failed, report
            CoupledStressInvariants corrected, trace, q, high, mid, low
            value = CoupledBarcelonaYield(trace, q, high, FC, FT, ALPHA, GAMMA)
            TestRecord Abs(value) < 4E-9 And q < 1E-12 And Abs(high - low) < 1E-12, _
                apexTag & " yield and hydrostatic invariants", passed, failed, report
        End If
    Next trialShear
    RunTestCoupledConcreteReturn = report & "TOTAL_COUPLED_CONCRETE_RETURN: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConcreteReturn = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит известный тензор упругой пробы для изолированной проверки возврата.
Private Function TestConcreteReturn(ByVal update As CCoupledConcreteReturn, ByRef elastic() As Double, _
        ByVal fc As Double, ByVal ft As Double, ByVal dilation As Double, ByRef corrected() As Double, _
        ByRef increment() As Double, ByRef multiplier As Double, ByRef multiplierDivisor As Double) As Boolean
    Dim trace As Double, q As Double, maximum As Double, middle As Double, minimum As Double, deviatorFactor As Double
    CoupledStressInvariants elastic, trace, q, maximum, middle, minimum
    TestConcreteReturn = update.Correct(elastic, trace, q, maximum, fc, ft, ft, dilation, _
        corrected, increment, multiplier, deviatorFactor, multiplierDivisor)
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет заданный пик, связь эффективной прочности с повреждением
' и энергию двух пластических ветвей независимой квадратурой трапециями.
Public Function RunTestCoupledEnvelope() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim envelope As New CCoupledConcreteEnvelope
    Dim stress As Double, effective As Double, damage As Double, energy As Double
    Const PEAK_PLASTIC As Double = 0.0009
    envelope.Initialize 15#, 30#, PEAK_PLASTIC, 80#, 150#, 0.27, PEAK_PLASTIC
    envelope.Evaluate 0#, stress, effective, damage, energy
    TestRecord stress = 15# And effective = 15# And damage = 0# And energy = 0#, _
        "compression yield and virgin modulus", passed, failed, report
    envelope.Evaluate PEAK_PLASTIC, stress, effective, damage, energy
    TestExpectNear stress, 30#, 1E-12, "calibrated compression peak", passed, failed, report
    TestExpectNear damage, 0.27, 1E-12, "calibrated cohesion degradation", passed, failed, report
    TestExpectNear effective * (1# - damage), stress, 1E-12, "compression effective strength identity", passed, failed, report
    TestExpectNear energy, (15# + 30#) * PEAK_PLASTIC * 0.5 * 150# / 80#, 1E-12, _
        "analytic prepeak dissipated energy", passed, failed, report
    envelope.Evaluate envelope.HalfStrengthPlasticStrain, stress, effective, damage, energy
    TestExpectNear stress, 15#, 1E-12, "half compression descending strength", passed, failed, report

    Const TENSION_ENERGY As Double = 0.12
    Const LENGTH As Double = 120#
    Const FT As Double = 3#
    Dim half As Double: half = Log(2#) * TENSION_ENERGY / (FT * LENGTH)
    envelope.Initialize FT, FT, 0#, TENSION_ENERGY, LENGTH, 0.5, half
    envelope.Evaluate half, stress, effective, damage, energy
    TestExpectNear stress, 0.5 * FT, 1E-12, "half tensile stress", passed, failed, report
    TestExpectNear damage, 0.5, 1E-12, "half tensile stress damage", passed, failed, report
    TestExpectNear effective, FT, 1E-12, "tensile effective cohesion from calibrated damage", passed, failed, report
    TestExpectNear energy, 0.5, 1E-12, "half tensile fracture work", passed, failed, report

    Const N As Long = 4000
    Dim dx As Double, i As Long, area As Double
    dx = 16# * TENSION_ENERGY / (FT * LENGTH) / N
    For i = 0 To N
        envelope.Evaluate i * dx, stress, effective, damage, energy
        If i = 0 Or i = N Then stress = stress * 0.5
        area = area + stress * dx
    Next i
    TestExpectNear area * LENGTH, TENSION_ENERGY, 0.00001, "independent fracture work integration", passed, failed, report
    Dim invalid As Boolean
    On Error Resume Next
    envelope.Initialize 15#, 30#, PEAK_PLASTIC, 0.01, 150#, 0.27, PEAK_PLASTIC
    invalid = Err.Number = vbObjectError + 1782
    On Error GoTo Failed
    TestRecord invalid, "incompatible prepeak work and energy rejected", passed, failed, report
    RunTestCoupledEnvelope = report & "TOTAL_COUPLED_ENVELOPE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledEnvelope = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет физические истории бетона и касательную в одной точке.
' Данные для проверки задаются явно в МПа, Н/мм и мм.
Public Function RunTestCoupledConcreteResponse() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim material As New CCoupledConcreteResponse, response As ICoupledPointResponse
    material.Initialize 2, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    Set response = material
    Dim strain(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim et As Double, ec As Double, status As ESolverFailureCode
    strain(6) = 0.00004
    status = response.EvaluatePoint(1, strain, stress, tangent)
    TestRecord status = sfcNone, "virgin elastic concrete response", passed, failed, report
    TestExpectNear stress(6), 0.5, 1E-12, "virgin shear", passed, failed, report
    TestExpectNear tangent(6, 6), 12500#, 1E-12, "virgin shear tangent", passed, failed, report

    strain(6) = 0.0009
    status = response.EvaluatePoint(1, strain, stress, tangent)
    TestRecord status = sfcNone, "mixed plastic-damage response", passed, failed, report
    If status <> sfcNone Then GoTo Done
    Dim base(1 To 6) As Double, derivative As Double, i As Long
    For i = 1 To 6: base(i) = stress(i): Next i
    derivative = tangent(6, 6)
    TestRecord base(6) > 0# And base(6) < 0.0009 * 12500#, "shear affected by multi-axial damage", passed, failed, report
    Dim trialTangent(1 To 6, 1 To 6) As Double, maxDifference As Double, iColumn As Variant
    trialTangent(1, 1) = 123#
    status = response.EvaluatePoint(1, strain, stress, trialTangent, ctmAxialGradient)
    For Each iColumn In Array(3, 5, 6)
        For i = 1 To 6
            If Abs(trialTangent(i, iColumn) - tangent(i, iColumn)) > maxDifference Then _
                maxDifference = Abs(trialTangent(i, iColumn) - tangent(i, iColumn))
        Next i
    Next iColumn
    TestRecord status = sfcNone And maxDifference = 0# And trialTangent(1, 1) = 123# And stress(6) = base(6), _
        "gradient probe computes the same required tangent columns and stress", passed, failed, report
    trialTangent(6, 6) = 456#
    status = response.EvaluatePoint(1, strain, stress, trialTangent, ctmStressOnly)
    TestRecord status = sfcNone And trialTangent(6, 6) = 456# And stress(6) = base(6), _
        "stress-only probe has identical stress with no tangent calculation", passed, failed, report

    strain(6) = 0.00091
    status = response.EvaluatePoint(1, strain, stress, trialTangent, ctmStressOnly)
    TestRecord status = sfcNone, "nearby rejected probe uses its own concrete strain", passed, failed, report
    strain(6) = 0.0009
    status = response.EvaluatePoint(1, strain, stress, trialTangent, ctmAxialGradient)
    TestRecord status = sfcNone, "axial columns evaluate again after a different trial", passed, failed, report
    status = response.EvaluatePoint(1, strain, stress, trialTangent)
    Dim j As Long, identical As Boolean: identical = (status = sfcNone)
    For i = 1 To 6
        If stress(i) <> base(i) Then identical = False
        For j = 1 To 6
            If tangent(i, j) <> trialTangent(i, j) Then identical = False
        Next j
    Next i
    TestRecord identical, "partial-to-full trial reuse preserves all stresses and tangent entries exactly", passed, failed, report

    Const DELTA As Double = 0.00000006
    strain(6) = 0.0009 + DELTA
    status = response.EvaluatePoint(1, strain, stress, tangent)
    Dim plus As Double: plus = stress(6)
    TestRecord status = sfcNone, "shear perturbation plus", passed, failed, report
    strain(6) = 0.0009 - DELTA
    status = response.EvaluatePoint(1, strain, stress, tangent)
    Dim checkDerivative As Double: checkDerivative = (plus - stress(6)) / (2# * DELTA)
    TestRecord status = sfcNone And Abs(derivative - checkDerivative) < 0.2, _
        "independent tangent increment " & CStr(derivative) & " vs " & CStr(checkDerivative), passed, failed, report

    material.AcceptedPointDamage 1, et, ec
    TestRecord et = 0# And ec = 0#, "trial and tangent leave committed energy untouched", passed, failed, report
    strain(6) = 0.0009
    status = response.EvaluatePoint(1, strain, stress, tangent)
    TestRecord status = sfcNone And Abs(stress(6) - base(6)) < 1E-10, "repeat trial gives identical stress", passed, failed, report
    strain(6) = 0#
    status = response.EvaluatePoint(2, strain, stress, tangent)
    TestRecord status = sfcNone, "other concrete location remains elastic", passed, failed, report
    strain(6) = 0.00091
    status = response.EvaluatePoint(1, strain, stress, tangent, recordTrial:=False)
    TestRecord status = sfcNone And stress(6) <> base(6), "directional probe returns the tensile-shear perturbation", passed, failed, report
    response.AcceptTrial
    material.CopyAcceptedStress 1, stress
    For i = 1 To 6
        If stress(i) <> base(i) Then Exit For
    Next i
    TestRecord i = 7, "accepted concrete stress is unchanged by the directional probe", passed, failed, report
    material.AcceptedPointDamage 1, et, ec
    TestRecord et > 0# And et < 1# And ec > 0# And ec < 1#, _
        "both histories advance in shear " & CStr(et) & " " & CStr(ec), passed, failed, report
    Dim committedTension As Double, committedCompression As Double
    committedTension = et: committedCompression = ec
    material.AcceptedPointDamage 2, et, ec
    TestRecord et = 0# And ec = 0#, "virgin integration point remains independent", passed, failed, report

    strain(6) = 0.00088
    status = response.EvaluatePoint(1, strain, stress, tangent)
    TestRecord status = sfcNone And stress(6) < base(6), "unloading from cracked concrete", passed, failed, report
    strain(6) = 0#
    status = response.EvaluatePoint(2, strain, stress, tangent)
    response.AcceptTrial
    material.AcceptedPointDamage 1, et, ec
    TestRecord et = committedTension And ec = committedCompression, "unloading preserves both damage histories", passed, failed, report
    response.ClearHistory
    material.AcceptedPointDamage 1, et, ec
    TestRecord et = 0# And ec = 0#, "independent combination clears concrete history", passed, failed, report
    TestHydrostaticConcreteResponse passed, failed, report
Done:
    RunTestCoupledConcreteResponse = report & "TOTAL_COUPLED_CONCRETE_RESPONSE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConcreteResponse = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет изотропное растяжение при нулевой начальной дилатансии.
' Растягивающая вершина за пределом требует объёмного пластического приращения.
Private Sub TestHydrostaticConcreteResponse(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim material As New CCoupledConcreteResponse, response As ICoupledPointResponse
    material.Initialize 1, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    Set response = material
    Dim strain(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    strain(1) = 0.0001: strain(2) = strain(1): strain(3) = strain(1)
    Dim status As ESolverFailureCode
    status = response.EvaluatePoint(1, strain, stress, tangent, ctmStressOnly)
    TestRecord status = sfcNone, "hydrostatic tension reaches a constitutively consistent return", passed, failed, report
    If status <> sfcNone Then Exit Sub
    TestRecord stress(1) > 0# And stress(1) < 5# And stress(1) = stress(2) And stress(1) = stress(3) And _
        stress(4) = 0# And stress(5) = 0# And stress(6) = 0#, _
        "hydrostatic return preserves symmetry and reduces elastic pressure", passed, failed, report
    Dim beta As Double, apex As Double, plastic As Double, expectedStress As Double
    beta = 16# * 0.88 / 3# - 1.12
    apex = 16# * 0.88 / (3# * 0.12 + beta) ' При q=0 эффективное давление следует из поверхности текучести.
    plastic = strain(1) - apex / 50000# ' Изотропное растяжение даёт модуль E/(1-2ν).
    expectedStress = apex * Exp(-3# * 125# * plastic / 0.12)
    TestExpectNear stress(1), expectedStress, 0.0000001, "hydrostatic damaged stress matches the scalar reference", passed, failed, report
    Dim et As Double, ec As Double
    material.AcceptedPointDamage 1, et, ec
    TestRecord et = 0# And ec = 0#, "hydrostatic trial keeps accepted history unchanged", passed, failed, report
    response.AcceptTrial
    material.AcceptedPointDamage 1, et, ec
    TestRecord et > 0# And ec = 0#, "hydrostatic tension advances only tension history", passed, failed, report
End Sub


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пропускает реальный пространственный бетон и две модели стали через
' общий решатель. Сверяет все шесть усилий и необратимость при разгрузке.
Public Function RunTestCoupledConcreteSection() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse
    Dim solver As New CCoupledNewtonSolver
    TestPrepareCoupledConcreteSection operators, concrete, solver
    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    path.Initialize 0#, 26320#, 0#, 479400#, 0#, 169200#, 0#, 1410000#, 0#, 15980#, 0#, -6674#
    branch.Initialize solver, path, 4, 60, 30, 0.0000005, 0.0000005, True
    Dim ok As Boolean, startTime As Double
    startTime = Timer
    ok = branch.Advance(1#)
    TestRecord ok, "six-action concrete section " & branch.ResultMeta.ResultComment & " " & branch.ResultMeta.DiagnosticDetails, passed, failed, report
    report = report & "SOLVE_SECONDS: " & CStr(Timer - startTime) & vbCrLf
    If Not ok Then GoTo Done
    Dim field() As Double, resultants() As Double, target() As Double, i As Long
    branch.CopyAccepted field, resultants
    path.TargetSix 1#, target
    For i = 1 To 6
        TestExpectNear resultants(i), target(i), 0.00002, "concrete-integrated component " & CStr(i), passed, failed, report
    Next i
    Dim et As Double, ec As Double, maximum As Double, criticalPoint As Long
    For i = 1 To operators.ConcretePointCount
        concrete.AcceptedPointDamage i, et, ec
        If et > maximum Then maximum = et: criticalPoint = i
    Next i
    TestRecord maximum > 0#, "concrete damage develops at integration point " & CStr(criticalPoint), passed, failed, report
    If criticalPoint = 0 Then GoTo Done
    concrete.AcceptedPointDamage criticalPoint, et, ec
    Dim tensionAtTarget As Double, compressionAtTarget As Double
    tensionAtTarget = et: compressionAtTarget = ec
    ok = solver.TryStep(0.95 * target(1), 0.95 * target(2), 0.95 * target(3), 0.95 * target(4), _
        0.95 * target(5), 0.95 * target(6), 0.0000005, 0.0000005, 30)
    TestRecord ok, "six-action unloading gives equilibrium " & solver.ResultMeta.ResultComment, passed, failed, report
    If ok Then
        solver.AcceptStep
        concrete.AcceptedPointDamage criticalPoint, et, ec
        TestRecord et >= maximum, "concrete point cannot heal during unloading", passed, failed, report
    End If
    Dim history As New Collection, savedField() As Double
    For i = 1 To branch.SavedFieldCount
        branch.CopySavedField i, savedField
        history.Add savedField
    Next i
    TestRecord solver.ReplayAcceptedFields(history, history.Count), "replayed accepted concrete and steel fields", passed, failed, report
    concrete.AcceptedPointDamage criticalPoint, et, ec
    TestRecord et = tensionAtTarget And ec = compressionAtTarget, "replay reproduces the exact accepted concrete damage", passed, failed, report
Done:
    RunTestCoupledConcreteSection = report & "TOTAL_COUPLED_CONCRETE_SECTION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConcreteSection = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Снимает восстановленный сложный режим с увеличенной на 6,4% нагрузкой.
' При численном отказе проверяет сохранение реально достигнутой точки.
Public Function RunTestCoupledHardShear() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, solver As New CCoupledNewtonSolver
    TestPrepareCoupledConcreteSection operators, concrete, solver
    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    Dim loads As Variant: loads = TestHardShearLoads()
    path.Initialize 0#, loads(0), 0#, loads(1), 0#, loads(2), 0#, loads(3), 0#, loads(4), 0#, loads(5)
    branch.Initialize solver, path, 4, 60, 30, 0.0000005, 0.0000005
    Dim started As Double: started = Timer
    Dim solved As Boolean: solved = branch.Advance(1#)
    report = report & "HARD_SHEAR_SOLVED: " & CStr(solved) & " | " & branch.ResultMeta.ResultComment & " | " & _
        branch.ResultMeta.DiagnosticDetails & vbCrLf & "HARD_SHEAR_SECONDS: " & CStr(Timer - started) & vbCrLf
    TestRecord (solved And branch.AcceptedLambda = 1#) Or _
        (Not solved And branch.AcceptedLambda >= 0.94 And branch.ResultMeta.InternalStatus = rsNumericalFailure), _
        "resolved target or documented numerical limit without physical FAIL", passed, failed, report
    Dim field() As Double, actual() As Double, expected() As Double, i As Long
    branch.CopyAccepted field, actual
    path.TargetSix branch.AcceptedLambda, expected
    For i = 1 To 6
        TestExpectNear actual(i), expected(i), 0.00005, "accepted harder case component " & CStr(i), passed, failed, report
    Next i
    RunTestCoupledHardShear = report & "TOTAL_COUPLED_HARD_SHEAR: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledHardShear = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет предельную силу пространственного бетона с независимым A·fc.
' Малый крутящий момент включает совместный режим и слабо возмущает
' контрольную задачу однородного одноосного сжатия.
Public Function RunTestCoupledConcreteCompressionPeak() As String
    RunTestCoupledConcreteCompressionPeak = TestRunConcretePeak(-720000#, 30#, "COMPRESSION")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет отдельный физический максимум растянутого бетона под N и T.
Public Function RunTestCoupledConcreteTensionPeak() As String
    RunTestCoupledConcreteTensionPeak = TestRunConcretePeak(72000#, 3#, "TENSION")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Собирает одно и то же сечение для двух предельных механизмов бетона.
Private Function TestRunConcretePeak(ByVal force As Double, ByVal independentStrength As Double, _
        ByVal mode As String) As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#
    Const FC As Double = 30#
    Const TORQUE As Double = 4000#
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim layout As New CTransverseRebarLayout: layout.DeclareNone
    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    concrete.Initialize FC, 3#, FC, 3#, 30000#, 30000#, rbMc2:=15#
    steel.Initialize 350#, 400#, 350#, 400#, 200000#, 200000#
    Dim provider As New CMaterialModelProvider
    provider.TestInitializeFromParameters concrete, steel
    Dim analysis As New CCoupledSectionAnalysis, spec As New CMaterialModelSpec
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / FC, 0.002, "A400", 30
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    Dim loads As New CSectionLoadState
    loads.Initialize FORCE, 0#, 0#, 0#, 0#
    loads.SetSpatialComponents TORQUE, 0#, 0#
    Dim started As Double, capacity As CCapacityResult
    started = Timer
    Set capacity = analysis.SolveCapacity(loads, "Auto", spec, 0.5, 1.25, 0.006, 30)
    report = report & "PEAK_SECONDS: " & CStr(Timer - started) & " | " & capacity.ResultMeta.ResultComment & " | " & _
        capacity.ResultMeta.DiagnosticDetails & vbCrLf
    TestRecord capacity.HasUltimatePoint And capacity.PeakLambda > 0#, "identifies real concrete peak", passed, failed, report
    If capacity.PeakLambda > 0# Then
        TestExpectNear capacity.PeakLambda, AREA * independentStrength / Abs(force), 0.02, _
            "comparison with uniform stress A*R", passed, failed, report
        TestRecord capacity.Status = "FAIL" And capacity.LambdaCapacity = capacity.PeakLambda, "real peak before design load fails", passed, failed, report
        TestExpectNear capacity.NUltimate, force * capacity.PeakLambda, 0.000001, "ultimate axial force", passed, failed, report
        TestExpectNear capacity.TUltimate, TORQUE * capacity.PeakLambda, 0.000001, "ultimate torque", passed, failed, report
        Dim peakState As CSectionStateResult: Set peakState = capacity.StateResult
        If peakState Is Nothing Then
            TestRecord False, "actual state at concrete peak is missing " & capacity.ResultMeta.ResultComment, passed, failed, report
        Else
            TestRecord peakState.StateType = sstCapacityState And peakState.Converged And peakState.WithinPhysicalRange, _
                "concrete peak retains the named equilibrium state", passed, failed, report
            TestExpectNear peakState.Nint, capacity.NUltimate, 0.05, "actual integrated N at the concrete peak", passed, failed, report
            TestRecord peakState.Twist <> 0# And Not (peakState.CoupledFields Is Nothing) And Not (peakState.CoupledMaterialCurves Is Nothing), _
                "concrete peak retains twist, element fields and material curves", passed, failed, report
        End If
    End If
    TestRunConcretePeak = report & "TOTAL_COUPLED_CONCRETE_" & mode & "_PEAK: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    TestRunConcretePeak = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, как выбранная по физическим деформациям мера продолжает
' реальную бетонную ветвь после трудного силового участка.
Public Function RunTestCoupledConcreteControlledPath() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, solver As New CCoupledNewtonSolver
    TestPrepareCoupledConcreteSection operators, concrete, solver
    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    path.Initialize 0#, 28000#, 0#, 510000#, 0#, 180000#, 0#, 1500000#, 0#, 17000#, 0#, -7100#
    branch.Initialize solver, path, 4, 60, 30, 0.0000005, 0.0000005
    Dim field() As Double, previous() As Double, actual() As Double, expected() As Double
    If Not branch.Advance(0.94) Then Err.Raise 5, , "Known first equilibrium failed. " & branch.ResultMeta.DiagnosticDetails
    branch.CopyAccepted previous, actual
    If Not branch.Advance(0.97) Then
        report = report & "FORCE_CONTROL_STOP: " & branch.ResultMeta.DiagnosticDetails & vbCrLf
        If branch.AcceptedLambda <= 0.94 Or branch.ResultMeta.InternalStatus <> rsNumericalFailure Then _
            Err.Raise 5, , "No confirmed second point for the deformation-controlled probe."
    End If
    branch.CopyAccepted field, actual
    Dim projection() As Double, control As Double, increment As Double, name As String
    If Not solver.SelectDeformationControl(previous, field, projection, control, increment, name) Then _
        Err.Raise 5, , "Could not select a physical deformation coordinate."
    report = report & "CONTROL: " & name & " | EPS=" & CStr(control) & " | DELTA=" & CStr(increment) & vbCrLf
    Dim base() As Double, direction() As Double, lambda As Double, lastLambda As Double
    Dim targetControl As Double, checkControl As Double, i As Long, stepIndex As Long
    Dim peak As Double, descending As Long, acceptedSteps As Long
    path.Components base, direction
    lastLambda = branch.AcceptedLambda: peak = lastLambda
    For stepIndex = 1 To 14
        targetControl = control + increment * stepIndex * 0.4
        If Not solver.TryDeformationStep(base, direction, projection, targetControl, lastLambda, lambda, _
                0.0000005, 0.0000005, 20) Then
            report = report & "CONTROL_STOP: " & solver.ResultMeta.ResultComment & " " & solver.ResultMeta.DiagnosticDetails & vbCrLf
            Exit For
        End If
        solver.CopyTrial field, actual
        path.TargetSix lambda, expected
        checkControl = 0#
        For i = 1 To UBound(field): checkControl = checkControl + field(i) * projection(i): Next i
        TestExpectNear checkControl, targetControl, 0.0000003, "controlled strain at step " & CStr(stepIndex), passed, failed, report
        For i = 1 To 6
            TestExpectNear actual(i), expected(i), 0.00005, "controlled section resultant " & CStr(i), passed, failed, report
        Next i
        If lambda <= lastLambda - 0.0001 Then descending = descending + 1 Else descending = 0
        If lambda > peak Then peak = lambda
        lastLambda = lambda: acceptedSteps = acceptedSteps + 1
        solver.AcceptStep
        report = report & "CONTROL_POINT: coordinate=" & CStr(checkControl) & " | lambda=" & CStr(lambda) & vbCrLf
        If descending >= 2 Then Exit For
    Next stepIndex
    TestRecord acceptedSteps > 0, "physical control advanced the nonlinear section", passed, failed, report
    report = report & "CONTROL_MAX_LAMBDA: " & CStr(peak) & " | DESCENDING: " & CStr(descending) & vbCrLf
    RunTestCoupledConcreteControlledPath = report & "TOTAL_COUPLED_CONCRETE_CONTROLLED_PATH: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConcreteControlledPath = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит одинаковую геометрию, материалы и решатель для двух уровней нагрузки.
Private Sub TestPrepareCoupledConcreteSection(ByVal operators As CCoupledElasticState, _
        ByVal concrete As CCoupledConcreteResponse, ByVal solver As CCoupledNewtonSolver)
    Dim section As CSectionModel: Set section = TestHardShearSection()
    Dim layout As New CTransverseRebarLayout
    layout.AddContourGroup section, "TestHoop", "", 8#, 100#, 10#, "A240"
    operators.PrepareModel section, 30000#, 0.2, 1, 1, 5, layout, 200000#
    concrete.Initialize operators.PointCount, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    Dim longitudinal As New CCoupledSteelModel, hoop As New CCoupledSteelModel, models As Object
    longitudinal.Initialize 200000#, 400#, 350#, 1500#, 1000#, 0.025, 0.018
    hoop.Initialize 200000#, 240#, 240#, 2000#, 0#, 0.02, 0.02
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "A400", longitudinal: models.Add "A240", hoop
    Dim material As New CCoupledReinforcedResponse
    material.Initialize operators, concrete, models
    solver.Initialize operators, material, 30000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Содержит одно тяжёлое сочетание в Н и Н·мм для обоих способов запуска.
Private Function TestHardShearLoads() As Variant
    TestHardShearLoads = Array(28000#, 510000#, 180000#, 1500000#, 17000#, -7100#)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит общую геометрию тяжёлого сдвига с двумя продольными стержнями.
Private Function TestHardShearSection() As CSectionModel
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, -30#, 12#, GEOM_PI * 36#, "A400"
    section.AddRebarElement 0#, 30#, 12#, GEOM_PI * 36#, "A400"
    Set TestHardShearSection = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает штатный пакет с пространственным и плоским сочетаниями.
' Сверяет результирующие усилия, изоляцию истории и статус профиля с пределом.
Public Function RunTestCoupledBatchRoute() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, -30#, 12#, GEOM_PI * 36#, "A400"
    section.AddRebarElement 0#, 30#, 12#, GEOM_PI * 36#, "A400"
    Dim layout As New CTransverseRebarLayout
    layout.AddContourGroup section, "TestHoop", "", 8#, 100#, 10#, "A400"

    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    concrete.Initialize 30#, 3#, 30#, 3#, 30000#, 30000#, rbMc2:=15#
    steel.Initialize 350#, 400#, 350#, 400#, 200000#, 200000#
    Dim provider As New CMaterialModelProvider
    provider.TestInitializeFromParameters concrete, steel
    Dim analysis As New CCoupledSectionAnalysis
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30

    Dim profiles As New CCalculationProfileCatalog, spec As New CMaterialModelSpec
    profiles.LoadFromWorkbook ThisWorkbook, False
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    profiles.ProfileById("PR4").SetStrengthSpec spec
    profiles.ProfileById("PR1").SetStrengthSpec spec
    Dim batch As New CBatchSectionCalculator
    batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles
    batch.ConfigureCoupledAnalysis analysis
    batch.AddCombination "Coupled1", 13160#, 239700#, 84600#, "PR4", "", torque:=705000#, qx:=7990#, qy:=-3337#
    batch.AddCombination "Plane", -10000#, 0#, 0#, "PR4", ""
    batch.AddCombination "Coupled2", 13160#, 239700#, 84600#, "PR1", "", torque:=705000#, qx:=7990#, qy:=-3337#
    batch.ApplyLoadReference 11#, -7#, 11#, -7#
    batch.Execute

    Dim first As CCombinationResult, second As CCombinationResult, plane As CCombinationResult
    Set first = batch.ResultAt(1): Set plane = batch.ResultAt(2): Set second = batch.ResultAt(3)
    TestRecord first.DirectStateMeta.InternalStatus = rsSuccess, "target equilibrium " & first.OverallMeta.ResultComment, passed, failed, report
    TestRecord plane.DirectStateMeta.InternalStatus = rsSuccess, "planar route remains available " & plane.OverallMeta.ResultComment, passed, failed, report
    TestRecord second.DirectStateMeta.InternalStatus = rsSuccess, "second independent target " & second.OverallMeta.ResultComment, passed, failed, report
    TestRecord second.CapacityMeta.InternalStatus = rsNumericalFailure And second.StrengthResult.Capacity.ReserveFactor = 0# And _
        (second.StrengthResult.Capacity.HasConfirmedLowerBound Or second.StrengthResult.Capacity.HasUltimatePoint), _
        "incomplete capacity retains the verified part of the path without reporting a full reserve", passed, failed, report
    TestRecord second.StabilityMeta.InternalStatus = rsNotApplicable And InStr(second.StabilityMeta.ResultComment, "Mx") > 0, _
        "tension preserves the standard stability status and scope", passed, failed, report
    TestRecord second.OverallMeta.InternalStatus = rsNumericalFailure, "incomplete profile does not report overall success", passed, failed, report

    Dim actual() As Double, repeated() As Double, target() As Double, i As Long
    Dim hasFirst As Boolean, hasSecond As Boolean
    hasFirst = first.CopyCoupledResultants(actual): hasSecond = second.CopyCoupledResultants(repeated)
    TestRecord hasFirst And hasSecond, "both accepted vectors are available", passed, failed, report
    If hasFirst And hasSecond Then
        Dim loads As New CSectionLoadState, x As Double, y As Double, area As Double
        section.ConcreteMeshCentroid x, y, area
        loads.Initialize 13160#, 239700#, 84600#, 11#, -7#
        loads.SetSpatialComponents 705000#, 7990#, -3337#
        loads.LoadAboutPoint x, y, target
        For i = 1 To 6
            TestExpectNear actual(i), target(i), 0.00002, "batch equilibrium component " & CStr(i), passed, failed, report
            TestRecord repeated(i) = actual(i), "independent material history component " & CStr(i), passed, failed, report
        Next i
    End If
    TestRecord Not plane.CopyCoupledResultants(actual), "planar result has no coupled vector", passed, failed, report
    RunTestCoupledBatchRoute = report & "TOTAL_COUPLED_BATCH_ROUTE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledBatchRoute = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет чистые и парные нагрузки с балансом всех усилий. Повторяет пакет
' на том же расчётчике в обратном порядке и под другим профилем прочности.
Public Function RunTestCoupledLoadMatrix() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel, hoop As CGeometryPath, layout As New CTransverseRebarLayout
    Set section = TestReinforcedAsymmetricSection(hoop)
    layout.AddGroup hoop, 8#, 100#, "A400"

    Dim provider As CMaterialModelProvider: Set provider = TestMatrixMaterials()
    Dim analysis As New CCoupledSectionAnalysis, profiles As New CCalculationProfileCatalog, spec As New CMaterialModelSpec
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30
    profiles.LoadFromWorkbook ThisWorkbook, False
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    Dim id As Variant, profile As CCalculationProfile
    For Each id In Array("PR4", "PR1")
        Set profile = profiles.ProfileById(CStr(id))
        profile.Initialize CStr(id), "Контроль матрицы нагрузок", "", True, False, False, False
        profile.SetStrengthSpec spec
    Next id
    Dim cases As Variant, components As Variant, amplitudes As Variant
    components = Array("N", "Mx", "My", "T", "Qx", "Qy")
    amplitudes = Array(28000#, 1300000#, -1700000#, 1800000#, 12500#, -10500#)
    cases = Array(Array("N", 0), Array("Mx", 1), Array("My", 2), Array("T", 3), Array("Qx", 4), Array("Qy", 5), _
        Array("Mx+T", 1, 3), Array("My+T", 2, 3), Array("Qx+T", 4, 3), Array("Qy+T", 5, 3), _
        Array("N+Qx", 0, 4), Array("N+Qy", 0, 5), Array("Mx+Qx", 1, 4), Array("Mx+Qy", 1, 5), _
        Array("My+Qx", 2, 4), Array("My+Qy", 2, 5), Array("AllSix", 0, 1, 2, 3, 4, 5))
    Dim batch As New CBatchSectionCalculator, cx As Double, cy As Double, area As Double
    Dim settings As New CSystemSettingsReader, units As New CUnitSystem
    batch.Initialize section, provider
    TestLoadMatrixSettings settings
    units.LoadFromSettings settings
    batch.ApplySettings settings, units
    Set batch.ProfileCatalog = profiles
    batch.ConfigureCoupledAnalysis analysis
    section.ConcreteMeshCentroid cx, cy, area
    Dim first() As CSectionStateResult: ReDim first(0 To UBound(cases))
    Dim expected() As Double, actual() As Double, result As CCombinationResult, state As CSectionStateResult
    Dim pass As Long, i As Long, row As Long, j As Long, caseData As Variant, spatial As Boolean
    Dim suffix As String, profileID As String
    For pass = 1 To 2
        batch.ClearCombinations
        If pass = 1 Then profileID = "PR4" Else profileID = "PR1"
        For row = 0 To UBound(cases)
            i = row
            If pass = 2 Then i = UBound(cases) - row
            caseData = cases(i)
            TestMatrixVector caseData, amplitudes, expected
            batch.AddCombination CStr(caseData(0)), expected(0), expected(1), expected(2), profileID, "", _
                torque:=expected(3), qx:=expected(4), qy:=expected(5)
        Next row
        batch.ApplyLoadReference cx, cy, cx, cy
        batch.Execute
        For row = 0 To UBound(cases)
            i = row
            If pass = 2 Then i = UBound(cases) - row
            caseData = cases(i): suffix = " | " & CStr(caseData(0)) & " | " & profileID
            TestMatrixVector caseData, amplitudes, expected
            spatial = (expected(3) <> 0# Or expected(4) <> 0# Or expected(5) <> 0#)
            Set result = batch.ResultAt(row + 1)
            Set state = result.StateRepository.FindState(sstStrengthState)
            TestRecord result.DirectStateMeta.InternalStatus = rsSuccess, "converged" & suffix & " | " & _
                result.DirectStateMeta.ResultComment, passed, failed, report
            TestRecord batch.CombinationID(row + 1) = CStr(caseData(0)) And batch.HasSpatialLoads(row + 1) = spatial, _
                "case identity and route" & suffix, passed, failed, report
            If state Is Nothing Then
                TestRecord False, "missing state" & suffix, passed, failed, report
            ElseIf state.Converged Then
                TestMatrixReadVector result, state, spatial, cx, cy, actual, suffix, passed, failed, report
                For j = 0 To 5
                    TestRecord Abs(actual(j) - expected(j)) <= 0.00005 * Abs(amplitudes(j)), _
                        "equilibrium " & CStr(components(j)) & suffix & " = " & CStr(actual(j)), passed, failed, report
                Next j
                If pass = 1 Then
                    Set first(i) = state
                ElseIf Not first(i) Is Nothing Then
                    TestRecord TestMatrixStatesMatch(first(i), state, section.ConcreteCount + section.RebarCount), _
                        "reversed batch gives identical material fields and state" & suffix, passed, failed, report
                End If
            End If
        Next row
    Next pass
    TestMatrixInteractions first, cases, section.ConcreteCount, passed, failed, report
    TestPerforatedHoopLimit section, hoop, provider, spec, cx, cy, passed, failed, report
    RunTestCoupledLoadMatrix = report & "TOTAL_COUPLED_LOAD_MATRIX: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledLoadMatrix = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет нелинейную связь в точках интегрирования на двух нагрузках,
' при которых бетон выходит из упругой области равновесного решения.
Public Function RunTestCoupledIntegratedInteraction() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel, hoop As CGeometryPath, layout As New CTransverseRebarLayout
    Set section = TestReinforcedAsymmetricSection(hoop)
    layout.AddGroup hoop, 8#, 100#, "A400"
    Dim provider As CMaterialModelProvider: Set provider = TestMatrixMaterials()
    Dim analysis As New CCoupledSectionAnalysis, spec As New CMaterialModelSpec
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"

    Dim cx As Double, cy As Double, area As Double
    section.ConcreteMeshCentroid cx, cy, area
    TestIntegratedShearPair analysis, spec, cx, cy, "Mx+T", 1300000#, 0#, 1.35, passed, failed, report
    TestIntegratedShearPair analysis, spec, cx, cy, "Qx+T", 0#, 12500#, 1.25, passed, failed, report
    RunTestCoupledIntegratedInteraction = report & "TOTAL_COUPLED_INTEGRATED_INTERACTION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledIntegratedInteraction = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет касательные напряжения бетона в одних и тех же точках.
' Чистые части проходят тот же связанный решатель и начинают с нулевой истории.
Private Sub TestIntegratedShearPair(ByVal analysis As CCoupledSectionAnalysis, ByVal spec As CMaterialModelSpec, _
        ByVal cx As Double, ByVal cy As Double, ByVal name As String, ByVal mx As Double, ByVal qx As Double, _
        ByVal factor As Double, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const TORQUE As Double = 1800000# ' Н·мм до умножения нагрузки на factor.
    Const MIN_INTERACTION As Double = 0.005 ' МПа: разрешимое изменение касательного напряжения.
    Dim combined() As Double, first() As Double, torsion() As Double, activeCombined As Long, activePart As Long
    TestIntegratedShearCase analysis, spec, cx, cy, name & " | пара", mx, TORQUE, qx, factor, _
        combined, activeCombined, passed, failed, report
    TestIntegratedShearCase analysis, spec, cx, cy, name & " | первая часть", mx, 0#, qx, factor, _
        first, activePart, passed, failed, report
    TestIntegratedShearCase analysis, spec, cx, cy, name & " | T", 0#, TORQUE, 0#, factor, _
        torsion, activePart, passed, failed, report

    TestRecord activeCombined > 0, "пластическая энергия бетона | " & name & " | точек=" & activeCombined, passed, failed, report
    Dim point As Long, component As Long, difference As Double
    For point = 1 To UBound(combined, 2)
        For component = 1 To 2
            difference = GeomMax(difference, Abs(combined(component, point) - first(component, point) - torsion(component, point)))
        Next component
    Next point
    TestRecord difference > MIN_INTERACTION, "нелинейное изменение касательных напряжений, МПа | " & name & " | " & _
        CStr(difference), passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет шесть независимо просуммированных усилий в точке задания нагрузки.
' Абсолютные допуски контролируют также компоненты с нулевой нагрузкой.
Private Sub TestIntegratedShearCase(ByVal analysis As CCoupledSectionAnalysis, ByVal spec As CMaterialModelSpec, _
        ByVal cx As Double, ByVal cy As Double, ByVal name As String, ByVal mx As Double, ByVal torque As Double, _
        ByVal qx As Double, ByVal factor As Double, ByRef shear() As Double, ByRef activePoints As Long, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim loads As New CSectionLoadState, resultants() As Double
    loads.Initialize 0#, mx, 0#, cx, cy
    loads.SetSpatialComponents torque, qx, 0#
    analysis.TestIntegratedState loads, spec, factor, resultants, shear, activePoints
    Dim components As Variant, expected As Variant, tolerance As Variant, i As Long
    components = Array("N", "Mx", "My", "T", "Qx", "Qy")
    expected = Array(0#, factor * mx, 0#, factor * torque, factor * qx, 0#)
    tolerance = Array(0.1, 15#, 15#, 15#, 0.1, 0.1) ' Н и Н·мм, порядок совпадает с вектором усилий.
    For i = LBound(components) To UBound(components)
        TestRecord Abs(resultants(i + 1) - CDbl(expected(i))) <= CDbl(tolerance(i)), _
            "равновесие " & CStr(components(i)) & " | " & name & " | " & CStr(resultants(i + 1)), passed, failed, report
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит три продольных стержня и хомут в стенке сечения. Все части
' преобразуются вместе, включая замкнутый путь поперечной арматуры.
Private Function TestReinforcedAsymmetricSection(ByRef hoop As CGeometryPath, Optional ByVal shiftX As Double = 0#, _
        Optional ByVal shiftY As Double = 0#, Optional ByVal rotation As Double = 0#, _
        Optional ByVal reflectX As Boolean = False) As CSectionModel
    Dim region As CConcreteRegion, section As CSectionModel, query As New CSectionGeometryQuery
    Set section = TestAsymmetricCoupledSection(region, shiftX, shiftY, rotation, reflectX)
    Dim bar As Variant, x As Double, y As Double, diameter As Double
    For Each bar In Array(Array(-76#, -54#, 12#), Array(68#, -47#, 12#), Array(-55#, 65#, 10#))
        x = CDbl(bar(0)): y = CDbl(bar(1)): diameter = CDbl(bar(2))
        query.TransformCoordinates x, y, shiftX, shiftY, rotation, reflectX
        section.AddRebarElement x, y, diameter, GEOM_PI * diameter ^ 2 / 4#, "A400"
    Next bar
    Dim path As New CGeometryPath
    path.Initialize "MatrixHoop"
    path.AddLine "Bottom", -85#, -65#, 20#, -65#
    path.AddLine "Right", 20#, -65#, 20#, 12#
    path.AddLine "Top", 20#, 12#, -85#, 12#
    path.AddLine "Left", -85#, 12#, -85#, -65#
    path.Seal True
    Set hoop = path.Transformed(shiftX, shiftY, rotation, reflectX)
    Set TestReinforcedAsymmetricSection = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Создаёт единый контрольный набор параметров бетона и стали.
Private Function TestMatrixMaterials() As CMaterialModelProvider
    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters, provider As New CMaterialModelProvider
    concrete.Initialize 30#, 3#, 30#, 3#, 30000#, 30000#, rbMc2:=15#
    steel.Initialize 350#, 400#, 350#, 400#, 200000#, 200000#
    provider.TestInitializeFromParameters concrete, steel
    Set TestMatrixMaterials = provider
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет шесть компонентов явных плоских путей с ненулевыми T, Qx и Qy.
' В том числе проверяет точный ноль направления при нулевом выбранном усилии.
Public Function RunTestCoupledCenteredLoadPaths() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim layout As New CTransverseRebarLayout, provider As CMaterialModelProvider
    layout.DeclareNone
    Set provider = TestMatrixMaterials()
    Dim origins As Variant, cases As Variant
    origins = Array(Array(250#, 125#), Array(-1000000#, 1000000#))
    cases = Array(Array("LambdaMx", False, True, False, 0#, 0#), _
        Array("LambdaMy", False, False, True, 620000#, 0#), _
        Array("LambdaMxy", False, True, True, 0#, 0#), _
        Array("LambdaMx", False, True, False, 620000#, -420000#), _
        Array("LambdaMy", False, False, True, 620000#, -420000#), _
        Array("LambdaMxy", False, True, True, 620000#, -420000#), _
        Array("LambdaN", True, False, False, 620000#, -420000#), _
        Array("LambdaNMxy", True, True, True, 620000#, -420000#))
    Dim origin As Variant, caseData As Variant, cx As Double, cy As Double, area As Double
    Dim section As CSectionModel, analysis As CCoupledSectionAnalysis
    For Each origin In origins
        Set section = New CSectionModel
        section.AddConcreteElement CDbl(origin(0)), CDbl(origin(1)), 10000#
        section.ConcreteMeshCentroid cx, cy, area
        Set analysis = New CCoupledSectionAnalysis
        analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30
        For Each caseData In cases
            TestCenteredLoadPathComponents analysis, caseData, cx, cy, passed, failed, report
        Next caseData
    Next origin
    TestCenteredLoadPathCapacity provider, passed, failed, report
    RunTestCoupledCenteredLoadPaths = report & "TOTAL_COUPLED_CENTERED_LOAD_PATHS: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledCenteredLoadPaths = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо масштабирует заданные пользователем усилия и переносит винт
' по разности координат. Фиксированные компоненты имеют точный ноль в Base.
Private Sub TestCenteredLoadPathComponents(ByVal analysis As CCoupledSectionAnalysis, ByVal caseData As Variant, _
        ByVal cx As Double, ByVal cy As Double, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const AXIAL As Double = 12345.6, TORQUE As Double = -80000#, SHEAR_X As Double = -110#, SHEAR_Y As Double = 270#
    Dim loads As New CSectionLoadState, path As CLoadPathVector
    Dim pointX As Double, pointY As Double, dx As Double, dy As Double
    pointX = cx - 7.4: pointY = cy + 17.3
    dx = pointX - cx: dy = pointY - cy
    loads.Initialize AXIAL, CDbl(caseData(4)), CDbl(caseData(5)), pointX, pointY
    loads.SetSpatialComponents TORQUE, SHEAR_X, SHEAR_Y
    Set path = analysis.TestCenteredLoadPath(loads, CStr(caseData(0)))
    Dim offset() As Double, direction() As Double, actual() As Double
    path.Components offset, direction
    Dim expected(1 To 6) As Double, factor As Variant, index As Long, label As String
    label = CStr(caseData(0)) & " | центр Y=" & CStr(cy) & " | Mx=" & CStr(caseData(4)) & " My=" & CStr(caseData(5))
    If Not CBool(caseData(1)) Then
        If Not CBool(caseData(2)) Or CDbl(caseData(4)) = 0# Then _
            TestRecord direction(2) = 0#, label & " | точно фиксированный Mx", passed, failed, report
        If Not CBool(caseData(3)) Or CDbl(caseData(5)) = 0# Then _
            TestRecord direction(3) = 0#, label & " | точно фиксированный My", passed, failed, report
    End If
    TestRecord direction(4) = 0# And direction(5) = 0# And direction(6) = 0#, _
        label & " | точно фиксированные T, Qx, Qy", passed, failed, report
    For Each factor In Array(0#, 0.375, 1#)
        expected(1) = AXIAL * CDbl(IIf(caseData(1), factor, 1#))
        expected(2) = CDbl(caseData(4)) * CDbl(IIf(caseData(2), factor, 1#)) + expected(1) * dy
        expected(3) = CDbl(caseData(5)) * CDbl(IIf(caseData(3), factor, 1#)) + expected(1) * dx
        expected(4) = TORQUE + dx * SHEAR_Y - dy * SHEAR_X
        expected(5) = SHEAR_X: expected(6) = SHEAR_Y
        path.TargetSix CDbl(factor), actual
        For index = 1 To 6
            TestRecord Abs(actual(index) - expected(index)) <= 0.000000001, label & " | lambda=" & CStr(factor) & _
                " | компонента " & CStr(index) & " | " & CStr(actual(index)) & " / " & CStr(expected(index)), passed, failed, report
        Next index
    Next factor
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подтверждает статус неприменимого поиска после решения постоянной части
' нагрузки для LambdaMx при нулевом заданном Mx и ненулевых N, T, Qx, Qy.
Private Sub TestCenteredLoadPathCapacity(ByVal provider As CMaterialModelProvider, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim section As CSectionModel, hoop As CGeometryPath, layout As New CTransverseRebarLayout
    Set section = TestReinforcedAsymmetricSection(hoop, -1000000#, 1000000#)
    layout.AddGroup hoop, 8#, 100#, "A400"
    Dim analysis As New CCoupledSectionAnalysis, spec As New CMaterialModelSpec
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    Dim cx As Double, cy As Double, area As Double, loads As New CSectionLoadState
    section.ConcreteMeshCentroid cx, cy, area
    loads.Initialize 12345.6, 0#, -42000#, cx - 7.4, cy + 17.3
    loads.SetSpatialComponents -80000#, -110#, 270#
    Dim capacity As CCapacityResult
    Set capacity = analysis.SolveCapacity(loads, "LambdaMx", spec, 1#, 2#, 0.0001, 30)
    TestRecord capacity.ResultMeta.InternalStatus = rsNotApplicable And capacity.ResultMeta.ResultCode = rcNotApplicable, _
        "Нулевой LambdaMx: статус поиска после решения постоянной нагрузки | " & capacity.ResultMeta.ResultComment, passed, failed, report
    TestRecord capacity.Status = "N/A" And InStr(1, capacity.ResultMeta.ResultComment, "ненулевой переменной частью", vbTextCompare) > 0, _
        "Нулевой LambdaMx: внешний статус и причина неприменимости", passed, failed, report
    TestRecord Not capacity.HasUltimatePoint And Not capacity.HasConfirmedLowerBound And capacity.ReserveFactor = 0#, _
        "Нулевой LambdaMx: предельная точка и запас отсутствуют", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет перенос, поворот на 90°, отражение и смену точки приведения
' полного вектора. Ортогональные преобразования сохраняют шаг сетки.
Public Function RunTestCoupledFrameInvariance() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim provider As CMaterialModelProvider: Set provider = TestMatrixMaterials()
    Dim spec As New CMaterialModelSpec: spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    Dim frames As Variant
    frames = Array(Array("Reference", 1#, 0#, 0#, 1#, 0#, 0#, 0#, False), _
        Array("Translated", 1#, 0#, 0#, 1#, 93000#, -71000#, 0#, False), _
        Array("QuarterTurn", 0#, -1#, 1#, 0#, 0#, 0#, GEOM_PI / 2#, False), _
        Array("MirrorX", 1#, 0#, 0#, -1#, 0#, 0#, 0#, True), _
        Array("NewLoadPoint", 1#, 0#, 0#, 1#, 0#, 0#, 0#, False))
    Dim components As Variant: components = Array("N", "Mx", "My", "T", "Qx", "Qy")
    Dim referenceSection As CSectionModel, referenceState As CSectionStateResult, refX As Double, refY As Double
    Dim referenceLambda As Double
    Dim frame As Variant
    For Each frame In frames
        Dim tag As String: tag = CStr(frame(0))
        Dim section As CSectionModel, hoop As CGeometryPath, layout As CTransverseRebarLayout
        Set section = TestReinforcedAsymmetricSection(hoop, CDbl(frame(5)), CDbl(frame(6)), CDbl(frame(7)), CBool(frame(8)))
        Set layout = New CTransverseRebarLayout: layout.AddGroup hoop, 8#, 100#, "A400"
        Dim cx As Double, cy As Double, area As Double
        section.ConcreteMeshCentroid cx, cy, area
        Dim loads As CSectionLoadState, analysis As CCoupledSectionAnalysis, expected As Variant, atPoint As Variant
        Set loads = TestFrameLoads(Array(28000#, 1300000#, -1700000#, 1800000#, 12500#, -10500#), frame, cx, cy, expected, atPoint)
        Set analysis = New CCoupledSectionAnalysis
        analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30
        Dim meta As CResultMeta, state As CSectionStateResult, field() As Double, resultants() As Double
        Set meta = analysis.SolveTarget(loads, "Auto", spec, field, resultants, True, state)
        TestRecord meta.InternalStatus = rsSuccess, tag & " converged: " & meta.ResultComment, passed, failed, report
        If Not state Is Nothing Then
            Dim i As Long
            For i = 0 To 5
                TestExpectNear resultants(i + 1), CDbl(expected(i)), 0.000002, _
                    tag & " equilibrium " & CStr(components(i)), passed, failed, report
            Next i
            If tag = "Reference" Then
                Set referenceSection = section: Set referenceState = state
                refX = cx: refY = cy
            ElseIf Not referenceState Is Nothing Then
                TestExpectNear state.Epsilon0 + state.KappaX * cy + state.KappaY * cx, _
                    referenceState.Epsilon0 + referenceState.KappaX * refY + referenceState.KappaY * refX, _
                    0.00001, tag & " strain at the concrete centroid", passed, failed, report
                TestExpectNear state.KappaX, frame(3) * referenceState.KappaY + frame(4) * referenceState.KappaX, _
                    0.00001, tag & " rotated kappaX", passed, failed, report
                TestExpectNear state.KappaY, frame(1) * referenceState.KappaY + frame(2) * referenceState.KappaX, _
                    0.00001, tag & " rotated kappaY", passed, failed, report
                TestExpectNear state.Twist, (frame(1) * frame(4) - frame(2) * frame(3)) * referenceState.Twist, _
                    0.00001, tag & " oriented twist", passed, failed, report
                TestFrameFields referenceSection, section, referenceState, state, frame, passed, failed, report
            End If
        End If
        TestFrameCapacity section, hoop, provider, spec, frame, cx, cy, referenceLambda, passed, failed, report
    Next frame
    RunTestCoupledFrameInvariance = report & "TOTAL_COUPLED_FRAME: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledFrameInvariance = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Независимо преобразует пары (My,Mx) и (Qx,Qy). T умножается на
' определитель ортогонального преобразования, а N сохраняет знак.
Private Function TestFrameLoads(ByVal baseline As Variant, ByVal frame As Variant, ByVal cx As Double, ByVal cy As Double, _
        ByRef aboutCentroid As Variant, ByRef aboutApplication As Variant) As CSectionLoadState
    Dim force As Double, mx As Double, my As Double, torque As Double, qx As Double, qy As Double
    Dim dx As Double, dy As Double, leverX As Double, leverY As Double
    force = baseline(0)
    my = frame(1) * baseline(2) + frame(2) * baseline(1): mx = frame(3) * baseline(2) + frame(4) * baseline(1)
    qx = frame(1) * baseline(4) + frame(2) * baseline(5): qy = frame(3) * baseline(4) + frame(4) * baseline(5)
    torque = (frame(1) * frame(4) - frame(2) * frame(3)) * baseline(3)
    dx = frame(1) * 9# + frame(2) * -7#: dy = frame(3) * 9# + frame(4) * -7#
    If CStr(frame(0)) = "NewLoadPoint" Then
        leverX = -13#: leverY = 8# ' Компенсируем перенос точки, сохраняя тот же физический силовой винт.
        mx = mx - force * leverY: my = my - force * leverX
        torque = torque - (leverX * qy - leverY * qx)
    End If
    dx = dx + leverX: dy = dy + leverY
    aboutApplication = Array(force, mx, my, torque, qx, qy)
    aboutCentroid = Array(force, mx + force * dy, my + force * dx, torque + dx * qy - dy * qx, qx, qy)
    Dim loads As New CSectionLoadState
    loads.Initialize force, mx, my, cx + dx, cy + dy
    loads.SetSpatialComponents torque, qx, qy
    Set TestFrameLoads = loads
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет тот же определяющий механизм и запас, а моменты предельной
' точки сверяет в пользовательской точке приложения каждого варианта.
Private Sub TestFrameCapacity(ByVal section As CSectionModel, ByVal hoop As CGeometryPath, ByVal provider As CMaterialModelProvider, _
        ByVal spec As CMaterialModelSpec, ByVal frame As Variant, ByVal cx As Double, ByVal cy As Double, ByRef referenceLambda As Double, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim aboutCentroid As Variant, atPoint As Variant, loads As CSectionLoadState, capacity As CCapacityResult
    Set loads = TestFrameLoads(Array(-100000#, 120000#, -90000#, 200000#, 240#, -120#), frame, cx, cy, aboutCentroid, atPoint)
    Set capacity = TestWeakHoopCapacity(section, hoop, provider, spec, loads, 0.000012)
    Dim tag As String: tag = CStr(frame(0))
    TestRecord capacity.HasUltimatePoint And capacity.IsTransverseLimit And capacity.Status = "FAIL", _
        tag & " hoop governs the physical limit: " & capacity.ResultMeta.ResultComment, passed, failed, report
    If Not capacity.HasUltimatePoint Then Exit Sub
    Dim lambda As Double: lambda = capacity.LambdaCapacity
    If tag = "Reference" Then
        referenceLambda = lambda
    Else
        TestRecord Abs(lambda - referenceLambda) <= 0.0001, tag & " same hoop reserve " & lambda, passed, failed, report
    End If
    Dim actual As Variant, i As Long
    actual = Array(capacity.NUltimate, capacity.MxUltimate, capacity.MyUltimate, capacity.TUltimate, capacity.QxUltimate, capacity.QyUltimate)
    For i = 0 To 5
        TestExpectNear CDbl(actual(i)), lambda * CDbl(atPoint(i)), 0.000002, _
            tag & " ultimate component " & CStr(i + 1), passed, failed, report
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет семь полей в соответствующих центрах. Сумма σzz·εzz и
' τ·γ по бетонным элементам проверяет инвариантность измеренной работы.
Private Sub TestFrameFields(ByVal referenceSection As CSectionModel, ByVal section As CSectionModel, _
        ByVal referenceState As CSectionStateResult, ByVal state As CSectionStateResult, ByVal frame As Variant, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim tag As String: tag = CStr(frame(0))
    Dim reference As CCoupledFieldSnapshot, moved As CCoupledFieldSnapshot
    Set reference = referenceState.CoupledFields: Set moved = state.CoupledFields
    Dim count As Long: count = referenceSection.ConcreteCount
    If count <> section.ConcreteCount Or reference Is Nothing Or moved Is Nothing Then
        TestRecord False, tag & " matching concrete mesh and saved fields", passed, failed, report
        Exit Sub
    End If
    Dim used() As Boolean: ReDim used(1 To count)
    Dim i As Long, j As Long, k As Long, x As Double, y As Double, newX As Double, newY As Double
    Dim a(0 To 6) As Double, b(0 To 6) As Double, expected As Variant
    Dim errorByField(0 To 6) As Double, baseWork As Double, movedWork As Double
    For i = 1 To count
        x = referenceSection.ConcreteX(i): y = referenceSection.ConcreteY(i)
        newX = frame(1) * x + frame(2) * y + frame(5): newY = frame(3) * x + frame(4) * y + frame(6)
        For j = 1 To count
            If Not used(j) Then
                If Abs(section.ConcreteX(j) - newX) < 0.000001 And Abs(section.ConcreteY(j) - newY) < 0.000001 Then Exit For
            End If
        Next j
        If j > count Then
            TestRecord False, tag & " missing matching concrete point " & CStr(i), passed, failed, report
            Exit Sub
        End If
        used(j) = True
        reference.Axial i, a(0), a(1): moved.Axial j, b(0), b(1)
        reference.Spatial i, a(2), a(3), a(4), a(5), a(6): moved.Spatial j, b(2), b(3), b(4), b(5), b(6)
        expected = Array(a(0), a(1), frame(1) * a(2) + frame(2) * a(3), frame(3) * a(2) + frame(4) * a(3), _
            frame(1) * a(4) + frame(2) * a(5), frame(3) * a(4) + frame(4) * a(5), a(6))
        For k = 0 To 6: errorByField(k) = GeomMax(errorByField(k), Abs(b(k) - expected(k))): Next k
        baseWork = baseWork + referenceSection.ConcreteArea(i) * (a(0) * a(1) + a(2) * a(4) + a(3) * a(5))
        movedWork = movedWork + section.ConcreteArea(j) * (b(0) * b(1) + b(2) * b(4) + b(3) * b(5))
    Next i
    Dim labels As Variant, tolerance As Variant
    labels = Array("epsilonZ", "sigmaZ", "tauXZ", "tauYZ", "gammaXZ", "gammaYZ", "warping")
    tolerance = Array(0.00000001, 0.00001, 0.00001, 0.00001, 0.00000001, 0.00000001, 0.0000001)
    For k = 0 To 6
        TestRecord errorByField(k) <= tolerance(k), tag & " field " & labels(k) & " max deviation " & errorByField(k), _
            passed, failed, report
    Next k
    TestExpectNear movedWork, baseWork, 0.00001, tag & " axial and shear work density integral", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' На той же геометрии проверяет предел хомута при шести ненулевых усилиях.
' Для контрольной стали заданы малые деформации текучести и предела.
Private Sub TestPerforatedHoopLimit(ByVal section As CSectionModel, ByVal hoop As CGeometryPath, _
        ByVal provider As CMaterialModelProvider, ByVal spec As CMaterialModelSpec, ByVal cx As Double, ByVal cy As Double, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const STRAIN_LIMIT As Double = 0.000012
    Const POISSON_RATIO As Double = 0.2
    Const MODULUS As Double = 30000#
    Const FORCE As Double = -100000#
    Dim loads As New CSectionLoadState, capacity As CCapacityResult
    loads.Initialize FORCE, 120000#, -90000#, cx, cy
    loads.SetSpatialComponents 200000#, 240#, -120#
    Set capacity = TestWeakHoopCapacity(section, hoop, provider, spec, loads, STRAIN_LIMIT)
    Dim lambda As Double: lambda = capacity.LambdaCapacity
    TestRecord capacity.HasUltimatePoint And capacity.IsTransverseLimit And capacity.Status = "FAIL", _
        "perforated section reaches the hoop limit: " & capacity.ResultMeta.ResultComment, passed, failed, report
    If Not capacity.HasUltimatePoint Then Exit Sub
    Const AREA As Double = 30800# - 441# * GEOM_PI ' Точное сечение по двум прямоугольникам за вычетом круга.
    Dim estimate As Double: estimate = STRAIN_LIMIT * MODULUS * AREA / (POISSON_RATIO * Abs(FORCE))
    TestRecord Abs(lambda - estimate) < 0.2, "hoop reserve " & lambda & " near the axial Poisson estimate " & estimate, passed, failed, report
    TestExpectNear capacity.NUltimate, lambda * FORCE, 1E-7, "perforated limit N", passed, failed, report
    TestExpectNear capacity.MxUltimate, lambda * 120000#, 1E-7, "perforated limit Mx", passed, failed, report
    TestExpectNear capacity.MyUltimate, -lambda * 90000#, 1E-7, "perforated limit My", passed, failed, report
    TestExpectNear capacity.TUltimate, lambda * 200000#, 1E-7, "perforated limit T", passed, failed, report
    TestExpectNear capacity.QxUltimate, lambda * 240#, 1E-7, "perforated limit Qx", passed, failed, report
    TestExpectNear capacity.QyUltimate, -lambda * 120#, 1E-7, "perforated limit Qy", passed, failed, report
    Dim state As CSectionStateResult: Set state = capacity.StateResult
    Dim control As String, utilization As Double, stress As Double
    If Not state Is Nothing Then
        TestRecord state.Converged And state.CopyCoupledTransverseControl(control, utilization, stress) And _
            InStr(control, "MatrixHoop (TestMatrixHoopGrade)") > 0 And utilization > 0.99 And utilization <= 1#, _
            "saved perforated limit state equilibrates just below the hoop threshold", passed, failed, report
    Else
        TestRecord False, "missing perforated limit state", passed, failed, report
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ищет предел с искусственно слабым хомутом при заданных усилиях.
Private Function TestWeakHoopCapacity(ByVal section As CSectionModel, ByVal hoop As CGeometryPath, _
        ByVal provider As CMaterialModelProvider, ByVal spec As CMaterialModelSpec, ByVal loads As CSectionLoadState, _
        ByVal strainLimit As Double) As CCapacityResult
    Dim layout As New CTransverseRebarLayout: layout.AddGroup hoop, 8#, 100#, "TestMatrixHoopGrade"
    Dim steel As New CCoupledSteelModel: steel.Initialize 200000#, 2#, 2#, 500#, 500#, strainLimit, strainLimit
    Dim grades As Object, supplementary As Object
    Set grades = CreateObject("Scripting.Dictionary"): grades.Add "TestMatrixHoopGrade", steel
    Set supplementary = CreateObject("Scripting.Dictionary"): supplementary.Add "ULS(I)", grades
    Dim analysis As New CCoupledSectionAnalysis
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30, supplementary
    Set TestWeakHoopCapacity = analysis.SolveCapacity(loads, "Auto", spec, 1#, 1.5, 0.0001, 30)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Загружает настройки без инженерного обнуления моментов. Даже малая
' добавка N·e нужна для точной проверки переноса чистой продольной силы.
Private Sub TestLoadMatrixSettings(ByVal settings As CSystemSettingsReader)
    Dim cell As Object, saved As Variant
    Set cell = TestInputCell(ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "Calculation.ZeroMomentPerDepth")
    saved = cell.Formula
    On Error GoTo Failed
    cell.Value2 = 0#
    settings.LoadFromWorkbook ThisWorkbook
    cell.Formula = saved
    Exit Sub
Failed:
    Dim failedNumber As Long, failedSource As String, failedDescription As String
    failedNumber = Err.Number: failedSource = Err.Source: failedDescription = Err.Description
    cell.Formula = saved
    Err.Raise failedNumber, failedSource, failedDescription
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Заполняет независимый вектор в порядке N, Mx, My, T, Qx, Qy.
Private Sub TestMatrixVector(ByVal caseData As Variant, ByVal amplitudes As Variant, ByRef vector() As Double)
    Dim i As Long, j As Long
    ReDim vector(0 To 5)
    For i = 1 To UBound(caseData)
        j = CLng(caseData(i)): vector(j) = CDbl(amplitudes(j))
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет усилия в центре бетонной сетки. Приводит к нему плоские моменты
' и проверяет, что в плоском состоянии нет пространственных полей.
Private Sub TestMatrixReadVector(ByVal result As CCombinationResult, ByVal state As CSectionStateResult, _
        ByVal spatial As Boolean, ByVal cx As Double, ByVal cy As Double, ByRef actual() As Double, _
        ByVal suffix As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim values() As Double, stored As Boolean, fields As CCoupledFieldSnapshot, j As Long
    ReDim actual(0 To 5)
    stored = result.CopyCoupledResultants(values)
    Set fields = state.CoupledFields
    TestRecord (stored = spatial) And ((Not fields Is Nothing) = spatial), _
        "accepted field contract" & suffix, passed, failed, report
    If stored Then
        For j = 0 To 5: actual(j) = values(j + 1): Next j
    Else
        actual(0) = state.Nint: actual(1) = state.Mxint - state.Nint * cy: actual(2) = state.Myint - state.Nint * cx
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет четыре обобщённые деформации и все семь полей каждого элемента.
Private Function TestMatrixStatesMatch(ByVal first As CSectionStateResult, ByVal second As CSectionStateResult, _
        ByVal elementCount As Long) As Boolean
    If first.InternalStatus <> second.InternalStatus Or first.Epsilon0 <> second.Epsilon0 Or _
            first.KappaX <> second.KappaX Or first.KappaY <> second.KappaY Or first.Twist <> second.Twist Then Exit Function
    Dim a As CCoupledFieldSnapshot, b As CCoupledFieldSnapshot
    Set a = first.CoupledFields: Set b = second.CoupledFields
    If a Is Nothing Or b Is Nothing Then
        TestMatrixStatesMatch = (a Is Nothing) And (b Is Nothing)
        Exit Function
    End If
    Dim i As Long, ax As Double, ay As Double, az As Double, av As Double, aw As Double
    Dim bx As Double, by As Double, bz As Double, bv As Double, bw As Double
    For i = 1 To elementCount
        a.Axial i, ax, ay: b.Axial i, bx, by
        If ax <> bx Or ay <> by Then Exit Function
        a.Spatial i, ax, ay, az, av, aw: b.Spatial i, bx, by, bz, bv, bw
        If ax <> bx Or ay <> by Or az <> bz Or av <> bv Or aw <> bw Then Exit Function
    Next i
    TestMatrixStatesMatch = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает касательные напряжения пары с суммой отдельных решений.
' Чистый N или M не прикладывает сдвиг и не создаёт пространственного поля.
Private Sub TestMatrixInteractions(ByRef states() As CSectionStateResult, ByVal cases As Variant, _
        ByVal concreteCount As Long, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const MIN_INTERACTION As Double = 0.005 ' МПа: превышает допуск равновесия для этой задачи.
    Dim i As Long, k As Long, data As Variant, first As Long, second As Long
    Dim bendingTorsion As Double, shearTorsion As Double
    For i = 6 To UBound(cases) - 1
        data = cases(i): first = CLng(data(1)): second = CLng(data(2))
        If Not states(i) Is Nothing And Not states(first) Is Nothing And Not states(second) Is Nothing Then
            Dim difference As Double, qx As Double, qy As Double, ax As Double, ay As Double, bx As Double, by As Double
            difference = 0#
            For k = 1 To concreteCount
                TestMatrixShear states(i), k, qx, qy
                TestMatrixShear states(first), k, ax, ay
                TestMatrixShear states(second), k, bx, by
                difference = GeomMax(difference, GeomMax(Abs(qx - ax - bx), Abs(qy - ay - by)))
            Next k
            report = report & "INTERACTION: " & data(0) & " max shear difference, MPa = " & CStr(difference) & vbCrLf
            Select Case CStr(data(0))
                Case "Mx+T": bendingTorsion = difference
                Case "Qx+T": shearTorsion = difference
            End Select
        End If
    Next i
    TestRecord bendingTorsion > MIN_INTERACTION, "Mx+T changes the shear field compared to the independent responses", passed, failed, report
    TestRecord shearTorsion > MIN_INTERACTION, "Qx+T changes the shear field compared to the independent responses", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает касательные напряжения пространственного состояния.
Private Sub TestMatrixShear(ByVal state As CSectionStateResult, ByVal i As Long, ByRef tx As Double, ByRef ty As Double)
    Dim fields As CCoupledFieldSnapshot: Set fields = state.CoupledFields
    Dim gx As Double, gy As Double, w As Double
    tx = 0#: ty = 0#
    If Not fields Is Nothing Then fields.Spatial i, tx, ty, gx, gy, w
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет, что замкнутый хомут с заданным малым пределом деформации
' определяет FAIL и запас раньше обычной продольной арматуры.
Public Function RunTestCoupledHoopCapacity() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, -30#, 12#, GEOM_PI * 36#, "A400"
    section.AddRebarElement 0#, 30#, 12#, GEOM_PI * 36#, "A400"
    Dim layout As New CTransverseRebarLayout
    layout.AddContourGroup section, "TestWeakHoop", "", 8#, 100#, 10#, "TestHoopGrade"
    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    concrete.Initialize 30#, 3#, 30#, 3#, 30000#, 30000#, rbMc2:=15#
    steel.Initialize 350#, 400#, 350#, 400#, 200000#, 200000#
    Dim provider As New CMaterialModelProvider: provider.TestInitializeFromParameters concrete, steel
    Dim hoop As New CCoupledSteelModel, grades As Object, supplementary As Object
    hoop.Initialize 200000#, 6#, 6#, 500#, 500#, 0.00004, 0.00004
    Set grades = CreateObject("Scripting.Dictionary"): grades.Add "TestHoopGrade", hoop
    Set supplementary = CreateObject("Scripting.Dictionary"): supplementary.Add "ULS(I)", grades
    Dim analysis As New CCoupledSectionAnalysis, spec As New CMaterialModelSpec
    analysis.Initialize section, layout, provider, 0.2, 0.12, 80#, 125#, 16# / 30#, 0.002, "A400", 30, supplementary
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"

    Dim loads As New CSectionLoadState
    loads.Initialize -216000#, 90000#, -40000#, 2#, -1#
    loads.SetSpatialComponents 250000#, 120#, -80#
    Dim meta As CResultMeta, state As CSectionStateResult, field() As Double, resultants() As Double
    Set meta = analysis.SolveTarget(loads, "Auto", spec, field, resultants, True, state)
    TestRecord meta.InternalStatus = rsCheckFailed And meta.ResultCode = rcTransverseStrainExceeded And _
        InStr(meta.ResultComment, "TestWeakHoop") > 0, _
        "direct hoop failure " & meta.ResultComment, passed, failed, report
    If Not state Is Nothing Then
        Dim control As String, utilization As Double, stress As Double
        TestRecord state.CopyCoupledTransverseControl(control, utilization, stress), _
            "current hoop has a separate result", passed, failed, report
        TestRecord InStr(control, "TestWeakHoop") > 0 And InStr(control, "TestHoopGrade") > 0 And _
            utilization > 1# And stress > 6#, "governing hoop material, signed stress and strain utilization", passed, failed, report
    Else
        TestRecord False, "accepted hoop result is missing " & meta.DiagnosticDetails, passed, failed, report
    End If

    Dim capacity As CCapacityResult, lambda As Double
    Set capacity = analysis.SolveCapacity(loads, "Auto", spec, 1#, 1.5, 0.0001, 30)
    lambda = capacity.LambdaCapacity
    TestRecord capacity.HasUltimatePoint And capacity.HasSteelLimitPoint And capacity.IsTransverseLimit And _
        capacity.Status = "FAIL" And InStr(capacity.LimitState, "TestWeakHoop") > 0, _
        "hoop governs capacity " & capacity.ResultMeta.ResultComment, passed, failed, report
    TestRecord lambda > 0.35 And lambda < 0.85, _
        "reserve near independent elastic Poisson estimate of 0.60, lambda=" & CStr(lambda), passed, failed, report
    TestExpectNear capacity.NUltimate, lambda * -216000#, 1E-7, "ultimate N", passed, failed, report
    TestExpectNear capacity.MxUltimate, lambda * 90000#, 1E-7, "ultimate Mx at load point", passed, failed, report
    TestExpectNear capacity.MyUltimate, lambda * -40000#, 1E-7, "ultimate My at load point", passed, failed, report
    TestExpectNear capacity.TUltimate, lambda * 250000#, 1E-7, "ultimate T at load point", passed, failed, report
    TestExpectNear capacity.QxUltimate, lambda * 120#, 1E-7, "ultimate Qx", passed, failed, report
    TestExpectNear capacity.QyUltimate, lambda * -80#, 1E-7, "ultimate Qy", passed, failed, report
    Dim limitState As CSectionStateResult: Set limitState = capacity.StateResult
    If limitState Is Nothing Then
        TestRecord False, "actual state at the hoop limit is missing " & capacity.ResultMeta.ResultComment, passed, failed, report
    Else
        TestRecord limitState.StateType = sstCapacityState And limitState.Converged And limitState.WithinPhysicalRange, _
            "safe side of hoop limit has its own equilibrium state", passed, failed, report
        TestRecord limitState.CopyCoupledTransverseControl(control, utilization, stress) And _
            InStr(control, "TestWeakHoop") > 0 And utilization > 0.99 And utilization <= 1#, _
            "saved governing hoop is immediately below the limit", passed, failed, report
        TestExpectNear limitState.Nint, capacity.NUltimate, 0.05, "actual axial equilibrium at the hoop limit", passed, failed, report
    End If
    Dim missingState As New CResultMeta, combined As New CCombinationResult
    missingState.SetSolverFailure sfcNumericalFailure, rkDirectState, "Unresolved direct state", True
    combined.SetDirectStateMeta missingState
    combined.StoreCapacityResult capacity
    TestRecord combined.Status = "FAIL" And combined.DirectStateMeta.InternalStatus = rsNumericalFailure And _
        combined.StrengthMeta.InternalStatus = rsCheckFailed And combined.GoverningUtil > 1#, _
        "confirmed hoop failure governs even if the requested direct state is unresolved", passed, failed, report
    Dim partialCapacity As New CCapacityResult, partialVector(1 To 6) As Double
    partialCapacity.InitializeCoupledLimit 1.02, partialVector, "Auto", "Test hoop", Nothing, transverseLimit:=True
    TestRecord partialCapacity.Status = "NumFail" And partialCapacity.ReserveFactor = 0# And _
        partialCapacity.HasSteelLimitPoint And Not partialCapacity.HasUltimatePoint And _
        partialCapacity.TransverseLimitStatus = "OK" And partialCapacity.CoupledLimitLambda = 1.02, _
        "passed hoop limit is distinct from unresolved concrete reserve", passed, failed, report
    combined.Clear "Auto"
    combined.SetDirectStateMeta meta
    combined.StoreCapacityResult partialCapacity
    TestRecord combined.Status = "FAIL" And combined.CapacityMeta.InternalStatus = rsNumericalFailure, _
        "confirmed direct hoop failure governs even if overall capacity is unresolved", passed, failed, report
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) > 0 Then
        Dim catalog As New CCalculationProfileCatalog, profile As CCalculationProfile, batch As New CBatchSectionCalculator
        catalog.LoadFromWorkbook ThisWorkbook, False
        Set profile = catalog.ProfileById("PR4")
        profile.Initialize "PR4", "Test hoop capacity", "", True, True, False, False
        profile.SetStrengthSpec spec
        batch.Initialize section, provider
        Dim settings As New CSystemSettingsReader, execution As New CExecutionReport
        TestInputCell(ThisWorkbook.Names.Item("rngSystemSettings").RefersToRange, "General.ExecutionReportEnabled").Value2 = "Yes"
        settings.LoadFromWorkbook ThisWorkbook: execution.Initialize ThisWorkbook, settings
        Set batch.ExecutionReport = execution
        Set batch.ProfileCatalog = catalog
        batch.ConfigureCoupledAnalysis analysis
        batch.AddCombination "Hoop", -216000#, 90000#, -40000#, "PR4", "", torque:=250000#, qx:=120#, qy:=-80#
        Dim eventScale As Double: eventScale = lambda / 1.2 ' Та же пропорциональная траектория, с арматурным событием за расчётной точкой.
        batch.AddCombination "HoopEvent", -216000# * eventScale, 90000# * eventScale, -40000# * eventScale, "PR4", "", _
            torque:=250000# * eventScale, qx:=120# * eventScale, qy:=-80# * eventScale
        batch.ApplyLoadReference 2#, -1#, 2#, -1#
        batch.Execute
        TestUnresolvedHoopState batch.ResultAt(2), eventScale, passed, failed, report
        execution.Save
        Dim fso As Object, stream As Object, reportText As String, heading As Long, ending As Long, detail As String
        Set fso = CreateObject("Scripting.FileSystemObject")
        Set stream = fso.OpenTextFile(execution.FilePath, 1, False, -1)
        reportText = stream.ReadAll: stream.Close
        Dim transverseHeading As String: transverseHeading = "Поперечная арматура — хомуты и ветви"
        heading = InStr(reportText, vbCrLf & transverseHeading & vbCrLf)
        If heading > 0 Then ending = InStr(heading + Len(vbCrLf & transverseHeading & vbCrLf & String$(72, "-")), _
            reportText, vbCrLf & String$(72, "-") & vbCrLf)
        If ending > heading Then detail = Mid$(reportText, heading, ending - heading)
        TestRecord InStr(detail, "Статус по деформации при заданной нагрузке: FAIL") > 0 And _
            InStr(detail, "Статус предела поперечной арматуры: FAIL") > 0, "detailed hoop subsection shows both statuses", passed, failed, report
        TestRecord InStr(detail, "TestWeakHoop") > 0 And InStr(detail, "TestHoopGrade") > 0 And _
            InStr(detail, "Коэффициент запаса хомута на выбранной траектории:") > 0 And _
            InStr(detail, "Контрольная точка при предельной нагрузке:") > 0, "detail gives the hoop grade, reserve and near-limit state", passed, failed, report
        TestRecord InStr(Left$(reportText, heading), "Коэффициент запаса хомута на выбранной траектории:") = 0 And _
            InStr(detail, "Множитель предела продольной арматуры") = 0, "reinforcement checks have separate details", passed, failed, report
        Dim writer As New CStrengthSummaryWriter, anchor As Object
        writer.WriteSummary ThisWorkbook, batch, section:=section
        Set anchor = ThisWorkbook.Names.Item("rngStrengthSummaryAnchor").RefersToRange
        Dim lambdaLabel As String: lambdaLabel = ChrW$(&H3BB)
        TestRecord anchor.Offset(0, 35).Value2 = lambdaLabel & "*All6" And anchor.Offset(1, 35).Value2 = lambdaLabel & "*All6", _
            "spatial Auto displays the effective six-component load path in both rows", passed, failed, report
        TestRecord anchor.Offset(-2, 41).Value2 = "T при " & lambdaLabel & "ult" And _
            anchor.Offset(-2, 42).Value2 = "Qx при " & lambdaLabel & "ult" And anchor.Offset(-2, 43).Value2 = "Qy при " & lambdaLabel & "ult", _
            "compact notes identify all three spatial ultimate components", passed, failed, report
        TestRecord anchor.Offset(0, 36).ShrinkToFit And Not anchor.Offset(0, 36).WrapText, _
            "capacity method fits the existing column without wrapping", passed, failed, report
        TestRecord anchor.Offset(-1, 41).Value2 = "Tult, N*mm" And anchor.Offset(-1, 42).Value2 = "Qx,ult, N" And _
            anchor.Offset(-1, 43).Value2 = "Qy,ult, N", "the three ultimate components follow My", passed, failed, report
        Dim cellCapacity As CCapacityResult: Set cellCapacity = batch.ResultAt(1).StrengthResult.Capacity
        Dim confirmed As CSectionStateResult: Set confirmed = batch.ResultAt(1).StateRepository.FindState(sstCapacityState)
        TestRecord Not confirmed Is Nothing, "confirmed governing steel limit keeps CapacityState", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 41).Value2), cellCapacity.TUltimate, 1E-7, "Results ultimate T", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 42).Value2), cellCapacity.QxUltimate, 1E-7, "Results ultimate Qx", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 43).Value2), cellCapacity.QyUltimate, 1E-7, "Results ultimate signed Qy", passed, failed, report
        TestRecord anchor.Offset(-4, 53).Value2 = "Хомуты и поперечные ветви" And anchor.Offset(-4, 53).MergeArea.Columns.Count = 12 And _
            anchor.Offset(-4, 53).Resize(5, 12).Borders.Item(10).LineStyle <> -4142, _
            "separate Results heading spans the fully framed hoop subsection", passed, failed, report
        TestRecord anchor.Offset(0, 58).Value2 = "FAIL" And anchor.Offset(0, 64).Value2 = "FAIL", _
            "separate current and reserve statuses are written", passed, failed, report
        TestRecord anchor.Offset(0, 58).Interior.Color = StatusFillColor("FAIL") And _
            anchor.Offset(0, 64).Interior.Color = StatusFillColor("FAIL"), "both hoop statuses use the standard failure fill", passed, failed, report
        TestRecord InStr(anchor.Offset(0, 53).Value2, "TestWeakHoop (TestHoopGrade)") = 1 And _
            InStr(anchor.Offset(0, 60).Value2, "TestWeakHoop (TestHoopGrade)") = 1, _
            "Results starts with the governing hoop and its grade in each check", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 59).Value2), lambda, 0.01, "reserve is written in the hoop subsection", passed, failed, report
        TestSavedHoopCapacityFields ThisWorkbook, section, provider, batch, eventScale, passed, failed, report
    End If
    RunTestCoupledHoopCapacity = report & "TOTAL_COUPLED_HOOP_CAPACITY: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledHoopCapacity = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Разделяет фактически рассчитанное арматурное событие и общий предел на той же траектории.
Private Sub TestUnresolvedHoopState(ByVal result As CCombinationResult, ByVal eventScale As Double, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim eventCapacity As CCapacityResult: Set eventCapacity = result.StrengthResult.Capacity
    TestRecord eventCapacity.ResultMeta.ResultCode = rcCoupledConcreteLimitUnresolved And _
        eventCapacity.Status = "NumFail" And eventCapacity.ReserveFactor = 0# And eventCapacity.HasSteelLimitPoint And _
        Not eventCapacity.HasUltimatePoint And eventCapacity.TransverseLimitStatus = "OK", _
        "calculated hoop event above design is distinct from unresolved overall capacity", passed, failed, report
    TestExpectNear eventCapacity.CoupledLimitLambda, 1.2, 0.01, "hoop event scaling agrees with the reference path", passed, failed, report
    Dim overall As CSectionStateResult: Set overall = eventCapacity.StateResult
    Dim published As CSectionStateResult: Set published = result.StateRepository.FindState(sstCapacityState)
    TestRecord overall Is Nothing And published Is Nothing And _
        eventCapacity.NUltimate = 0# And eventCapacity.MxUltimate = 0# And eventCapacity.MyUltimate = 0# And _
        eventCapacity.TUltimate = 0# And eventCapacity.QxUltimate = 0# And eventCapacity.QyUltimate = 0#, _
        "unresolved steel event supplies neither the ultimate vector nor the global CapacityState", passed, failed, report
    Dim eventState As CSectionStateResult: Set eventState = eventCapacity.SteelLimitStateResult
    If eventState Is Nothing Then
        TestRecord False, "hoop event equilibrium state is missing", passed, failed, report
    Else
        Dim control As String, utilization As Double, stress As Double
        TestRecord eventState.Converged And eventState.WithinPhysicalRange And _
            eventState.CopyCoupledTransverseControl(control, utilization, stress) And _
            utilization > 0.99 And utilization <= 1# And stress > 6#, _
            "separate hoop event keeps the accepted near-limit material state", passed, failed, report
    End If
    Dim actual() As Double, expected As Variant, i As Long, hasVector As Boolean
    hasVector = eventCapacity.CopySteelLimitVector(actual)
    TestRecord hasVector, "hoop event has its dedicated vector", passed, failed, report
    If Not hasVector Then Exit Sub
    expected = Array(-216000#, 90000#, -40000#, 250000#, 120#, -80#)
    For i = 1 To 6
        TestExpectNear actual(i), CDbl(expected(i - 1)) * eventScale * eventCapacity.CoupledLimitLambda, 1E-7, _
            "signed hoop event component " & CStr(i), passed, failed, report
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет подраздел хомута при выводе на лист и восстановлении из файла,
' а также предельные поля обоих носителей по ElementID и ключу состояния.
Private Sub TestSavedHoopCapacityFields(ByVal book As Object, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
        ByVal eventScale As Double, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim settings As Object: Set settings = book.Names.Item("rngSystemSettings").RefersToRange
    TestInputCell(settings, "Results.SnapshotName").Value2 = "TestSavedHoopCapacityFields"
    Dim selector As Object, combination As Object, savedSelector As Variant, savedCombination As Variant
    Set selector = TestInputCell(settings, "Output.CombinationMode")
    Set combination = TestInputCell(settings, "Output.CombinationID")
    savedSelector = selector.Value2: savedCombination = combination.Value2
    selector.Value2 = "User": combination.Value2 = "Hoop"
    Dim props As New CSectionPropertiesCalculator, writer As New CNDMResultsWriter, summary As New CBatchResultWriter
    props.PrepareSnapshot section, provider
    Dim snap As New CCalculationSnapshot, store As New CCalculationSnapshotStore
    Set summary.CaptureSnapshot = snap: Set writer.CaptureSnapshot = snap
    summary.WriteSummary book, batch, section:=section
    writer.WriteResults book, section, props, batch
    snap.AddTable "Calculation.Profiles", book.Names.Item("rngCalculationProfiles").RefersToRange.Value2
    TestInputCell(settings, "Results.SnapshotStorage").Value2 = "File"
    store.SaveFile book, snap
    Dim anchor As Object: Set anchor = book.Names.Item("rngStrengthSummaryAnchor").RefersToRange
    Dim eventRow As Object: Set eventRow = anchor.Offset(1, 0)
    Dim eventLambda As Double: eventLambda = batch.ResultAt(2).StrengthResult.Capacity.CoupledLimitLambda
    Dim reserve As Double: reserve = CDbl(anchor.Offset(0, 59).Value2)
    Dim mode As Variant, values As Variant, n As Long, shear As Object, runId As Double
    n = section.ConcreteCount + section.RebarCount
    For Each mode In Array("Excel", "File")
        TestInputCell(settings, "Results.SnapshotStorage").Value2 = CStr(mode)
        anchor.Offset(0, 53).Resize(1, 12).ClearContents ' Проверяем вывод заново, без значений от другого носителя.
        eventRow.Offset(0, 31).Resize(1, 3).Value2 = 1# ' Подмена плоскости и общих предельных величин до восстановления снимка.
        eventRow.Offset(0, 37).Resize(1, 14).Value2 = 1#
        eventRow.Offset(0, 59).Resize(1, 6).ClearContents
        If mode = "Excel" Then
            UpdateSectionResultsForWorkbook book, snap
        Else
            anchor.Worksheet.ChartObjects("chtNDMCoupledLoadTrace").Left = anchor.Offset(0, 58).Left ' Сохранённая позиция при подтаблице из 57 столбцов.
            UpdateSectionResultsForWorkbook book
        End If
        TestRecord anchor.Offset(0, 58).Value2 = "FAIL" And anchor.Offset(0, 64).Value2 = "FAIL", _
            CStr(mode) & " restores both independent hoop statuses on Results", passed, failed, report
        TestRecord anchor.Offset(0, 35).Value2 = ChrW$(&H3BB) & "*All6" And eventRow.Offset(0, 35).Value2 = ChrW$(&H3BB) & "*All6", _
            CStr(mode) & " restores the effective six-component load path", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 59).Value2), reserve, 1E-10, _
            CStr(mode) & " restores the hoop reserve on Results", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 50).Value2), reserve, 1E-10, _
            CStr(mode) & " uses the hoop reserve as the section reserve", passed, failed, report
        TestLoadTraceChart book, "Hoop", True, True, True, passed, failed, report
        TestRecord anchor.Worksheet.ChartObjects("chtNDMCoupledLoadTrace").Left >= anchor.Offset(0, 66).Left, _
            CStr(mode) & " keeps the loading chart to the right of the hoop table", passed, failed, report
        TestRecord InStr(CStr(anchor.Offset(0, 53).Value2), "TestWeakHoop (TestHoopGrade)") = 1 And _
            InStr(CStr(anchor.Offset(0, 60).Value2), "TestWeakHoop (TestHoopGrade)") = 1, _
            CStr(mode) & " restores the hoop and its grade in both checks", passed, failed, report
        TestHoopStrainDetails book, anchor, CStr(mode), passed, failed, report
        TestSavedUnresolvedHoopEvent book, eventRow, eventScale, eventLambda, CStr(mode), passed, failed, report
        Dim expected As Variant
        expected = Array(-216000# * reserve, 306000# * reserve, -472000# * reserve, 249960# * reserve, 120# * reserve, -80# * reserve)
        TestSavedCoupledEquilibrium book, "Hoop", "CapacityState", expected, Nothing, CStr(mode), passed, failed, report
        values = ReadElementResultTable(book, "Hoop", "CapacityState")
        If HasResultTableRows(values) Then
            runId = CDbl(values(2, ResultColumn(values, "RunID")))
            Set shear = ReadCoupledFieldValues(book, vqTauXZ, runId, "Hoop", "PR4", "CapacityState")
            TestRecord UBound(values, 1) = n + 1 And shear.Count = n, CStr(mode) & " keeps capacity element and shear fields", passed, failed, report
            TestRecord Val(ReadLoadCasePropertyText(book, "Hoop", "State.CapacityState.Twist")) <> 0#, _
                CStr(mode) & " keeps the nonzero capacity twist", passed, failed, report
            TestRecord Len(ReadLoadCasePropertyText(book, "Hoop", "State.CapacityState.Rebar.2.RebarDiagramId")) > 0, _
                CStr(mode) & " links the distinct hoop material curve at capacity", passed, failed, report
            Dim pathSet As Collection: Set pathSet = ReadSavedTransversePaths(book, runId, "Hoop", "PR4", "CapacityState")
            TestRecord Not pathSet Is Nothing, CStr(mode) & " keeps the accepted path at capacity", passed, failed, report
            Set pathSet = ReadSavedTransversePaths(book, runId, "Hoop", "PR4", "CrackedState")
            TestRecord pathSet Is Nothing, CStr(mode) & " separates named states of one load case", passed, failed, report
        Else
            TestRecord False, CStr(mode) & " has no elements in the hoop capacity state", passed, failed, report
        End If
    Next mode
    TestTraceWithoutElementResults book, snap, passed, failed, report
    selector.Value2 = savedSelector: combination.Value2 = savedCombination
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет раздельные утверждения отчёта после восстановления листа или файла.
Private Sub TestSavedUnresolvedHoopEvent(ByVal book As Object, ByVal row As Object, ByVal eventScale As Double, _
        ByVal eventLambda As Double, ByVal mode As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const EVENT_ID As String = "HoopEvent"
    TestRecord row.Value2 = EVENT_ID And row.Offset(0, 51).Value2 = "NumFail" And _
        book.Application.WorksheetFunction.CountA(row.Offset(0, 31).Resize(1, 3), row.Offset(0, 37).Resize(1, 14)) = 0, _
        mode & " leaves the global ultimate plane, vector and reserve unpublished", passed, failed, report
    TestRecord row.Offset(0, 58).Value2 = "OK" And row.Offset(0, 64).Value2 = "OK" And _
        InStr(CStr(row.Offset(0, 60).Value2), "TestWeakHoop (TestHoopGrade)") = 1, _
        mode & " preserves the separate hoop event and both hoop statuses", passed, failed, report
    TestRecord ReadLoadCasePropertyText(book, EVENT_ID, "Coupled.Capacity.ResultCode") = "COUPLED_CONCRETE_LIMIT_UNRESOLVED" And _
        ReadLoadCasePropertyText(book, EVENT_ID, "Coupled.Transverse.Limit.Status") = "OK", _
        mode & " preserves separate machine and hoop event statuses", passed, failed, report
    TestRecord Len(ReadLoadCasePropertyText(book, EVENT_ID, "State.CapacityState.Equilibrium.Reference")) = 0 And _
        Not HasResultTableRows(ReadElementResultTable(book, EVENT_ID, "CapacityState")), _
        mode & " has no global CapacityState for the unresolved steel event", passed, failed, report
    If eventLambda <= 0# Then Exit Sub
    TestExpectNear CDbl(row.Offset(0, 59).Value2), eventLambda, 1E-10, mode & " restores the hoop event multiplier", passed, failed, report
    Dim key As Variant, actual As String
    For Each key In Array("Coupled.Capacity.SteelLambda", "Coupled.Transverse.Limit.Lambda")
        actual = ReadLoadCasePropertyText(book, EVENT_ID, CStr(key))
        If IsNumeric(actual) Then
            TestExpectNear CDbl(actual), eventLambda, 1E-10, mode & " keeps multiplier " & key, passed, failed, report
        Else
            TestRecord False, mode & " is missing multiplier " & key, passed, failed, report
        End If
    Next key
    TestExpectNear CDbl(row.Offset(0, 63).Value2), 0.00004, 1E-12, mode & " keeps the hoop event strain limit", passed, failed, report
    TestRecord CDbl(row.Offset(0, 61).Value2) > 6# And CDbl(row.Offset(0, 62).Value2) > 0.0000396 And _
        CDbl(row.Offset(0, 62).Value2) <= 0.00004, mode & " keeps stress and admissible strain at the separate hoop event", passed, failed, report
    Dim components As Variant, reference As Variant, i As Long, component As String
    components = Array("N", "Mx", "My", "T", "Qx", "Qy")
    reference = Array(-216000#, 90000#, -40000#, 250000#, 120#, -80#)
    For i = LBound(components) To UBound(components)
        component = CStr(components(i))
        TestRecord Len(ReadLoadCasePropertyText(book, EVENT_ID, "Coupled.Capacity." & component & "Ultimate")) = 0 And _
            IsNumeric(ReadLoadCasePropertyText(book, "Hoop", "Coupled.Capacity." & component & "Ultimate")), _
            mode & " keeps ultimate " & component & " only for the confirmed governing limit", passed, failed, report
        actual = ReadLoadCasePropertyText(book, EVENT_ID, "Coupled.Capacity.SteelLimit." & component)
        If IsNumeric(actual) Then
            TestExpectNear CDbl(actual), CDbl(reference(i)) * eventScale * eventLambda, 1E-7, _
                mode & " keeps the signed steel event component " & component, passed, failed, report
        Else
            TestRecord False, mode & " is missing steel event component " & component, passed, failed, report
        End If
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет автоматический вывод сохранённой истории без подробного НДС.
' Схема сечения и выбор носителя не определяют доступность графика.
Private Sub TestTraceWithoutElementResults(ByVal book As Object, ByVal source As CCalculationSnapshot, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim traceOnly As New CCalculationSnapshot, key As Variant
    traceOnly.RunID = source.RunID
    For Each key In source.Keys
        If key = "rngNDMElementResults" Or key = COUPLED_FIELD_TABLE Then
            traceOnly.AddTable CStr(key), source.TableHeader(CStr(key))
        Else
            traceOnly.AddTable CStr(key), source.Table(CStr(key))
        End If
    Next key
    Dim settings As Object: Set settings = book.Names.Item("rngSystemSettings").RefersToRange
    Dim selectedName As Object: Set selectedName = TestInputCell(settings, "Results.SnapshotName")
    Dim enabled As Object: Set enabled = TestInputCell(settings, "Plot.Enabled")
    Dim autoUpdate As Object: Set autoUpdate = TestInputCell(settings, "Plot.AutoUpdateAfterCalculation")
    Dim savedName As Variant, savedMode As Variant, savedEnabled As Variant, savedAuto As Variant
    savedName = selectedName.Value2: savedEnabled = enabled.Value2: savedAuto = autoUpdate.Value2
    Dim storage As Object: Set storage = TestInputCell(settings, "Results.SnapshotStorage")
    savedMode = storage.Value2
    selectedName.Value2 = "TestTraceWithoutElementResults"
    Dim mode As Variant, plotMode As Variant, note As String, warning As Boolean
    For Each mode In Array("Excel", "File")
        storage.Value2 = mode
        For Each plotMode In Array(Array("Yes", "Yes"), Array("No", "Yes"), Array("Yes", "No"))
            enabled.Value2 = plotMode(0): autoUpdate.Value2 = plotMode(1)
            ClearCoupledLoadChart book
            note = SaveAndDisplayCalculationSnapshot(book, traceOnly, warning)
            TestRecord Not HasResultTableRows(ReadElementResultTable(book, "Hoop", "StrengthState")), _
                CStr(mode) & ", Plot.Enabled=" & plotMode(0) & ", Plot.AutoUpdateAfterCalculation=" & plotMode(1) & _
                    " has no saved element state", passed, failed, report
            TestLoadTraceChart book, "Hoop", True, True, True, passed, failed, report
        Next plotMode
    Next mode
    selectedName.Value2 = savedName: storage.Value2 = savedMode
    enabled.Value2 = savedEnabled: autoUpdate.Value2 = savedAuto
    note = SaveAndDisplayCalculationSnapshot(book, source, warning)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет два сохранённых НДС со знаковой деформацией предела
' тестового класса стали. Вычисляет использование независимо от вывода.
Private Sub TestHoopStrainDetails(ByVal book As Object, ByVal anchor As Object, ByVal mode As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const HOOP_LIMIT As Double = 0.00004
    Dim direct As Variant, limit As Variant
    direct = anchor.Offset(0, 53).Resize(1, 6).Value2
    limit = anchor.Offset(0, 59).Resize(1, 6).Value2
    TestRecord direct(1, 2) > 6# And limit(1, 3) > 6#, mode & " displays stresses along the hoop from both states", passed, failed, report
    TestExpectNear CDbl(direct(1, 4)), HOOP_LIMIT, 1E-12, mode & " displays the hoop grade's tension limit", passed, failed, report
    TestExpectNear CDbl(limit(1, 5)), HOOP_LIMIT, 1E-12, mode & " displays the same limit in the limit state", passed, failed, report
    TestExpectNear CDbl(direct(1, 5)), Abs(direct(1, 3)) / HOOP_LIMIT, 1E-12, mode & " displays the independently computed utilization", passed, failed, report
    TestRecord direct(1, 3) > HOOP_LIMIT And limit(1, 4) > 0.99 * HOOP_LIMIT And limit(1, 4) <= HOOP_LIMIT, _
        mode & " keeps the over-limit design strain and the safe side of the capacity bracket distinct", passed, failed, report
    TestExpectNear CDbl(ReadLoadCasePropertyText(book, "Hoop", "Coupled.Transverse.Strain")), direct(1, 3), _
        1E-12, mode & " saves the actual design strain as a property", passed, failed, report
    TestExpectNear CDbl(ReadLoadCasePropertyText(book, "Hoop", "Coupled.Transverse.Limit.StrainLimit")), limit(1, 5), _
        1E-12, mode & " saves the governing steel limit as a property", passed, failed, report
    TestRecord ReadLoadCasePropertyText(book, "Hoop", "Coupled.Transverse.Limit.Control") = Replace$(limit(1, 2), " — предел деформации", "") And _
        ReadLoadCasePropertyText(book, "Hoop", "Coupled.Transverse.Limit.Status") = limit(1, 6), _
        mode & " documents the same governing point and FAIL in the structured properties", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет проверки СП в плоском и пространственном сочетаниях при одинаковых
' N, Mx, My и при их нулевой проекции. Профиль не запрашивает связанное НДС.
Public Function RunTestCoupledTraditionalChecks() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(300#, 400#))
    Dim x As Variant, y As Variant
    For Each x In Array(-85#, 85#)
        For Each y In Array(-150#, 150#)
            section.AddRebarElement CDbl(x), CDbl(y), 20#, GEOM_PI * 100#, "A400"
        Next y
    Next x
    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    concrete.Initialize 18#, 1.3, 26#, 1.9, 32000#, 32000#, rbMc2:=15#
    steel.Initialize 350#, 400#, 350#, 400#, 200000#, 200000#
    Dim provider As New CMaterialModelProvider: provider.TestInitializeFromParameters concrete, steel
    Dim profiles As New CCalculationProfileCatalog
    profiles.LoadFromWorkbook ThisWorkbook, False
    Dim checkProfile As CCalculationProfile: Set checkProfile = profiles.ProfileById("PR4")
    checkProfile.Initialize "PR4", "Test traditional checks", "", False, False, True, True
    checkProfile.SetStabilityValueSet "ULS(I)"
    Dim materialSpec As New CMaterialModelSpec
    materialSpec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    checkProfile.SetCrackInitiationSpec materialSpec
    materialSpec.Initialize "SLS(II)", "TwoLine", "Ignore", "TwoLine"
    checkProfile.SetCrackedStateSpec materialSpec

    Dim batch As New CBatchSectionCalculator, settings As New CSystemSettingsReader, units As New CUnitSystem
    settings.LoadFromWorkbook ThisWorkbook: units.LoadFromSettings settings
    batch.Initialize section, provider
    Set batch.ProfileCatalog = profiles
    batch.ApplySettings settings, units
    Dim sp35Error As String, sp35 As Variant
    sp35 = ReadSP35Table721FromWorkbook(ThisWorkbook, sp35Error)
    batch.SetSP35Table721 sp35, sp35Error
    batch.AddCombination "PlanarChecks", -10000#, 55000000#, 0#, "PR4", ""
    batch.AddCombination "SpatialChecks", -10000#, 55000000#, 0#, "PR4", "", torque:=1000000#, qx:=700#
    Dim path As Variant
    For Each path In Array("Auto", "LambdaAll6", "LambdaT", "LambdaNMxy")
        batch.AddCombination "ZeroNormal" & CStr(path), 0#, 0#, 0#, "PR4", "", CStr(path), torque:=1000000#, qx:=700#
    Next path
    batch.Execute

    Dim planar As CCombinationResult, spatial As CCombinationResult
    Set planar = batch.ResultAt(1): Set spatial = batch.ResultAt(2)
    TestRecord planar.StabilityMeta.Calculated And planar.StabilityMeta.InternalStatus = rsSuccess, _
        "reference standard stability " & planar.StabilityMeta.ResultComment, passed, failed, report
    TestRecord spatial.StabilityMeta.InternalStatus = planar.StabilityMeta.InternalStatus And InStr(spatial.StabilityMeta.ResultComment, "Mx") > 0, _
        "same stability result under its stated load scope", passed, failed, report
    TestRecord spatial.StabilityResult.DesignMx = planar.StabilityResult.DesignMx, "same second-order bending input", passed, failed, report
    TestRecord planar.CrackWidthMeta.Calculated And planar.CrackResult.Width.CrackWidth > 0#, _
        "reference calculation produces a normal crack width " & planar.CrackSummaryMeta.ResultComment, passed, failed, report
    TestRecord spatial.CrackWidthMeta.InternalStatus = planar.CrackWidthMeta.InternalStatus And _
        spatial.CrackResult.Width.CrackWidth = planar.CrackResult.Width.CrackWidth, "identical normal crack width after projection", passed, failed, report
    TestRecord InStr(spatial.CrackSummaryMeta.ResultComment, "наклонных") > 0, "summary discloses absent diagonal crack width", passed, failed, report
    Dim i As Long, result As CCombinationResult, caseId As String
    For i = 3 To 5
        Set result = batch.ResultAt(i)
        caseId = batch.CombinationID(i)
        TestRecord result.CrackResult.Formation.ConfirmedNotCracked And _
            result.CrackFormationMeta.ResultCode = rcCriterionNotReached And _
            result.NormalCrackStatus = "OK", caseId & " confirms zero N, Mx, My without input error", passed, failed, report
        TestRecord result.CrackWidthMeta.ResultCode = rcCrackNotFormed And Not result.CrackWidthMeta.Calculated, _
            caseId & " skips normal width", passed, failed, report
        TestRecord result.LongitudinalCrackMeta.InternalStatus = rsNotApplicable And _
            InStr(result.LongitudinalCrackMeta.ResultComment, "Сжатого бетона") > 0, _
            caseId & " independently checks whether concrete is in compression", passed, failed, report
    Next i
    Set result = batch.ResultAt(6)
    TestRecord result.CrackFormationMeta.InternalStatus = rsInvalidInput And result.NormalCrackStatus = "InputErr", _
        "explicit NMxy with no normal component remains invalid", passed, failed, report
    RunTestCoupledTraditionalChecks = report & "TOTAL_COUPLED_TRADITIONAL: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledTraditionalChecks = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает жёсткость при двух классах с независимым интегралом EA.
' Симметричные пары исключают изгиб, коэффициент Пуассона бетона равен нулю.
Public Function RunTestCoupledSteelClasses() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const EB As Double = 30000#, ES_PRIMARY As Double = 200000#, ES_SECONDARY As Double = 165000#
    Const LOAD_N As Double = 30000#
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim barArea As Double: barArea = GEOM_PI * 36#
    section.AddRebarElement -30#, 0#, 12#, barArea, "A400"
    section.AddRebarElement 30#, 0#, 12#, barArea, "A400"
    section.AddRebarElement 0#, -30#, 12#, barArea, "A240"
    section.AddRebarElement 0#, 30#, 12#, barArea, "A240"
    Dim layout As New CTransverseRebarLayout: layout.DeclareNone "Test no hoops"

    Dim concrete As New CConcreteMaterialParameters, steel As New CSteelMaterialParameters
    concrete.Initialize 30#, 3#, 30#, 3#, EB, EB, rbMc2:=15#
    steel.Initialize 350#, 400#, 350#, 400#, ES_PRIMARY, ES_PRIMARY
    Dim provider As New CMaterialModelProvider: provider.TestInitializeFromParameters concrete, steel
    Dim secondary As New CCoupledSteelModel
    secondary.Initialize ES_SECONDARY, 240#, 220#, 0#, 0#, 0.021, 0.018
    Dim steelByClass As Object: Set steelByClass = CreateObject("Scripting.Dictionary")
    steelByClass.CompareMode = vbTextCompare
    steelByClass.Add "a240", secondary
    Dim extra As Object: Set extra = CreateObject("Scripting.Dictionary")
    extra.Add "ULS(I)", steelByClass
    Dim analysis As New CCoupledSectionAnalysis
    analysis.Initialize section, layout, provider, 0#, 0.12, 80#, 125#, 0.6, 0.002, "A400", 30, extra

    Dim spec As New CMaterialModelSpec, loads As New CSectionLoadState
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    loads.Initialize LOAD_N, 0#, 0#, 0#, 0#
    Dim field() As Double, resultants() As Double, meta As CResultMeta
    Set meta = analysis.SolveTarget(loads, "Auto", spec, field, resultants)
    TestRecord meta.InternalStatus = rsSuccess, "mixed class equilibrium " & meta.ResultComment, passed, failed, report
    If meta.InternalStatus = rsSuccess Then
        Dim ea As Double
        ea = EB * (180# * 120# - 4# * barArea) + 2# * barArea * (ES_PRIMARY + ES_SECONDARY) ' Сталь замещает тот же объём бетона.
        TestExpectNear field(1) / (LOAD_N / ea), 1#, 0.0000001, "independent mixed-grade axial stiffness", passed, failed, report
        TestExpectNear resultants(1), LOAD_N, 0.0000001, "mixed-grade resultant", passed, failed, report
    End If
    spec.Initialize "SLS(II)", "ThreeLine", "UseDiagram", "TwoLine"
    Set meta = analysis.SolveTarget(loads, "Auto", spec, field, resultants)
    TestRecord meta.InternalStatus = rsInvalidConfiguration And InStr(meta.ResultComment, "A240") > 0 And _
        InStr(meta.ResultComment, "SLS(II)") > 0, "class is required for the actual material value set", passed, failed, report
    RunTestCoupledSteelClasses = report & "TOTAL_COUPLED_STEEL_CLASSES: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledSteelClasses = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проходит обычные команды формирования и расчёта из изолированной книги.
' Проверяет адрес отсутствующей физики и сохранение шести усилий в Excel и File.
Public Function RunTestCoupledWorkbookInput() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestCoupledWorkbookInput", "Тест меняет входы и Results. Запустите его из изолированной копии книги."

    Dim book As Object: Set book = ThisWorkbook
    Dim calibration As Object: Set calibration = book.Names.Item("rngCoupledMaterialSettings").RefersToRange
    Dim grades As Object: Set grades = book.Names.Item("rngCoupledSteelClasses").RefersToRange
    Dim settings As Object: Set settings = book.Names.Item("rngSystemSettings").RefersToRange
    Dim profiles As Object: Set profiles = book.Names.Item("rngCalculationProfiles").RefersToRange
    Dim loads As Object: Set loads = book.Names.Item("rngLoadCombinations").RefersToRange
    TestRecord calibration.Columns.Count = 6 And calibration.Rows.Count = 8, "named calibration layout", passed, failed, report
    TestRecord calibration.Cells(2, 2).Interior.Color = RGB(255, 242, 230), "calibration input fill", passed, failed, report
    TestRecord Not calibration.WrapText, "calibration text is not wrapped", passed, failed, report
    TestRecord calibration.Borders.Item(10).LineStyle <> -4142, "calibration has right border", passed, failed, report
    TestRecord grades.Rows.Count = 9 And grades.Columns.Count = 7, "steel class catalog layout", passed, failed, report
    TestRecord grades.Cells(2, 7).Interior.Color = RGB(255, 242, 230), "steel class input fill", passed, failed, report
    TestRecord grades.Borders.Item(10).LineStyle <> -4142, "steel class catalog has right border", passed, failed, report

    TestInputCell(settings, "Geometry.Source").Value2 = "Generated"
    TestInputCell(settings, "General.DiagramExtension").Value2 = "Yes"
    TestInputCell(settings, "Output.CombinationMode").Value2 = "Worst"
    TestInputCell(settings, "Results.SnapshotName").Value2 = "TestCoupledWorkbookInput"
    Dim pr4 As Long
    For pr4 = 3 To profiles.Columns.Count
        If CStr(profiles.Cells(2, pr4).Value2) = "PR4" Then Exit For
    Next pr4
    TestInputCell(profiles, "Calculation.Strength.DirectState", 2, pr4).Value2 = "Yes"
    TestInputCell(profiles, "Calculation.Strength.Capacity", 2, pr4).Value2 = "No"
    TestInputCell(profiles, "Calculation.Stability.Enabled", 2, pr4).Value2 = "No"
    TestInputCell(profiles, "Calculation.Crack.Width", 2, pr4).Value2 = "No"
    TestInputCell(profiles, "MaterialModel.Strength.ValueSet", 2, pr4).Value2 = "ULS(I)"
    TestInputCell(profiles, "MaterialModel.Strength.ConcreteDiagram", 2, pr4).Value2 = "ThreeLine"
    TestInputCell(profiles, "MaterialModel.Strength.ConcreteTension", 2, pr4).Value2 = "UseDiagram"
    TestInputCell(profiles, "MaterialModel.Strength.SteelDiagram", 2, pr4).Value2 = "TwoLine"
    TestInputCell(profiles, "Visualization.State", 2, pr4).Value2 = "StrengthState"
    TestInputCell(profiles, "Visualization.Quantity", 2, pr4).Value2 = "TauXZ"

    Dim firstCase(1 To 2, 1 To 10) As Variant
    firstCase(1, 1) = "Spatial": firstCase(1, 2) = -0.05: firstCase(1, 3) = 0.001: firstCase(1, 4) = -0.002
    firstCase(1, 5) = "PR4": firstCase(1, 8) = 0.001: firstCase(1, 9) = 0.002: firstCase(1, 10) = -0.003
    firstCase(2, 1) = "Plane": firstCase(2, 2) = 0.05: firstCase(2, 5) = "PR4"
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Resize(2, 10).Value2 = firstCase

    Dim unitSettings As New CSystemSettingsReader, units As New CUnitSystem
    unitSettings.LoadFromWorkbook book: units.LoadFromSettings unitSettings
    Const LOAD_OFFSET_X_MM As Double = 5#, LOAD_OFFSET_Y_MM As Double = -3#
    TestInputCell(settings, "Load.ReferenceOffsetX").Value2 = LOAD_OFFSET_X_MM / units.InputLengthToInternal(1#)
    TestInputCell(settings, "Load.ReferenceOffsetY").Value2 = LOAD_OFFSET_Y_MM / units.InputLengthToInternal(1#)
    Dim expected As Variant, loadN As Double, loadQx As Double, loadQy As Double
    loadN = units.InputForceToInternal(firstCase(1, 2))
    loadQx = units.InputShearToInternal(firstCase(1, 9)): loadQy = units.InputShearToInternal(firstCase(1, 10))
    expected = Array(units.InternalForceToOutput(loadN), _
        units.InternalMomentMxToOutput(units.InputMomentMxToInternal(firstCase(1, 3)) + LOAD_OFFSET_Y_MM * loadN), _
        units.InternalMomentMyToOutput(units.InputMomentMyToInternal(firstCase(1, 4)) + LOAD_OFFSET_X_MM * loadN), _
        units.InternalTorqueToOutput(units.InputTorqueToInternal(firstCase(1, 8)) + LOAD_OFFSET_X_MM * loadQy - LOAD_OFFSET_Y_MM * loadQx), _
        units.InternalShearToOutput(loadQx), units.InternalShearToOutput(loadQy))

    Dim geometry As Object: Set geometry = book.Names.Item("rngRectanglesGeometry").RefersToRange
    Dim transverse As Object: Set transverse = TestInputCell(geometry, "Group1")
    transverse.Value2 = "Ветвь"
    transverse.Offset(0, 1).Value2 = 8# / units.InputLengthToInternal(1#)
    transverse.Offset(0, 2).Value2 = 120# / units.InputLengthToInternal(1#)
    transverse.Offset(0, 4).Value2 = "TestHoop"
    Dim mode As Variant, section As CSectionModel, ignored As String, gt As Object
    Set gt = TestInputCell(calibration, "Coupled.Gt")
    For Each mode In Array("Excel", "File")
        TestInputCell(settings, "Geometry.Type").Value2 = "Rectangles"
        TestInputCell(settings, "Results.SnapshotStorage").Value2 = CStr(mode)
        ignored = FormSectionForWorkbook(book)
        Set section = ReadFormedSectionFromResults(book, includeMaterialDetails:=True)
        TestRecord section.RecipeType = "Rectangles", CStr(mode) & " stores recipe identity on formation", passed, failed, report
        TestRecord Len(section.RebarSteelClass(1)) > 0, CStr(mode) & " restores longitudinal steel class", passed, failed, report
        Dim cx As Double, cy As Double, area As Double, lengthFactor As Double
        section.ConcreteMeshCentroid cx, cy, area
        lengthFactor = units.InputLengthToInternal(1#)
        transverse.Offset(0, 6).Value2 = CStr((cx - 10#) / lengthFactor) & ";" & CStr(cy / lengthFactor) & ";" & _
            CStr((cx + 10#) / lengthFactor) & ";" & CStr(cy / lengthFactor)
        TestInputCell(settings, "Geometry.Type").Value2 = "Circle" ' Изменение выбора следующей геометрии не меняет уже построенное сечение.
        Dim hoopAnchor As Object: Set hoopAnchor = book.Names.Item("rngStrengthSummaryAnchor").RefersToRange
        TestInputCell(profiles, "Calculation.Strength.Capacity", 2, pr4).Value2 = "Yes"
        gt.ClearContents
        ignored = RunSectionCalculationForWorkbook(book)
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "OverallStatus") = "InputErr", CStr(mode) & " missing Gt is an input error", passed, failed, report
        TestRecord Len(ReadLoadCasePropertyText(book, "Spatial", "State.StrengthState.Equilibrium.Reference")) = 0 And _
            Len(ReadLoadCasePropertyText(book, "Spatial", "State.CapacityState.Equilibrium.Reference")) = 0, _
            CStr(mode) & " input failure saves no stale equilibrium", passed, failed, report
        TestRecord hoopAnchor.Offset(0, 58).Value2 = "InputErr" And hoopAnchor.Offset(0, 64).Value2 = "InputErr" And _
            Len(hoopAnchor.Offset(0, 59).Value2) = 0, CStr(mode) & " hoop subdivision reports failed preparation for both requested checks", passed, failed, report
        TestRecord InStr(1, ignored, gt.Worksheet.Name & "!" & gt.Address(False, False), vbTextCompare) > 0, CStr(mode) & " missing Gt names its cell", passed, failed, report
        TestRecord ReadLoadCasePropertyText(book, "Plane", "DirectStateStatus") = "OK", CStr(mode) & " planar case unaffected by calibration", passed, failed, report

        TestInputCell(profiles, "Calculation.Strength.Capacity", 2, pr4).Value2 = "No"
        gt.Value2 = 0.12
        TestInputCell(calibration, "Coupled.Gc").Value2 = 80#
        TestInputCell(calibration, "Coupled.Localization").Value2 = 125#
        TestInputCell(calibration, "Coupled.YieldRatio").Value2 = 0.6
        TestInputCell(calibration, "Coupled.PeakStrain").Value2 = 0.002
        TestInputCell(calibration, "Coupled.SteelClass").Value2 = section.RebarSteelClass(1)
        grades.Cells(2, 1).Value2 = "TestHoop"
        grades.Cells(2, 2).Value2 = "ULS(I)"
        grades.Cells(2, 3).ClearContents
        grades.Cells(2, 4).Value2 = 240#
        grades.Cells(2, 5).Value2 = 220#
        grades.Cells(2, 6).Value2 = 0.021
        grades.Cells(2, 7).Value2 = 0.018
        ignored = RunSectionCalculationForWorkbook(book)
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "OverallStatus") = "InputErr", CStr(mode) & " missing hoop modulus is an input error", passed, failed, report
        TestRecord hoopAnchor.Offset(0, 58).Value2 = "InputErr" And hoopAnchor.Offset(0, 64).Value2 = "N/A", _
            CStr(mode) & " invalid hoop steel keeps its input error in the hoop subdivision", passed, failed, report
        TestRecord InStr(1, ignored, grades.Worksheet.Name & "!" & grades.Cells(2, 3).Address(False, False), vbTextCompare) > 0, _
            CStr(mode) & " invalid class value names its cell", passed, failed, report
        grades.Cells(2, 3).Value2 = 180000#
        Dim started As Double: started = Timer
        ignored = RunSectionCalculationForWorkbook(book)
        Dim elapsed As Double: elapsed = Timer - started
        If elapsed < 0# Then elapsed = elapsed + 86400#
        report = report & "WORKBOOK_SECONDS " & CStr(mode) & " (two cases): " & CStr(elapsed) & vbCrLf
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "DirectStateStatus") = "OK", CStr(mode) & " produces linked equilibrium " & ignored, passed, failed, report
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "Coupled.StateSolver") = "SplineNewton", CStr(mode) & " saves actual solver", passed, failed, report
        TestRecord ReadLoadCasePropertyText(book, "ALL", "Geometry.RecipeType") = "Rectangles", CStr(mode) & " same saved recipe after repeated calculation", passed, failed, report
        TestRecord hoopAnchor.Offset(-4, 53).Value2 = "Хомуты и поперечные ветви", _
            CStr(mode) & " has separate hoop heading", passed, failed, report
        TestRecord hoopAnchor.Offset(0, 58).Value2 = "OK" And Len(hoopAnchor.Offset(1, 58).Value2) = 0, _
            CStr(mode) & " writes hoop status only for spatial case", passed, failed, report
        TestRecord hoopAnchor.Offset(0, 64).Value2 = "N/A" And Len(hoopAnchor.Offset(0, 59).Value2) = 0, _
            CStr(mode) & " distinguishes direct OK from a switched-off reserve search", passed, failed, report
        TestStrengthModelColumns hoopAnchor, CStr(mode), passed, failed, report
        If CStr(mode) = "File" Then
            Dim savedStrength As Variant, reader As CResultReadSource
            Set reader = CurrentResultSource(book)
            savedStrength = reader.Table("Strength.Data")
            TestRecord UBound(savedStrength, 2) = STRENGTH_SNAPSHOT_COLUMNS And savedStrength(1, 59) = "OK", _
                "File stores separate hoop status and schema", passed, failed, report
            hoopAnchor.Offset(0, 58).ClearContents
            hoopAnchor.Offset(0, 10).Resize(1, 4).Value2 = 1#
            hoopAnchor.Offset(0, 27).Value2 = "no"
            UpdateSectionResultsForWorkbook book
            TestStrengthModelColumns hoopAnchor, "File после восстановления", passed, failed, report
            TestRecord hoopAnchor.Offset(0, 58).Value2 = "OK", "File restores a removed hoop status", passed, failed, report
            TestStrengthSnapshotWidths book, reader, hoopAnchor, passed, failed, report
        End If
        Dim components As Variant, i As Long, actual As String
        components = Array("N", "Mx", "My", "T", "Qx", "Qy")
        For i = LBound(components) To UBound(components)
            actual = ReadLoadCasePropertyText(book, "Spatial", "Coupled.Internal." & components(i))
            If IsNumeric(actual) Then
                TestExpectNear CDbl(actual), CDbl(expected(i)), 0.00002, CStr(mode) & " saves internal " & components(i), passed, failed, report
            Else
                TestRecord False, CStr(mode) & " has no numerical value for " & components(i), passed, failed, report
            End If
        Next i
        TestSavedCoupledEquilibrium book, "Spatial", "StrengthState", expected, units, CStr(mode), passed, failed, report
        TestRecord Len(ReadLoadCasePropertyText(book, "Plane", "State.StrengthState.Equilibrium.Reference")) = 0 And _
            Len(ReadLoadCasePropertyText(book, "Spatial", "State.CapacityState.Equilibrium.Reference")) = 0, _
            CStr(mode) & " equilibrium output distinguishes planar and disabled states", passed, failed, report
        TestSavedCoupledFields book, CStr(mode), section.ConcreteCount + section.RebarCount, passed, failed, report
        TestSavedCoupledMaterialCurves book, CStr(mode), grades.Cells(2, 3), units, passed, failed, report
        TestLoadChartSelection book, settings, TestInputCell(profiles, "Visualization.Quantity", 2, pr4), _
            CStr(mode), passed, failed, report
    Next mode
    TestInputCell(settings, "Load.ReferenceOffsetX").Value2 = 0#
    TestInputCell(settings, "Load.ReferenceOffsetY").Value2 = 0#
    TestSavedAutoCADTransverse book, settings, units, passed, failed, report
    TestSavedCurvedTransversePaths book, units, passed, failed, report
    RunTestCoupledWorkbookInput = report & "TOTAL_COUPLED_WORKBOOK_INPUT: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledWorkbookInput = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Различает плоские предельные деформации и фактические результаты связанного НДС.
Private Sub TestStrengthModelColumns(ByVal anchor As Object, ByVal mode As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim spatial As Variant, planar As Variant, i As Long, nonempty As Long, positiveLimits As Long
    spatial = anchor.Offset(0, 10).Resize(1, 4).Value2
    planar = anchor.Offset(1, 10).Resize(1, 4).Value2
    For i = 1 To 4
        If Not IsEmpty(spatial(1, i)) Then nonempty = nonempty + 1
        If IsNumeric(planar(1, i)) Then
            If Abs(CDbl(planar(1, i))) > 0# Then positiveLimits = positiveLimits + 1
        End If
    Next i
    TestRecord anchor.Value2 = "Spatial" And nonempty = 0, _
        mode & " не подставляет плоские пределы в связанное НДС", passed, failed, report
    TestRecord anchor.Offset(1, 0).Value2 = "Plane" And positiveLimits = 4, _
        mode & " сохраняет четыре предельные деформации плоской модели", passed, failed, report
    TestRecord anchor.Offset(0, 27).Value2 = "-" And anchor.Offset(1, 27).Value2 = "no", _
        mode & " различает неприменимое продолжение и разрешённое, но не использованное", passed, failed, report
    TestRecord InStr(1, CStr(anchor.Offset(-2, 10).Value2), "плоской модели", vbTextCompare) > 0, _
        mode & " указывает применимость в шапке предельной деформации", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет шесть усилий в центре бетонной сетки и приведённые невязки.
' Переданный ожидаемый вектор получен отдельно из внешних нагрузок.
Private Sub TestSavedCoupledEquilibrium(ByVal book As Object, ByVal id As String, ByVal state As String, _
        ByVal expected As Variant, ByVal units As CUnitSystem, ByVal mode As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim data As Variant, prefix As String, label As String
    data = ReadAnchoredResultTable(book, "rngNDMSectionProperties")
    prefix = "State." & state & ".Equilibrium."
    label = mode & " " & state
    TestRecord HasResultTableRows(data), label & " has a properties table", passed, failed, report
    If Not HasResultTableRows(data) Then Exit Sub

    Dim dataRow As Long, idCol As Long, keyCol As Long, valueCol As Long, key As String
    idCol = ResultColumn(data, "LoadCase"): keyCol = ResultColumn(data, "Parameter"): valueCol = ResultColumn(data, "Value")
    Dim values As Object: Set values = CreateObject("Scripting.Dictionary")
    For dataRow = 2 To UBound(data, 1)
        If SafeText(data(dataRow, idCol)) = id Then
            key = SafeText(data(dataRow, keyCol))
            If Left$(key, Len(prefix)) = prefix Then values.Add Mid$(key, Len(prefix) + 1), data(dataRow, valueCol)
        End If
    Next dataRow
    TestRecord values.Count = 29, label & " saves the complete equilibrium block", passed, failed, report
    TestRecord values("Reference") = "ConcreteMeshCentroid" And Len(SafeText(values("Diagnostic"))) > 0, _
        label & " identifies moment reference and convergence diagnostics", passed, failed, report
    Dim hasScales As Boolean: hasScales = IsNumeric(values("ResidualArea")) And IsNumeric(values("ResidualLength"))
    TestRecord hasScales, label & " saves numerical normalization scales", passed, failed, report
    If Not hasScales Then Exit Sub

    Dim residualArea As Double, residualLength As Double, forceUnitScale As Double, momentUnitScale As Double
    Dim stressUnitScale As Double
    residualArea = CDbl(values("ResidualArea")): residualLength = CDbl(values("ResidualLength"))
    forceUnitScale = 1#: momentUnitScale = 1#: stressUnitScale = 1#
    If Not units Is Nothing Then
        residualArea = units.OutputAreaToInternal(residualArea)
        residualLength = units.OutputLengthToInternal(residualLength)
        forceUnitScale = units.InternalShearToOutput(1#)
        momentUnitScale = units.InternalTorqueToOutput(1#)
        stressUnitScale = units.InternalStressToOutput(1#)
    End If
    TestRecord residualArea > 0# And residualLength > 0#, label & " saves the physical normalization scales", passed, failed, report
    If residualArea <= 0# Or residualLength <= 0# Then Exit Sub

    Dim components As Variant, component As String, i As Long, target As Variant, restored As Variant
    Dim absolute As Variant, normalized As Variant, expectedNormalized As Double, maxNormalized As Double
    components = Array("N", "Mx", "My", "T", "Qx", "Qy")
    For i = 0 To 5
        component = components(i)
        target = values("Target." & component): restored = values("Restored." & component)
        absolute = values("AbsoluteResidual." & component): normalized = values("NormalizedResidual." & component)
        If IsNumeric(target) And IsNumeric(restored) And IsNumeric(absolute) And IsNumeric(normalized) Then
            TestRecord Abs(CDbl(target) - CDbl(expected(i))) <= 1E-10 * GeomMax(1#, Abs(CDbl(expected(i)))), _
                label & " saves the correct centered target " & component & " actual=" & target & " expected=" & expected(i), _
                passed, failed, report
            TestRecord Abs(CDbl(absolute) - Abs(CDbl(restored) - CDbl(target))) <= 8E-15 * GeomMax(1#, Abs(CDbl(target))), _
                label & " reports the absolute residual " & component, passed, failed, report
            If i >= 1 And i <= 3 Then
                expectedNormalized = CDbl(absolute) / momentUnitScale / residualArea / residualLength
            Else
                expectedNormalized = CDbl(absolute) / forceUnitScale / residualArea
            End If
            TestRecord Abs(CDbl(normalized) / stressUnitScale - expectedNormalized) <= 1E-15, _
                label & " uses the documented stress scale " & component, passed, failed, report
            maxNormalized = GeomMax(maxNormalized, CDbl(normalized))
        Else
            TestRecord False, label & " lacks numerical equilibrium values for " & component, passed, failed, report
        End If
    Next i
    If IsNumeric(values("MaxNormalizedResidual")) Then
        TestRecord Abs(CDbl(values("MaxNormalizedResidual")) - maxNormalized) <= 1E-15 * stressUnitScale, _
            label & " saves the worst integrated residual", passed, failed, report
        TestRecord maxNormalized / stressUnitScale < 0.0001, label & " has a small integrated residual", passed, failed, report
    Else
        TestRecord False, label & " lacks the worst integrated residual", passed, failed, report
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Нативное обновление сначала убирает график для плоской строки, затем
' восстанавливает его для пространственной строки из того же снимка.
Private Sub TestLoadChartSelection(ByVal book As Object, ByVal settings As Object, ByVal quantity As Object, ByVal mode As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim selector As Object, combination As Object, oldMode As Variant, oldID As Variant, oldQuantity As Variant
    Set selector = TestInputCell(settings, "Output.CombinationMode")
    Set combination = TestInputCell(settings, "Output.CombinationID")
    oldMode = selector.Formula: oldID = combination.Formula: oldQuantity = quantity.Formula
    quantity.Value2 = "Stress"
    selector.Value2 = "User": combination.Value2 = "Plane"
    UpdateSectionResultsForWorkbook book
    Dim chart As Object, hasChart As Boolean
    For Each chart In book.Worksheets("Results").ChartObjects
        If chart.Name = "chtNDMCoupledLoadTrace" Then hasChart = True: Exit For
    Next chart
    TestRecord Not hasChart, mode & " planar selection clears the load trace", passed, failed, report
    combination.Value2 = "Spatial"
    UpdateSectionResultsForWorkbook book
    TestLoadTraceChart book, "Spatial", True, False, False, passed, failed, report
    selector.Formula = oldMode: combination.Formula = oldID: quantity.Formula = oldQuantity
    UpdateSectionResultsForWorkbook book
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет сохранённые кривые с аналитическими длинами дуги и окружности.
Private Sub TestSavedCurvedTransversePaths(ByVal book As Object, ByVal units As CUnitSystem, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const RADIUS As Double = 25#
    Const RUN_ID As Double = 46100#
    Dim hoop As New CGeometryPath, arc As New CGeometryPath, layout As New CTransverseRebarLayout
    hoop.Initialize "RoundHoop": hoop.AddArc "FullCircle", 1250#, -600#, RADIUS, 0#, -2# * GEOM_PI: hoop.Seal True
    arc.Initialize "QuarterBranch": arc.AddArc "Quarter", 1250#, -600#, RADIUS, 0#, GEOM_PI / 2#: arc.Seal
    layout.AddGroup hoop, 8#, 120#, "TestHoop": layout.AddGroup arc, 8#, 120#, "TestHoop"
    Dim raw() As Variant, data As Variant, row As Long: ReDim raw(1 To 3, 1 To 24): row = 2
    PutTransversePathRows raw, row, RUN_ID, "Curved", "PR4", "StrengthState", layout, units
    data = raw
    Dim paths As Collection: Set paths = ReadSavedTransversePaths(book, RUN_ID, "Curved", "PR4", "StrengthState", data)
    If paths Is Nothing Then
        TestRecord False, "curved transverse axes are missing", passed, failed, report
        Exit Sub
    End If
    If paths.Count <> 2 Then
        TestRecord False, "curved transverse axes have incorrect grouping", passed, failed, report
        Exit Sub
    End If
    Set hoop = paths(1): Set arc = paths(2)
    TestExpectNear hoop.Length, 2# * GEOM_PI * RADIUS, 0.000000001, "saved full circle matches analytic circumference", passed, failed, report
    TestExpectNear arc.Length, GEOM_PI * RADIUS / 2#, 0.000000001, "saved arc matches analytic quarter length", passed, failed, report
    Dim edge As TRegionEdge: hoop.CopySegment 1, edge
    TestRecord hoop.IsClosed And Not arc.IsClosed And edge.Sweep = -2# * GEOM_PI, "saved full clockwise axis keeps exact closure and sign", passed, failed, report
    Dim warning As String
    On Error Resume Next
    Set paths = ReadSavedTransversePaths(book, RUN_ID + 1#, "Curved", "PR4", "StrengthState", data)
    warning = Err.Description: Err.Clear
    On Error GoTo 0
    TestRecord InStr(warning, "другому запуску") > 0 And InStr(warning, "Results!") > 0, _
        "stale paths are rejected with an address", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет расчёт из сохранённой геометрии с меткой импорта AutoCAD.
' Сетка моделирует сохранённый импорт, без обращения к открытому чертежу.
Private Sub TestSavedAutoCADTransverse(ByVal book As Object, ByVal settings As Object, ByVal units As CUnitSystem, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const CX As Double = 1250#, CY As Double = -600#
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#, CX, CY))
    section.AddRebarElement CX, CY - 30#, 12#, GEOM_PI * 36#, "Rebar"
    section.AddRebarElement CX, CY + 30#, 12#, GEOM_PI * 36#, "Rebar"
    section.SourceType = "AutoCADImport"
    Dim properties As CSectionPropertiesCalculator: Set properties = PrepareSectionSnapshot(section)
    Dim target As Object, kind As Object, secondKind As Object
    Set target = book.Names.Item("rngAutoCADTransverseRebar").RefersToRange
    Set kind = TestInputCell(target, "Group1"): Set secondKind = TestInputCell(target, "Group2")
    TestRecord target.Address = "$A$167:$K$173" And Not target.WrapText And target.Borders.Item(10).LineStyle <> -4142, _
        "native AutoCAD input layout, no wrapping and right border", passed, failed, report
    TestRecord kind.Interior.Color = RGB(255, 242, 230) And kind.HorizontalAlignment = -4108, _
        "native AutoCAD input fill and alignment", passed, failed, report
    Dim cell As Object, choice As Variant
    For Each cell In kind.Resize(3, 1).Cells
        For Each choice In Array("Нет", "Контур", "Ветвь")
            cell.Value2 = CStr(choice)
            TestRecord cell.Validation.Value, "native AutoCAD scheme list accepts " & CStr(choice), passed, failed, report
        Next choice
        For Each choice In Array("Нет,Контур,Ветвь", "Нет;Контур;Ветвь", "__INVALID__")
            cell.Value2 = CStr(choice)
            TestRecord Not cell.Validation.Value, "native AutoCAD scheme list rejects " & CStr(choice), passed, failed, report
        Next choice
        cell.ClearContents
    Next cell
    TestInputCell(settings, "Results.SnapshotName").Value2 = "TestSavedAutoCADTransverse"
    TestInputCell(settings, "Geometry.Source").Value2 = "Generated" ' Выбор для следующего формирования намеренно расходится со снимком.
    Dim mode As Variant, writer As New CNDMResultsWriter, preview As CCalculationSnapshot, store As New CCalculationSnapshotStore
    Dim lengthFactor As Double: lengthFactor = units.InputLengthToInternal(1#)
    For Each mode In Array("Excel", "File")
        TestInputCell(settings, "Results.SnapshotStorage").Value2 = CStr(mode)
        If mode = "Excel" Then
            Set preview = writer.WriteGeometryPreviewForDisplay(book, section, properties, units, False)
        Else
            Set preview = writer.PrepareGeometryPreviewSnapshot(book, section, properties, units)
            store.SaveFile book, preview
            book.Names.Item("rngNDMSectionGeometry").RefersToRange.Offset(1, 0).Resize(1, 8).ClearContents
        End If
        kind.ClearContents: secondKind.ClearContents
        Dim message As String, label As String: label = CStr(mode) & " imported section "
        message = RunSectionCalculationForWorkbook(book)
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "OverallStatus") = "InputErr" And _
            InStr(message, kind.Worksheet.Name & "!" & kind.Address(False, False)) > 0, _
            label & "requires an addressed transverse declaration", passed, failed, report
        TestRecord ReadLoadCasePropertyText(book, "Plane", "DirectStateStatus") = "OK", _
            label & "still solves the planar case", passed, failed, report

        kind.Value2 = "Нет"
        message = RunSectionCalculationForWorkbook(book)
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "DirectStateStatus") = "OK" And _
            Len(ReadLoadCasePropertyText(book, "Spatial", "Coupled.Transverse.Status")) = 0, _
            label & "accepts explicit absence", passed, failed, report
        Dim hoopSummary As Object: Set hoopSummary = book.Names.Item("rngStrengthSummaryAnchor").RefersToRange
        TestRecord hoopSummary.Value2 = "Spatial" And hoopSummary.Offset(0, 58).Value2 = "N/A" And _
            hoopSummary.Offset(0, 64).Value2 = "N/A", label & "reports the hoop checks as inapplicable for explicit None", passed, failed, report
        kind.Value2 = "Контур"
        kind.Offset(0, 1).Value2 = 8# / lengthFactor: kind.Offset(0, 2).Value2 = 120# / lengthFactor
        kind.Offset(0, 3).Value2 = 10# / lengthFactor: kind.Offset(0, 4).Value2 = "TestHoop"
        secondKind.Value2 = "Ветвь"
        secondKind.Offset(0, 1).Value2 = 8# / lengthFactor: secondKind.Offset(0, 2).Value2 = 100# / lengthFactor
        secondKind.Offset(0, 4).Value2 = "TestHoop"
        secondKind.Offset(0, 6).Value2 = CStr((CX - 10#) / lengthFactor) & ";" & CStr(CY / lengthFactor) & ";" & _
            CStr((CX + 10#) / lengthFactor) & ";" & CStr(CY / lengthFactor)
        message = RunSectionCalculationForWorkbook(book)
        TestRecord ResultsGeometrySource(book) = "AutoCADImport" And ReadLoadCasePropertyText(book, "Spatial", "DirectStateStatus") = "OK", _
            label & "solves the saved contour and WCS branch", passed, failed, report
        TestRecord InStr(ReadLoadCasePropertyText(book, "Spatial", "Coupled.Transverse.Control"), "TestHoop") > 0 And _
            ReadLoadCasePropertyText(book, "Spatial", "Coupled.Transverse.Status") = "OK", _
            label & "uses and reports the separate hoop grade", passed, failed, report
        TestSavedTransverseOverlay book, label, kind, CX, CY, passed, failed, report
    Next mode
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет открытый и замкнутый пути с геометрией примера после изменения Config.
Private Sub TestSavedTransverseOverlay(ByVal book As Object, ByVal label As String, ByVal kind As Object, _
        ByVal cx As Double, ByVal cy As Double, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    kind.Value2 = "Нет"
    Dim data As Variant: data = ReadElementResultTable(book, "Spatial", "StrengthState")
    Dim runID As Double: runID = CDbl(data(2, ResultColumn(data, "RunID")))
    Dim paths As Collection: Set paths = ReadSavedTransversePaths(book, runID, "Spatial", "PR4", "StrengthState")
    If paths Is Nothing Then
        TestRecord False, label & "missing the accepted path after editing inputs", passed, failed, report
        Exit Sub
    End If
    TestRecord paths.Count = 2, label & "keeps both accepted paths after editing inputs", passed, failed, report
    If paths.Count = 2 Then
        Dim closedPath As CGeometryPath, openPath As CGeometryPath, edge As TRegionEdge
        Set closedPath = paths(1): Set openPath = paths(2)
        TestRecord closedPath.IsClosed And Not openPath.IsClosed And openPath.SegmentCount = 1, _
            label & "distinguishes the closed hoop from the open branch", passed, failed, report
        openPath.CopySegment 1, edge
        TestRecord Abs(edge.X1 - cx + 10#) < 0.000001 And Abs(edge.X2 - cx - 10#) < 0.000001 And Abs(edge.Y1 - cy) < 0.000001, _
            label & "stores the accepted branch in the original WCS coordinates", passed, failed, report
        Dim minimumX As Double, maximumX As Double
        closedPath.ProjectionBounds 1#, 0#, minimumX, maximumX
        TestRecord Abs(minimumX - cx + 76#) < 0.000001 And Abs(maximumX - cx - 76#) < 0.000001, _
            label & "hoop axis has the expected 14 mm setback", passed, failed, report
    End If
    Set paths = ReadSavedTransversePaths(book, runID, "Plane", "PR4", "StrengthState")
    TestRecord paths Is Nothing, label & "planar case has no transverse overlay", passed, failed, report
    Dim sys As Object: Set sys = book.Names.Item("rngSystemSettings").RefersToRange
    TestInputCell(sys, "Output.CombinationMode").Value2 = "User"
    TestInputCell(sys, "Output.CombinationID").Value2 = "Spatial"
    UpdateSectionResultsForWorkbook book
    Dim chart As Object, shape As Object, shapeCount As Long
    Set chart = book.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
    For Each shape In chart.Shapes
        If Left$(shape.Name, Len("NDMPlot_Transverse")) = "NDMPlot_Transverse" Then shapeCount = shapeCount + 1
    Next shape
    TestRecord shapeCount = 2, label & "actual Excel chart creates one figure per accepted group", passed, failed, report
    TestInputCell(sys, "Output.CombinationMode").Value2 = "Worst"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Восстанавливает совместимую старую таблицу и проверяет отказ до очистки
' Results, если шапка и данные относятся к разным форматам.
Private Sub TestStrengthSnapshotWidths(ByVal book As Object, ByVal source As CResultReadSource, _
        ByVal anchor As Object, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim current As CCalculationSnapshot: Set current = source.LoadAll
    Dim legacy As New CCalculationSnapshot, legacyHoop As New CCalculationSnapshot
    Dim legacyHoopStates As New CCalculationSnapshot, damaged As New CCalculationSnapshot
    Dim key As Variant, values As Variant, cropped As Variant, hoopControl As String, hoopUtil As Double
    hoopControl = CStr(anchor.Offset(0, 53).Value2): hoopUtil = CDbl(anchor.Offset(0, 57).Value2)
    legacy.RunID = current.RunID: legacyHoop.RunID = current.RunID
    legacyHoopStates.RunID = current.RunID: damaged.RunID = current.RunID
    For Each key In current.Keys
        values = current.Table(CStr(key))
        If key = "Strength.Headers" Or key = "Strength.Data" Then
            cropped = TestCropStrengthTable(values)
            legacy.AddTable CStr(key), cropped
            legacyHoop.AddTable CStr(key), TestCropStrengthTable(values, LEGACY_HOOP_STRENGTH_SNAPSHOT_COLUMNS)
            legacyHoopStates.AddTable CStr(key), TestCropStrengthTable(values, LEGACY_HOOP_STATES_STRENGTH_SNAPSHOT_COLUMNS)
            If key = "Strength.Headers" Then
                damaged.AddTable CStr(key), cropped
            Else
                damaged.AddTable CStr(key), values
            End If
        Else
            legacy.AddTable CStr(key), values: legacyHoop.AddTable CStr(key), values
            legacyHoopStates.AddTable CStr(key), values: damaged.AddTable CStr(key), values
        End If
    Next key
    Dim errText As String
    On Error Resume Next
    UpdateSectionResultsForWorkbook book, damaged
    errText = Err.Description
    Err.Clear
    On Error GoTo 0
    TestRecord InStr(errText, "число столбцов") > 0 And anchor.Offset(0, 58).Value2 = "OK", _
        "mismatched snapshot widths are rejected before Results changes", passed, failed, report
    UpdateSectionResultsForWorkbook book, legacy
    TestRecord anchor.Value2 = "Spatial" And Len(anchor.Offset(0, 58).Value2) = 0, _
        "49-column strength snapshot restores existing cells and leaves the hoop subsection empty", passed, failed, report
    UpdateSectionResultsForWorkbook book, legacyHoop
    TestRecord anchor.Offset(0, 53).Value2 = hoopControl And anchor.Offset(0, 57).Value2 = hoopUtil And _
        anchor.Offset(0, 58).Value2 = "OK" And Len(anchor.Offset(0, 55).Value2) = 0 And Len(anchor.Offset(0, 62).Value2) = 0, _
        "57-column strength snapshot restores saved hoop cells and leaves unavailable strains blank", passed, failed, report
    UpdateSectionResultsForWorkbook book, legacyHoopStates
    TestRecord anchor.Offset(0, 58).Value2 = "OK" And IsNumeric(anchor.Offset(0, 55).Value2) And _
        Len(anchor.Offset(0, 41).Value2 & anchor.Offset(0, 42).Value2 & anchor.Offset(0, 43).Value2) = 0, _
        "62-column strength snapshot restores hoop strains and leaves unavailable ultimate components blank", passed, failed, report
    UpdateSectionResultsForWorkbook book
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует таблицу в структуру снимка из 49, 57 или 62 столбцов.
Private Function TestCropStrengthTable(ByRef values As Variant, _
        Optional ByVal width As Long = LEGACY_STRENGTH_SNAPSHOT_COLUMNS) As Variant
    Dim result() As Variant, row As Long, column As Long, endColumn As Long
    endColumn = LEGACY_STRENGTH_SNAPSHOT_COLUMNS
    If width = LEGACY_HOOP_STATES_STRENGTH_SNAPSHOT_COLUMNS Then endColumn = width
    ReDim result(1 To UBound(values, 1), 1 To width)
    For row = 1 To UBound(values, 1)
        For column = 1 To 41
            result(row, column) = values(row, column)
        Next column
        For column = 42 To endColumn
            result(row, column) = values(row, column + 3)
        Next column
    Next row
    If width = LEGACY_HOOP_STRENGTH_SNAPSHOT_COLUMNS Then
        Dim sourceColumns As Variant: sourceColumns = Array(54, 55, 58, 59, 60, 61, 65)
        For row = 1 To UBound(values, 1)
            For column = LBound(sourceColumns) To UBound(sourceColumns)
                result(row, 51 + column) = values(row, CLng(sourceColumns(column)))
            Next column
        Next row
    End If
    TestCropStrengthTable = result
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет рабочий выбор пространственного поля из фактических Results и File.
Private Sub TestSavedCoupledFields(ByVal book As Object, ByVal mode As String, ByVal elementCount As Long, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim base As Variant: base = ReadElementResultTable(book, "Spatial", "StrengthState")
    If Not HasResultTableRows(base) Then
        TestRecord False, mode & " stores elements of actual coupled state", passed, failed, report
        Exit Sub
    End If
    TestRecord UBound(base, 1) = elementCount + 1, mode & " stores elements of actual coupled state", passed, failed, report
    TestRecord IsNumeric(ReadLoadCasePropertyText(book, "Spatial", "Coupled.Twist")), mode & " stores twist", passed, failed, report
    Dim runID As Double: runID = CDbl(base(2, ResultColumn(base, "RunID")))
    Dim tau As Object, gamma As Object, warping As Object
    Set tau = ReadCoupledFieldValues(book, vqTauXZ, runID, "Spatial", "PR4", "StrengthState")
    Set gamma = ReadCoupledFieldValues(book, vqGamma, runID, "Spatial", "PR4", "StrengthState")
    Set warping = ReadCoupledFieldValues(book, vqWarping, runID, "Spatial", "PR4", "StrengthState")
    TestRecord tau.Count = elementCount And gamma.Count = elementCount And warping.Count = elementCount, _
        mode & " keeps all spatial fields", passed, failed, report
    Dim id As Variant, nonzeroTau As Boolean, nonzeroWarping As Boolean
    For Each id In tau.Keys
        If Abs(CDbl(tau(id))) > 0.000000001 Then nonzeroTau = True
        If Abs(CDbl(warping(id))) > 0.000000001 Then nonzeroWarping = True
    Next id
    TestRecord nonzeroTau And nonzeroWarping, mode & " has physical shear and warping", passed, failed, report
    Dim system As Object: Set system = book.Names.Item("rngSystemSettings").RefersToRange
    TestInputCell(system, "Output.CombinationID").Value2 = "Spatial"
    TestInputCell(system, "Output.CombinationMode").Value2 = "User"
    Dim cfg As New CSystemSettingsReader, plot As New CSectionPlotDataReader
    cfg.LoadFromWorkbook book
    plot.LoadFromWorkbook book, cfg
    TestRecord plot.Count = elementCount And plot.ResultUnit = "MPa", mode & " plot reads the actual shear field", passed, failed, report
    TestExpectNear plot.ResultValue(1), CDbl(tau(plot.ElementID(1))), 0.0000001, mode & " plot matches saved shear", passed, failed, report
    Dim paths As Collection: Set paths = plot.TransversePaths
    TestRecord Not paths Is Nothing, mode & " plot reads the saved transverse branch", passed, failed, report
    UpdateSectionResultsForWorkbook book
    Dim chart As Object, shape As Object, drawn As Boolean
    Set chart = book.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
    For Each shape In chart.Shapes
        If Left$(shape.Name, Len("NDMPlot_Transverse")) = "NDMPlot_Transverse" Then drawn = True
    Next shape
    TestRecord drawn, mode & " actual Excel chart draws the transverse path", passed, failed, report
    TestInputCell(system, "Output.CombinationMode").Value2 = "Worst"
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет узлы фактических кривых, сохранённые ссылки и независимость
' опубликованных диаграмм от последующих изменений Config.
Private Sub TestSavedCoupledMaterialCurves(ByVal book As Object, ByVal mode As String, ByVal hoopModulusCell As Object, _
        ByVal units As CUnitSystem, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim data As Variant: data = ReadAnchoredResultTable(book, "rngNDMMaterialDiagrams")
    Dim concreteId As String, mainId As String, hoopId As String
    concreteId = ReadLoadCasePropertyText(book, "Spatial", "State.StrengthState.ConcreteDiagramId")
    mainId = ReadLoadCasePropertyText(book, "Spatial", "State.StrengthState.Rebar.1.RebarDiagramId")
    hoopId = ReadLoadCasePropertyText(book, "Spatial", "State.StrengthState.Rebar.2.RebarDiagramId")
    TestRecord Len(concreteId) > 0 And Len(mainId) > 0 And Len(hoopId) > 0 And mainId <> hoopId, _
        mode & " identifies concrete and both distinct steel curves", passed, failed, report
    Dim settings As New CSystemSettingsReader, provider As New CMaterialModelProvider
    settings.LoadFromWorkbook book: provider.Initialize settings, units
    Dim fc As Double, ft As Double, eb As Double, peakStress As Double, peakStrain As Double
    Dim hasTensionPeak As Boolean, hasHoopYield As Boolean, hasMain As Boolean, hasPlanar As Boolean
    fc = units.InternalStressToOutput(provider.ConcreteCompressionResistanceFromValueSet(mvsULS))
    ft = units.InternalStressToOutput(provider.ConcreteTensionResistanceFromValueSet(mvsULS))
    eb = provider.ConcreteElasticModulusFromValueSet(mvsULS)
    Dim row As Long, sigma As Double, strain As Double
    For row = 2 To UBound(data, 1)
        If CStr(data(row, 2)) = concreteId And CStr(data(row, 7)) = "UniaxialCalibration" Then
            sigma = CDbl(data(row, 11)): strain = CDbl(data(row, 10))
            If sigma < peakStress Then peakStress = sigma: peakStrain = strain
            If Abs(sigma - ft) < 0.00000001 And Abs(strain - provider.ConcreteTensionResistanceFromValueSet(mvsULS) / eb) < 0.000000001 Then hasTensionPeak = True
        ElseIf CStr(data(row, 2)) = hoopId And CStr(data(row, 8)) = "Rebar TestHoop" Then
            If Abs(CDbl(data(row, 11)) - units.InternalStressToOutput(240#)) < 0.00000001 And _
                    Abs(CDbl(data(row, 10)) - 240# / 180000#) < 0.000000001 Then hasHoopYield = True
        ElseIf CStr(data(row, 2)) = mainId And CStr(data(row, 7)) = "MonotonicCalibration" Then
            hasMain = True
        ElseIf CStr(data(row, 7)) = "Physical" Or CStr(data(row, 7)) = "StateExtension" Then
            hasPlanar = True
        End If
    Next row
    TestExpectNear peakStress, -fc, 0.0000001, mode & " concrete compression peak stress", passed, failed, report
    TestExpectNear peakStrain, -0.002, 0.0000001, mode & " concrete compression peak strain", passed, failed, report
    TestRecord hasTensionPeak And hasHoopYield And hasMain And hasPlanar, mode & " preserves tested material laws and legacy curves", passed, failed, report
    hoopModulusCell.Value2 = 160000#
    Dim stored As Variant: stored = ReadAnchoredResultTable(book, "rngNDMMaterialDiagrams")
    Dim identical As Boolean: identical = (UBound(stored, 1) = UBound(data, 1))
    If identical Then
        For row = 2 To UBound(data, 1)
            If CStr(stored(row, 2)) <> CStr(data(row, 2)) Or CStr(stored(row, 10)) <> CStr(data(row, 10)) Or _
                    CStr(stored(row, 11)) <> CStr(data(row, 11)) Then identical = False: Exit For
        Next row
    End If
    TestRecord identical, mode & " saved curves survive a change to Config", passed, failed, report
    hoopModulusCell.Value2 = 180000#
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит тестовую ячейку по ключу строки и ожидаемому столбцу значения.
Private Function TestInputCell(ByVal table As Object, ByVal key As String, _
        Optional ByVal keyColumn As Long = 1, Optional ByVal valueColumn As Long = 2) As Object
    Dim row As Long
    For row = 1 To table.Rows.Count
        If StrComp(CStr(table.Cells(row, keyColumn).Value2), key, vbTextCompare) = 0 Then
            Set TestInputCell = table.Cells(row, valueColumn)
            Exit Function
        End If
    Next row
    Err.Raise 5, "TestInputCell", "Ключ не найден: " & key
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Измеряет один штатный запуск на тяжёлом сочетании. Подготовка входной
' книги отделена от времени чтения, расчёта и записи Results со схемой.
Public Function RunTestCoupledWorkbookTiming(ByVal calculation As String, ByVal storage As String) As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestCoupledWorkbookTiming", "Запустите замер из изолированной копии книги."
    If calculation <> "Direct" And calculation <> "Capacity" And calculation <> "Both" Then _
        Err.Raise 5, "RunTestCoupledWorkbookTiming", "Укажите Direct, Capacity или Both."
    If storage <> "Excel" And storage <> "File" Then _
        Err.Raise 5, "RunTestCoupledWorkbookTiming", "Укажите Excel или File."

    Dim book As Object: Set book = ThisWorkbook
    Dim loads As Variant: loads = TestHardShearLoads()
    Dim units As CUnitSystem: Set units = TestPrepareCoupledTimingBook(book, calculation, storage, loads)
    Dim started As Double, elapsed As Double, message As String
    started = Timer
    message = RunSectionCalculationForWorkbook(book)
    elapsed = Timer - started
    If elapsed < 0# Then elapsed = elapsed + 86400#
    report = "TIMING_SECONDS: " & CStr(elapsed) & " | " & calculation & " | " & storage & vbCrLf
    report = report & "TIMING_MESSAGE: " & Replace(message, vbCrLf, " | ") & vbCrLf

    Const CASE_ID As String = "HardShear"
    Dim directStatus As String, capacityStatus As String, lower As String, limit As String, capacityCode As String
    directStatus = ReadLoadCasePropertyText(book, CASE_ID, "DirectStateStatus")
    capacityStatus = ReadLoadCasePropertyText(book, CASE_ID, "CapacityStatus")
    lower = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity.LowerLambda")
    limit = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity.PeakLambda")
    If Len(limit) = 0 Then limit = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity.SteelLambda")
    capacityCode = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity.ResultCode")
    report = report & "TIMING_STATES: direct=" & directStatus & "; capacity=" & capacityStatus & _
        "; confirmed lower=" & lower & "; limit=" & limit & "; code=" & capacityCode & vbCrLf
    report = report & "TIMING_LIMIT_REASON: " & ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity.LimitState") & vbCrLf

    Dim components As Variant: components = Array("N", "Mx", "My", "T", "Qx", "Qy")
    Dim values As Variant
    values = Array(units.InternalForceToOutput(loads(0)), units.InternalMomentMxToOutput(loads(1)), _
        units.InternalMomentMyToOutput(loads(2)), units.InternalTorqueToOutput(loads(3)), _
        units.InternalShearToOutput(loads(4)), units.InternalShearToOutput(loads(5)))
    Dim i As Long, actual As String, lambda As Double
    If calculation <> "Capacity" Then
        TestRecord directStatus = "OK", "workbook balances the requested hard-shear case", passed, failed, report
        For i = 0 To 5
            actual = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Internal." & components(i))
            TestCoupledTimingNumber actual, CDbl(values(i)), "direct " & components(i), passed, failed, report
        Next i
        Dim chart As Object: Set chart = book.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
        Dim shape As Object, overlayCount As Long
        For Each shape In chart.Shapes
            If Left$(shape.Name, Len("NDMPlot_TransversePath")) = "NDMPlot_TransversePath" Then overlayCount = overlayCount + 1
        Next shape
        TestRecord overlayCount = 1 And InStr(message, "Схема не обновлена") = 0, _
            "automatic chart draws the closed hoop from the new snapshot", passed, failed, report
        If directStatus = "OK" Then TestCoupledActivePlotLegends book, passed, failed, report
    End If
    If calculation <> "Direct" Then
        If IsNumeric(limit) Then
            lambda = CDbl(limit)
            TestRecord lambda > 0#, "positive confirmed capacity factor", passed, failed, report
            For i = 0 To 5
                actual = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity." & components(i) & "Ultimate")
                TestCoupledTimingNumber actual, lambda * CDbl(values(i)), "capacity " & components(i), passed, failed, report
            Next i
        Else
            TestRecord capacityStatus = "NumFail" And IsNumeric(lower) And Len(capacityCode) > 0, _
                "unresolved capacity gives a confirmed lower bound and a reason", passed, failed, report
            For i = 0 To 5
                actual = ReadLoadCasePropertyText(book, CASE_ID, "Coupled.Capacity." & components(i) & "Ultimate")
                TestRecord Len(actual) = 0, "no unconfirmed limit component " & components(i), passed, failed, report
            Next i
        End If
    End If
    TestLoadTraceChart book, CASE_ID, calculation <> "Capacity", calculation <> "Direct", IsNumeric(limit), passed, failed, report, _
        (directStatus = "OK" Or directStatus = "FAIL")
    RunTestCoupledWorkbookTiming = report & "TOTAL_COUPLED_WORKBOOK_TIMING: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledWorkbookTiming = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет продвижение штатного поиска HardShear через переходы бетона.
' Точное значение конечного предела этим порогом не предписывается.
Public Function RunTestHardShearYieldRecovery() As String
    Dim report As String, passed As Long, failed As Long, lower As String, crossed As Boolean
    report = RunTestCoupledWorkbookTiming("Both", "Excel") & vbCrLf
    If InStr(1, report, "RUNTIME ERROR:", vbBinaryCompare) > 0 Then
        RunTestHardShearYieldRecovery = report
        Exit Function
    End If
    lower = ReadLoadCasePropertyText(ThisWorkbook, "HardShear", "Coupled.Capacity.LowerLambda")
    If IsNumeric(lower) Then crossed = (CDbl(lower) >= 1.1029)
    TestRecord crossed, "штатный поиск HardShear сохраняет равновесие после двух переходов бетона", passed, failed, report
    RunTestHardShearYieldRecovery = report & "TOTAL_HARD_SHEAR_YIELD_RECOVERY: passed=" & CStr(passed) & "; failed=" & CStr(failed)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет значения нативных рядов с последовательностью из активного
' носителя и различает маркер предела от последнего успешного шага.
Private Sub TestLoadTraceChart(ByVal book As Object, ByVal combination As String, ByVal wantDirect As Boolean, ByVal wantCapacity As Boolean, _
        ByVal wantLimit As Boolean, ByRef passed As Long, ByRef failed As Long, ByRef report As String, Optional ByVal targetReached As Boolean = True)
    Dim container As Object, chart As Object
    On Error Resume Next
    Set container = book.Worksheets("Results").ChartObjects("chtNDMCoupledLoadTrace")
    On Error GoTo 0
    TestRecord Not container Is Nothing, "native load trace for " & combination, passed, failed, report
    If container Is Nothing Then Exit Sub
    Set chart = container.Chart
    Dim data As Variant, source As CResultReadSource
    Set source = CurrentResultSource(book)
    If source Is Nothing Then
        data = ReadPhysicalResultTable(book, "rngNDMSectionProperties")
    Else
        data = source.Table("LoadCase.Properties")
    End If
    Dim direct As Object, capacity As Object, series As Object, marker As String, hasTarget As Boolean
    For Each series In chart.SeriesCollection
        Select Case CStr(series.Name)
            Case "Заданные усилия": Set direct = series
            Case "Поиск предела": Set capacity = series
            Case "Заданная точка": hasTarget = True
            Case "Подтверждённый максимум", "Предельная деформация арматуры", "Последняя точка поиска": marker = CStr(series.Name)
        End Select
    Next series
    Dim directMatches As Boolean
    directMatches = ((Not direct Is Nothing) = wantDirect And hasTarget = (wantDirect And targetReached))
    TestRecord directMatches, "native target marker belongs to the direct branch", passed, failed, report
    If Not direct Is Nothing Then
        TestRecord TestTraceSeriesPoints(direct, data, combination, "Coupled.Trace.Direct."), _
            "native direct series matches accepted trace", passed, failed, report
    End If
    If wantCapacity Then
        TestRecord Not capacity Is Nothing, "native capacity branch is shown independently", passed, failed, report
        If Not capacity Is Nothing Then TestRecord TestTraceSeriesPoints(capacity, data, combination, "Coupled.Trace.Capacity."), _
            "native capacity series matches accepted trace", passed, failed, report
        If wantLimit Then
            TestRecord marker = "Подтверждённый максимум" Or marker = "Предельная деформация арматуры", _
                "native marker identifies the confirmed limit", passed, failed, report
        Else
            TestRecord marker = "Последняя точка поиска", "native marker does not claim a limit for NumFail", passed, failed, report
        End If
    Else
        TestRecord capacity Is Nothing And Len(marker) = 0, "direct-only chart has no capacity branch", passed, failed, report
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет фактический порядок, значения и отсутствие сглаживания.
Private Function TestTraceSeriesPoints(ByVal series As Object, ByRef data As Variant, ByVal combination As String, _
        ByVal prefix As String) As Boolean
    Dim x As Variant, y As Variant, row As Long, point As Long, key As String, stem As String, index As Long
    x = series.XValues: y = series.Values
    If Not IsArray(x) Or Not IsArray(y) Or series.Smooth Then Exit Function
    For row = 2 To UBound(data, 1) - 1
        If CStr(data(row, 2)) = combination Then
            key = CStr(data(row, 3))
            If Left$(key, Len(prefix)) = prefix And Right$(key, 7) = ".Lambda" Then
                stem = Left$(key, Len(key) - 7)
                If CStr(data(row + 1, 3)) <> stem & ".Offset" Then Exit Function
                If CDbl(data(row + 1, 4)) = 1# Then
                    index = LBound(x) + point
                    If index > UBound(x) Then Exit Function
                    If CDbl(x(index)) <> CDbl(Mid$(stem, Len(prefix) + 1)) Then Exit Function
                    If Abs(CDbl(y(LBound(y) + point)) - CDbl(data(row, 4))) > 1E-12 Then Exit Function
                    point = point + 1
                End If
            End If
        End If
    Next row
    TestTraceSeriesPoints = point = UBound(x) - LBound(x) + 1
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Переключает поле на готовом снимке и проверяет подписи диаграммы и
' знаковые группы. Замер длительности закончен до этой проверки.
Private Sub TestCoupledActivePlotLegends(ByVal book As Object, ByRef passed As Long, _
        ByRef failed As Long, ByRef report As String)
    Dim profiles As Object: Set profiles = book.Names.Item("rngCalculationProfiles").RefersToRange
    Dim column As Long
    For column = 3 To profiles.Columns.Count
        If CStr(profiles.Cells(2, column).Value2) = "PR4" Then Exit For
    Next column
    Dim selection As Object: Set selection = TestInputCell(profiles, "Visualization.Quantity", 2, column)
    Dim settings As New CSystemSettingsReader: settings.LoadFromWorkbook book
    Dim quantity As Variant, reader As CSectionPlotDataReader, signedGroups As Boolean
    Dim index As Long, expected As String, chart As Object, shape As Object, legend As String
    For Each quantity In Array("TauXZ", "Tau", "Warping", "Stress")
        selection.Value2 = CStr(quantity)
        UpdateSectionPlotForWorkbook book, True
        If quantity <> "Stress" Then
            Set reader = New CSectionPlotDataReader
            reader.LoadFromWorkbook book, settings
            signedGroups = True
            For index = 1 To reader.Count
                Select Case Sgn(reader.ResultValue(index))
                    Case -1: expected = "Compression"
                    Case 1: expected = "Tension"
                    Case Else: expected = "NearZero"
                End Select
                If reader.PhysicalState(index) <> expected Then signedGroups = False: Exit For
            Next index
            TestRecord signedGroups, CStr(quantity) & " colours use the actual displayed sign", passed, failed, report
        End If
        Set chart = book.Worksheets("Расчет").ChartObjects("chtNDMSectionPlot").Chart
        legend = vbNullString
        For Each shape In chart.Shapes
            If Left$(shape.Name, Len("NDMPlot_LegendText")) = "NDMPlot_LegendText" Then
                legend = legend & "|" & shape.TextFrame.Characters().Text & "|"
            End If
        Next shape
        Select Case CStr(quantity)
            Case "TauXZ", "Warping"
                TestRecord InStr(legend, "|Отрицательные|") > 0 And InStr(legend, "|Положительные|") > 0 And _
                    InStr(legend, "|Сжатие|") = 0 And InStr(legend, "|Растяжение|") = 0, _
                    CStr(quantity) & " draws a signed legend", passed, failed, report
            Case "Tau"
                TestRecord InStr(legend, "|0|") > 0 And InStr(legend, "|Максимум|") > 0 And _
                    InStr(legend, "|Отрицательные|") = 0 And InStr(legend, "|Сжатие|") = 0, _
                    "shear magnitude draws a scale from zero to the maximum", passed, failed, report
            Case "Stress"
                TestRecord InStr(legend, "|Сжатие|") > 0 And InStr(legend, "|Растяжение|") > 0, _
                    "longitudinal stress keeps its material-state legend", passed, failed, report
        End Select
    Next quantity
    selection.Value2 = "TauXZ"
    UpdateSectionPlotForWorkbook book, True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет точную геометрию тяжёлой задачи штатным писателем. Материалы
' берёт из Config с одним общим коэффициентом наклона плато стали.
Private Function TestPrepareCoupledTimingBook(ByVal book As Object, ByVal calculation As String, ByVal storage As String, _
        ByVal forces As Variant) As CUnitSystem
    Dim settings As Object: Set settings = book.Names.Item("rngSystemSettings").RefersToRange
    TestInputCell(settings, "Results.SnapshotStorage").Value2 = storage
    TestInputCell(settings, "Results.SnapshotName").Value2 = "HardShearWorkbook"
    TestInputCell(settings, "Output.CombinationMode").Value2 = "Worst"
    TestInputCell(settings, "Load.ReferenceOffsetX").Value2 = 0#
    TestInputCell(settings, "Load.ReferenceOffsetY").Value2 = 0#
    TestInputCell(settings, "General.ExecutionReportEnabled").Value2 = "Yes"
    TestInputCell(settings, "General.PlateauSlopeFactor").Value2 = 0.0075
    TestInputCell(settings, "Solver.MaxIterations").Value2 = 30
    TestInputCell(settings, "Capacity.SolverMaxIterations").Value2 = 30
    TestInputCell(settings, "Capacity.InitialLambda").Value2 = 1#
    TestInputCell(settings, "Capacity.MaxLambda").Value2 = 4#
    TestInputCell(settings, "Capacity.ToleranceLambda").Value2 = 0.01

    Dim reader As New CSystemSettingsReader, units As New CUnitSystem
    reader.LoadFromWorkbook book: units.LoadFromSettings reader
    Dim calibration As Object: Set calibration = book.Names.Item("rngCoupledMaterialSettings").RefersToRange
    TestInputCell(calibration, "Coupled.Poisson").Value2 = 0.2
    TestInputCell(calibration, "Coupled.Gt").Value2 = 0.12
    TestInputCell(calibration, "Coupled.Gc").Value2 = 80#
    TestInputCell(calibration, "Coupled.Localization").Value2 = 125#
    TestInputCell(calibration, "Coupled.YieldRatio").Value2 = 16# / 30#
    TestInputCell(calibration, "Coupled.PeakStrain").Value2 = 0.002
    TestInputCell(calibration, "Coupled.SteelClass").Value2 = "A400"

    Dim concrete As Object: Set concrete = book.Names.Item("rngConcreteMaterialParameters").RefersToRange
    Dim steel As Object: Set steel = book.Names.Item("rngSteelMaterialParameters").RefersToRange
    TestInputCell(concrete, "Concrete.E").Value2 = 30000# / units.InputStressToInternal(1#)
    TestInputCell(concrete, "Concrete.E", 1, 3).Value2 = 30000# / units.InputStressToInternal(1#)
    TestInputCell(concrete, "Concrete.R.ULS(I)").Value2 = 30# / units.InputStressToInternal(1#)
    TestInputCell(concrete, "Concrete.R.ULS(I)", 1, 3).Value2 = 3# / units.InputStressToInternal(1#)
    TestInputCell(steel, "Steel.E").Value2 = 200000# / units.InputStressToInternal(1#)
    TestInputCell(steel, "Steel.E", 1, 3).Value2 = 200000# / units.InputStressToInternal(1#)
    TestInputCell(steel, "Steel.R.ULS(I)").Value2 = 350# / units.InputStressToInternal(1#)
    TestInputCell(steel, "Steel.R.ULS(I)", 1, 3).Value2 = 400# / units.InputStressToInternal(1#)
    TestInputCell(steel, "Steel.TwoLine.Es2").Value2 = 0.018
    TestInputCell(steel, "Steel.TwoLine.Es2", 1, 3).Value2 = 0.025
    Dim grades As Object: Set grades = book.Names.Item("rngCoupledSteelClasses").RefersToRange
    grades.Cells(2, 1).Value2 = "A240": grades.Cells(2, 2).Value2 = "ULS(I)"
    grades.Cells(2, 3).Value2 = 200000#: grades.Cells(2, 4).Value2 = 240#: grades.Cells(2, 5).Value2 = 240#
    grades.Cells(2, 6).Value2 = 0.02: grades.Cells(2, 7).Value2 = 0.02

    Dim profiles As Object: Set profiles = book.Names.Item("rngCalculationProfiles").RefersToRange
    Dim pr4 As Long
    For pr4 = 3 To profiles.Columns.Count
        If CStr(profiles.Cells(2, pr4).Value2) = "PR4" Then Exit For
    Next pr4
    If pr4 > profiles.Columns.Count Then Err.Raise 5, "TestPrepareCoupledTimingBook", "Профиль PR4 не найден."
    Dim flag As Variant
    For Each flag In Array("Calculation.Strength.DirectState", "Calculation.Strength.Capacity", "Calculation.Stability.Enabled", "Calculation.Crack.Width")
        TestInputCell(profiles, CStr(flag), 2, pr4).Value2 = "No"
    Next flag
    If calculation <> "Capacity" Then TestInputCell(profiles, "Calculation.Strength.DirectState", 2, pr4).Value2 = "Yes"
    If calculation <> "Direct" Then TestInputCell(profiles, "Calculation.Strength.Capacity", 2, pr4).Value2 = "Yes"
    TestInputCell(profiles, "MaterialModel.Strength.ValueSet", 2, pr4).Value2 = "ULS(I)"
    TestInputCell(profiles, "MaterialModel.Strength.ConcreteTension", 2, pr4).Value2 = "UseDiagram"
    TestInputCell(profiles, "MaterialModel.Strength.SteelDiagram", 2, pr4).Value2 = "TwoLine"
    TestInputCell(profiles, "Visualization.State", 2, pr4).Value2 = "StrengthState"
    TestInputCell(profiles, "Visualization.Quantity", 2, pr4).Value2 = "TauXZ"
    If calculation = "Capacity" Then TestInputCell(profiles, "Visualization.State", 2, pr4).Value2 = "CapacityState"

    Dim loads As Object: Set loads = book.Names.Item("rngLoadCombinations").RefersToRange
    Dim data(1 To 1, 1 To 10) As Variant
    data(1, 1) = "HardShear": data(1, 5) = "PR4"
    data(1, 2) = forces(0) / units.InputForceToInternal(1#)
    data(1, 3) = forces(1) / units.InputMomentMxToInternal(1#)
    data(1, 4) = forces(2) / units.InputMomentMyToInternal(1#)
    data(1, 8) = forces(3) / units.InputTorqueToInternal(1#)
    data(1, 9) = forces(4) / units.InputShearToInternal(1#)
    data(1, 10) = forces(5) / units.InputShearToInternal(1#)
    loads.Offset(1, 0).Resize(loads.Rows.Count - 1, loads.Columns.Count).ClearContents
    loads.Cells(2, 1).Resize(1, 10).Value2 = data

    Dim geometry As Object: Set geometry = book.Names.Item("rngRectanglesGeometry").RefersToRange
    Dim group As Object: Set group = TestInputCell(geometry, "Group1")
    group.Value2 = "Контур"
    group.Offset(0, 1).Value2 = 8# / units.InputLengthToInternal(1#)
    group.Offset(0, 2).Value2 = 100# / units.InputLengthToInternal(1#)
    group.Offset(0, 3).Value2 = 10# / units.InputLengthToInternal(1#)
    group.Offset(0, 4).Value2 = "A240": group.Offset(0, 5).ClearContents
    TestInputCell(geometry, "Group2").ClearContents
    TestInputCell(geometry, "Group3").ClearContents

    Dim section As CSectionModel: Set section = TestHardShearSection()
    section.RecipeType = "Rectangles"
    Dim properties As CSectionPropertiesCalculator: Set properties = PrepareSectionSnapshot(section)
    Dim writer As New CNDMResultsWriter, preview As CCalculationSnapshot, store As New CCalculationSnapshotStore
    If storage = "Excel" Then
        Set preview = writer.WriteGeometryPreviewForDisplay(book, section, properties, units, False)
    Else
        Set preview = writer.PrepareGeometryPreviewSnapshot(book, section, properties, units)
        store.SaveFile book, preview
    End If
    Set TestPrepareCoupledTimingBook = units
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает числовое поле Results с усилием в штатных выходных единицах.
Private Sub TestCoupledTimingNumber(ByVal actual As String, ByVal expected As Double, ByVal label As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If IsNumeric(actual) Then
        TestExpectNear CDbl(actual), expected, 0.00005, label, passed, failed, report
    Else
        TestRecord False, label & " was not published", passed, failed, report
    End If
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет совместное принятие трёх историй, отказ в одной боковой
' точке и точное воспроизведение по сохранённым полям той же траектории.
Public Function RunTestCoupledSpatialTransactions() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim operators As New CCoupledElasticState, solver As New CCoupledNewtonSolver
    Dim center As New CTestCoupledPointResponse, plusSide As New CTestCoupledPointResponse, minusSide As New CTestCoupledPointResponse
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5
    center.TestInitialize operators.PointCount
    plusSide.TestInitialize operators.PointCount
    minusSide.TestInitialize operators.PointCount
    solver.Initialize operators, center, 30000#, plusSide, minusSide

    Dim history As New Collection, field() As Double, resultants() As Double
    If Not solver.TryStep(9000#, 1000#, -1200#, 900#, 650#, -520#, 1E-8, 1E-8, 25) Then _
        Err.Raise 5, "RunTestCoupledSpatialTransactions", "Не найдено первое пространственное равновесие. " & solver.ResultMeta.ResultComment
    solver.CopyTrial field, resultants
    history.Add field
    solver.AcceptStep
    Dim firstCenter As Double, firstPlus As Double, firstMinus As Double
    firstCenter = center.TestAcceptedPeak: firstPlus = plusSide.TestAcceptedPeak: firstMinus = minusSide.TestAcceptedPeak
    TestRecord center.TestAcceptedSteps = 1 And plusSide.TestAcceptedSteps = 1 And minusSide.TestAcceptedSteps = 1, _
        "первая фиксация записывает три истории", passed, failed, report
    TestRecord Abs(firstPlus - firstMinus) > 1E-12, "боковые истории различаются", passed, failed, report

    minusSide.TestFail = True ' Отказ в точке 2 после успешной записи центра и положительной стороны.
    TestRecord Not solver.TryStep(14000#, 2000#, -3200#, 1500#, 970#, -850#, 1E-8, 1E-8, 25), _
        "отказ одной боковой точки отклоняет весь шаг", passed, failed, report
    solver.DiscardStep
    TestRecord center.TestAcceptedSteps = 1 And plusSide.TestAcceptedSteps = 1 And minusSide.TestAcceptedSteps = 1 And _
        center.TestAcceptedPeak = firstCenter And plusSide.TestAcceptedPeak = firstPlus And minusSide.TestAcceptedPeak = firstMinus, _
        "отклонённая проба сохраняет все три истории", passed, failed, report

    minusSide.TestFail = False
    If Not solver.TryStep(14000#, 2000#, -3200#, 1500#, 970#, -850#, 1E-8, 1E-8, 25) Then _
        Err.Raise 5, "RunTestCoupledSpatialTransactions", "Повтор после отказа не подтвердил равновесие. " & solver.ResultMeta.ResultComment
    solver.CopyTrial field, resultants
    history.Add field
    solver.AcceptStep
    Dim lastCenter As Double, lastPlus As Double, lastMinus As Double
    lastCenter = center.TestAcceptedPeak: lastPlus = plusSide.TestAcceptedPeak: lastMinus = minusSide.TestAcceptedPeak
    TestRecord center.TestAcceptedSteps = 2 And plusSide.TestAcceptedSteps = 2 And minusSide.TestAcceptedSteps = 2, _
        "успешный повтор принимает поле только один раз", passed, failed, report
    TestRecord solver.ReplayAcceptedFields(history, 1), "повтор первого состояния", passed, failed, report
    TestRecord center.TestAcceptedSteps = 1 And plusSide.TestAcceptedSteps = 1 And minusSide.TestAcceptedSteps = 1 And _
        center.TestAcceptedPeak = firstCenter And plusSide.TestAcceptedPeak = firstPlus And minusSide.TestAcceptedPeak = firstMinus, _
        "первое принятое поле восстанавливает каждую историю", passed, failed, report
    TestRecord solver.ReplayAcceptedFields(history, 2), "повтор полной последовательности", passed, failed, report
    TestRecord center.TestAcceptedSteps = 2 And plusSide.TestAcceptedSteps = 2 And minusSide.TestAcceptedSteps = 2 And _
        center.TestAcceptedPeak = lastCenter And plusSide.TestAcceptedPeak = lastPlus And minusSide.TestAcceptedPeak = lastMinus, _
        "полная последовательность восстанавливает каждую историю", passed, failed, report
    solver.ResetHistory
    TestRecord center.TestAcceptedSteps = 0 And plusSide.TestAcceptedSteps = 0 And minusSide.TestAcceptedSteps = 0 And _
        center.TestAcceptedPeak = 0# And plusSide.TestAcceptedPeak = 0# And minusSide.TestAcceptedPeak = 0#, _
        "новое сочетание получает три ненагруженных истории", passed, failed, report
    RunTestCoupledSpatialTransactions = report & "TOTAL_COUPLED_SPATIAL_TRANSACTIONS: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledSpatialTransactions = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет непрерывность номинального сдвига вблизи поверхности текучести
' и независимое локальное замыкание возврата при двух дилатансиях.
Public Function RunTestRegularYieldOnset() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const E As Double = 30000#
    Const NU As Double = 0.2
    Const CC As Double = 16#
    Const FT As Double = 3#
    Const ALPHA As Double = 0.12
    Const GAMMA As Double = 3#
    Const ECC As Double = 0.1
    Const PROBE_LEVEL As Double = 0.00000001
    Const MARGIN As Double = 0.000000000001
    Dim material As New CCoupledConcreteResponse, response As ICoupledPointResponse, projection As New CCoupledConcreteReturn
    material.Initialize 2, E, NU, CC, 30#, FT, 0.002, 0.12, 80#, 125#
    Set response = material
    projection.Initialize E, NU, ALPHA, GAMMA, ECC
    Dim tensor(1 To 6, 1 To 6) As Double
    CoupledElasticTensor E, NU, tensor
    Dim beta As Double, yieldShear As Double, tau(1 To 2) As Double, apparent(1 To 2) As Double
    beta = CC / FT * (1# - ALPHA) - (1# + ALPHA)
    yieldShear = CC * (1# - ALPHA) / (Sqr(3#) + beta)
    tau(1) = yieldShear * (1# + PROBE_LEVEL - MARGIN)
    tau(2) = yieldShear * (1# + PROBE_LEVEL + MARGIN)

    Dim point As Long, trialYield As Double, reserve As Double, historyT As Double, historyC As Double
    Dim strain(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double, status As ESolverFailureCode
    For point = 1 To 2
        trialYield = (Sqr(3#) + beta) * tau(point) / (1# - ALPHA) - CC
        reserve = trialYield - PROBE_LEVEL * CC
        TestRecord Abs(reserve - CDbl(2 * point - 3) * MARGIN * CC) < 2E-14, _
            "yield residual brackets reference probe level " & CStr(point), passed, failed, report
        strain(6) = tau(point) / tensor(6, 6)
        status = response.EvaluatePoint(point, strain, stress, tangent, ctmStressOnly, recordTrial:=False)
        TestRecord status = sfcNone, "virgin shear response " & CStr(point), passed, failed, report
        apparent(point) = stress(6)
        material.AcceptedPointDamage point, historyT, historyC
        material.CopyAcceptedStress point, stress
        TestRecord historyT = 0# And historyC = 0# And stress(6) = 0#, _
            "read-only probe leaves virgin history " & CStr(point), passed, failed, report
        report = report & "ONSET_SAMPLE;" & CStr(point) & ";" & TestOnsetNumber(tau(point)) & ";" & _
            TestOnsetNumber(reserve) & ";" & TestOnsetNumber(apparent(point)) & vbCrLf
    Next point
    Dim signal As Double, inputGap As Double
    signal = apparent(2) - apparent(1): inputGap = tau(2) - tau(1)
    TestRecord Abs(signal) < 20# * inputGap, "regular return is continuous near yield", passed, failed, report
    report = report & "ONSET_GAP;" & TestOnsetNumber(inputGap) & ";" & TestOnsetNumber(signal) & vbCrLf

    Dim dilation As Variant
    For Each dilation In Array(0#, 0.3)
        For point = 1 To 2
            TestRegularOnsetReturn projection, tensor, tau(point), CC, FT, ALPHA, GAMMA, ECC, CDbl(dilation), point, passed, failed, report
        Next point
    Next dilation
    TestYieldBranchReadTransaction material, response, E, NU, FT, passed, failed, report
    RunTestRegularYieldOnset = report & "TOTAL_REGULAR_YIELD_ONSET: passed=" & CStr(passed) & "; failed=" & CStr(failed)
    Exit Function
Failed:
    RunTestRegularYieldOnset = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет одностороннюю касательную у начала растягивающего возврата
' и готовый буфер, оставшийся после читающих проб другой деформации.
Private Sub TestYieldBranchReadTransaction(ByVal material As CCoupledConcreteResponse, ByVal response As ICoupledPointResponse, _
        ByVal modulus As Double, ByVal poisson As Double, ByVal tensileStrength As Double, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim strain(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim saved(1 To 2) As Double, point As Long, status As ESolverFailureCode
    For point = 1 To 2
        strain(1) = point * 0.25 * tensileStrength / modulus
        strain(2) = -poisson * strain(1): strain(3) = strain(2)
        status = response.EvaluatePoint(point, strain, stress, tangent, ctmStressOnly)
        If status <> sfcNone Then Err.Raise 5, , "Не создано пробное поле для проверки буфера бетона."
        saved(point) = stress(1)
    Next point

    strain(1) = (1# - 0.00000008) * tensileStrength / modulus
    strain(2) = -poisson * strain(1): strain(3) = strain(2)
    Dim gradient(1 To 6) As Double, centerF As Double, leftF As Double, rightF As Double, delta As Double
    Dim centerX As Double, leftStress As Double, slope As Double, hadPlastic As Boolean, principalSign As Long, substeps As Long
    centerX = strain(1): delta = 0.000000002 * (1# + Abs(centerX))
    If Not material.TryReadTrialYield(1, strain, centerF, gradient) Then Err.Raise 5, , "Не прочитан исходный критерий."
    strain(1) = centerX + delta
    If Not material.TryReadTrialYield(1, strain, rightF, gradient, False) Then Err.Raise 5, , "Не прочитана правая оценка критерия."
    strain(1) = centerX - delta
    If Not material.TryReadTrialYield(1, strain, leftF, gradient, False) Then Err.Raise 5, , "Не прочитана левая оценка критерия."
    TestRecord leftF < centerF And centerF < 0# And rightF > 0#, _
        "численное возмущение пересекает поверхность текучести", passed, failed, report
    TestRecord Abs((rightF - leftF) / (2# * delta) - gradient(1)) < 0.01, _
        "градиент Ftrial согласуется с разностью упругой пробы", passed, failed, report
    status = response.EvaluatePoint(1, strain, stress, tangent, ctmStressOnly, False)
    If status <> sfcNone Then Err.Raise 5, , "Не получено напряжение левой пробы."
    leftStress = stress(1)
    strain(1) = centerX
    Dim validBranch As Boolean: validBranch = material.TryReadConcreteBranch(1, strain, hadPlastic, principalSign, substeps)
    TestRecord validBranch And Not hadPlastic And principalSign = 1 And substeps = 1, _
        "исходная станция остаётся на упругой ветви", passed, failed, report
    status = response.EvaluatePoint(1, strain, stress, tangent, ctmFullBranchConsistent, False)
    If status <> sfcNone Then Err.Raise 5, , "Не построена касательная согласованной ветви."
    slope = (stress(1) - leftStress) / delta
    TestRecord Abs(tangent(1, 1) - slope) < 0.001, _
        "касательная использует возмущение своей ветви", passed, failed, report
    status = response.EvaluatePoint(1, strain, stress, tangent, ctmFullBranchConsistent, True)
    TestRecord status = sfcInvalidConfiguration, "режим касательной ветви отклоняет запись", passed, failed, report
    response.AcceptTrial
    For point = 1 To 2
        material.CopyAcceptedStress point, stress
        TestRecord stress(1) = saved(point), "сохранена ранее записанная проба " & CStr(point), passed, failed, report
    Next point

    strain(1) = 1.02 * tensileStrength / modulus
    strain(2) = -poisson * strain(1): strain(3) = strain(2)
    For point = 1 To 2
        status = response.EvaluatePoint(point, strain, stress, tangent, ctmStressOnly)
        If status <> sfcNone Then Err.Raise 5, , "Не создано принятых напряжений после растягивающего возврата."
    Next point
    response.AcceptTrial
    material.CopyAcceptedStress 1, stress
    Dim acceptedStress As Double: acceptedStress = stress(1)
    status = response.EvaluatePoint(1, strain, stress, tangent, ctmFullBranchConsistent, False)
    TestRecord status = sfcNone And stress(1) = acceptedStress, _
        "касательная построена при точно принятой деформации", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Восстанавливает поверхность и уравнения обратного Эйлера из возвращённых
' напряжений, матрицы упругости и градиента гиперболического потенциала.
Private Sub TestRegularOnsetReturn(ByVal projection As CCoupledConcreteReturn, ByRef tensor() As Double, _
        ByVal tau As Double, ByVal cc As Double, ByVal ft As Double, ByVal alpha As Double, ByVal gamma As Double, _
        ByVal eccentricity As Double, ByVal dilation As Double, ByVal point As Long, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim elastic(1 To 6) As Double, corrected(1 To 6) As Double, increment(1 To 6) As Double, flow(1 To 6) As Double
    elastic(6) = tau
    Dim multiplier As Double, factor As Double, multiplierDivisor As Double, ok As Boolean, tag As String
    tag = TestOnsetNumber(dilation) & ";" & CStr(point)
    ok = projection.Correct(elastic, 0#, Sqr(3#) * tau, tau, cc, ft, ft, dilation, _
         corrected, increment, multiplier, factor, multiplierDivisor)
    TestRecord ok And multiplier > 0# And factor < 1# And multiplierDivisor = 1#, "active local return " & tag, passed, failed, report
    If Not ok Then Exit Sub
    Dim trace As Double, q As Double, high As Double, middle As Double, low As Double, yieldValue As Double
    Dim unused As Double, backward As Double, maxBackward As Double, maxFlow As Double, i As Long, j As Long
    CoupledStressInvariants corrected, trace, q, high, middle, low
    yieldValue = CoupledBarcelonaYield(trace, q, high, cc, ft, alpha, gamma)
    unused = CoupledConcretePotential(corrected, trace, q, dilation, ft, eccentricity, flow)
    For i = 1 To 6
        backward = corrected(i) - elastic(i)
        For j = 1 To 6: backward = backward + tensor(i, j) * increment(j): Next j
        maxBackward = GeomMax(maxBackward, Abs(backward))
        maxFlow = GeomMax(maxFlow, Abs(increment(i) - multiplier * flow(i)))
    Next i
    TestRecord Abs(yieldValue) <= 2E-12, "local F closes " & tag, passed, failed, report
    TestRecord maxBackward < 2E-12 And maxFlow < 2E-13, "local backward and flow equations close " & tag, passed, failed, report
    If dilation > 0# Then
        TestRecord trace < 0#, "positive dilation reduces pressure trace " & tag, passed, failed, report
    Else
        TestRecord trace = 0#, "zero dilation preserves pressure trace " & tag, passed, failed, report
    End If
    report = report & "ONSET_RETURN;" & tag & ";" & TestOnsetNumber(yieldValue) & ";" & TestOnsetNumber(maxBackward) & _
        ";" & TestOnsetNumber(maxFlow) & ";" & TestOnsetNumber(trace) & ";" & TestOnsetNumber(multiplier) & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пишет десятичные числа языка отчёта вне зависимости от локали Excel.
Private Function TestOnsetNumber(ByVal value As Double) As String
    TestOnsetNumber = Trim$(Str$(value))
End Function


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет запись всех трёх фактических историй при равновесии без
' поправки после двух разных пространственных шагов.
Public Function RunTestCoupledConsistentZeroStep() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim operators As New CCoupledElasticState, solver As New CCoupledNewtonSolver
    Dim center As New CCoupledConcreteResponse, plusSide As New CCoupledConcreteResponse, minusSide As New CCoupledConcreteResponse
    operators.PrepareModel section, 30000#, 0.2, 1, 1, 5
    center.Initialize operators.PointCount, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    plusSide.Initialize operators.PointCount, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    minusSide.Initialize operators.PointCount, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    solver.Initialize operators, center, 30000#, plusSide, minusSide

    If Not solver.TryStep(-9000#, 1000#, -1200#, 900#, 650#, -520#, 1E-8, 1E-8, 25) Then _
        Err.Raise 5, , "Не найдено первое пространственное равновесие. " & solver.ResultMeta.ResultComment
    solver.AcceptStep
    Dim first(1 To 6, 1 To 3) As Double, second(1 To 6, 1 To 3) As Double, last(1 To 6, 1 To 3) As Double
    TestCopyConcreteTriplet center, plusSide, minusSide, first
    If Not solver.TryStep(-14000#, 2000#, -3200#, 1500#, 970#, -850#, 1E-8, 1E-8, 25) Then _
        Err.Raise 5, , "Не найдено второе пространственное равновесие. " & solver.ResultMeta.ResultComment
    solver.AcceptStep
    TestCopyConcreteTriplet center, plusSide, minusSide, second

    Dim baseLoad(1 To 6) As Double, direction(1 To 6) As Double, projection() As Double
    Dim acceptedField() As Double, trialField() As Double, resultants() As Double, lambda As Double
    baseLoad(1) = -14000#: baseLoad(2) = 2000#: baseLoad(3) = -3200#
    baseLoad(4) = 1500#: baseLoad(5) = 970#: baseLoad(6) = -850#
    direction(1) = 10000#
    ReDim projection(1 To operators.DegreeCount)
    projection(1) = 1#
    solver.CopyAccepted acceptedField
    If Not solver.TryDeformationStep(baseLoad, direction, projection, acceptedField(1), 0#, lambda, 1E-8, 1E-8, 25, _
            consistentControl:=True) Then _
        Err.Raise 5, , "Не подтверждена нулевая поправка. " & solver.ResultMeta.ResultComment

    solver.CopyTrial trialField, resultants
    Dim d As Long, fieldError As Double
    For d = 1 To UBound(acceptedField)
        fieldError = GeomMax(fieldError, Abs(trialField(d) - acceptedField(d)))
    Next d
    TestRecord fieldError = 0# And lambda = 0#, "поле и множитель совпадают точно при нулевой поправке", passed, failed, report
    TestCopyConcreteTriplet center, plusSide, minusSide, last
    Dim side As Long, component As Long, difference As Double
    Dim acceptedDifference As Double, stationsChanged As Long, priorMaterialPreserved As Boolean
    priorMaterialPreserved = True
    For side = 1 To 3
        difference = 0#
        For component = 1 To 6
            difference = GeomMax(difference, Abs(first(component, side) - second(component, side)))
            If second(component, side) <> last(component, side) Then priorMaterialPreserved = False
        Next component
        If difference > 1E-7 Then stationsChanged = stationsChanged + 1
    Next side
    TestRecord stationsChanged = 3, "первое и второе поля дают разные напряжения на каждой станции", passed, failed, report
    TestRecord priorMaterialPreserved, "подтверждённая проба пока сохраняет принятую историю", passed, failed, report

    solver.AcceptStep
    TestCopyConcreteTriplet center, plusSide, minusSide, last
    For side = 1 To 3
        For component = 1 To 6
            acceptedDifference = GeomMax(acceptedDifference, Abs(last(component, side) - second(component, side)))
        Next component
    Next side
    TestRecord acceptedDifference = 0#, "принятие нулевой поправки публикует три совпадающие материальные истории", passed, failed, report
    RunTestCoupledConsistentZeroStep = report & "TOTAL_COUPLED_CONSISTENT_ZERO_STEP: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConsistentZeroStep = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает шесть принятых напряжений первой бетонной точки на каждой станции.
Private Sub TestCopyConcreteTriplet(ByVal center As CCoupledConcreteResponse, ByVal plusSide As CCoupledConcreteResponse, _
        ByVal minusSide As CCoupledConcreteResponse, ByRef values() As Double)
    Dim stress(1 To 6) As Double, i As Long
    center.CopyAcceptedStress 1, stress
    For i = 1 To 6: values(i, 1) = stress(i): Next i
    plusSide.CopyAcceptedStress 1, stress
    For i = 1 To 6: values(i, 2) = stress(i): Next i
    minusSide.CopyAcceptedStress 1, stress
    For i = 1 To 6: values(i, 3) = stress(i): Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет два максимума на подшагах, сохранённую нагрузку и принятую историю.
Public Function RunTestCoupledRetainedShoulders() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const STEEL_AREA As Double = 100#, DESIGN_FORCE As Double = 410000#
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, Sqr(4# * STEEL_AREA / GEOM_PI), STEEL_AREA, "A400"
    Dim steel As New CCoupledSteelModel, models As Object
    steel.Initialize 200000#, 40#, 40#, 0#, 0#, 0.02, 0.02
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "A400", steel
    Dim operators As New CCoupledElasticState, concrete As New CTestCoupledPointResponse
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModels:=models
    concrete.TestInitialize operators.PointCount
    concrete.TestUseTwinPeakAxialLaw
    Dim material As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    material.Initialize operators, concrete, models
    solver.Initialize operators, material, 30000#

    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    direction(1) = DESIGN_FORCE
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    Dim history As New Collection, trace As New CCoupledLoadTrace, field() As Double, resultants() As Double
    Dim coordinate As Variant, strain As Double, lambda As Double, previous As Double, factor As Double
    For Each coordinate In Array(0.0004, 0.0006)
        strain = CDbl(coordinate)
        If Not solver.TryDeformationStep(offset, direction, control, strain, lambda, factor, 1E-8, 1E-8, 30) Then _
            Err.Raise 5, , "Не получено исходное равновесие. " & solver.ResultMeta.DiagnosticDetails
        solver.CopyTrial field, resultants
        solver.AcceptStep
        history.Add field
        trace.Add 1#, factor
        previous = lambda: lambda = factor
    Next coordinate
    TestExpectNear previous, 0.75, 1E-8, "первое исходное поле", passed, failed, report
    TestExpectNear lambda, 0.8, 1E-8, "второе исходное поле", passed, failed, report

    Dim branch As New CCoupledDeformationContinuation
    branch.Trace solver, material, offset, direction, history, previous, lambda, resultants, 1.5, 0.0001, 30, 1E-8, 1E-8, trace
    Dim saved As CCoupledLoadTrace: Set saved = branch.SavedTrace
    Dim greatest As Double
    greatest = TestTwinPeakRetainedTrace(branch, solver, passed, failed, report)
    Dim meta As CResultMeta: Set meta = branch.ResultMeta
    TestRecord InStr(1, meta.ResultComment, "локальный максимум при lambda = ") > 0, _
        "в диагностике указан больший подтверждённый максимум", passed, failed, report
    TestRecord InStr(1, meta.DiagnosticDetails, "равновесие при lambda = 1 подтверждено: True", vbTextCompare) > 0, _
        "точное сочетание проверено из уточнённого восходящего участка", passed, failed, report

    Dim accuracy As Double, designVerified As Boolean, limitFields As Collection
    TestRecord meta.InternalStatus = rsNumericalFailure And Not branch.CopyPeak(factor, accuracy, designVerified), _
        "локальный пик железобетона не опубликован как предел", passed, failed, report
    TestRecord Not branch.CopyLimitState(limitFields, resultants), _
        "локальный пик не публикует предельное поле", passed, failed, report
    solver.CopyAccepted field
    TestExpectNear concrete.TestAcceptedPeak, field(1), 1E-12, _
        "отклонённая проба не изменила принятую материальную историю", passed, failed, report
    TestRecord concrete.TestAcceptedSteps = saved.Count, "индексы истории и следа совпадают", passed, failed, report
    RunTestCoupledRetainedShoulders = report & "TOTAL_COUPLED_RETAINED_SHOULDERS: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledRetainedShoulders = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Находит больший максимум заданного полинома из изолированного корня его производной на t в [2;3].
' Деформация равна 0,0004+0,0001t; множитель равен (-218t⁴+1309t³-2428t²+1379t+450)/600.
Private Function TestTwinPeakMaximum() As Double
    Dim low As Double, high As Double, t As Double, slope As Double, i As Long
    low = 2#: high = 3#
    For i = 1 To 48
        t = 0.5 * (low + high)
        slope = ((-872# * t + 3927#) * t - 4856#) * t + 1379#
        If slope > 0# Then low = t Else high = t
    Next i
    t = 0.5 * (low + high)
    TestTwinPeakMaximum = (((-218# * t + 1309#) * t - 2428#) * t + 1379#) * t / 600# + 0.75
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет все сохранённые поля по явному закону, строгий ход деформации, максимум и точное принятое состояние.
Private Function TestTwinPeakRetainedTrace(ByVal branch As CCoupledDeformationContinuation, ByVal solver As CCoupledNewtonSolver, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String) As Double
    Dim saved As CCoupledLoadTrace: Set saved = branch.SavedTrace
    Dim fields As Collection: Set fields = branch.AcceptedFields
    If saved Is Nothing Or fields Is Nothing Then
        TestRecord False, "доступна полная уточнённая история полей и нагрузок", passed, failed, report
        Exit Function
    End If
    TestRecord saved.Count = fields.Count And saved.Count > 2, "все уточнённые поля согласованы с нагрузочным следом", passed, failed, report
    Dim i As Long, fraction As Double, factor As Double, previous As Double, t As Double, expected As Double
    Dim field() As Double, accepted() As Double, same As Boolean
    For i = 1 To GeomMin(saved.Count, fields.Count)
        field = fields.Item(i): saved.ReadPoint i, fraction, factor
        t = (field(1) - 0.0004) / 0.0001
        expected = (((-218# * t + 1309#) * t - 2428#) * t + 1379#) * t / 600# + 0.75
        TestRecord Abs(factor - expected) <= 0.000000001, "каждое сохранённое поле удовлетворяет независимому полиному " & CStr(i), passed, failed, report
        TestRecord fraction = 1# And field(1) > previous, "каждое принятое поле продвигается при полной постоянной части " & CStr(i), passed, failed, report
        report = report & "TEST_TWIN_PEAK_FIELD;" & CStr(i) & ";" & CStr(field(1)) & ";" & CStr(factor) & vbCrLf
        previous = field(1)
        TestTwinPeakRetainedTrace = GeomMax(TestTwinPeakRetainedTrace, factor)
    Next i
    TestRecord Abs(TestTwinPeakRetainedTrace - TestTwinPeakMaximum()) <= 0.0001, _
        "уточнённый максимум совпадает с независимым корнем производной до 0.0001", passed, failed, report
    TestRecord Abs(branch.BestAdmissibleLambda - TestTwinPeakRetainedTrace) <= 0.000000000001, _
        "граница подтверждения равна максимуму всех сохранённых равновесий", passed, failed, report
    solver.CopyAccepted accepted: same = (UBound(accepted) = UBound(field))
    If same Then
        For i = 1 To UBound(field)
            If accepted(i) <> field(i) Then same = False: Exit For
        Next i
    End If
    TestRecord same, "последнее сохранённое поле точно совпадает с принятым решателем", passed, failed, report
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет прямую пробу при единице до и после замены принятой истории.
Public Function RunTestCoupledDesignWitness() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    TestDesignWitnessAfterRefinement True, passed, failed, report
    TestDesignWitnessAfterRefinement False, passed, failed, report
    TestDesignWitnessInPrefix passed, failed, report
    RunTestCoupledDesignWitness = report & "TOTAL_COUPLED_DESIGN_WITNESS: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledDesignWitness = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Уточняет два плеча кубической силовой ветви с возможным отказом в единственной полосе деформаций.
Private Sub TestDesignWitnessAfterRefinement(ByVal blocksRefined As Boolean, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const AREA As Double = 21600#, LEFT_STRAIN As Double = 0.00075, DESIGN_STRAIN As Double = 0.0008114013519
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    Dim operators As New CCoupledElasticState, material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5
    material.TestInitialize operators.PointCount
    material.TestUseDescendingAxialLaw
    If blocksRefined Then material.TestBlockRefinedDesign
    solver.Initialize operators, material, 30000#
    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    offset(1) = 4# * AREA: direction(1) = 15# * AREA
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    Dim history As New Collection, trace As New CCoupledLoadTrace
    Dim field() As Double, resultants() As Double, trialResults() As Double
    Dim previous As Double, lambda As Double, prior As Double, eps As Variant
    For Each eps In Array(0.0003, LEFT_STRAIN)
        If Not solver.TryDeformationStep(offset, direction, control, CDbl(eps), lambda, prior, 1E-8, 1E-8, 30) Then _
            Err.Raise 5, , "Не получено исходное поле кубической ветви."
        solver.CopyTrial field, resultants
        solver.AcceptStep
        history.Add field
        trace.Add 1#, prior
        previous = lambda: lambda = prior
    Next eps
    TestExpectNear lambda, (30000# * LEFT_STRAIN - 1E+10# * LEFT_STRAIN ^ 3 - 4#) / 15#, 1E-8, _
        "левая принятая нагрузка", passed, failed, report
    If Not solver.TryStep(19# * AREA, 0#, 0#, 0#, 0#, 0#, 1E-8, 1E-8, 30) Then _
        Err.Raise 5, , "Силовая проба с исходного поля не подтверждена."
    solver.CopyTrial field, trialResults
    TestExpectNear field(1), DESIGN_STRAIN, 1E-11, "точная сила на грубом плече", passed, failed, report
    solver.DiscardStep

    Dim branch As New CCoupledDeformationContinuation, peak As Double, accuracy As Double, designVerified As Boolean
    branch.Trace solver, Nothing, offset, direction, history, previous, lambda, resultants, 1.4, 0.0125, 30, 1E-8, 1E-8, trace
    If Not branch.CopyPeak(peak, accuracy, designVerified) Then _
        Err.Raise 5, , "Не подтверждён максимум: " & branch.ResultMeta.DiagnosticDetails
    TestExpectNear peak, (30000# * 0.000975 - 1E+10# * 0.000975 ^ 3 - 4#) / 15#, 1E-8, _
        "сохранён проверенный локальный максимум", passed, failed, report
    TestRecord peak - accuracy > 1# And accuracy <= 0.0125, _
        "максимум отделён от единицы по погрешности", passed, failed, report
    TestRecord designVerified = Not blocksRefined, _
        "подтверждение сочетания относится к уточнённой истории, отказ=" & CStr(blocksRefined), passed, failed, report

    Dim capacity As New CCapacityResult, actual() As Double, atPeak As Collection
    If Not branch.CopyLimitState(atPeak, actual) Then Err.Raise 5, , "У максимума отсутствует подтверждённое поле."
    capacity.InitializeCoupledPeak peak, accuracy, designVerified, actual, "Auto", branch.ResultMeta.DiagnosticDetails
    If blocksRefined Then
        TestRecord capacity.Status = "NumFail" And capacity.LambdaCapacity = 0#, _
            "пик выше единицы без равновесия сочетания даёт NumFail", passed, failed, report
    Else
        TestRecord capacity.Status = "OK" And capacity.LambdaCapacity = peak, _
            "подтверждённое сочетание и больший пик дают OK", passed, failed, report
    End If
    Dim i As Long, priorIndex As Long
    For i = 1 To atPeak.Count
        field = atPeak.Item(i)
        If field(1) < DESIGN_STRAIN Then priorIndex = i
    Next i
    If priorIndex < 1 Then Err.Raise 5, , "Не найден левый уточнённый узел."
    field = atPeak.Item(priorIndex)
    TestExpectNear field(1), 0.00080625, 1E-12, "выбранный уточнённый левый узел", passed, failed, report
    If Not solver.ReplayAcceptedFields(atPeak, priorIndex) Then Err.Raise 5, , "Не восстановлено уточнённое плечо."
    Dim exact As Boolean: exact = solver.TryStep(19# * AREA, 0#, 0#, 0#, 0#, 0#, 1E-8, 1E-8, 30)
    TestRecord exact = Not blocksRefined, _
        "независимая точная проба с сохранённого левого узла, отказ=" & CStr(blocksRefined), passed, failed, report
    solver.DiscardStep
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет точное силовое равновесие в неизменившемся префиксе истории.
Private Sub TestDesignWitnessInPrefix(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Const AREA As Double = 21600#, DESIGN_FORCE As Double = 19# * AREA
    Dim operators As New CCoupledElasticState, material As New CTestCoupledPointResponse, solver As New CCoupledNewtonSolver
    operators.PrepareModel TestBuildSection(TestFixtureBox(180#, 120#)), 30000#, 0#, 1, 1, 5
    material.TestInitialize operators.PointCount
    material.TestUseDescendingAxialLaw
    material.TestBlockRefinedDesign
    solver.Initialize operators, material, 30000#
    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    direction(1) = DESIGN_FORCE
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    Dim history As New Collection, trace As New CCoupledLoadTrace
    Dim field() As Double, actual() As Double, lambda As Double
    If Not solver.TryStep(DESIGN_FORCE, 0#, 0#, 0#, 0#, 0#, 1E-8, 1E-8, 30) Then _
        Err.Raise 5, , "Не подтверждена точная исходная сила."
    solver.CopyTrial field, actual
    solver.AcceptStep
    history.Add field: trace.Add 1#, 1#
    If Not solver.TryDeformationStep(offset, direction, control, 0.0009, 1#, lambda, 1E-8, 1E-8, 30) Then _
        Err.Raise 5, , "Не подтверждён второй узел префикса."
    solver.CopyTrial field, actual
    solver.AcceptStep
    history.Add field: trace.Add 1#, lambda

    Dim branch As New CCoupledDeformationContinuation, peak As Double, accuracy As Double, designVerified As Boolean
    branch.Trace solver, Nothing, offset, direction, history, 1#, lambda, actual, 1.4, 0.005, 30, 1E-8, 1E-8, trace
    TestRecord branch.CopyPeak(peak, accuracy, designVerified) And designVerified, _
        "точное сочетание из сохранённого префикса подтверждено", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет нижнюю границу с сохранёнными равновесиями после замены грубого плеча.
Public Function RunTestCoupledDiscardedShoulder() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const STEEL_AREA As Double = 100#, DESIGN_FORCE As Double = 410000#
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, Sqr(4# * STEEL_AREA / GEOM_PI), STEEL_AREA, "A400"
    Dim steel As New CCoupledSteelModel, models As Object
    steel.Initialize 200000#, 40#, 40#, 0#, 0#, 0.02, 0.02
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "A400", steel
    Dim operators As New CCoupledElasticState, concrete As New CTestCoupledPointResponse
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModels:=models
    concrete.TestInitialize operators.PointCount
    concrete.TestUseDiscardedShoulderLaw
    Dim material As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    material.Initialize operators, concrete, models
    solver.Initialize operators, material, 30000#

    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    direction(1) = DESIGN_FORCE
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    Dim history As New Collection, trace As New CCoupledLoadTrace, field() As Double, resultants() As Double
    Dim coordinate As Variant, factor As Double, lambda As Double, previous As Double
    For Each coordinate In Array(0.0004, 0.0006)
        If Not solver.TryDeformationStep(offset, direction, control, CDbl(coordinate), lambda, factor, 1E-8, 1E-8, 30) Then _
            Err.Raise 5, , "Не получено исходное равновесие. " & solver.ResultMeta.DiagnosticDetails
        solver.CopyTrial field, resultants
        solver.AcceptStep
        history.Add field
        trace.Add 1#, factor
        previous = lambda: lambda = factor
    Next coordinate
    TestExpectNear previous, 0.75, 1E-8, "грубое левое поле", passed, failed, report
    TestExpectNear lambda, 1.2, 1E-8, "грубый максимум до уточнения", passed, failed, report

    Dim branch As New CCoupledDeformationContinuation
    branch.Trace solver, material, offset, direction, history, previous, lambda, resultants, 1.5, 0.0001, 30, 1E-8, 1E-8, trace
    Dim saved As CCoupledLoadTrace: Set saved = branch.SavedTrace
    Dim greatest As Double
    greatest = TestTwinPeakRetainedTrace(branch, solver, passed, failed, report)
    TestRecord greatest < 1.2, "отменённый грубый множитель 1.2 не является сохранённой границей", passed, failed, report

    Dim meta As CResultMeta: Set meta = branch.ResultMeta
    Dim capacity As New CCapacityResult
    capacity.InitializeCoupledUnresolved meta, "Auto", branch.BestAdmissibleLambda
    TestRecord capacity.HasConfirmedLowerBound And Not capacity.HasUltimatePoint And _
        capacity.ResultMeta.InternalStatus = rsNumericalFailure, _
        "сохраняется численный исход без физического предела", passed, failed, report
    TestExpectNear capacity.ConfirmedLowerLambda, greatest, 1E-8, _
        "публикуемая нижняя граница взята из принятой ветви", passed, failed, report
    solver.CopyAccepted field
    TestExpectNear concrete.TestAcceptedPeak, field(1), 1E-12, "материальная история совпадает с полем", passed, failed, report
    TestRecord concrete.TestAcceptedSteps = saved.Count, "история и след имеют одинаковое число шагов", passed, failed, report
    RunTestCoupledDiscardedShoulder = report & "TOTAL_COUPLED_DISCARDED_SHOULDER: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledDiscardedShoulder = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подтверждает равновесное поле при отказе касательной в одной из боковых историй.
' Проверка использует шесть интегралов, меру деформации и данные трёх принятых буферов.
Public Function RunTestCoupledTangentProof() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#, STRAIN As Double = 0.0004
    Const FORCE As Double = AREA * (30000# * STRAIN + 1E+10# * STRAIN ^ 3)
    Const ABS_TOL As Double = 1E-10, REL_TOL As Double = 1E-9
    Const TARGET_LAMBDA As Double = 0.6
    Dim operators As New CCoupledElasticState, solver As New CCoupledNewtonSolver
    Dim center As New CTestCoupledPointResponse, plusSide As New CTestCoupledPointResponse, minusSide As New CTestCoupledPointResponse
    operators.PrepareModel TestBuildSection(TestFixtureBox(180#, 120#)), 30000#, 0#, 1, 1, 5
    center.TestInitialize operators.PointCount
    plusSide.TestInitialize operators.PointCount
    minusSide.TestInitialize operators.PointCount
    solver.Initialize operators, center, 30000#, plusSide, minusSide
    If Not solver.TryStep(FORCE, 0#, 0#, 0#, 0#, 0#, ABS_TOL, REL_TOL, 30) Then _
        Err.Raise 5, , "Не получено исходное равновесие. " & solver.ResultMeta.DiagnosticDetails
    solver.AcceptStep
    Dim original() As Double, peaks(1 To 3) As Double
    solver.CopyAccepted original
    peaks(1) = center.TestAcceptedPeak: peaks(2) = plusSide.TestAcceptedPeak: peaks(3) = minusSide.TestAcceptedPeak

    Dim target(1 To 6) As Double, baseLoad(1 To 6) As Double, direction(1 To 6) As Double, projection() As Double
    target(1) = FORCE: baseLoad(1) = (1# - TARGET_LAMBDA) * FORCE: direction(1) = FORCE
    ReDim projection(1 To operators.DegreeCount)
    projection(1) = 1#
    plusSide.TestTangentFailure = sfcNumericalFailure
    Dim ok As Boolean, before As Long, trialLambda As Double
    before = plusSide.TestStressOnlyProbes
    ok = solver.TryStep(FORCE, 0#, 0#, 0#, 0#, 0#, ABS_TOL, REL_TOL, 30)
    TestRecord TestTangentProofWrench(ok, solver, operators, original, target), _
        "силовой повтор подтверждает исходное поле и шесть усилий", passed, failed, report
    TestRecord plusSide.TestStressOnlyProbes > before, "силовой повтор проверяет отклик без касательной", passed, failed, report
    solver.DiscardStep
    TestRecord Not TestTangentProofHasCandidate(solver), "отклонённая силовая проба не доступна к принятию", passed, failed, report
    TestRecord TestTangentProofHistory(solver, center, plusSide, minusSide, original, peaks), _
        "после силовой пробы неизменны три принятые истории", passed, failed, report

    before = plusSide.TestStressOnlyProbes
    ok = solver.TryDeformationStep(baseLoad, direction, projection, original(1), TARGET_LAMBDA, trialLambda, _
        ABS_TOL, REL_TOL, 30, consistentControl:=True)
    TestRecord TestTangentProofWrench(ok, solver, operators, original, target) And _
        Abs(trialLambda - TARGET_LAMBDA) < 1E-12, "деформационный повтор подтверждает ту же меру, множитель и шесть усилий", passed, failed, report
    TestRecord plusSide.TestStressOnlyProbes > before, "деформационный повтор проверяет отклик без касательной", passed, failed, report
    solver.DiscardStep
    TestRecord Not TestTangentProofHasCandidate(solver), "отклонённая деформационная проба не доступна к принятию", passed, failed, report
    TestRecord TestTangentProofHistory(solver, center, plusSide, minusSide, original, peaks), _
        "после деформационной пробы неизменны три принятые истории", passed, failed, report

    ok = solver.TryStep(1.1 * FORCE, 0#, 0#, 0#, 0#, 0#, ABS_TOL, REL_TOL, 30)
    TestRecord Not ok And solver.ResultMeta.InternalStatus = rsNumericalFailure And Not TestTangentProofHasCandidate(solver), _
        "поле при другой продольной силе не принимается", passed, failed, report
    ok = solver.TryDeformationStep(baseLoad, direction, projection, original(1) + 0.00005, TARGET_LAMBDA, trialLambda, _
        ABS_TOL, REL_TOL, 30, consistentControl:=True)
    TestRecord Not ok And solver.ResultMeta.InternalStatus = rsNumericalFailure And Not TestTangentProofHasCandidate(solver), _
        "поле при другой управляемой деформации не принимается", passed, failed, report
    TestRecord TestTangentProofHistory(solver, center, plusSide, minusSide, original, peaks), _
        "две отклонённые пробы не изменяют принятые буферы", passed, failed, report

    plusSide.TestTangentFailure = sfcInvalidInput
    before = plusSide.TestStressOnlyProbes
    ok = solver.TryStep(FORCE, 0#, 0#, 0#, 0#, 0#, ABS_TOL, REL_TOL, 30)
    Dim policy As New CResultStatusPolicy
    TestRecord Not ok And policy.ExternalStatus(solver.ResultMeta) = "InputErr" And Not TestTangentProofHasCandidate(solver), _
        "прерывающая ошибка ввода сохраняет статус InputErr", passed, failed, report
    TestRecord plusSide.TestStressOnlyProbes = before, "ошибка ввода не запускает повтор без касательной", passed, failed, report
    TestRecord TestTangentProofHistory(solver, center, plusSide, minusSide, original, peaks), _
        "после ошибки ввода неизменны три принятые истории", passed, failed, report

    plusSide.TestTangentFailure = sfcNone
    If Not solver.TryStep(1.1 * FORCE, 0#, 0#, 0#, 0#, 0#, ABS_TOL, REL_TOL, 30) Then _
        Err.Raise 5, , "Нормальный расчёт после отката не восстановился. " & solver.ResultMeta.DiagnosticDetails
    solver.AcceptStep
    TestRecord center.TestAcceptedSteps = 2 And plusSide.TestAcceptedSteps = 2 And minusSide.TestAcceptedSteps = 2, _
        "следующая допустимая проба закрепляет три истории по одному разу", passed, failed, report
    RunTestCoupledTangentProof = report & "TOTAL_COUPLED_TANGENT_PROOF: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledTangentProof = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет подтверждённый винт в масштабе МПа и совпадение всех координат.
Private Function TestTangentProofWrench(ByVal solved As Boolean, ByVal solver As CCoupledNewtonSolver, _
        ByVal operators As CCoupledElasticState, ByRef original() As Double, ByRef target() As Double) As Boolean
    If Not solved Then Exit Function
    If solver.ResultMeta.InternalStatus <> rsSuccess Then Exit Function
    Dim field() As Double, actual() As Double, i As Long
    solver.CopyTrial field, actual
    For i = LBound(original) To UBound(original)
        If field(i) <> original(i) Then Exit Function
    Next i
    TestTangentProofWrench = (operators.ResultantResidualNorm(actual, target) <= 0.00000002)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет принятое поле, максимум деформации и число фиксаций каждой истории.
Private Function TestTangentProofHistory(ByVal solver As CCoupledNewtonSolver, ByVal center As CTestCoupledPointResponse, _
        ByVal plusSide As CTestCoupledPointResponse, ByVal minusSide As CTestCoupledPointResponse, _
        ByRef original() As Double, ByRef peaks() As Double) As Boolean
    If center.TestAcceptedSteps <> 1 Or plusSide.TestAcceptedSteps <> 1 Or minusSide.TestAcceptedSteps <> 1 Then Exit Function
    If center.TestAcceptedPeak <> peaks(1) Or plusSide.TestAcceptedPeak <> peaks(2) Or minusSide.TestAcceptedPeak <> peaks(3) Then Exit Function
    Dim field() As Double, i As Long
    solver.CopyAccepted field
    For i = LBound(original) To UBound(original)
        If field(i) <> original(i) Then Exit Function
    Next i
    TestTangentProofHistory = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Определяет доступность пробного поля через штатную проверку готовности.
Private Function TestTangentProofHasCandidate(ByVal solver As CCoupledNewtonSolver) As Boolean
    Dim field() As Double, actual() As Double
    On Error GoTo ReadFailed
    solver.CopyTrial field, actual
    TestTangentProofHasCandidate = True
    Exit Function
ReadFailed:
    Dim number As Long, source As String, description As String
    number = Err.Number: source = Err.Source: description = Err.Description
    On Error GoTo 0
    If number <> 5 Or source <> "CCoupledNewtonSolver" Then Err.Raise number, source, description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сравнивает допуск принятого деформационного шага с нагрузкой найденного состояния.
Public Function RunTestCoupledFinalLoadTolerance() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#, MODULUS As Double = 30000#
    Const CUBIC As Double = 1E+10, STRAIN As Double = 0.0000018
    Const ABS_TOL As Double = 1E-8, REL_TOL As Double = 0.0000005
    Const TARGET_LAMBDA As Double = 1# + CUBIC * STRAIN ^ 2 / MODULUS
    Dim operators As New CCoupledElasticState, solver As New CCoupledNewtonSolver
    Dim material As New CTestCoupledPointResponse
    operators.PrepareModel TestBuildSection(TestFixtureBox(180#, 120#)), MODULUS, 0#, 1, 1, 5
    material.TestInitialize operators.PointCount
    solver.Initialize operators, material, MODULUS
    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    direction(1) = AREA * MODULUS * STRAIN
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    Dim lambda As Double, field() As Double, resultants() As Double, target(1 To 6) As Double
    If Not solver.TryDeformationStep(offset, direction, control, STRAIN, 3#, lambda, ABS_TOL, REL_TOL, 12) Then _
        Err.Raise 5, , "Не найдена точка кубической ветви. " & solver.ResultMeta.DiagnosticDetails
    solver.CopyTrial field, resultants
    target(1) = direction(1) * lambda
    Dim actualTolerance As Double: actualTolerance = ABS_TOL + REL_TOL * Abs(target(1) / AREA)
    Dim residual As Double: residual = operators.ResultantResidualNorm(resultants, target)
    TestRecord Abs(field(1) - STRAIN) * MODULUS <= actualTolerance, _
        "заданная деформация подтверждена", passed, failed, report
    TestRecord residual <= actualTolerance, "остаток в допуске найденной нагрузки " & CStr(residual) & _
        " <= " & CStr(actualTolerance), passed, failed, report
    TestExpectNear lambda, TARGET_LAMBDA, 1E-9, "множитель по независимому кубическому закону", passed, failed, report
    TestRecord material.TestAcceptedSteps = 0, "до принятия история остаётся исходной", passed, failed, report
    solver.DiscardStep
    RunTestCoupledFinalLoadTolerance = report & "TOTAL_COUPLED_FINAL_LOAD_TOLERANCE: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledFinalLoadTolerance = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет продолжение после локального пика, где малый шаг попадает
' в допуск уравнения управления без изменения механического поля.
Public Function RunTestCoupledControlProgress() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const STEEL_AREA As Double = 100#, DESIGN_FORCE As Double = 410000#
    Const ABS_TOL As Double = 0.03, REL_TOL As Double = 1E-8
    Const TAIL_STRAIN As Double = 0.0007
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, Sqr(4# * STEEL_AREA / GEOM_PI), STEEL_AREA, "A400"
    Dim steel As New CCoupledSteelModel, models As Object
    steel.Initialize 200000#, 40#, 40#, 0#, 0#, 0.02, 0.02
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "A400", steel
    Dim operators As New CCoupledElasticState, concrete As New CTestCoupledPointResponse
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModels:=models
    concrete.TestInitialize operators.PointCount
    concrete.TestUseTwinPeakAxialLaw
    concrete.TestFreezeAxialTailAt = TAIL_STRAIN - 0.00000001
    Dim material As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    material.Initialize operators, concrete, models
    solver.Initialize operators, material, 30000#

    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    direction(1) = DESIGN_FORCE
    ReDim control(1 To operators.DegreeCount)
    control(1) = 1#
    Dim history As New Collection, trace As New CCoupledLoadTrace, field() As Double, resultants() As Double
    Dim coordinate As Variant, factor As Double, previous As Double, lambda As Double
    For Each coordinate In Array(0.0004, 0.0006)
        If Not solver.TryDeformationStep(offset, direction, control, CDbl(coordinate), lambda, factor, 1E-8, REL_TOL, 30) Then _
            Err.Raise 5, , "Не получено исходное равновесие. " & solver.ResultMeta.DiagnosticDetails
        solver.CopyTrial field, resultants
        solver.AcceptStep
        history.Add field
        trace.Add 1#, factor
        previous = lambda: lambda = factor
    Next coordinate
    TestExpectNear previous, 0.75, 1E-8, "исходная первая нагрузка", passed, failed, report
    TestExpectNear lambda, 0.8, 1E-8, "исходная вторая нагрузка", passed, failed, report

    Dim branch As New CCoupledDeformationContinuation
    branch.Trace solver, material, offset, direction, history, previous, lambda, resultants, 1.5, 0.0001, 30, ABS_TOL, REL_TOL, trace
    Dim saved As CCoupledLoadTrace: Set saved = branch.SavedTrace
    Dim status As CResultMeta: Set status = branch.ResultMeta
    TestRecord status.InternalStatus = rsNumericalFailure, "предел сечения остаётся неподтверждённым", passed, failed, report
    TestExpectNear branch.BestAdmissibleLambda, 1.1, 0.0005, "подтверждён пройденный локальный максимум", passed, failed, report
    TestRecord saved.Count = 5, "стационарные повторы не добавлены к пяти принятым состояниям", passed, failed, report
    TestRecord concrete.TestAcceptedSteps = 5, "стационарные повторы не записаны в материал", passed, failed, report
    Dim i As Long, offsetFraction As Double, expected As Variant, expectedCoordinate As Variant
    Dim firstRepetition As Long
    expected = Array(0.75, 0.82, 0.8, 1.1, 0.7)
    expectedCoordinate = Array(0.0004, 0.0005, 0.0006, 0.00065, 0.0007)
    For i = 1 To GeomMin(saved.Count, 5)
        saved.ReadPoint i, offsetFraction, lambda
        TestExpectNear lambda, CDbl(expected(i - 1)), 0.0005, "принятая нагрузка " & CStr(i), passed, failed, report
        TestExpectNear concrete.TestAcceptedAxialAt(i), CDbl(expectedCoordinate(i - 1)), 0.000000001, _
            "фактическая координата принятого поля " & CStr(i), passed, failed, report
    Next i
    For i = 2 To concrete.TestAcceptedSteps
        If concrete.TestAcceptedAxialAt(i) <= concrete.TestAcceptedAxialAt(i - 1) Then
            firstRepetition = i
            Exit For
        End If
    Next i
    TestRecord firstRepetition = 0, "фактические координаты всех закреплённых шагов строго возрастают", passed, failed, report
    solver.CopyAccepted field
    TestExpectNear field(1), TAIL_STRAIN, 0.000000001, "оставлено последнее фактически достигнутое поле", passed, failed, report

    Dim trial() As Double, recovered As Boolean
    recovered = solver.TryDeformationStep(offset, direction, control, field(1) + 0.00000078125, lambda, factor, _
        ABS_TOL, REL_TOL, 30)
    TestRecord recovered, "решатель допускает остаток управления внутри заданного допуска", passed, failed, report
    If recovered Then
        solver.CopyTrial trial, resultants
        TestRecord trial(1) = field(1) And factor = lambda, _
            "проба внутри допуска остаётся тем же полем и множителем", passed, failed, report
        solver.DiscardStep
    Else
        TestRecord False, "не получено исходное стационарное поле", passed, failed, report
    End If
    RunTestCoupledControlProgress = report & "TOTAL_COUPLED_CONTROL_PROGRESS: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledControlProgress = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет решение около единицы с точным пределом осевой ветви.
' Слабое кручение сохраняет пространственный путь без влияния на осевую диаграмму.
Public Function RunTestCoupledSteelAtDesign() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const AREA As Double = 21600#, BAR_AREA As Double = 1000#, STEEL_LIMIT As Double = 0.001
    Const LIMIT_FORCE As Double = (AREA - BAR_AREA) * (30000# * STEEL_LIMIT + 1E+10 * STEEL_LIMIT ^ 3) + _
        BAR_AREA * (100# + 2000# * (STEEL_LIMIT - 100# / 200000#))
    Const ABS_TOL As Double = 0.0000005, REL_TOL As Double = 0.0000005
    Dim section As CSectionModel, operators As New CCoupledElasticState
    Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, 12#, BAR_AREA, "TestGrade"
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModulus:=200000#
    Dim concrete As New CTestCoupledPointResponse, steel As New CCoupledSteelModel, models As Object
    concrete.TestInitialize operators.PointCount
    steel.Initialize 200000#, 100#, 100#, 2000#, 2000#, STEEL_LIMIT, STEEL_LIMIT
    Set models = CreateObject("Scripting.Dictionary")
    models.Add "TestGrade", steel
    Dim response As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    response.Initialize operators, concrete, models
    solver.Initialize operators, response, 30000#

    Dim path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    Dim designForce As Double, expected As Variant, initialFactor As Variant
    Dim low As Double, high As Double, onOffset As Boolean, eventFound As Boolean
    Dim numericalRatio As Double, meta As CResultMeta, field() As Double, resultants() As Double
    For Each expected In Array(0.98, 0.999999, 1#, 1.00000002, 1.02)
        designForce = LIMIT_FORCE / CDbl(expected)
        solver.ResetHistory
        If Not solver.TryStep(designForce, 0#, 0#, 300000#, 0#, 0#, ABS_TOL, REL_TOL, 60) Then _
            Err.Raise 5, , "Не получено контрольное прямое равновесие."
        solver.CopyTrial field, resultants
        numericalRatio = response.TrialSteelUtilization
        solver.DiscardStep
        path.Initialize 0#, designForce, 0#, 0#, 0#, 0#, 0#, 300000#, 0#, 0#, 0#, 0#
        For Each initialFactor In Array(1#, 0.5)
            branch.Initialize solver, path, 4, 64, 60, ABS_TOL, REL_TOL, True
            branch.TraceSteelLimit response, CDbl(initialFactor), 1#, 0.003, preferDeformationOnStall:=True
            low = 0#: high = 0#
            eventFound = branch.CopySteelLimit(low, high, onOffset)
            Set meta = branch.ResultMeta
            report = report & "STEEL_NEAR: exact=" & CStr(expected) & "; initial=" & CStr(initialFactor) & _
                "; directU-1=" & CStr(numericalRatio - 1#) & "; event=" & CStr(eventFound) & "; lo=" & CStr(low) & _
                "; hi=" & CStr(high) & "; accepted=" & CStr(branch.AcceptedLambda) & "; code=" & CStr(meta.ResultCode) & vbCrLf
            report = report & "STEEL_NEAR_DETAIL: " & meta.DiagnosticDetails & vbCrLf
            If expected >= 1# Then
                TestRecord meta.InternalStatus <> rsCheckFailed, "допустимая ветвь не имеет подтверждённого предела ниже расчётной нагрузки", _
                    passed, failed, report
            ElseIf expected < 0.99 Then
                TestRecord eventFound And low <= expected And high >= expected And high < 1# And meta.InternalStatus = rsCheckFailed, _
                    "определённый арматурный предел ниже расчётной нагрузки", passed, failed, report
            Else
                TestRecord Not eventFound Or (low <= expected And high >= expected), _
                    "при близком физическом пределе найденный интервал содержит аналитическое значение", passed, failed, report
            End If
        Next initialFactor
    Next expected
    RunTestCoupledSteelAtDesign = report & "TOTAL_COUPLED_STEEL_AT_DESIGN: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledSteelAtDesign = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
