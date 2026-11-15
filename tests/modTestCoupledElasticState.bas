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
    branch.Initialize solver, path, 1, 100, 20, 1E-8, 1E-8
    Dim field() As Double, resultants() As Double, expected() As Double, ok As Boolean
    ok = branch.Advance(0#)
    TestRecord ok, "constant part reached " & branch.ResultMeta.DiagnosticDetails, passed, failed, report
    If Not ok Then GoTo Done
    branch.CopyAccepted field, resultants
    TestExpectNear resultants(1), FORCE0, 1E-7, "preload resultant", passed, failed, report
    TestRecord branch.AcceptedOffsetFraction = 1# And branch.AcceptedLambda = 0# And branch.RejectedAttempts > 0, _
        "preload reduced an unsafe increment and reached its exact target", passed, failed, report

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
    material.TestFail = True
    ok = branch.Advance(1.1)
    TestRecord Not ok And branch.ResultMeta.InternalStatus = rsNumericalFailure And branch.AcceptedLambda = 1#, _
        "minimum step failure stays numerical at last accepted lambda", passed, failed, report
    branch.CopyAccepted field, resultants
    TestRecord firstDegree = field(1) And acceptedBefore = material.TestAcceptedSteps, _
        "failed increments leave accepted section and history unchanged", passed, failed, report
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
Done:
    RunTestCoupledContinuation = report & "TOTAL_COUPLED_CONTINUATION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledContinuation = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function


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
    response.AcceptTrial
    TestExpectNear material.AcceptedSteelUtilization, 0.004 / 0.025, 1E-10, "accepted utilization", passed, failed, report

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
    Dim multiplier As Double, value As Double, trace As Double, q As Double, high As Double, mid As Double, low As Double
    Dim i As Long, j As Long, caseNo As Long, dilation As Double
    Dim beta As Double: beta = FC / FT * (1# - ALPHA) - (1# + ALPHA)
    Dim ok As Boolean
    elastic(3) = 1#
    ok = TestConcreteReturn(update, elastic, FC, FT, 0#, corrected, increment, multiplier)
    TestRecord ok And corrected(3) = 1# And multiplier = 0# And increment(3) = 0#, _
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
        ok = TestConcreteReturn(update, elastic, FC, FT, dilation, corrected, increment, multiplier)
        TestRecord ok And multiplier > 0#, "return " & CStr(caseNo), passed, failed, report
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
    ok = TestConcreteReturn(update, elastic, FC, FT, 0#, corrected, increment, multiplier)
    TestRecord Not ok, "hydrostatic tension without volumetric flow reports inability to return", passed, failed, report
    For i = 1 To 3: elastic(i) = 4.5: Next i
    elastic(6) = 0.025
    Dim smallDilation As Variant, apexStress As Double, apexVolume As Double
    apexStress = FC * (1# - ALPHA) / (3# * ALPHA + beta)
    apexVolume = (elastic(1) - apexStress) / (E / (1# - 2# * NU))
    For Each smallDilation In Array(0.000001, 1E-170)
        dilation = CDbl(smallDilation)
        ok = TestConcreteReturn(update, elastic, FC, FT, dilation, corrected, increment, multiplier)
        TestRecord ok And multiplier > 0#, "near-apex return for dilation " & CStr(dilation), passed, failed, report
        If ok Then
            TestExpectNear corrected(1), apexStress, 1E-9, "near-apex analytical pressure", passed, failed, report
            TestExpectNear multiplier * dilation, apexVolume, 1E-12, "near-apex volumetric flow", passed, failed, report
            TestExpectNear tensor(6, 6) * increment(6) + corrected(6), elastic(6), 1E-12, _
                "near-apex engineering shear balance", passed, failed, report
        End If
    Next smallDilation
    ok = TestConcreteReturn(update, elastic, FC, FT, 0#, corrected, increment, multiplier)
    TestRecord Not ok, "near-apex deviator does not enable volumetric flow at zero dilation", passed, failed, report
    RunTestCoupledConcreteReturn = report & "TOTAL_COUPLED_CONCRETE_RETURN: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledConcreteReturn = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит известный тензор упругой пробы для изолированной проверки возврата.
Private Function TestConcreteReturn(ByVal update As CCoupledConcreteReturn, ByRef elastic() As Double, _
        ByVal fc As Double, ByVal ft As Double, ByVal dilation As Double, ByRef corrected() As Double, _
        ByRef increment() As Double, ByRef multiplier As Double) As Boolean
    Dim trace As Double, q As Double, maximum As Double, middle As Double, minimum As Double, deviatorFactor As Double
    CoupledStressInvariants elastic, trace, q, maximum, middle, minimum
    TestConcreteReturn = update.Correct(elastic, trace, q, maximum, fc, ft, ft, dilation, _
        corrected, increment, multiplier, deviatorFactor)
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
    response.AcceptTrial
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
    TestRecord capacity.HasUltimatePoint And capacity.IsTransverseLimit And capacity.Status = "FAIL" And InStr(capacity.LimitState, "TestWeakHoop") > 0, _
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
        batch.ApplyLoadReference 2#, -1#, 2#, -1#
        batch.Execute
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
        TestRecord anchor.Offset(-4, 50).Value2 = transverseHeading And anchor.Offset(-4, 50).MergeArea.Columns.Count = 7 And _
            anchor.Offset(-4, 50).Resize(5, 7).Borders.Item(10).LineStyle <> -4142, _
            "separate Results heading spans the fully framed hoop subsection", passed, failed, report
        TestRecord anchor.Offset(0, 53).Value2 = "FAIL" And anchor.Offset(0, 56).Value2 = "FAIL", _
            "separate current and reserve statuses are written", passed, failed, report
        TestRecord anchor.Offset(0, 53).Interior.Color = StatusFillColor("FAIL") And _
            anchor.Offset(0, 56).Interior.Color = StatusFillColor("FAIL"), "both hoop statuses use the standard failure fill", passed, failed, report
        TestRecord InStr(anchor.Offset(0, 50).Value2, "TestWeakHoop (TestHoopGrade)") = 1 And _
            InStr(anchor.Offset(0, 55).Value2, "TestWeakHoop (TestHoopGrade)") = 1, _
            "Results starts with the governing hoop and its grade in each check", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 54).Value2), lambda, 0.01, "reserve is written in the hoop subsection", passed, failed, report
        TestSavedHoopCapacityFields ThisWorkbook, section, provider, batch, passed, failed, report
    End If
    RunTestCoupledHoopCapacity = report & "TOTAL_COUPLED_HOOP_CAPACITY: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledHoopCapacity = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет подраздел хомута при выводе на лист и восстановлении из файла,
' а также предельные поля обоих носителей по ElementID и ключу состояния.
Private Sub TestSavedHoopCapacityFields(ByVal book As Object, ByVal section As CSectionModel, _
        ByVal provider As CMaterialModelProvider, ByVal batch As CBatchSectionCalculator, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim settings As Object: Set settings = book.Names.Item("rngSystemSettings").RefersToRange
    TestInputCell(settings, "Results.SnapshotName").Value2 = "TestSavedHoopCapacityFields"
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
    Dim reserve As Double: reserve = CDbl(anchor.Offset(0, 54).Value2)
    Dim mode As Variant, values As Variant, n As Long, shear As Object, runId As Double
    n = section.ConcreteCount + section.RebarCount
    For Each mode In Array("Excel", "File")
        TestInputCell(settings, "Results.SnapshotStorage").Value2 = CStr(mode)
        anchor.Offset(0, 50).Resize(1, 7).ClearContents ' Проверяем вывод заново, без значений от другого носителя.
        If mode = "Excel" Then
            UpdateSectionResultsForWorkbook book, snap
        Else
            UpdateSectionResultsForWorkbook book
        End If
        TestRecord anchor.Offset(0, 53).Value2 = "FAIL" And anchor.Offset(0, 56).Value2 = "FAIL", _
            CStr(mode) & " restores both independent hoop statuses on Results", passed, failed, report
        TestExpectNear CDbl(anchor.Offset(0, 54).Value2), reserve, 1E-10, _
            CStr(mode) & " restores the hoop reserve on Results", passed, failed, report
        TestRecord InStr(CStr(anchor.Offset(0, 50).Value2), "TestWeakHoop (TestHoopGrade)") = 1 And _
            InStr(CStr(anchor.Offset(0, 55).Value2), "TestWeakHoop (TestHoopGrade)") = 1, _
            CStr(mode) & " restores the hoop and its grade in both checks", passed, failed, report
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
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет действующие проверки СП в плоском и пространственном сочетаниях
' при одинаковых N, Mx, My. В профиле отсутствует запрос связанного НДС.
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
    TestInputCell(settings, "Load.ReferenceOffsetX").Value2 = 0#
    TestInputCell(settings, "Load.ReferenceOffsetY").Value2 = 0#
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
    Dim expected As Variant
    expected = Array(units.InternalForceToOutput(units.InputForceToInternal(firstCase(1, 2))), _
        units.InternalMomentMxToOutput(units.InputMomentMxToInternal(firstCase(1, 3))), _
        units.InternalMomentMyToOutput(units.InputMomentMyToInternal(firstCase(1, 4))), _
        units.InternalTorqueToOutput(units.InputTorqueToInternal(firstCase(1, 8))), _
        units.InternalShearToOutput(units.InputShearToInternal(firstCase(1, 9))), _
        units.InternalShearToOutput(units.InputShearToInternal(firstCase(1, 10))))

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
        gt.ClearContents
        ignored = RunSectionCalculationForWorkbook(book)
        TestRecord ReadLoadCasePropertyText(book, "Spatial", "OverallStatus") = "InputErr", CStr(mode) & " missing Gt is an input error", passed, failed, report
        TestRecord InStr(1, ignored, gt.Worksheet.Name & "!" & gt.Address(False, False), vbTextCompare) > 0, CStr(mode) & " missing Gt names its cell", passed, failed, report
        TestRecord ReadLoadCasePropertyText(book, "Plane", "DirectStateStatus") = "OK", CStr(mode) & " planar case unaffected by calibration", passed, failed, report

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
        Dim hoopAnchor As Object: Set hoopAnchor = book.Names.Item("rngStrengthSummaryAnchor").RefersToRange
        TestRecord hoopAnchor.Offset(-4, 50).Value2 = "Поперечная арматура — хомуты и ветви", _
            CStr(mode) & " has separate hoop heading", passed, failed, report
        TestRecord hoopAnchor.Offset(0, 53).Value2 = "OK" And Len(hoopAnchor.Offset(1, 53).Value2) = 0, _
            CStr(mode) & " writes hoop status only for spatial case", passed, failed, report
        If CStr(mode) = "File" Then
            Dim savedStrength As Variant, reader As CResultReadSource
            Set reader = CurrentResultSource(book)
            savedStrength = reader.Table("Strength.Data")
            TestRecord UBound(savedStrength, 2) = STRENGTH_SNAPSHOT_COLUMNS And savedStrength(1, 54) = "OK", _
                "File stores separate hoop status and schema", passed, failed, report
            hoopAnchor.Offset(0, 53).ClearContents
            UpdateSectionResultsForWorkbook book
            TestRecord hoopAnchor.Offset(0, 53).Value2 = "OK", "File restores a removed hoop status", passed, failed, report
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
        TestSavedCoupledFields book, CStr(mode), section.ConcreteCount + section.RebarCount, passed, failed, report
        TestSavedCoupledMaterialCurves book, CStr(mode), grades.Cells(2, 3), units, passed, failed, report
    Next mode
    TestSavedAutoCADTransverse book, settings, units, passed, failed, report
    TestSavedCurvedTransversePaths book, units, passed, failed, report
    RunTestCoupledWorkbookInput = report & "TOTAL_COUPLED_WORKBOOK_INPUT: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledWorkbookInput = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

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
        TestRecord hoopSummary.Value2 = "Spatial" And hoopSummary.Offset(0, 53).Value2 = "N/A" And _
            hoopSummary.Offset(0, 56).Value2 = "N/A", label & "reports the hoop checks as inapplicable for explicit None", passed, failed, report
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
    Dim legacy As New CCalculationSnapshot, damaged As New CCalculationSnapshot, key As Variant, values As Variant, cropped As Variant
    legacy.RunID = current.RunID: damaged.RunID = current.RunID
    For Each key In current.Keys
        values = current.Table(CStr(key))
        If key = "Strength.Headers" Or key = "Strength.Data" Then
            cropped = TestCropStrengthTable(values)
            legacy.AddTable CStr(key), cropped
            If key = "Strength.Headers" Then
                damaged.AddTable CStr(key), cropped
            Else
                damaged.AddTable CStr(key), values
            End If
        Else
            legacy.AddTable CStr(key), values: damaged.AddTable CStr(key), values
        End If
    Next key
    Dim errText As String
    On Error Resume Next
    UpdateSectionResultsForWorkbook book, damaged
    errText = Err.Description
    Err.Clear
    On Error GoTo 0
    TestRecord InStr(errText, "число столбцов") > 0 And anchor.Offset(0, 53).Value2 = "OK", _
        "mismatched snapshot widths are rejected before Results changes", passed, failed, report
    UpdateSectionResultsForWorkbook book, legacy
    TestRecord anchor.Value2 = "Spatial" And Len(anchor.Offset(0, 53).Value2) = 0, _
        "49-column strength snapshot restores existing cells and leaves the hoop subsection empty", passed, failed, report
    UpdateSectionResultsForWorkbook book
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует сохранённую таблицу в допустимую структуру из 49 столбцов.
Private Function TestCropStrengthTable(ByRef values As Variant) As Variant
    Dim result() As Variant, row As Long, column As Long
    ReDim result(1 To UBound(values, 1), 1 To LEGACY_STRENGTH_SNAPSHOT_COLUMNS)
    For row = 1 To UBound(values, 1)
        For column = 1 To LEGACY_STRENGTH_SNAPSHOT_COLUMNS
            result(row, column) = values(row, column)
        Next column
    Next row
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
    RunTestCoupledWorkbookTiming = report & "TOTAL_COUPLED_WORKBOOK_TIMING: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledWorkbookTiming = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
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
