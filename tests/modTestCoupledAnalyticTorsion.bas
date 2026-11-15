Attribute VB_Name = "modTestCoupledAnalyticTorsion"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Сравнивает чистое упругое кручение с решениями Сен-Венана для
' прямоугольника 2:1 и равностороннего треугольника. Материал однородный.

Private Const E_MODULUS As Double = 32000#
Private Const POISSON As Double = 0.2
Private Const TORQUE_NMM As Double = 2500000#
Private Const FOURIER_TERMS As Long = 64

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет восстановленные усилия, жёсткость и оба поля касательных напряжений.
Public Function RunTestCoupledAnalyticTorsion() As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Const RECTANGLE_WIDTH As Double = 200#, RECTANGLE_HEIGHT As Double = 100#, TRIANGLE_SIDE As Double = 160#
    Dim geometry As CGeometryRegion, spanLevels As Variant, cellLevels As Variant
    Set geometry = TestFixtureBox(RECTANGLE_WIDTH, RECTANGLE_HEIGHT)
    spanLevels = Array(2, 4, 6): cellLevels = Array(10, 14, 16)
    TestTorsionConvergence geometry, False, RECTANGLE_WIDTH, RECTANGLE_HEIGHT, spanLevels, cellLevels, _
        0.0005, 0.01, passed, failed, report
    Dim altitude As Double: altitude = Sqr(3#) * TRIANGLE_SIDE / 2#
    Set geometry = TestTorsionTriangle(TRIANGLE_SIDE, altitude)
    spanLevels = Array(1, 1, 1): cellLevels = Array(12, 24, 48) ' Кубическое поле уже содержится в базисе одного интервала.
    TestTorsionConvergence geometry, True, TRIANGLE_SIDE, altitude, spanLevels, cellLevels, _
        0.003, 0.004, passed, failed, report
    RunTestCoupledAnalyticTorsion = report & "TOTAL_COUPLED_ANALYTIC_TORSION: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledAnalyticTorsion = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_COUPLED_ANALYTIC_TORSION: passed=" & passed & "; failed=" & (failed + 1)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит точный линейный контур с началом в центре тяжести треугольника.
Private Function TestTorsionTriangle(ByVal sideLength As Double, ByVal altitude As Double) As CGeometryRegion
    Dim outline As New CGeometryPath, query As New CSectionGeometryQuery
    Dim baseY As Double, tipY As Double
    baseY = -altitude / 3#: tipY = 2# * altitude / 3#
    outline.Initialize "TriangleTorsion"
    outline.AddLine "Base", -sideLength / 2#, baseY, sideLength / 2#, baseY
    outline.AddLine "Right", sideLength / 2#, baseY, 0#, tipY
    outline.AddLine "Left", 0#, tipY, -sideLength / 2#, baseY
    outline.Seal True
    Set TestTorsionTriangle = TestFixtureGeometry(query.RegionFromPath(outline))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет три разрешения. В прямоугольнике уточняются базис и квадратура,
' в треугольнике меняется только квадратура полиномиального решения.
Private Sub TestTorsionConvergence(ByVal geometry As CGeometryRegion, ByVal triangle As Boolean, _
        ByVal width As Double, ByVal height As Double, ByVal spanLevels As Variant, ByVal cellLevels As Variant, _
        ByVal maximumJError As Double, ByVal maximumFieldError As Double, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim label As String, referenceJ As Double
    If triangle Then
        label = "Triangle": referenceJ = Sqr(3#) * width ^ 4 / 80#
    Else
        label = "Rectangle 2:1": referenceJ = TestRectangleTorsionJ(width, height)
    End If
    Dim mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder, section As CSectionModel
    mesh.BuildMesh geometry, width / 20#, height / 20#, cellGeometry:="Contour"
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    geometry.BuildContours section.Contours
    Dim level As Long, jError As Double, fieldError As Double, previousJ As Double, previousField As Double
    For level = LBound(spanLevels) To UBound(spanLevels)
        Dim spans As Long, cells As Long
        spans = CLng(spanLevels(level)): cells = CLng(cellLevels(level))
        If Not TestSolveTorsionLevel(section, triangle, width, height, spans, cells, referenceJ, _
                label, jError, fieldError, passed, failed, report) Then Exit Sub
        If level > LBound(spanLevels) Then _
            TestRecordAnalyticTorsion (jError < previousJ And fieldError < previousField), _
                label & " convergence " & spans & " x " & cells, passed, failed, report
        previousJ = jError: previousField = fieldError
    Next level
    TestRecordAnalyticTorsion jError < maximumJError, label & " finest J relative error=" & CStr(jError), passed, failed, report
    TestRecordAnalyticTorsion fieldError < maximumFieldError, label & " finest stress relative L2=" & CStr(fieldError), passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Восстанавливает крутящий момент интегрированием τ на материальных точках
' и сравнивает поле при заданном T с независимым решением при θ=T/GJ.
Private Function TestSolveTorsionLevel(ByVal section As CSectionModel, ByVal triangle As Boolean, _
        ByVal width As Double, ByVal height As Double, ByVal spans As Long, ByVal cells As Long, ByVal referenceJ As Double, _
        ByVal label As String, ByRef jError As Double, ByRef fieldError As Double, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String) As Boolean
    Dim solver As New CCoupledElasticState, deformation() As Double, resultants() As Double, stress() As Double
    solver.PrepareModel section, E_MODULUS, POISSON, spans, spans, cells
    Dim solved As Boolean
    solved = solver.Solve(0#, 0#, 0#, TORQUE_NMM, 0#, 0#, deformation, resultants, stress)
    label = label & " | spans=" & spans & " | cells=" & cells
    TestRecordAnalyticTorsion solved, label & " pure torque equilibrium", passed, failed, report
    If Not solved Then
        report = report & "DIAGNOSTIC: " & solver.ResultMeta.ResultComment & vbCrLf
        Exit Function
    End If
    Dim shearModulus As Double, jNumeric As Double, gTwist As Double
    shearModulus = E_MODULUS / (2# * (1# + POISSON))
    jNumeric = TORQUE_NMM / (shearModulus * deformation(4))
    gTwist = TORQUE_NMM / referenceJ ' МПа/мм; закрутка из аналитического J, а не из поля решателя.
    Dim point As Long, x As Double, y As Double, area As Double, tauX As Double, tauY As Double
    Dim torque As Double, integratedQx As Double, integratedQy As Double, errorNorm As Double, referenceNorm As Double
    For point = 1 To UBound(stress, 2)
        solver.Point point, x, y, area
        If triangle Then
            TestTriangleTorsionStress x, y, height, tauX, tauY
        Else
            TestRectangleTorsionStress x, y, width, height, tauX, tauY
        End If
        tauX = tauX * gTwist: tauY = tauY * gTwist
        torque = torque + area * (x * stress(6, point) - y * stress(5, point)) ' Н*мм; знак по правой тройке осей.
        integratedQx = integratedQx + area * stress(5, point)
        integratedQy = integratedQy + area * stress(6, point)
        errorNorm = errorNorm + area * ((stress(5, point) - tauX) ^ 2 + (stress(6, point) - tauY) ^ 2)
        referenceNorm = referenceNorm + area * (tauX ^ 2 + tauY ^ 2)
    Next point
    TestRecordAnalyticTorsion Abs(torque - TORQUE_NMM) < 1E-8 * TORQUE_NMM And _
        (Abs(integratedQx) + Abs(integratedQy)) * width < 1E-8 * TORQUE_NMM And _
        Abs(resultants(1)) * width + Abs(resultants(2)) + Abs(resultants(3)) < 1E-8 * TORQUE_NMM, _
        label & " independent T integral and zero parasitic resultants", passed, failed, report
    jError = Abs(jNumeric / referenceJ - 1#)
    fieldError = Sqr(errorNorm / referenceNorm)
    report = report & "TORSION: " & label & " | J_ref_mm4=" & CStr(referenceJ) & " | J_num_mm4=" & CStr(jNumeric) & _
        " | stress_L2=" & CStr(fieldError) & " | recovered_T_Nmm=" & CStr(torque) & vbCrLf
    TestSolveTorsionLevel = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Интегрирует функцию Прандтля для прямоугольника. J в мм4;
' ширина вдоль X, меньшая сторона вдоль Y, только нечётные гармоники.
Private Function TestRectangleTorsionJ(ByVal width As Double, ByVal height As Double) As Double
    Dim term As Long, harmonic As Long, exponent As Double, correction As Double
    For term = 1 To FOURIER_TERMS
        harmonic = 2 * term - 1
        exponent = Exp(-harmonic * GEOM_PI * width / height)
        correction = correction + (1# - exponent) / ((1# + exponent) * CDbl(harmonic) ^ 5) ' tanh(nπB/2H)/n⁵.
    Next term
    TestRectangleTorsionJ = width * height ^ 3 / 3# - 64# * height ^ 4 / GEOM_PI ^ 5 * correction
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает τxz/(Gθ) и τyz/(Gθ) из производных функции Прандтля.
' Отношения гиперболических функций записаны без растущих экспонент.
Private Sub TestRectangleTorsionStress(ByVal x As Double, ByVal y As Double, ByVal width As Double, _
        ByVal height As Double, ByRef tauX As Double, ByRef tauY As Double)
    Dim term As Long, harmonic As Long, sign As Double, omega As Double, positive As Double, negative As Double
    Dim seriesX As Double, seriesY As Double, coefficient As Double, halfWidth As Double
    sign = 1#: halfWidth = width / 2#
    For term = 1 To FOURIER_TERMS
        harmonic = 2 * term - 1: omega = harmonic * GEOM_PI / height
        positive = Exp(-omega * (halfWidth - x)): negative = Exp(-omega * (halfWidth + x))
        coefficient = sign / ((1# + Exp(-omega * width)) * CDbl(harmonic) ^ 2)
        seriesX = seriesX + coefficient * (positive + negative) * Sin(omega * y)
        seriesY = seriesY + coefficient * (positive - negative) * Cos(omega * y)
        sign = -sign
    Next term
    coefficient = 8# * height / GEOM_PI ^ 2
    tauX = -2# * y + coefficient * seriesX: tauY = coefficient * seriesY
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Для равностороннего треугольника функция Прандтля ψ=2Gθ·d0·d1·d2/h.
' d0, d1, d2 — внутренние нормальные расстояния до сторон, мм.
Private Sub TestTriangleTorsionStress(ByVal x As Double, ByVal y As Double, ByVal height As Double, _
        ByRef tauX As Double, ByRef tauY As Double)
    Dim d0 As Double, d1 As Double, d2 As Double, relativeY As Double, rootThreeX As Double
    d0 = y + height / 3#: relativeY = 2# * height / 3# - y: rootThreeX = Sqr(3#) * x
    d1 = (relativeY - rootThreeX) / 2#: d2 = (relativeY + rootThreeX) / 2#
    tauX = 2# / height * (d1 * d2 - d0 * (d1 + d2) / 2#) ' (∂ψ/∂y)/(Gθ).
    tauY = 3# * x * d0 / height ' (−∂ψ/∂x)/(Gθ).
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет все независимые критерии в формате штатного журнала тестов.
Private Sub TestRecordAnalyticTorsion(ByVal agreed As Boolean, ByVal text As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If agreed Then
        passed = passed + 1: report = report & "PASS: " & text & vbCrLf
    Else
        failed = failed + 1: report = report & "FAIL: " & text & vbCrLf
    End If
End Sub
