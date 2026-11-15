Attribute VB_Name = "modTestCoupledExperiment"
Option Explicit

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставление связанного НДС с опытами Onsongo TBU3 и Kani.
' Геометрия и стали восстановлены по первичным публикациям в literature/Coupled6D.
' Параметры, отсутствующие в источниках, фиксируются как допущения расчёта.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прослеживает чистый Qy с заданием среднего инженерного сдвига yz.
' Опытная кривая сравнивается с принятыми точками без экстраполяции.
Public Function RunTestCoupledKaniShear(ByVal spansText As String, ByVal shearText As String) As String
    RunTestCoupledKaniShear = TraceKaniShearTest(spansText, shearText, False)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет траекторию с оценками энергий из формул (16.3–16.4)
' Poliotti и Bairán (2019). Характеристическую длину сохраняет.
Public Function RunTestCoupledKaniPublishedEnergy(ByVal spansText As String, ByVal shearText As String) As String
    RunTestCoupledKaniPublishedEnergy = TraceKaniShearTest(spansText, shearText, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет точки равновесия и причины уменьшения шага при заданном сдвиге.
Private Function TraceKaniShearTest(ByVal spansText As String, ByVal shearText As String, ByVal publishedEnergies As Boolean) As String
    On Error GoTo Failed
    Const EC As Double = 25600#
    Const FC As Double = 28.2
    Const FT As Double = 1.75
    Const NOMINAL_INCREMENT As Double = 0.0002
    Const MAX_ATTEMPTS As Long = 120
    Const FORCE_SCALE As Double = 100000# ' Множитель равновесия отсчитывается в долях 100 кН.
    Const AREA As Double = 305# * 610#
    Dim report As String, tracePath As String, passed As Long, failed As Long
    Dim spans As Long: spans = CLng(spansText)
    Dim maximumShear As Double: maximumShear = Val(shearText)
    If spans < 1 Or spans > 3 Or maximumShear <= 0# Or maximumShear > 0.008 Then _
        Err.Raise 5, , "Тестовые границы: базис 1–3, инженерный сдвиг от 0 до 0,008."
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) > 0 Then tracePath = ThisWorkbook.Path & "\Kani_Trace.csv"
    AddExperimentTraceTest report, tracePath, "attempt;gamma_yz;Qy_kN;hoop_stress_MPa;hoop_utilization;stress_residual_MPa;seconds"
    Dim gt As Double, gc As Double
    gt = 0.12: gc = 80#
    If publishedEnergies Then
        gt = 0.073 * FC ^ 0.18
        gc = (FC / FT) ^ 2 * gt
    End If
    AddExperimentTraceTest report, tracePath, "KANI_ENERGIES_N_PER_MM: " & ExperimentNumberTest(gt) & " | " & ExperimentNumberTest(gc)
    Dim section As CSectionModel, layout As CTransverseRebarLayout, steel As Object
    Set section = BuildKaniSectionTest(layout, steel)
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, response As New CCoupledReinforcedResponse
    operators.PrepareModel section, EC, 0.2, spans, spans, 5, layout, steelModels:=steel
    concrete.Initialize operators.PointCount, EC, 0.2, 0.6 * FC, FC, FT, 0.0022, gt, gc, 125#
    response.Initialize operators, concrete, steel
    Dim solver As New CCoupledNewtonSolver, shearX() As Double, shearY() As Double
    solver.Initialize operators, response, EC
    operators.MeanShearProjections shearX, shearY
    Dim baseLoad(1 To 6) As Double, direction(1 To 6) As Double
    direction(6) = FORCE_SCALE
    AddExperimentTraceTest report, tracePath, "KANI_POINTS: " & CStr(operators.PointCount) & " | DOF: " & CStr(operators.DegreeCount)

    Dim coordinate As Double, increment As Double, candidate As Double, trialLambda As Double, reached As Double
    Dim field() As Double, actual() As Double, attempt As Long, i As Long, errorNorm As Double, difference As Double
    Dim gamma As Double, expected As Double, hoopPoint As Long, started As Double, stopReason As String
    increment = NOMINAL_INCREMENT: stopReason = "AttemptLimit"
    For attempt = 1 To MAX_ATTEMPTS
        candidate = coordinate + increment
        If candidate > maximumShear Then candidate = maximumShear
        started = Timer
        If solver.TryDeformationStep(baseLoad, direction, shearY, candidate, reached, trialLambda, 0.0000005, 0.00000002, 24) Then
            solver.CopyTrial field, actual: solver.AcceptStep
            reached = trialLambda: coordinate = candidate: errorNorm = 0#: gamma = 0#
            For i = 1 To operators.DegreeCount: gamma = gamma + shearY(i) * field(i): Next i
            For i = 1 To 6
                expected = reached * direction(i)
                difference = Abs(actual(i) - expected) / AREA
                If i >= 2 And i <= 4 Then difference = difference / 610#
                If difference > errorNorm Then errorNorm = difference
            Next i
            hoopPoint = response.AcceptedTransverseCriticalPoint
            AddExperimentTraceTest report, tracePath, "Kani" & CStr(attempt) & ";" & ExperimentNumberTest(gamma) & ";" & _
                ExperimentNumberTest(actual(6) / 1000#) & ";" & ExperimentNumberTest(response.AcceptedSteelStress(hoopPoint)) & ";" & _
                ExperimentNumberTest(response.AcceptedTransverseSteelUtilization) & ";" & ExperimentNumberTest(errorNorm) & ";" & ExperimentNumberTest(Timer - started)
            If errorNorm < 0.000005 And Abs(gamma - candidate) < 0.000000001 Then passed = passed + 1 Else failed = failed + 1
            If coordinate = maximumShear Then
                stopReason = "TargetReached"
                Exit For
            End If
            increment = increment * 1.25
            If increment > NOMINAL_INCREMENT Then increment = NOMINAL_INCREMENT
        Else
            AddExperimentTraceTest report, tracePath, "KANI_REJECTED: " & CStr(attempt) & " | gamma: " & ExperimentNumberTest(candidate) & _
                " | " & solver.ResultMeta.ResultComment & " | " & solver.ResultMeta.DiagnosticDetails
            increment = increment * 0.5
            If increment < NOMINAL_INCREMENT / 32# Then
                stopReason = "MinimumShearIncrement"
                Exit For
            End If
        End If
    Next attempt
    AddExperimentTraceTest report, tracePath, "KANI_STOP: " & stopReason
    If passed = 0 Then failed = failed + 1
    TraceKaniShearTest = report & "TOTAL_KANI_SHEAR_EQUILIBRIUM: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    TraceKaniShearTest = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит номинальное сечение Kani. Положение осей масштабировано по рис. 4a
' Poliotti и Bairán. Фактические радиусы загибов и диаграммы после текучести неизвестны.
Private Function BuildKaniSectionTest(ByRef layout As CTransverseRebarLayout, ByRef steel As Object) As CSectionModel
    Dim query As New CSectionGeometryQuery, region As New CGeometryRegion
    region.Initialize query.RegionFromPath(query.AdjacentRectangles("Kani_Outer", Array(Array(-152.5, -305#, 305#, 610#))))
    Dim mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder, section As CSectionModel
    mesh.BuildMesh region, 20#, 20#, cellGeometry:="Contour"
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    region.BuildContours section.Contours

    Dim i As Long, x As Double
    For i = 0 To 3
        x = -105# + 70# * i
        section.AddRebarElement x, -254#, 25#, GEOM_PI * 156.25, "KANI_LONG"
        section.AddRebarElement x, 254#, 25#, GEOM_PI * 156.25, "KANI_LONG"
    Next i
    Set layout = New CTransverseRebarLayout
    layout.AddContourGroup section, "KANI_HOOP", vbNullString, 10#, 100#, 25#, "KANI_HOOP"
    Set steel = CreateObject("Scripting.Dictionary")
    steel.Add "KANI_LONG", ExperimentBilinearSteelTest(200000#, 442#, 0#)
    steel.Add "KANI_HOOP", ExperimentBilinearSteelTest(200000#, 400#, 0#)
    Set BuildKaniSectionTest = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Воспроизводит предварительный момент и натурные приращения Mx/T.
' Протокол хранит даже последнюю допустимую точку перед численным отказом.
Public Function RunTestCoupledTBU3(ByVal spansText As String, ByVal maximumStageText As String) As String
    RunTestCoupledTBU3 = TraceTBU3Test(spansText, maximumStageText, False)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' До пятой стадии задаёт усилия, затем управляет закручиванием до
' запрошенной стадии либо до явно зафиксированной остановки.
Public Function RunTestCoupledTBU3Twist(ByVal spansText As String, ByVal maximumStageText As String) As String
    RunTestCoupledTBU3Twist = TraceTBU3Test(spansText, maximumStageText, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт начальную траекторию TBU3 рабочему алгоритму подтверждения
' первого максимума с повторным прохождением обоих плеч.
Public Function RunTestCoupledTBU3Peak(ByVal spansText As String, ByVal maximumStageText As String) As String
    RunTestCoupledTBU3Peak = TraceTBU3Test(spansText, maximumStageText, True, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает принятые точки силового и деформационного контроля отдельно.
Private Function TraceTBU3Test(ByVal spansText As String, ByVal maximumStageText As String, ByVal followTwist As Boolean, _
        Optional ByVal confirmPeak As Boolean = False) As String
    On Error GoTo Failed
    Const EC As Double = 30000#
    Const NU As Double = 0.2
    Const FC As Double = 34.8
    Const FT As Double = 1.95
    Const MX_OFFSET As Double = -29600000#
    Const DELTA_MX As Double = -14870000#
    Const DELTA_T As Double = 10350000#
    Dim report As String, passed As Long, failures As Long, tracePath As String
    Dim spans As Long: spans = CLng(spansText)
    Dim maximumStage As Long: maximumStage = CLng(maximumStageText)
    If spans < 1 Or spans > 3 Or maximumStage < 1 Or maximumStage > 21 Then Err.Raise 5, , "Тестовые границы: базис 1–3, стадия 1–21."
    Dim forceStage As Long: forceStage = maximumStage
    If followTwist And forceStage > 5 Then forceStage = 5
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) > 0 Then tracePath = ThisWorkbook.Path & "\TBU3_Trace.csv"
    AddExperimentTraceTest report, tracePath, "stage;actual_stage;Mx_kNm;T_kNm;kappa_per_m;twist_rad_per_m;hoop_utilization;steel_utilization;stress_residual_MPa;elapsed_s;status"

    Dim layout As CTransverseRebarLayout, steel As Object, section As CSectionModel
    Set section = BuildTBU3SectionTest(layout, steel)
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, response As New CCoupledReinforcedResponse
    operators.PrepareModel section, EC, NU, spans, spans, 5, layout, steelModels:=steel
    concrete.Initialize operators.PointCount, EC, NU, FC * 0.6, FC, FT, 0.0045, 0.12, 80#, 125#
    response.Initialize operators, concrete, steel
    Dim solver As New CCoupledNewtonSolver, path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    solver.Initialize operators, response, EC
    path.Initialize 0#, 0#, MX_OFFSET, DELTA_MX, 0#, 0#, 0#, DELTA_T
    branch.Initialize solver, path, 1, 12, 24, 0.0000005, 0.00000002, saveAcceptedFields:=confirmPeak

    Dim stage As Long, solved As Boolean, started As Double, reached As Double, meta As CResultMeta
    Dim field() As Double, achieved() As Double, target() As Double, errorNorm As Double
    Dim eps0 As Double, kx As Double, ky As Double, twist As Double, row As String
    For stage = 0 To forceStage
        started = Timer
        solved = branch.Advance(CDbl(stage))
        Set meta = branch.ResultMeta
        reached = branch.AcceptedLambda
        row = CStr(stage) & ";" & ExperimentNumberTest(reached)
        If branch.HasAcceptedVariableState Or branch.AcceptedOffsetFraction = 1# And solved Then
            branch.CopyAccepted field, achieved
            path.TargetSix reached, target
            operators.GeneralizedDeformation field, eps0, kx, ky, twist
            errorNorm = TBU3ForceErrorTest(achieved, target)
            row = row & ";" & ExperimentNumberTest(-achieved(2) / 1E+6) & ";" & ExperimentNumberTest(achieved(4) / 1E+6) & _
                ";" & ExperimentNumberTest(-kx * 1000#) & ";" & ExperimentNumberTest(twist * 1000#) & _
                ";" & ExperimentNumberTest(response.AcceptedTransverseSteelUtilization) & ";" & ExperimentNumberTest(response.AcceptedSteelUtilization) & _
                ";" & ExperimentNumberTest(errorNorm)
            If errorNorm < 0.000005 Then passed = passed + 1 Else failures = failures + 1
        Else
            row = row & ";;;;;;;"
        End If
        row = row & ";" & ExperimentNumberTest(Timer - started)
        If solved Then
            AddExperimentTraceTest report, tracePath, row & ";Accepted"
        Else
            AddExperimentTraceTest report, tracePath, row & ";Unresolved"
            AddExperimentTraceTest report, tracePath, "DIAGNOSTICS: " & meta.ResultComment & " | " & meta.DiagnosticDetails
            Exit For
        End If
    Next stage
    AddExperimentTraceTest report, tracePath, "LAST_REQUESTED_STAGE_ACCEPTED: " & CStr(solved)
    If followTwist And branch.HasAcceptedVariableState And reached < maximumStage Then
        If confirmPeak Then
            ConfirmTBU3PeakTest operators, solver, path, branch, response, maximumStage, report, tracePath, passed, failures
        Else
            ContinueTBU3ByTwistTest operators, solver, path, branch, response, maximumStage, report, tracePath, passed, failures
        End If
    End If
    If passed = 0 Then failures = failures + 1
    TraceTBU3Test = report & "TOTAL_TBU3_FORCE_EQUILIBRIUM: passed=" & passed & "; failed=" & failures
    Exit Function
Failed:
    TraceTBU3Test = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет рабочий критерий максимума отдельно от наблюдаемого падения
' в диагностических точках. Сохраняет статус, точность и интегралы пика.
Private Sub ConfirmTBU3PeakTest(ByVal operators As CCoupledElasticState, ByVal solver As CCoupledNewtonSolver, _
        ByVal path As CLoadPathVector, ByVal branch As CCoupledLoadContinuation, ByVal response As CCoupledReinforcedResponse, _
        ByVal requestedStage As Long, ByRef report As String, ByVal tracePath As String, ByRef passed As Long, ByRef failures As Long)
    Const STAGE_TOLERANCE As Double = 0.005 ' 0,052 кН·м на траектории кручения TBU3.
    Dim baseLoad() As Double, direction() As Double, accepted() As Double, resultants() As Double
    path.Components baseLoad, direction
    branch.CopyAccepted accepted, resultants
    Dim controlled As New CCoupledDeformationContinuation
    AddExperimentTraceTest report, tracePath, "PEAK_SEARCH_STARTED: " & ExperimentNumberTest(branch.AcceptedLambda)
    controlled.Trace solver, response, baseLoad, direction, branch.AcceptedFields, branch.PreviousAcceptedLambda, _
        branch.AcceptedLambda, resultants, CDbl(requestedStage), STAGE_TOLERANCE, 24, 0.0000005, 0.00000002

    Dim meta As CResultMeta: Set meta = controlled.ResultMeta
    AddExperimentTraceTest report, tracePath, "PEAK_DIAGNOSTICS: " & meta.ResultComment & " | " & meta.DiagnosticDetails
    AddExperimentTraceTest report, tracePath, "PEAK_BEST_ADMISSIBLE_STAGE: " & ExperimentNumberTest(controlled.BestAdmissibleLambda)
    Dim peak As Double, accuracy As Double, designVerified As Boolean
    If Not controlled.CopyPeak(peak, accuracy, designVerified) Then
        AddExperimentTraceTest report, tracePath, "PEAK_CONFIRMED: False"
        Exit Sub
    End If
    AddExperimentTraceTest report, tracePath, "PEAK_CONFIRMED: True"
    Dim fields As Collection
    If Not controlled.CopyLimitState(fields, resultants) Then
        failures = failures + 1
        Exit Sub
    End If
    Dim target() As Double, eps0 As Double, kx As Double, ky As Double, twist As Double, errorNorm As Double
    path.TargetSix peak, target
    accepted = fields(fields.Count)
    operators.GeneralizedDeformation accepted, eps0, kx, ky, twist
    errorNorm = TBU3ForceErrorTest(resultants, target)
    AddExperimentTraceTest report, tracePath, "PEAK_STAGE_M_T_TWIST_ERROR_ACCURACY;" & ExperimentNumberTest(peak) & ";" & _
        ExperimentNumberTest(-resultants(2) / 1E+6) & ";" & ExperimentNumberTest(resultants(4) / 1E+6) & ";" & _
        ExperimentNumberTest(twist * 1000#) & ";" & ExperimentNumberTest(errorNorm) & ";" & ExperimentNumberTest(accuracy)
    If errorNorm < 0.000005 And accuracy <= STAGE_TOLERANCE Then passed = passed + 1 Else failures = failures + 1
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет переход через отказ силового шага по управляемому углу,
' сохраняя постоянный момент, отношение приращений и принятую историю.
Private Sub ContinueTBU3ByTwistTest(ByVal operators As CCoupledElasticState, ByVal solver As CCoupledNewtonSolver, _
        ByVal path As CLoadPathVector, ByVal branch As CCoupledLoadContinuation, ByVal response As CCoupledReinforcedResponse, _
        ByVal requestedStage As Long, ByRef report As String, ByVal tracePath As String, ByRef passed As Long, ByRef failures As Long)
    Const MAX_ATTEMPTS As Long = 200 ' Лимит опытного прогона, включая уменьшения шага.
    Dim baseLoad() As Double, direction() As Double, projection() As Double
    path.Components baseLoad, direction
    ReDim projection(1 To operators.DegreeCount): projection(4) = 1#
    Dim field() As Double, achieved() As Double, target() As Double, eps0 As Double, kx As Double, ky As Double, twist As Double
    branch.CopyAccepted field, achieved
    operators.GeneralizedDeformation field, eps0, kx, ky, twist
    If twist <= 0# Then
        AddExperimentTraceTest report, tracePath, "TWIST_STOP: NonpositiveInitialTwist"
        Exit Sub
    End If
    Dim increment As Double, initialIncrement As Double, length As Double, reached As Double, coordinate As Double
    length = field(4) / twist ' Обобщённая координата равна относительному углу, умноженному на длину масштаба.
    initialIncrement = 0.0002 * length / 1000#
    increment = initialIncrement: reached = branch.AcceptedLambda: coordinate = field(4)
    Dim attempt As Long, trialLambda As Double, errorNorm As Double, row As String, started As Double, stopReason As String
    stopReason = "AttemptLimit"
    AddExperimentTraceTest report, tracePath, "TWIST_REQUESTED_STAGE: " & CStr(requestedStage)
    For attempt = 1 To MAX_ATTEMPTS
        started = Timer
        If solver.TryDeformationStep(baseLoad, direction, projection, coordinate + increment, reached, trialLambda, _
                0.0000005, 0.00000002, 24) Then
            solver.CopyTrial field, achieved: solver.AcceptStep
            coordinate = field(4): reached = trialLambda
            path.TargetSix reached, target
            operators.GeneralizedDeformation field, eps0, kx, ky, twist
            errorNorm = TBU3ForceErrorTest(achieved, target)
            row = "Twist" & CStr(attempt) & ";" & ExperimentNumberTest(reached) & ";" & ExperimentNumberTest(-achieved(2) / 1E+6) & _
                ";" & ExperimentNumberTest(achieved(4) / 1E+6) & ";" & ExperimentNumberTest(-kx * 1000#) & ";" & ExperimentNumberTest(twist * 1000#) & _
                ";" & ExperimentNumberTest(response.AcceptedTransverseSteelUtilization) & ";" & ExperimentNumberTest(response.AcceptedSteelUtilization) & _
                ";" & ExperimentNumberTest(errorNorm) & ";" & ExperimentNumberTest(Timer - started) & ";Accepted"
            AddExperimentTraceTest report, tracePath, row
            If errorNorm < 0.000005 Then passed = passed + 1 Else failures = failures + 1
            If reached >= requestedStage Then
                stopReason = "TargetReached"
                Exit For
            End If
            If increment < initialIncrement Then increment = increment * 1.25
            If increment > initialIncrement Then increment = initialIncrement
        Else
            AddExperimentTraceTest report, tracePath, "TWIST_REJECTED: " & CStr(attempt) & " | " & ExperimentNumberTest(Timer - started) & _
                " | " & solver.ResultMeta.ResultComment & " | " & solver.ResultMeta.DiagnosticDetails
            increment = increment * 0.5
            If increment < initialIncrement / 32# Then
                stopReason = "MinimumTwistIncrement"
                Exit For
            End If
        End If
    Next attempt
    AddExperimentTraceTest report, tracePath, "TWIST_CONTROL_REACHED_STAGE: " & ExperimentNumberTest(reached)
    AddExperimentTraceTest report, tracePath, "TWIST_STOP: " & stopReason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выдаёт напряжения однородной коробки при шести ненулевых усилиях для
' сравнения с независимым треугольным КЭ-расчётом sectionproperties.
Public Function RunTestCoupledTBU3Elastic() As String
    On Error GoTo Failed
    Const E As Double = 30000#
    Const G As Double = E / 2# ' В обеих программах контроль проводится при нулевом коэффициенте Пуассона.
    Const FORCE As Double = 100000#
    Const MX As Double = 15000000#
    Const MY As Double = -11000000#
    Const TORQUE As Double = 10000000#
    Const QX As Double = 28000#
    Const QY As Double = -19000#
    Dim section As CSectionModel: Set section = BuildTBU3BoxTest()
    Dim points As Variant: points = Array(Array(0#, 167#), Array(0#, -167#), Array(216#, 0#), Array(-216#, 0#), _
        Array(210#, 170#), Array(-210#, -170#))
    Dim spans As Long, point As Variant, passed As Long, failed As Long, report As String
    Dim solver As CCoupledElasticState, deformation() As Double, resultants() As Double, stress() As Double
    Dim field() As Double, raw As Variant, strain(1 To 6) As Double, warping As Double
    Dim path As New CLoadPathVector, target() As Double
    path.Initialize 0#, FORCE, 0#, MX, 0#, MY, 0#, TORQUE, 0#, QX, 0#, QY
    path.TargetSix 1#, target
    For spans = 1 To 3
        Set solver = New CCoupledElasticState
        solver.PrepareModel section, E, 0#, spans, spans, 7
        If solver.Solve(FORCE, MX, MY, TORQUE, QX, QY, deformation, resultants, stress, raw) Then
            field = raw
            report = report & "ELASTIC_RIGIDITY;" & CStr(spans) & ";" & ExperimentNumberTest(TORQUE / (G * deformation(4))) & vbCrLf
            If TBU3ForceErrorTest(resultants, target) < 0.0000001 Then passed = passed + 1 Else failed = failed + 1
            For Each point In points
                solver.KinematicsAt CDbl(point(0)), CDbl(point(1)), field, strain, warping
                report = report & "ELASTIC_STRESS;" & CStr(spans) & ";" & ExperimentNumberTest(CDbl(point(0))) & ";" & ExperimentNumberTest(CDbl(point(1))) & _
                    ";" & ExperimentNumberTest(E * strain(3)) & ";" & ExperimentNumberTest(G * strain(5)) & ";" & ExperimentNumberTest(G * strain(6)) & vbCrLf
            Next point
        Else
            failed = failed + 1
            report = report & "FAIL: spans=" & CStr(spans) & " " & solver.ResultMeta.ResultComment & vbCrLf
        End If
    Next spans
    RunTestCoupledTBU3Elastic = report & "TOTAL_TBU3_ELASTIC_EQUILIBRIUM: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledTBU3Elastic = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит номинальную коробку и измеренные площади стали. Позиции осей
' оценены по рис. 6.39. Малая длина стыка хомута отдельно не моделируется.
Private Function BuildTBU3SectionTest(ByRef layout As CTransverseRebarLayout, ByRef steel As Object) As CSectionModel
    Dim section As CSectionModel: Set section = BuildTBU3BoxTest()

    Dim i As Long, dx As Double
    For i = -3 To 3
        dx = 206# * CDbl(i) / 3#
        section.AddRebarElement dx, -154#, Sqr(4# * 510# / GEOM_PI), 510#, "TBU3_8"
    Next i
    For i = -1 To 1
        section.AddRebarElement 216# * i, 167#, Sqr(4# * 129# / GEOM_PI), 129#, "TBU3_4"
        section.AddRebarElement -218#, 78# * i, Sqr(4# * 71# / GEOM_PI), 71#, "TBU3_3"
        section.AddRebarElement 218#, 78# * i, Sqr(4# * 71# / GEOM_PI), 71#, "TBU3_3"
    Next i
    Set layout = New CTransverseRebarLayout
    Dim diameter As Double: diameter = Sqr(4# * 129# / GEOM_PI)
    layout.AddContourGroup section, "TBU3_HOOP", vbNullString, diameter, 150#, 32# - diameter / 2#, "TBU3_HOOP"

    Set steel = CreateObject("Scripting.Dictionary")
    steel.Add "TBU3_8", ExperimentBilinearSteelTest(194200#, 436#, 9000#)
    steel.Add "TBU3_4", ExperimentBilinearSteelTest(201900#, 393#, 4900#)
    steel.Add "TBU3_3", ExperimentBilinearSteelTest(205100#, 552#, 9000#)
    steel.Add "TBU3_HOOP", ExperimentBilinearSteelTest(198500#, 379#, 5200#)
    Set BuildTBU3SectionTest = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Использует одинаковые контуры для упругой и нелинейной постановок.
Private Function BuildTBU3BoxTest() As CSectionModel
    Dim query As New CSectionGeometryQuery, outer As CConcreteRegion, inner As CConcreteRegion, geometry As New CGeometryRegion
    Set outer = query.RegionFromPath(query.AdjacentRectangles("TBU3_Outer", Array(Array(-254#, -205#, 508#, 410#))))
    Set inner = query.RegionFromPath(query.AdjacentRectangles("TBU3_Opening", Array(Array(-178#, -129#, 356#, 258#))))
    geometry.Initialize query.SubtractContainedRegion(outer, inner)
    Dim mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder, section As CSectionModel
    mesh.BuildMesh geometry, 20#, 18#, cellGeometry:="Contour"
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    geometry.BuildContours section.Contours
    Set BuildTBU3BoxTest = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Двухлинейная аппроксимация измеренной стали. После текучести модуль
' упрочнения включается сразу. Предел 0,05 принят только для диагностики.
Private Function ExperimentBilinearSteelTest(ByVal modulus As Double, ByVal strength As Double, ByVal hardening As Double) As CCoupledSteelModel
    Dim steel As New CCoupledSteelModel
    steel.Initialize modulus, strength, strength, hardening, hardening, 0.05, 0.05
    Set ExperimentBilinearSteelTest = steel
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Приводит остатки шести усилий к МПа по номинальной площади коробки.
Private Function TBU3ForceErrorTest(ByRef achieved() As Double, ByRef target() As Double) As Double
    Const AREA As Double = 508# * 410# - 356# * 258#
    Dim i As Long, difference As Double
    For i = 1 To 6
        difference = Abs(achieved(i) - target(i)) / AREA
        If i >= 2 And i <= 4 Then difference = difference / 508#
        If difference > TBU3ForceErrorTest Then TBU3ForceErrorTest = difference
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Русская локаль Excel не меняет десятичную точку в машинном протоколе.
Private Function ExperimentNumberTest(ByVal value As Double) As String
    ExperimentNumberTest = Trim$(Str$(value))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отмечает завершённую стадию также на диске, чтобы длительный контроль
' можно было наблюдать без обращения к пользовательскому экземпляру Excel.
Private Sub AddExperimentTraceTest(ByRef report As String, ByVal path As String, ByVal row As String)
    report = report & row & vbCrLf
    If Len(path) = 0 Then Exit Sub
    Dim file As Integer: file = FreeFile
    Open path For Append As #file
    Print #file, row
    Close #file
End Sub
