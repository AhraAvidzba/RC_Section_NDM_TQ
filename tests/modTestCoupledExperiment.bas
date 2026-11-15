Attribute VB_Name = "modTestCoupledExperiment"
Option Explicit

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставление связанного НДС с измерениями на коробчатой балке Onsongo TBU3.
' Геометрия и стали восстановлены по Bairán, рис. 6.39, и Yamamoto, разд. 4.2.
' Параметры, отсутствующие в источниках, фиксируются как допущения расчёта.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Воспроизводит предварительный момент и натурные приращения Mx/T.
' Протокол хранит даже последнюю допустимую точку перед численным отказом.
Public Function RunTestCoupledTBU3(ByVal spansText As String, ByVal maximumStageText As String) As String
    RunTestCoupledTBU3 = TraceTBU3Test(spansText, maximumStageText, False)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отбивает ту же начальную траекторию и затем управляет закручиванием.
Public Function RunTestCoupledTBU3Twist(ByVal spansText As String, ByVal maximumStageText As String) As String
    RunTestCoupledTBU3Twist = TraceTBU3Test(spansText, maximumStageText, True)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает принятые точки силового и деформационного контроля отдельно.
Private Function TraceTBU3Test(ByVal spansText As String, ByVal maximumStageText As String, ByVal followTwist As Boolean) As String
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
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) > 0 Then tracePath = ThisWorkbook.Path & "\TBU3_Trace.csv"
    AddTBU3TraceTest report, tracePath, "stage;actual_stage;Mx_kNm;T_kNm;kappa_per_m;twist_rad_per_m;hoop_utilization;steel_utilization;stress_residual_MPa;elapsed_s;status"

    Dim layout As CTransverseRebarLayout, steel As Object, section As CSectionModel
    Set section = BuildTBU3SectionTest(layout, steel)
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, response As New CCoupledReinforcedResponse
    operators.PrepareModel section, EC, NU, spans, spans, 5, layout, steelModels:=steel
    concrete.Initialize operators.PointCount, EC, NU, FC * 0.6, FC, FT, 0.0045, 0.12, 80#, 125#
    response.Initialize operators, concrete, steel
    Dim solver As New CCoupledNewtonSolver, path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    solver.Initialize operators, response, EC
    path.Initialize 0#, 0#, MX_OFFSET, DELTA_MX, 0#, 0#, 0#, DELTA_T
    branch.Initialize solver, path, 1, 12, 24, 0.0000005, 0.00000002

    Dim stage As Long, solved As Boolean, started As Double, reached As Double, meta As CResultMeta
    Dim field() As Double, achieved() As Double, target() As Double, errorNorm As Double
    Dim eps0 As Double, kx As Double, ky As Double, twist As Double, row As String
    For stage = 0 To maximumStage
        started = Timer
        solved = branch.Advance(CDbl(stage))
        Set meta = branch.ResultMeta
        reached = branch.AcceptedLambda
        row = CStr(stage) & ";" & TBU3NumberTest(reached)
        If branch.HasAcceptedVariableState Or branch.AcceptedOffsetFraction = 1# And solved Then
            branch.CopyAccepted field, achieved
            path.TargetSix reached, target
            operators.GeneralizedDeformation field, eps0, kx, ky, twist
            errorNorm = TBU3ForceErrorTest(achieved, target)
            row = row & ";" & TBU3NumberTest(-achieved(2) / 1E+6) & ";" & TBU3NumberTest(achieved(4) / 1E+6) & _
                ";" & TBU3NumberTest(-kx * 1000#) & ";" & TBU3NumberTest(twist * 1000#) & _
                ";" & TBU3NumberTest(response.AcceptedTransverseSteelUtilization) & ";" & TBU3NumberTest(response.AcceptedSteelUtilization) & _
                ";" & TBU3NumberTest(errorNorm)
            If errorNorm < 0.000005 Then passed = passed + 1 Else failures = failures + 1
        Else
            row = row & ";;;;;;;"
        End If
        row = row & ";" & TBU3NumberTest(Timer - started)
        If solved Then
            AddTBU3TraceTest report, tracePath, row & ";Accepted"
        Else
            AddTBU3TraceTest report, tracePath, row & ";Unresolved"
            AddTBU3TraceTest report, tracePath, "DIAGNOSTICS: " & meta.ResultComment & " | " & meta.DiagnosticDetails
            Exit For
        End If
    Next stage
    AddTBU3TraceTest report, tracePath, "LAST_REQUESTED_STAGE_ACCEPTED: " & CStr(solved)
    If followTwist And branch.HasAcceptedVariableState Then _
        ContinueTBU3ByTwistTest operators, solver, path, branch, response, report, tracePath, passed, failures
    If passed = 0 Then failures = failures + 1
    TraceTBU3Test = report & "TOTAL_TBU3_FORCE_EQUILIBRIUM: passed=" & passed & "; failed=" & failures
    Exit Function
Failed:
    TraceTBU3Test = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет переход через отказ силового шага по управляемому углу,
' сохраняя постоянный момент, отношение приращений и принятую историю.
Private Sub ContinueTBU3ByTwistTest(ByVal operators As CCoupledElasticState, ByVal solver As CCoupledNewtonSolver, _
        ByVal path As CLoadPathVector, ByVal branch As CCoupledLoadContinuation, ByVal response As CCoupledReinforcedResponse, _
        ByRef report As String, ByVal tracePath As String, ByRef passed As Long, ByRef failures As Long)
    Dim baseLoad() As Double, direction() As Double, projection() As Double
    path.Components baseLoad, direction
    ReDim projection(1 To operators.DegreeCount): projection(4) = 1#
    Dim field() As Double, achieved() As Double, target() As Double, eps0 As Double, kx As Double, ky As Double, twist As Double
    branch.CopyAccepted field, achieved
    operators.GeneralizedDeformation field, eps0, kx, ky, twist
    If twist <= 0# Then Exit Sub
    Dim increment As Double, initialIncrement As Double, length As Double, reached As Double, coordinate As Double
    length = field(4) / twist ' Обобщённая координата равна относительному углу, умноженному на длину масштаба.
    initialIncrement = 0.0002 * length / 1000#
    increment = initialIncrement: reached = branch.AcceptedLambda: coordinate = field(4)
    Dim attempt As Long, trialLambda As Double, errorNorm As Double, row As String, started As Double
    For attempt = 1 To 32
        started = Timer
        If solver.TryDeformationStep(baseLoad, direction, projection, coordinate + increment, reached, trialLambda, _
                0.0000005, 0.00000002, 24) Then
            solver.CopyTrial field, achieved: solver.AcceptStep
            coordinate = field(4): reached = trialLambda
            path.TargetSix reached, target
            operators.GeneralizedDeformation field, eps0, kx, ky, twist
            errorNorm = TBU3ForceErrorTest(achieved, target)
            row = "Twist" & CStr(attempt) & ";" & TBU3NumberTest(reached) & ";" & TBU3NumberTest(-achieved(2) / 1E+6) & _
                ";" & TBU3NumberTest(achieved(4) / 1E+6) & ";" & TBU3NumberTest(-kx * 1000#) & ";" & TBU3NumberTest(twist * 1000#) & _
                ";" & TBU3NumberTest(response.AcceptedTransverseSteelUtilization) & ";" & TBU3NumberTest(response.AcceptedSteelUtilization) & _
                ";" & TBU3NumberTest(errorNorm) & ";" & TBU3NumberTest(Timer - started) & ";Accepted"
            AddTBU3TraceTest report, tracePath, row
            If errorNorm < 0.000005 Then passed = passed + 1 Else failures = failures + 1
            If reached >= 7# Then Exit For
            If increment < initialIncrement Then increment = increment * 1.25
        Else
            AddTBU3TraceTest report, tracePath, "TWIST_REJECTED: " & CStr(attempt) & " | " & TBU3NumberTest(Timer - started) & _
                " | " & solver.ResultMeta.ResultComment & " | " & solver.ResultMeta.DiagnosticDetails
            increment = increment * 0.5
            If increment < initialIncrement / 32# Then Exit For
        End If
    Next attempt
    AddTBU3TraceTest report, tracePath, "TWIST_CONTROL_REACHED_STAGE: " & TBU3NumberTest(reached)
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
            report = report & "ELASTIC_RIGIDITY;" & CStr(spans) & ";" & TBU3NumberTest(TORQUE / (G * deformation(4))) & vbCrLf
            If TBU3ForceErrorTest(resultants, target) < 0.0000001 Then passed = passed + 1 Else failed = failed + 1
            For Each point In points
                solver.KinematicsAt CDbl(point(0)), CDbl(point(1)), field, strain, warping
                report = report & "ELASTIC_STRESS;" & CStr(spans) & ";" & TBU3NumberTest(CDbl(point(0))) & ";" & TBU3NumberTest(CDbl(point(1))) & _
                    ";" & TBU3NumberTest(E * strain(3)) & ";" & TBU3NumberTest(G * strain(5)) & ";" & TBU3NumberTest(G * strain(6)) & vbCrLf
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
    steel.Add "TBU3_8", TBU3SteelModelTest(194200#, 436#, 9000#)
    steel.Add "TBU3_4", TBU3SteelModelTest(201900#, 393#, 4900#)
    steel.Add "TBU3_3", TBU3SteelModelTest(205100#, 552#, 9000#)
    steel.Add "TBU3_HOOP", TBU3SteelModelTest(198500#, 379#, 5200#)
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
Private Function TBU3SteelModelTest(ByVal modulus As Double, ByVal strength As Double, ByVal hardening As Double) As CCoupledSteelModel
    Dim steel As New CCoupledSteelModel
    steel.Initialize modulus, strength, strength, hardening, hardening, 0.05, 0.05
    Set TBU3SteelModelTest = steel
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
Private Function TBU3NumberTest(ByVal value As Double) As String
    TBU3NumberTest = Trim$(Str$(value))
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Отмечает завершённую стадию также на диске, чтобы длительный контроль
' можно было наблюдать без обращения к пользовательскому экземпляру Excel.
Private Sub AddTBU3TraceTest(ByRef report As String, ByVal path As String, ByVal row As String)
    report = report & row & vbCrLf
    If Len(path) = 0 Then Exit Sub
    Dim file As Integer: file = FreeFile
    Open path For Append As #file
    Print #file, row
    Close #file
End Sub
