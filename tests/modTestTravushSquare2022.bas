Attribute VB_Name = "modTestTravushSquare2022"
Option Explicit

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставление этапа БВ-КВ-410(2) со статьями 2018 и 2022 годов.
' В модели из статьи шаг 100 мм. Допустим другой шаг для проверки
' чувствительности. Фактическое положение осей хомута не опубликовано.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Прослеживает прямое НДС в сечении k. Суммирование усилий проверяется
' на принятой ветви. Признак повреждения сопоставляется с визуальной трещиной.
Public Function RunTestTravushSquareStage(ByVal pitchText As String, ByVal spansText As String) As String
    On Error GoTo Failed
    Const WIDTH As Double = 200#
    Const EC As Double = 43000# ' СП 63, табл. 6.11 для В100. Экспериментальный модуль неизвестен.
    Const NU As Double = 0.2
    Const FC As Double = 47.5  ' Расчётное сопротивление из п. 2 статьи 2022 года.
    Const FT As Double = 2.2
    Const STEEL_LONG As Double = 210#
    Const STEEL_HOOP As Double = 170#
    Const LOCALIZATION As Double = 125# ' Принято для расчётной чувствительности.
    Const PEAK_STRAIN As Double = 0.002 ' Для расчётной диаграммы. Опытная диаграмма не дана.
    Const DAMAGE_THRESHOLD As Double = 0.000001
    Const Q_STAGE As Double = 15250#
    Const SUPPORT_STAGE As Double = 30.5
    Const SUPPORT_CRACK As Double = 13.3 ' Опытная обобщённая нагрузка по табл. 1 статьи 2018 года.
    Dim passed As Long, failed As Long, report As String
    Dim pitch As Double, spans As Long
    pitch = Val(pitchText): spans = CLng(spansText)
    If pitch <= 0# Or pitch > 200# Or spans < 1 Or spans > 3 Then _
        Err.Raise 5, , "Задайте расчётный шаг хомута в пределах 0 < s <= 200 мм и число интервалов базиса 1–3."

    Dim section As CSectionModel, layout As CTransverseRebarLayout
    Set section = BuildTravushSquareSectionTest(pitch, layout)
    Dim steel As Object, longitudinal As New CCoupledSteelModel, hoop As New CCoupledSteelModel
    longitudinal.Initialize 200000#, STEEL_LONG, STEEL_LONG, 0#, 0#, 0.05, 0.05 ' Плато, диагностический предел.
    hoop.Initialize 200000#, STEEL_HOOP, STEEL_HOOP, 0#, 0#, 0.05, 0.05
    Set steel = CreateObject("Scripting.Dictionary")
    steel.Add "TRAVUSH_LONG", longitudinal
    steel.Add "TRAVUSH_HOOP", hoop

    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, reinforcement As New CCoupledReinforcedResponse
    Dim gt As Double, gc As Double
    gt = 0.073 * FC ^ 0.18: gc = (FC / FT) ^ 2 * gt ' Оценки Poliotti и Bairán (2019), а не опытные энергии.
    operators.PrepareModel section, EC, NU, spans, spans, 5, layout, steelModels:=steel
    concrete.Initialize operators.PointCount, EC, NU, 0.6 * FC, FC, FT, PEAK_STRAIN, gt, gc, LOCALIZATION
    reinforcement.Initialize operators, concrete, steel

    Dim solver As New CCoupledNewtonSolver, path As New CLoadPathVector, branch As New CCoupledLoadContinuation
    solver.Initialize operators, reinforcement, EC
    path.Initialize 0#, 0#, 0#, Q_STAGE * 300#, 0#, 0#, 0#, Q_STAGE * 410#, 0#, 0#, 0#, Q_STAGE
    branch.Initialize solver, path, 4, 20, 24, 0.0000005, 0.00000002
    report = "FIXTURE=TRAVUSH_SQUARE_K;M_KNM=4.575;T_KNM=6.2525;Q_KN=15.25" & vbCrLf
    report = report & "SETUP;pitch_mm=" & TravushNumberTest(pitch) & ";pitch_matches_article_model=" & CStr(pitch = 100#) & _
        ";spans=" & CStr(spans) & _
        ";hoop_axis_from_face_mm=22;steel_limit=0.05;Ec_MPa=43000;localization_mm=125" & vbCrLf
    report = report & "MESH;concrete_points=" & CStr(operators.ConcretePointCount) & _
        ";all_points=" & CStr(operators.PointCount) & ";DOF=" & CStr(operators.DegreeCount) & vbCrLf
    report = report & "EXPERIMENT_CRACK;R_sup_kN=13.3;lambda_nominal=" & TravushNumberTest(SUPPORT_CRACK / SUPPORT_STAGE) & vbCrLf
    report = report & "stage;lambda;R_sup_kN;N_kN;Mx_kNm;My_kNm;T_kNm;Qx_kN;Qy_kN;" & _
        "twist_rad_per_m;curvature_x_per_m;tension_indicator;compression_indicator;" & _
        "steel_long_max_MPa;hoop_max_MPa;stress_residual_MPa" & vbCrLf

    Dim requested As Variant, meta As CResultMeta
    Dim actual() As Double, target() As Double, field() As Double
    Dim eps0 As Double, kx As Double, ky As Double, twist As Double, tension As Double, compression As Double
    Dim steelStress As Double, hoopStress As Double, reached As Double, errorNorm As Double, i As Long, damageSeen As Boolean
    For Each requested In Array(0.25, 0.35, 0.375, 0.4, SUPPORT_CRACK / SUPPORT_STAGE, 0.45, 0.5, 0.65, 0.8, 1#)
        If Not branch.Advance(CDbl(requested)) Then
            Set meta = branch.ResultMeta
            report = report & "STOP;requested=" & TravushNumberTest(CDbl(requested)) & ";accepted=" & _
                TravushNumberTest(branch.AcceptedLambda) & ";remaining=" & TravushNumberTest(CDbl(requested) - branch.AcceptedLambda) & _
                ";" & meta.ResultComment & " " & meta.DiagnosticDetails & vbCrLf
            Exit For
        End If
        reached = branch.AcceptedLambda
        branch.CopyAccepted field, actual
        path.TargetSix reached, target
        operators.GeneralizedDeformation field, eps0, kx, ky, twist
        TravushMaterialExtremesTest operators, section.RebarCount, concrete, reinforcement, tension, compression, steelStress, hoopStress
        errorNorm = TravushBalanceErrorTest(actual, target, WIDTH)
        report = report & "Accepted;" & TravushNumberTest(reached) & ";" & TravushNumberTest(SUPPORT_STAGE * reached)
        For i = 1 To 6
            If i >= 2 And i <= 4 Then
                report = report & ";" & TravushNumberTest(actual(i) / 1000000#)
            Else
                report = report & ";" & TravushNumberTest(actual(i) / 1000#)
            End If
        Next i
        report = report & ";" & TravushNumberTest(twist * 1000#) & ";" & TravushNumberTest(kx * 1000#) & _
            ";" & TravushNumberTest(tension) & ";" & TravushNumberTest(compression) & ";" & TravushNumberTest(steelStress) & _
            ";" & TravushNumberTest(hoopStress) & ";" & TravushNumberTest(errorNorm) & vbCrLf
        If Not damageSeen And tension > DAMAGE_THRESHOLD Then
            damageSeen = True
            report = report & "FIRST_TENSION_DAMAGE_SAMPLE;R_sup_kN=" & TravushNumberTest(SUPPORT_STAGE * reached) & vbCrLf
        End If
        If errorNorm <= 0.000005 Then passed = passed + 1 Else failed = failed + 1
    Next requested
    If passed = 0 Then failed = failed + 1
    report = report & "TARGET_REACHED=" & CStr(reached = 1#) & vbCrLf
    RunTestTravushSquareStage = report & "TOTAL_TRAVUSH_SQUARE_EQUILIBRIUM: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestTravushSquareStage = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Строит квадрат и восемь Ø6 по рис. 2. Ось хомута принята на 6 мм
' наружу от осей продольных стержней при одинаковом диаметре.
Private Function BuildTravushSquareSectionTest(ByVal pitch As Double, ByRef layout As CTransverseRebarLayout) As CSectionModel
    Dim query As New CSectionGeometryQuery, region As New CGeometryRegion
    region.Initialize query.RegionFromPath(query.AdjacentRectangles("TravushSquareOuter", Array(Array(-100#, -100#, 200#, 200#))))
    Dim mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder, section As CSectionModel
    mesh.BuildMesh region, 20#, 20#, cellGeometry:="Contour"
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    region.BuildContours section.Contours

    Dim column As Long, row As Long
    For row = -1 To 1
        For column = -1 To 1
            If row <> 0 Or column <> 0 Then _
                section.AddRebarElement 72# * column, 72# * row, 6#, 9# * GEOM_PI, "TRAVUSH_LONG"
        Next column
    Next row
    Set layout = New CTransverseRebarLayout
    layout.AddContourGroup section, "TravushHoop", vbNullString, 6#, pitch, 19#, "TRAVUSH_HOOP"
    Set BuildTravushSquareSectionTest = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Раздельно снимает наибольшие меры повреждения в бетоне и напряжения
' в продольной стали и хомуте в принятом состоянии.
Private Sub TravushMaterialExtremesTest(ByVal operators As CCoupledElasticState, ByVal rebarCount As Long, _
        ByVal concrete As CCoupledConcreteResponse, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef tension As Double, ByRef compression As Double, ByRef longStress As Double, ByRef hoopStress As Double)
    tension = 0#: compression = 0#: longStress = 0#: hoopStress = 0#
    Dim i As Long, tensionAtPoint As Double, compressionAtPoint As Double
    For i = 1 To operators.ConcretePointCount
        concrete.AcceptedPointDamage i, tensionAtPoint, compressionAtPoint
        If tensionAtPoint > tension Then tension = tensionAtPoint
        If compressionAtPoint > compression Then compression = compressionAtPoint
    Next i
    Dim stress As Double
    For i = operators.ConcretePointCount + 1 To operators.ConcretePointCount + rebarCount
        stress = Abs(reinforcement.AcceptedSteelStress(i))
        If stress > longStress Then longStress = stress
    Next i
    For i = operators.ConcretePointCount + rebarCount + 1 To operators.PointCount
        stress = Abs(reinforcement.AcceptedSteelStress(i))
        If stress > hoopStress Then hoopStress = stress
    Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Приводит невязки сил по площади, а моментов по площади и стороне.
Private Function TravushBalanceErrorTest(ByRef actual() As Double, ByRef target() As Double, ByVal width As Double) As Double
    Dim i As Long, difference As Double
    For i = 1 To 6
        difference = Abs(actual(i) - target(i)) / (width * width)
        If i >= 2 And i <= 4 Then difference = difference / width
        If difference > TravushBalanceErrorTest Then TravushBalanceErrorTest = difference
    Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число отчёта с точкой для его чтения независимо от локали.
Private Function TravushNumberTest(ByVal value As Double) As String
    TravushNumberTest = Trim$(Str$(value))
End Function
