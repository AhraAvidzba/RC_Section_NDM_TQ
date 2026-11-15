Attribute VB_Name = "modTestCoupledReinforcedPrism"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ
' ==========================================================================
' Передаёт 48 измеренных винтов OpenSees рабочему решателю сечения.
' Идентичность нагрузок проверяется независимо от различий объёмной модели.
' Записывает только отчёт рядом с изолированной копией, в книгу не пишет.

Private Const PRISM_WIDTH As Double = 200#, PRISM_DEPTH As Double = 160#
Private Const CONCRETE_E As Double = 30000#, CONCRETE_NU As Double = 0.2
Private Const STEEL_E As Double = 200000#, REFERENCE_STRAIN_LIMIT As Double = 0.05 ' Внешняя модель не задаёт разрыв; предел здесь служит только масштабу деформации.
Private Const STAGES As Long = 48, POINTS_PER_SPAN As Long = 4

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подтверждает последовательные загруженные винты с общей историей материала.
Public Function RunTestCoupledReinforcedPrism(ByVal inputPath As String, Optional ByVal spansText As String = "1") As String
    RunTestCoupledReinforcedPrism = TestReplayPrismHistory(inputPath, spansText, 1#)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Допускает уточнение только текущего интервала между исходными винтами.
Public Function RunTestCoupledReinforcedPrismSubsteps(ByVal inputPath As String, Optional ByVal spansText As String = "1") As String
    RunTestCoupledReinforcedPrismSubsteps = TestReplayPrismHistory(inputPath, spansText, 0.0625)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет физические точки сравнения и ход деления интервала в протоколе.
Private Function TestReplayPrismHistory(ByVal inputPath As String, ByVal spansText As String, ByVal minimumStep As Double) As String
    Dim stream As Object, completed As Long, filePath As String
    On Error GoTo Failed
    If InStr(1, ThisWorkbook.Path, "RC_NDM_ModeSuite_", vbTextCompare) = 0 Then _
        Err.Raise 5, "RunTestCoupledReinforcedPrism", "Запустите проверку из изолированной копии книги."
    Dim spans As Long: spans = CLng(spansText)
    If spans < 1 Or spans > 3 Then Err.Raise 5, "RunTestCoupledReinforcedPrism", "Число интервалов базиса должно быть от 1 до 3."
    Dim loads() As Double, signature As String
    TestReadPrismLoads inputPath, loads, signature

    Dim section As CSectionModel, layout As CTransverseRebarLayout, steel As Object
    Set section = TestPrismSection(layout, steel)
    Dim operators As New CCoupledElasticState, frame As New CSectionNumericalFrame
    frame.Initialize section
    operators.PrepareModel section, CONCRETE_E, CONCRETE_NU, spans, spans, POINTS_PER_SPAN, layout, steelModels:=steel
    Dim center As New CCoupledConcreteResponse, response As New CCoupledReinforcedResponse, solver As New CCoupledNewtonSolver
    TestInitializePrismResponse operators, steel, center, response, solver
    Dim shearX() As Double, shearY() As Double
    operators.MeanShearProjections shearX, shearY

    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    filePath = ThisWorkbook.Path & Application.PathSeparator & "ReinforcedPrismVBA.csv"
    Set stream = fileSystem.CreateTextFile(filePath, False, True)
    stream.WriteLine "META;NDM_PRISM_VBA_2;" & signature & ";" & CStr(spans) & ";" & CStr(POINTS_PER_SPAN) & ";" & CStr(operators.PointCount) & ";" & _
        TestPrismNumber(frame.OriginX) & ";" & TestPrismNumber(frame.OriginY) & ";" & TestPrismNumber(frame.LengthScale)
    stream.WriteLine "CONCRETE;30000;0.2;18;30;3;0.002;0.1;10;100"
    stream.WriteLine "STEEL;L500;12;200000;500;2000;0.05;unconfined_longitudinal_replacement"
    stream.WriteLine "STEEL;H240;8;200000;240;2000;0.05;smeared_transverse_replacement_at_pitch_100"

    Dim load As New CSectionLoadState, stage As Long, field() As Double, target() As Double, actual() As Double, meta As CResultMeta
    Dim previous() As Double: ReDim previous(1 To 6)
    Dim outcome As String: outcome = "ReachedAll48"
    For stage = 1 To STAGES
        stream.WriteLine "TARGET;" & CStr(stage) & ";" & TestPrismNumber(loads(1, stage)) & ";" & _
            TestPrismNumber(loads(2, stage)) & ";" & TestPrismNumber(loads(3, stage)) & ";" & _
            TestPrismNumber(loads(4, stage)) & ";" & TestPrismNumber(loads(5, stage)) & ";" & _
            TestPrismNumber(loads(6, stage)) & ";" & TestPrismNumber(loads(7, stage))
        load.Initialize loads(2, stage), loads(3, stage), loads(4, stage), 0#, 0#
        load.SetSpatialComponents loads(5, stage), loads(6, stage), loads(7, stage)
        load.LoadAboutPoint frame.OriginX, frame.OriginY, target
        If Not TestReachPrismTarget(solver, previous, target, minimumStep, stream, stage, field, actual) Then
            Set meta = solver.ResultMeta
            outcome = "StoppedUnbalanced"
            stream.WriteLine "STOP;" & CStr(stage) & ";" & CStr(meta.InternalStatus) & ";" & CStr(meta.ResultCode) & ";" & _
                TestPrismDiagnostic(meta.ResultComment) & ";" & TestPrismDiagnostic(meta.DiagnosticDetails)
            Exit For
        End If
        previous = target
        load.MoveResultantsToLoadPoint frame.OriginX, frame.OriginY, actual
        TestWritePrismStep stream, stage, loads(1, stage), operators, field, shearX, shearY, actual, response
        TestWritePrismPoints stream, stage, section.RebarCount, operators, field, center, response
        completed = stage
    Next stage
    stream.WriteLine "END;" & outcome & ";" & CStr(completed) & ";48"
    stream.Close
    TestReplayPrismHistory = "PRISM_TRACE: " & filePath & vbCrLf & "PRISM_OUTCOME: " & outcome & _
        "; accepted=" & CStr(completed) & "/48" & vbCrLf & "TOTAL_REINFORCED_PRISM_EXPORT: passed=1; failed=0"
    Exit Function
Failed:
    TestReplayPrismHistory = "RUNTIME ERROR: " & Err.Source & " " & Err.Description & vbCrLf & _
        "TOTAL_REINFORCED_PRISM_EXPORT: passed=0; failed=1"
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Делит только неравновесную часть текущего интервала. Поле и результат
' снимает на точном исходном винте, промежуточные поля изменяют историю.
Private Function TestReachPrismTarget(ByVal solver As CCoupledNewtonSolver, ByRef previous() As Double, _
        ByRef target() As Double, ByVal minimumStep As Double, ByVal stream As Object, ByVal stage As Long, _
        ByRef field() As Double, ByRef actual() As Double) As Boolean
    Dim position As Double, increment As Double, trial As Double, usedIncrement As Double, minimumTried As Double
    Dim accepted As Long, rejected As Long, i As Long, trialLoad(1 To 6) As Double
    increment = 1#: minimumTried = 1#
    Do While position < 1#
        trial = position + increment
        If trial > 1# Then trial = 1#
        usedIncrement = trial - position
        If usedIncrement < minimumTried Then minimumTried = usedIncrement
        For i = 1 To 6
            trialLoad(i) = target(i)
            If trial < 1# Then trialLoad(i) = previous(i) + trial * (target(i) - previous(i))
        Next i
        If solver.TryStep(trialLoad(1), trialLoad(2), trialLoad(3), trialLoad(4), trialLoad(5), trialLoad(6), _
                0.0000005, 0.00000002, 32) Then
            If trial = 1# Then solver.CopyTrial field, actual
            solver.AcceptStep
            position = trial
            accepted = accepted + 1
        Else
            rejected = rejected + 1
            If solver.ResultMeta.InternalStatus <> rsNumericalFailure Then Exit Do
            increment = usedIncrement * 0.5
            If increment < minimumStep Then Exit Do
        End If
    Loop
    stream.WriteLine "ROUTE;" & CStr(stage) & ";" & CStr(accepted) & ";" & CStr(rejected) & ";" & _
        TestPrismNumber(position) & ";" & TestPrismNumber(minimumTried)
    TestReachPrismTarget = (position = 1#)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает все исходные нагрузки до решателя с проверкой номера, формата и порядка.
Private Sub TestReadPrismLoads(ByVal inputPath As String, ByRef loads() As Double, ByRef signature As String)
    Dim fileSystem As Object: Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Dim source As Object: Set source = fileSystem.OpenTextFile(inputPath, 1, False, 0)
    Dim header As Variant: header = Split(source.ReadLine, ";")
    If UBound(header) <> 1 Then Err.Raise 5, "TestReadPrismLoads", "Неверная шапка контрольной истории."
    If header(0) <> "NDM_PRISM_INPUT_1" Or Len(header(1)) <> 64 Then Err.Raise 5, "TestReadPrismLoads", "Неверная версия или подпись истории."
    signature = header(1)
    If source.ReadLine <> "step;factor;N;Mx;My;T;Qx;Qy" Then Err.Raise 5, "TestReadPrismLoads", "Обязателен установленный порядок усилий."
    ReDim loads(1 To 7, 1 To STAGES)
    Dim stepIndex As Long, i As Long, values As Variant
    For stepIndex = 1 To STAGES
        values = Split(source.ReadLine, ";")
        If UBound(values) <> 7 Then Err.Raise 5, "TestReadPrismLoads", "Неполный вектор нагрузки " & CStr(stepIndex) & "."
        If values(0) <> CStr(stepIndex) Then Err.Raise 5, "TestReadPrismLoads", "Изменён порядок стадий."
        For i = 1 To 7
            If Not TryReadStrictDouble(values(i), loads(i, stepIndex)) Then Err.Raise 5, "TestReadPrismLoads", "Неверное число на стадии " & CStr(stepIndex) & "."
        Next i
        If Abs(loads(1, stepIndex) - stepIndex / STAGES) > 0.000000000001 Then _
            Err.Raise 5, "TestReadPrismLoads", "Неверный фактор стадии " & CStr(stepIndex) & "."
    Next stepIndex
    If Not source.AtEndOfStream Then Err.Raise 5, "TestReadPrismLoads", "Ожидалось ровно 48 стадий."
    source.Close
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит общий контур, четыре продольных стержня и точную прямоугольную ось хомута.
Private Function TestPrismSection(ByRef layout As CTransverseRebarLayout, ByRef steel As Object) As CSectionModel
    Dim query As New CSectionGeometryQuery, geometry As New CGeometryRegion, section As New CSectionModel
    geometry.Initialize query.RectangleRegion(0#, 0#, PRISM_WIDTH, PRISM_DEPTH)
    section.AddConcreteElement 0#, 0#, PRISM_WIDTH * PRISM_DEPTH, width:=PRISM_WIDTH, height:=PRISM_DEPTH
    geometry.BuildContours section.Contours
    Dim x As Variant, y As Variant
    For Each x In Array(-50#, 50#)
        For Each y In Array(-40#, 40#)
            section.AddRebarElement CDbl(x), CDbl(y), 12#, GEOM_PI * 36#, "L500"
        Next y
    Next x
    Dim path As New CGeometryPath
    path.Initialize "H240"
    path.AddLine "Bottom", -50#, -40#, 50#, -40#
    path.AddLine "Right", 50#, -40#, 50#, 40#
    path.AddLine "Top", 50#, 40#, -50#, 40#
    path.AddLine "Left", -50#, 40#, -50#, -40#
    path.Seal True
    Set layout = New CTransverseRebarLayout
    layout.AddGroup path, 8#, 100#, "H240"
    Set steel = CreateObject("Scripting.Dictionary")
    steel.CompareMode = vbTextCompare
    Dim longModel As New CCoupledSteelModel, hoopModel As New CCoupledSteelModel
    longModel.Initialize STEEL_E, 500#, 500#, 2000#, 2000#, REFERENCE_STRAIN_LIMIT, REFERENCE_STRAIN_LIMIT
    hoopModel.Initialize STEEL_E, 240#, 240#, 2000#, 2000#, REFERENCE_STRAIN_LIMIT, REFERENCE_STRAIN_LIMIT
    steel.Add "L500", longModel
    steel.Add "H240", hoopModel
    Set TestPrismSection = section
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подключает те же три независимые истории, что рабочее НДС НДМ.
Private Sub TestInitializePrismResponse(ByVal operators As CCoupledElasticState, ByVal steel As Object, _
        ByVal center As CCoupledConcreteResponse, ByVal response As CCoupledReinforcedResponse, ByVal solver As CCoupledNewtonSolver)
    Dim plusConcrete As New CCoupledConcreteResponse, minusConcrete As New CCoupledConcreteResponse
    Dim plusSide As New CCoupledReinforcedResponse, minusSide As New CCoupledReinforcedResponse
    Dim models As Variant, concrete As CCoupledConcreteResponse
    For Each models In Array(center, plusConcrete, minusConcrete)
        Set concrete = models
        concrete.Initialize operators.PointCount, CONCRETE_E, CONCRETE_NU, 18#, 30#, 3#, 0.002, 0.1, 10#, 100#
    Next models
    response.Initialize operators, center, steel
    plusSide.Initialize operators, plusConcrete, steel
    minusSide.Initialize operators, minusConcrete, steel
    solver.Initialize operators, response, CONCRETE_E, plusSide, minusSide
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет согласованный винт, обобщённые деформации и средний сдвиг.
Private Sub TestWritePrismStep(ByVal stream As Object, ByVal stage As Long, ByVal factor As Double, _
        ByVal operators As CCoupledElasticState, ByRef field() As Double, ByRef shearX() As Double, ByRef shearY() As Double, _
        ByRef actual() As Double, ByVal response As CCoupledReinforcedResponse)
    Dim eps0 As Double, kx As Double, ky As Double, twist As Double, gx As Double, gy As Double, i As Long
    operators.GeneralizedDeformation field, eps0, kx, ky, twist
    For i = 1 To operators.DegreeCount
        gx = gx + shearX(i) * field(i): gy = gy + shearY(i) * field(i)
    Next i
    stream.WriteLine "STEP;" & CStr(stage) & ";" & TestPrismNumber(factor) & TestPrismVector(actual) & _
        ";" & TestPrismNumber(eps0) & ";" & TestPrismNumber(kx) & ";" & TestPrismNumber(ky) & ";" & TestPrismNumber(twist) & _
        ";" & TestPrismNumber(gx) & ";" & TestPrismNumber(gy) & ";" & TestPrismNumber(response.AcceptedSteelUtilization) & _
        ";" & TestPrismNumber(response.AcceptedTransverseSteelUtilization)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пишет сырые принятые поля для независимого интегрирования бетонной поправки и стали.
Private Sub TestWritePrismPoints(ByVal stream As Object, ByVal stage As Long, ByVal longitudinalCount As Long, _
        ByVal operators As CCoupledElasticState, ByRef field() As Double, _
        ByVal concrete As CCoupledConcreteResponse, ByVal response As CCoupledReinforcedResponse)
    Dim point As Long, firstSteel As Long, firstHoop As Long, i As Long
    firstSteel = operators.ConcretePointCount + 1: firstHoop = firstSteel + longitudinalCount
    Dim x As Double, y As Double, weight As Double, tEnergy As Double, cEnergy As Double
    Dim stress(1 To 6) As Double, strain(1 To 6) As Double, direction(1 To 6) As Double
    Dim code As String, steelStress As Double, steelStrain As Double, fy As Double
    For point = 1 To operators.PointCount
        operators.Point point, x, y, weight
        operators.PointStrain point, field, strain
        concrete.CopyAcceptedStress point, stress
        concrete.AcceptedPointDamage point, tEnergy, cEnergy
        code = "C": steelStress = 0#: steelStrain = 0#: fy = 0#
        For i = 1 To 6: direction(i) = 0#: Next i
        If point >= firstSteel Then
            code = operators.SteelClassAt(point)
            If point >= firstHoop Then fy = 240# Else fy = 500#
            operators.SteelProjection point, direction
            steelStress = response.AcceptedSteelStress(point)
            For i = 1 To 6: steelStrain = steelStrain + direction(i) * strain(i): Next i
        End If
        stream.WriteLine "POINT;" & CStr(stage) & ";" & CStr(point) & ";" & code & ";" & TestPrismNumber(x) & ";" & _
            TestPrismNumber(y) & ";" & TestPrismNumber(weight) & TestPrismVector(strain) & TestPrismVector(stress) & _
            TestPrismVector(direction) & ";" & TestPrismNumber(steelStrain) & ";" & TestPrismNumber(steelStress) & _
            ";" & TestPrismNumber(fy) & ";" & TestPrismNumber(tEnergy) & ";" & TestPrismNumber(cEnergy)
    Next point
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Преобразует числовой массив в поля отчёта без округления до фиксированных десятичных знаков.
Private Function TestPrismVector(ByRef values() As Double) As String
    Dim i As Long
    For i = LBound(values) To UBound(values): TestPrismVector = TestPrismVector & ";" & TestPrismNumber(values(i)): Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число с инвариантным десятичным разделителем.
Private Function TestPrismNumber(ByVal value As Double) As String
    TestPrismNumber = Replace$(Trim$(Str$(value)), ",", ".")
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет полный текст причины в одном поле отчёта.
Private Function TestPrismDiagnostic(ByVal value As String) As String
    TestPrismDiagnostic = Replace(Replace(Replace(value, ";", "|"), vbCr, " "), vbLf, " ")
End Function
