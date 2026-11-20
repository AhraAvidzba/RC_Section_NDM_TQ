Attribute VB_Name = "modTestCoupledEnvelopeControl"
Option Explicit

Private mAdjointRootTest As String, mAdjointCapturedTest As Boolean
Private mAdjointStageCapturedTest As Boolean

' ==========================================================================
' ДЛЯ ТЕСТОВ. Принятый путь H и максимум однородного бетонного сечения
' ==========================================================================
' Три независимые истории CDP проверяются без Excel-ввода. Одноосный предел
' сопоставляется с заданным fc*A; это проверка указанного материала и постановки.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет журнал H, независимость копий, читающую оценку lambda и поиск физического пика без арматуры.
Public Function RunTestCoupledEnvelopeControl(Optional ByVal evidenceRoot As String = "") As String
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    mAdjointRootTest = evidenceRoot: mAdjointCapturedTest = False: mAdjointStageCapturedTest = False
    Dim region As CGeometryRegion: Set region = TestFixtureBox(180#, 120#)
    Dim mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder, section As CSectionModel
    mesh.BuildMesh region, 9#, 6#, cellGeometry:="Contour"
    Set section = builder.BuildFromGenerated(mesh, Nothing)
    region.BuildContours section.Contours
    Dim operators As New CCoupledElasticState
    operators.PrepareModel section, 30000#, 0.2, 1, 1, 5
    Dim concrete(1 To 3) As CCoupledConcreteResponse, response(1 To 3) As CCoupledReinforcedResponse
    Dim models As Object: Set models = CreateObject("Scripting.Dictionary")
    Dim side As Long
    For side = 1 To 3
        Set concrete(side) = New CCoupledConcreteResponse
        concrete(side).Initialize operators.PointCount, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
        Set response(side) = New CCoupledReinforcedResponse
        response(side).Initialize operators, concrete(side), models
    Next side
    Dim solver As New CCoupledNewtonSolver
    solver.Initialize operators, response(2), 30000#, response(3), response(1)
    Dim previous() As Double, before As Double, after As Double
    TestExpect Not solver.TryAcceptedEnvelopeSegment(previous, before, after), "reset has no accepted H pair", passed, failed, report
    Dim offset(1 To 6) As Double, direction(1 To 6) As Double, control() As Double
    ReDim control(1 To operators.DegreeCount): control(1) = -1#
    direction(1) = -720000#
    Dim fields As New Collection, trace As New CCoupledLoadTrace, field() As Double, actual() As Double
    Dim strain As Variant, lambda As Double, previousLambda As Double, factor As Double, solved As Boolean
    For Each strain In Array(0.0008, 0.0016)
        solved = solver.TryDeformationStep(offset, direction, control, CDbl(strain), lambda, factor, 0.0000005, 0.0000005, 30)
        TestExpect solved, "prescribed compression equilibrium " & CStr(strain), passed, failed, report
        If Not solved Then GoTo Done
        solver.CopyTrial field, actual: solver.AcceptStep
        fields.Add field: trace.Add 1#, factor
        previousLambda = lambda: lambda = factor
    Next strain
    TestExpect solver.TryAcceptedEnvelopeSegment(previous, before, after), "accepted H pair available", passed, failed, report
    Dim firstH As Double, secondH As Double
    TestExpect solver.TryReadAcceptedEnvelopeAt(1, 2, firstH), "first H path entry available", passed, failed, report
    TestExpect solver.TryReadAcceptedEnvelopeAt(2, 2, secondH), "second H path entry available", passed, failed, report
    TestExpect firstH = before And secondH = after And secondH > firstH, "path agrees exactly with growing accepted pair", passed, failed, report
    TestExpect Not solver.TryReadAcceptedEnvelopeAt(0, 2, secondH), "zero H path index rejected", passed, failed, report
    TestExpect Not solver.TryReadAcceptedEnvelopeAt(1, 3, secondH), "foreign H path count rejected", passed, failed, report
    Dim original As Double: original = previous(1): previous(1) = original + 1#
    TestExpect solver.TryAcceptedEnvelopeSegment(previous, before, after), "accepted pair can be read repeatedly", passed, failed, report
    TestExpect previous(1) = original, "returned field copy does not mutate accepted pair", passed, failed, report
    TestExpect solver.ReplayAcceptedFields(fields, 1), "previous prefix can be replayed", passed, failed, report
    TestExpect solver.TryReadEnvelopeCoordinate(True, firstH), "previous prefix H available", passed, failed, report
    TestExpect firstH = before, "previous H equals independent native replay", passed, failed, report
    TestExpect solver.ReplayAcceptedFields(fields, 2), "complete prefix can be replayed", passed, failed, report
    TestExpect solver.TryReadEnvelopeCoordinate(True, secondH), "complete prefix H available", passed, failed, report
    TestExpect secondH = after, "complete H equals independent native replay", passed, failed, report
    Dim seed() As Double, projection() As Double, degree As Long, probes As Long, trialLambda As Double, uncertainty As Double
    seed = field
    For degree = 1 To solver.PhysicalDegreeCount: seed(degree) = field(degree) + 0.1 * (field(degree) - previous(degree)): Next degree
    Dim revision As Long: revision = solver.AcceptedRevision
    solved = solver.TryEnvelopeStep(offset, direction, previous, lambda, seed, lambda, after + 0.1 * (after - before), _
        trialLambda, projection, 0.0000005, 0.0000005, 30, probes, True)
    TestExpect solved, "growing H trial with retained control", passed, failed, report
    If solved Then
        TestExpect solver.TrialLambdaUncertainty(0.0000005, 0.0000005, uncertainty), "H adjoint uncertainty available", passed, failed, report
        TestExpect uncertainty > 0# And uncertainty < 1#, "H uncertainty is finite and distinguishes the load scale", passed, failed, report
        TestExpect solver.AcceptedRevision = revision, "H solution and adjoint preserve accepted revision", passed, failed, report
        solver.DiscardStep
        TestExpect solver.TryReadEnvelopeCoordinate(True, secondH), "accepted H available after discarded trial", passed, failed, report
        TestExpect secondH = after, "discarded H trial leaves accepted coordinate exact", passed, failed, report
    End If
    Dim branch As New CCoupledDeformationContinuation
    branch.Trace solver, response(2), offset, direction, fields, previousLambda, lambda, actual, 2#, 0.001, 30, 0.0000005, 0.0000005, trace
    report = report & "BRANCH_DETAILS: " & branch.ResultMeta.DiagnosticDetails & vbCrLf
    Dim peak As Double, accuracy As Double, design As Boolean
    TestExpect branch.CopyPeak(peak, accuracy, design), "three-station concrete peak certified: " & branch.ResultMeta.ResultComment, passed, failed, report
    If branch.CopyPeak(peak, accuracy, design) Then
        TestExpect Abs(peak - 30# * 21600# / 720000#) <= 0.002, "peak agrees with independent uniaxial fc*A", passed, failed, report
        TestExpect accuracy <= 0.001, "peak meets requested absolute lambda accuracy", passed, failed, report
    End If
Done:
    RunTestCoupledEnvelopeControl = report & "TOTAL_COUPLED_ENVELOPE_CONTROL: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledEnvelopeControl = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет первую неразрешённую сопряжённую систему для независимой проверки ранга; массивы владельца не меняет.
' Вызывается только из явно инструментированной частной компиляции решателя.
Public Sub TestCaptureHAdjoint(ByRef matrix() As Double, ByRef rightHandSide() As Double, ByRef field() As Double, _
        ByVal lambda As Double, ByVal modulus As Double)
    If mAdjointCapturedTest Or Len(mAdjointRootTest) = 0 Then Exit Sub
    Dim handle As Integer: handle = FreeFile
    Open mAdjointRootTest & "\Native\HAdjointFailure.csv" For Output As #handle
    Dim row As Long, column As Long, count As Long: count = UBound(rightHandSide)
    Print #handle, "DIM;" & CStr(count)
    Print #handle, "STATE;" & CStr(lambda) & ";" & CStr(modulus)
    For row = 1 To count
        Print #handle, "B;" & CStr(row) & ";" & CStr(IIf(row = count, 1#, 0#))
        Print #handle, "OBSERVED_B;" & CStr(row) & ";" & CStr(rightHandSide(row))
        For column = 1 To count: Print #handle, "A;" & CStr(column) & ";" & CStr(row) & ";" & CStr(matrix(column, row)): Next column
    Next row
    For row = 1 To UBound(field): Print #handle, "X;" & CStr(row) & ";" & CStr(field(row)): Next row
    Close #handle
    mAdjointCapturedTest = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Различает отказ материальной касательной и отказ линейного решения, не изменяя ни поле, ни статус рабочего расчёта.
Public Sub TestCaptureHAdjointStage(ByVal stage As String, ByVal failureCode As Long, ByVal point As Long, ByVal side As Long, _
        ByVal lambda As Double, ByRef field() As Double)
    If mAdjointStageCapturedTest Or Len(mAdjointRootTest) = 0 Then Exit Sub
    Dim handle As Integer: handle = FreeFile
    Open mAdjointRootTest & "\Native\HAdjointFailureStage.csv" For Output As #handle
    Print #handle, "STAGE;" & stage & ";" & CStr(failureCode) & ";" & CStr(point) & ";" & CStr(side) & ";" & CStr(lambda)
    Dim degree As Long
    For degree = 1 To UBound(field): Print #handle, "X;" & CStr(degree) & ";" & CStr(field(degree)): Next degree
    Close #handle
    mAdjointStageCapturedTest = True
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Учитывает проверку с конкретной причиной отказа без изменения ожидаемого результата.
Private Sub TestExpect(ByVal condition As Boolean, ByVal name As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If condition Then passed = passed + 1: report = report & "OK: " & name & vbCrLf Else failed = failed + 1: report = report & "FAIL: " & name & vbCrLf
End Sub
