Attribute VB_Name = "modTestCoupledASDMaterial"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: рабочий спектральный материал и его связь с расчётом сечения
' ==========================================================================
' Подготовленные таблицы и истории читает только этот тестовый слой.
' Напряжения сохраняются для независимого повторения в нативном OpenSees.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет две калибровки, пробную историю, диаграммы и штатный SolveTarget.
Public Function RunTestCoupledASDMaterial(ByVal folder As String) As String
    On Error GoTo Failed
    Dim passed As Long, failed As Long, report As String, name As Variant
    For Each name In Array("default", "ndm_uniaxial")
        TestMaterialHistory folder, CStr(name), passed, failed, report
    Next name
    TestTransverseControl folder, passed, failed, report
    TestSectionAnalysis folder, passed, failed, report
    RunTestCoupledASDMaterial = report & "TOTAL_COUPLED_ASD_MATERIAL: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledASDMaterial = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Загружает закреплённые подготовленные ветви без файлового доступа рабочего материала.
Public Function TestReadFlow(ByVal folder As String, ByVal name As String) As CCoupledASDFlow
    Dim handle As Integer, line As String, head As Variant, flow As New CCoupledASDFlow
    Dim tension As CCoupledConcreteBackbone, compression As CCoupledConcreteBackbone
    On Error GoTo Failed
    handle = FreeFile: Open folder & "/" & name & ".csv" For Input As #handle
    Line Input #handle, line: head = Split(line, ";")
    If UBound(head) <> 2 Then Err.Raise 5, , "Некорректная шапка тестового материала."
    If head(0) <> "ASDFlow1" Then Err.Raise 5, , "Неизвестный формат тестовых таблиц."
    Set tension = TestReadBranch(handle, "T"): Set compression = TestReadBranch(handle, "C")
    If Not EOF(handle) Then Err.Raise 5, , "Лишние строки тестовой калибровки."
    Close #handle: handle = 0
    flow.Initialize tension, compression, Val(head(1)), Val(head(2))
    Set TestReadFlow = flow
    Exit Function
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    Err.Raise number, "modTestCoupledASDMaterial.TestReadFlow", reason
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт в неизменяемую ветвь массивы и проверяет копирование исходных данных.
Private Function TestReadBranch(ByVal handle As Integer, ByVal expected As String) As CCoupledConcreteBackbone
    Dim line As String, head As Variant, row As Variant, count As Long, i As Long
    Line Input #handle, line: head = Split(line, ";")
    If UBound(head) <> 3 Then Err.Raise 5, , "Некорректная шапка ветви."
    If head(0) <> "BRANCH" Or head(1) <> expected Then Err.Raise 5, , "Неверный порядок ветвей."
    count = CLng(head(2))
    Dim strain() As Double, nominal() As Double, effective() As Double
    ReDim strain(1 To count): ReDim nominal(1 To count): ReDim effective(1 To count)
    For i = 1 To count
        Line Input #handle, line: row = Split(line, ";")
        If UBound(row) <> 2 Then Err.Raise 5, , "Неверная строка подготовленной ветви."
        strain(i) = Val(row(0)): nominal(i) = Val(row(1)): effective(i) = Val(row(2))
    Next i
    Dim branch As New CCoupledConcreteBackbone, copiedStrain() As Double, copiedStress() As Double
    branch.Initialize strain, nominal, effective, count, Val(head(3))
    strain(count) = -1#: nominal(count) = -1#: effective(count) = -1#
    branch.CopyCurve 1#, copiedStrain, copiedStress
    If copiedStrain(count) < 0# Or copiedStress(count) < 0# Then Err.Raise 5, , "Ветвь сохранила изменяемую ссылку на массив."
    Set TestReadBranch = branch
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет рабочую оболочку с чистым обновлением, включая отклонённые пробы и нулевой шаг.
Private Sub TestMaterialHistory(ByVal folder As String, ByVal name As String, ByRef passed As Long, _
        ByRef failed As Long, ByRef report As String)
    Dim flow As CCoupledASDFlow: Set flow = TestReadFlow(folder, name)
    Dim source As Integer, output As Integer, probeOutput As Integer, line As String, row As Variant, path As String
    Dim concrete As CCoupledConcreteResponse, response As ICoupledPointResponse
    Dim strain(1 To 6) As Double, priorStrain(1 To 6) As Double, priorEffective(1 To 6) As Double
    Dim effective(1 To 6) As Double, nominal(1 To 6) As Double, plastic(1 To 6) As Double
    Dim stress(1 To 6) As Double, probe(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim priorT As Double, priorC As Double, newT As Double, newC As Double, dt As Double, dc As Double
    Dim code As ESolverFailureCode, i As Long, error As Double, pointCount As Long
    On Error GoTo Failed
    source = FreeFile: Open folder & "/History.csv" For Input As #source
    output = FreeFile: Open folder & "/" & name & ".Native.csv" For Output As #output
    probeOutput = FreeFile: Open folder & "/" & name & ".Probes.csv" For Output As #probeOutput
    Do While Not EOF(source)
        Line Input #source, line: row = Split(line, ";")
        If UBound(row) <> 7 Then Err.Raise 5, , "Неверный размер тестового тензора."
        If path <> row(0) Then
            path = row(0): priorT = 0#: priorC = 0#
            For i = 1 To 6: priorStrain(i) = 0#: priorEffective(i) = 0#: Next i
            Set concrete = New CCoupledConcreteResponse: concrete.InitializeASD 1, flow
            Set response = concrete
            TestRecord concrete.LawIdentifier = "ASD-Closed-Cdf0", name & " actual law identifier", passed, failed, report
        End If
        For i = 1 To 6: strain(i) = Val(row(i + 1)): Next i
        If Not flow.TryStep(strain, priorStrain, priorEffective, priorT, priorC, newT, newC, _
                effective, nominal, plastic, dt, dc) Then Err.Raise 5, , "Не замкнулась контрольная история материала."
        code = response.EvaluatePoint(1, strain, stress, tangent, ctmStressOnly)
        If code <> sfcNone Then Err.Raise 5, , "Оболочка отказала на контрольной истории."
        error = 0#
        For i = 1 To 6: error = GeomMax(error, Abs(nominal(i) - stress(i))): Next i
        TestRecord error < 0.00000001, name & "/" & path & "/" & row(1) & " point response", passed, failed, report
        response.AcceptTrial
        Print #output, line & TestVector(stress)
        For i = 1 To 6: priorStrain(i) = strain(i): priorEffective(i) = effective(i): Next i
        priorT = newT: priorC = newC
        If CLng(row(1)) Mod 3 = 0 Then
            For i = 1 To 6: probe(i) = strain(i) + 0.00003 * CDbl(i): Next i
            code = response.EvaluatePoint(1, probe, stress, tangent, ctmFull, False)
            If code <> sfcNone Then Err.Raise 5, , "Не выполнена отклонённая материальная проба."
            Print #probeOutput, path & ";" & row(1) & TestVector(probe) & TestVector(stress)
            code = response.EvaluatePoint(1, strain, stress, tangent, ctmStressOnly)
            If code <> sfcNone Then Err.Raise 5, , "Не восстановлено принятое напряжение."
            error = 0#
            For i = 1 To 6: error = GeomMax(error, Abs(nominal(i) - stress(i))): Next i
            TestRecord error < 0.00000001, name & "/" & path & " rejected trial preserves history", passed, failed, report
        End If
        pointCount = pointCount + 1
    Loop
    Close #source: source = 0: Close #output: output = 0: Close #probeOutput: probeOutput = 0
    report = report & "ASD_NATIVE_POINTS|" & name & "|" & pointCount & vbCrLf
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If source <> 0 Then Close #source
    If output <> 0 Then Close #output
    If probeOutput <> 0 Then Close #probeOutput
    On Error GoTo 0
    Err.Raise number, "modTestCoupledASDMaterial.TestMaterialHistory", name & "/" & path & ": " & reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет средний сдвиг независимым интегралом деформаций, знаки, масштаб силы и исключение служебных координат.
Private Sub TestTransverseControl(ByVal folder As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim geometry As CGeometryRegion, mesh As New CFiberMeshBuilder, builder As New CSectionModelBuilder, section As CSectionModel
    Set geometry = TestFixtureBox(180#, 120#): mesh.BuildMesh geometry, 18#, 12#, cellGeometry:="Contour"
    Set section = builder.BuildFromGenerated(mesh, Nothing): geometry.BuildContours section.Contours
    Dim operators As New CCoupledElasticState, concrete As New CCoupledConcreteResponse, solver As New CCoupledNewtonSolver
    operators.PrepareModel section, 30000#, 0.2, 1, 1, 5
    concrete.InitializeASD operators.PointCount, TestReadFlow(folder, "default")
    solver.Initialize operators, concrete, 30000#
    Dim quadrature As New CCoupledSectionQuadrature: quadrature.Initialize section, 1, 1
    Dim x() As Double, y() As Double, area() As Double, indices() As Long, values() As Double, dx() As Double, dy() As Double
    quadrature.Prepare 5, x, y, area, indices, values, dx, dy
    If UBound(area) <> operators.ConcretePointCount Then Err.Raise 5, , "Не совпал порядок точек контрольной квадратуры."
    Dim previous() As Double, current() As Double, direction(1 To 6) As Double, projection() As Double, reference() As Double
    ReDim previous(1 To operators.DegreeCount): ReDim current(1 To operators.DegreeCount)
    Dim i As Long, point As Long, strain(1 To 6) As Double, meanX As Double, meanY As Double
    For i = 1 To solver.PhysicalDegreeCount: current(i) = 0.00001 * Sin(CDbl(i)): Next i
    For i = solver.PhysicalDegreeCount + 1 To operators.DegreeCount: current(i) = 1000000#: Next i
    For point = 1 To UBound(area)
        operators.PointStrain point, current, strain
        meanX = meanX + area(point) * strain(5) / 21600#: meanY = meanY + area(point) * strain(6) / 21600#
    Next point
    Dim choice As Long, expected As Double, control As Double, increment As Double, name As String, ok As Boolean, error As Double
    For choice = 1 To 5
        direction(1) = 9E+200: direction(4) = -7E+200
        Select Case choice
            Case 1: direction(5) = 1#: direction(6) = 0#: expected = meanX
            Case 2: direction(5) = 0#: direction(6) = 1#: expected = meanY
            Case 3: direction(5) = 3#: direction(6) = 4#: expected = 0.6 * meanX + 0.8 * meanY
            Case 4: direction(5) = -3#: direction(6) = 4#: expected = -0.6 * meanX + 0.8 * meanY
            Case 5: direction(5) = 3E+200: direction(6) = 4E+200: expected = 0.6 * meanX + 0.8 * meanY
        End Select
        ok = solver.SelectTransverseDeformationControl(previous, current, direction, projection, control, increment, name)
        TestRecord ok And Abs(expected) > 0#, "transverse control exists " & CStr(choice), passed, failed, report
        TestRecord Abs(control - Abs(expected)) < 0.000000000001, "transverse actual area mean " & CStr(choice), passed, failed, report
        TestRecord Abs(increment - Abs(expected)) < 0.000000000001, "transverse positive increment " & CStr(choice), passed, failed, report
        error = 0#
        For i = solver.PhysicalDegreeCount + 1 To operators.DegreeCount: error = GeomMax(error, Abs(projection(i))): Next i
        TestRecord error = 0#, "transverse control ignores gauge multipliers " & CStr(choice), passed, failed, report
        If choice = 3 Then reference = projection
        If choice = 5 Then
            error = 0#
            For i = 1 To UBound(projection): error = GeomMax(error, Abs(projection(i) - reference(i))): Next i
            TestRecord error = 0#, "transverse direction invariant to large force scale", passed, failed, report
        End If
    Next choice
    direction(5) = 3#: direction(6) = 4#
    ok = solver.SelectTransverseDeformationControl(current, previous, direction, projection, control, increment, name)
    error = 0#
    For i = 1 To UBound(projection): error = GeomMax(error, Abs(projection(i) + reference(i))): Next i
    TestRecord ok And control = 0# And increment > 0# And error = 0#, "transverse reverse traversal retains positive measure", passed, failed, report
    TestRecord Not solver.SelectTransverseDeformationControl(current, current, direction, projection, control, increment, name), _
        "transverse unchanged field has no progressing measure", passed, failed, report
    direction(5) = 0#: direction(6) = 0#
    TestRecord Not solver.SelectTransverseDeformationControl(previous, current, direction, projection, control, increment, name), _
        "transverse absent variable shear retains general control selection", passed, failed, report
    Dim gradient(1 To 6) As Double, side As Long, component As Long, value As Double, tensorWeight As Double, expectedSquare As Double
    For point = 1 To UBound(area)
        operators.PointStrain point, current, strain: operators.AxialGradientAtPoint point, current, gradient
        For side = -1 To 1
            For component = 1 To 6
                value = strain(component) + CDbl(side) * 0.001 * gradient(component)
                tensorWeight = 1#: If component > 3 Then tensorWeight = 0.5
                expectedSquare = expectedSquare + area(point) * tensorWeight * value * value / (3# * 21600#)
            Next component
        Next side
    Next point
    expected = Sqr(expectedSquare)
    ok = solver.SelectSpatialDirectionControl(previous, current, current, projection, control, increment)
    TestRecord ok, "spatial direction measure exists", passed, failed, report
    TestRecord Abs(control - expected) < 1E-12, "spatial measure independent three-station tensor norm", passed, failed, report
    TestRecord Abs(increment - expected) < 1E-12, "spatial measure positive increment", passed, failed, report
    error = 0#
    For i = solver.PhysicalDegreeCount + 1 To operators.DegreeCount: error = GeomMax(error, Abs(projection(i))): Next i
    TestRecord error = 0#, "spatial measure ignores gauge multipliers", passed, failed, report
    reference = projection
    Dim movement() As Double: movement = current
    For i = 1 To solver.PhysicalDegreeCount: movement(i) = -1E+200 * current(i): Next i
    ok = solver.SelectSpatialDirectionControl(previous, current, movement, projection, control, increment)
    error = 0#
    For i = 1 To UBound(projection): error = GeomMax(error, Abs(projection(i) - reference(i))): Next i
    TestRecord ok And error < 1E-12, "spatial measure independent of direction scale and sign", passed, failed, report
    ok = solver.SelectSpatialDirectionControl(current, previous, movement, projection, control, increment)
    TestRecord ok And control = 0# And Abs(increment - expected) < 1E-12, "spatial reversed traversal retains positive increment", passed, failed, report
    For i = 1 To solver.PhysicalDegreeCount: movement(i) = 0#: Next i
    TestRecord Not solver.SelectSpatialDirectionControl(previous, current, movement, projection, control, increment), _
        "spatial gauge-only direction is inapplicable", passed, failed, report
    TestRecord Not solver.SelectSpatialDirectionControl(current, current, current, projection, control, increment), _
        "spatial unchanged field has no forward measure", passed, failed, report
    TestRecord Not solver.SelectRejectedDeformationControl(previous, current, projection, control, increment, name), _
        "spatial recovery requires an actual rejected direction", passed, failed, report
    Dim fields As New Collection, trace As New CCoupledLoadTrace, continuation As New CCoupledDeformationContinuation
    Dim base(1 To 6) As Double, actual(1 To 6) As Double
    fields.Add previous: fields.Add current: trace.Add 1#, 0.8
    continuation.Trace solver, Nothing, base, direction, fields, 0.8, 0.9, actual, 2#, 0.001, 4, 0.0000005, 0.0000005, trace
    TestRecord continuation.ResultMeta.InternalStatus = rsNumericalFailure And _
        InStr(continuation.ResultMeta.ResultComment, "Количество записей нагрузочной траектории") > 0, _
        "continuation rejects mismatched load trace before evaluating fields", passed, failed, report
    Dim cleanFields As New Collection, cleanTrace As New CCoupledLoadTrace, snapshot() As Double, observedActual() As Double
    Dim factor As Variant, originalFirst As Double
    For Each factor In Array(0.5, 1#)
        If Not solver.TryStep(100# * factor, 0#, 0#, 0#, 0#, 0#, 0.0000005, 0.0000005, 4) Then _
            Err.Raise 5, , "Не подтверждено упругое поле контрольного наблюдателя."
        solver.CopyTrial snapshot, observedActual: solver.AcceptStep
        cleanFields.Add snapshot: cleanTrace.Add 1#, CDbl(factor)
        If factor = 0.5 Then originalFirst = snapshot(1)
    Next factor
    Erase direction: direction(1) = 100#
    Dim observer As New CCoupledProgressObserverTest
    observer.TestInitialize vbNullString, True
    continuation.Trace solver, Nothing, base, direction, cleanFields, 0.5, 1#, observedActual, 1.1, 0.001, 4, _
        0.0000005, 0.0000005, cleanTrace, observer
    snapshot = cleanFields(1)
    TestRecord observer.TestCallCount > 0, "continuation actually reports progress", passed, failed, report
    TestRecord cleanFields.Count = 2 And snapshot(1) = originalFirst, "progress observer cannot modify accepted fields", passed, failed, report
    TestRecord cleanTrace.Count = 2 And continuation.SavedTrace.Count = 2, "progress observer cannot modify load history", passed, failed, report
    TestRecord continuation.ResultMeta.ResultCode = rcSearchBoundReached, "progress observer cannot assign a physical limit", passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет общий путь шести усилий с разными классами стали и сохранением фактических диаграмм.
Private Sub TestSectionAnalysis(ByVal folder As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim flow As CCoupledASDFlow: Set flow = TestReadFlow(folder, "default")
    Dim geometry As CGeometryRegion, mesh As New CFiberMeshBuilder
    Set geometry = TestFixtureBox(180#, 120#)
    mesh.BuildMesh geometry, 18#, 12#, cellGeometry:="Contour"
    Dim builder As New CSectionModelBuilder, section As CSectionModel
    Set section = builder.BuildFromGenerated(mesh, Nothing): geometry.BuildContours section.Contours
    section.AddRebarElement 0#, -30#, 12#, GEOM_PI * 36#, "A400"
    section.AddRebarElement 0#, 30#, 12#, GEOM_PI * 36#, "A400"
    Dim layout As New CTransverseRebarLayout
    layout.AddContourGroup section, "TestASDHoop", "", 8#, 100#, 10#, "TestASDHoopGrade"
    Dim c As New CConcreteMaterialParameters, s As New CSteelMaterialParameters, provider As New CMaterialModelProvider
    c.Initialize flow.PeakStress(True), flow.PeakStress(False), flow.PeakStress(True), flow.PeakStress(False), 30000#, 30000#, rbMc2:=15#
    s.Initialize 350#, 400#, 350#, 400#, 200000#, 200000#: provider.TestInitializeFromParameters c, s
    Dim steel As New CCoupledSteelModel, grades As Object, supplementary As Object
    steel.Initialize 200000#, 240#, 240#, 500#, 500#, 0.025, 0.025
    Set grades = CreateObject("Scripting.Dictionary"): grades.Add "TestASDHoopGrade", steel
    Set supplementary = CreateObject("Scripting.Dictionary"): supplementary.Add "ULS(I)", grades
    Dim analysis As New CCoupledSectionAnalysis, spec As New CMaterialModelSpec
    Dim discretization As New CCoupledDiscretization: discretization.Initialize 2, 2, 5, True
    analysis.Initialize section, layout, provider, 0.2, 0#, 0#, 0#, 0#, 0#, "A400", 30, supplementary, flow, discretization
    spec.Initialize "ULS(I)", "ThreeLine", "UseDiagram", "TwoLine"
    Dim loads As New CSectionLoadState, state As CSectionStateResult, meta As CResultMeta
    Dim field() As Double, resultants() As Double
    loads.Initialize -21600#, 9000#, -4000#, 0#, 0#: loads.SetSpatialComponents 25000#, 12#, -8#
    Set meta = analysis.SolveTarget(loads, "Auto", spec, field, resultants, True, state)
    TestRecord meta.InternalStatus = rsSuccess Or meta.InternalStatus = rsSuccessWithWarning, _
        "section all six equilibrium " & meta.ResultComment & " " & meta.DiagnosticDetails, passed, failed, report
    If state Is Nothing Then Err.Raise 5, , "Нет принятого пространственного состояния."
    TestRecord Not (state.CoupledFields Is Nothing), "section output observer retains fields", passed, failed, report
    Dim curves As CCoupledMaterialSnapshot: Set curves = state.CoupledMaterialCurves
    TestRecord curves.ConcreteLawIdentifier = flow.LawIdentifier And curves.SeriesCount = 3, _
        "section actual concrete and two steel classes", passed, failed, report
    Dim strain() As Double, stress() As Double, positiveStrain() As Double, positiveStress() As Double
    curves.CopyCurve 1, strain, stress: flow.CopyCurve False, positiveStrain, positiveStress
    TestRecord strain(UBound(strain)) = positiveStrain(UBound(positiveStrain)) And _
        stress(UBound(stress)) = positiveStress(UBound(positiveStress)), "section prepared diagram tail retained", passed, failed, report
    TestRecord Abs(resultants(1) + 21600#) < 0.1 And Abs(resultants(4) - 25000#) < 0.1 And _
        Abs(resultants(6) + 8#) < 0.1, "section forces and sign convention", passed, failed, report
    c.Initialize 30#, 3#, 30#, 3#, 30000#, 30000#, rbMc2:=15#: provider.TestInitializeFromParameters c, s
    analysis.Initialize section, layout, provider, 0.2, 0#, 0#, 0#, 0#, 0#, "A400", 30, supplementary, flow
    Set meta = analysis.SolveTarget(loads, "Auto", spec, field, resultants)
    TestRecord InStr(meta.ResultComment, "Прочности профиля") > 0 And meta.InternalStatus <> rsSuccess, _
        "section rejects inconsistent material tables " & meta.ResultComment, passed, failed, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выводит шесть компонент с инвариантным десятичным разделителем.
Private Function TestVector(ByRef values() As Double) As String
    Dim i As Long
    For i = 1 To 6: TestVector = TestVector & ";" & Trim$(Str$(values(i))): Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет именованный результат, не меняя ожидания после неудачи.
Private Sub TestRecord(ByVal ok As Boolean, ByVal name As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    report = report & IIf(ok, "PASS ", "FAIL ") & name & vbCrLf
End Sub
