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
    TestSectionAnalysis folder, passed, failed, report
    RunTestCoupledASDMaterial = report & "TOTAL_COUPLED_ASD_MATERIAL: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledASDMaterial = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Загружает закреплённые подготовленные ветви без файлового доступа рабочего материала.
Private Function TestReadFlow(ByVal folder As String, ByVal name As String) As CCoupledASDFlow
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
    analysis.Initialize section, layout, provider, 0.2, 0#, 0#, 0#, 0#, 0#, "A400", 30, supplementary, flow
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
