Attribute VB_Name = "modTestCoupledASDTangent"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: алгоритмическая касательная спектрального материала
' ==========================================================================
' Экспортирует все 36 компонент и исходную принятую историю для независимой
' проверки собственных векторов и правила цепочки вне VBA. Не меняет диаграммы.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет обе калибровки, возврат на разностный путь и точные ключи пробной истории.
Public Function RunTestCoupledASDTangent(ByVal folder As String) As String
    On Error GoTo Failed
    Dim passed As Long, failed As Long, report As String, name As Variant
    TestBackboneSlopes passed, failed, report
    For Each name In Array("default", "ndm_uniaxial")
        TestTangentHistory folder, CStr(name), passed, failed, report
    Next name
    RunTestCoupledASDTangent = report & "TOTAL_COUPLED_ASD_TANGENT: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestCoupledASDTangent = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет ординаты, наклоны, отрицательный хвост и отсечение численным нижним уровнем.
Private Sub TestBackboneSlopes(ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    Dim x(1 To 4) As Double, n(1 To 4) As Double, e(1 To 4) As Double
    x(2) = 0.001: x(3) = 0.003: x(4) = 0.005
    n(2) = 10#: n(3) = 8#: n(4) = 2#: e(2) = 12#: e(3) = 18#: e(4) = 20#
    Dim branch As New CCoupledConcreteBackbone: branch.Initialize x, n, e, 4, 0.01
    Dim coordinate As Variant, nominal As Double, effective As Double, ns As Double, es As Double
    Dim slopeN As Double, slopeE As Double, expectedN As Double, expectedE As Double
    For Each coordinate In Array(-0.001, 0#, 0.0005, 0.002, 0.004, 0.006)
        branch.Stresses CDbl(coordinate), nominal, effective
        branch.StressesAndSlopes CDbl(coordinate), ns, es, slopeN, slopeE
        Select Case CDbl(coordinate)
            Case Is <= 0#: expectedN = 0#: expectedE = 0#
            Case Is < 0.001: expectedN = 10000#: expectedE = 12000#
            Case Is < 0.003: expectedN = -1000#: expectedE = 3000#
            Case Is < 0.005: expectedN = -3000#: expectedE = 1000#
            Case Else: expectedN = 0#: expectedE = 1000#
        End Select
        TestRecord nominal = ns And effective = es And Abs(slopeN - expectedN) < 0.000001 And _
            Abs(slopeE - expectedE) < 0.000001, "backbone ordinates and slopes " & CStr(coordinate), passed, failed, report
    Next coordinate
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Передаёт каждой пробе только её принятую историю, сверяет оболочку и сохраняет полный тензор производных.
Private Sub TestTangentHistory(ByVal folder As String, ByVal name As String, ByRef passed As Long, _
        ByRef failed As Long, ByRef report As String)
    Dim flow As CCoupledASDFlow: Set flow = modTestCoupledASDMaterial.TestReadFlow(folder, name)
    Dim source As Integer, output As Integer, line As String, row As Variant, path As String
    Dim priorStrain(1 To 6) As Double, priorEffective(1 To 6) As Double, strain(1 To 6) As Double
    Dim nominal(1 To 6) As Double, effective(1 To 6) As Double, plastic(1 To 6) As Double
    Dim tangent(1 To 6, 1 To 6) As Double, actual(1 To 6, 1 To 6) As Double, stress(1 To 6) As Double
    Dim oldT As Double, oldC As Double, newT As Double, newC As Double, dt As Double, dc As Double
    Dim analytic As Boolean, available As Long, fallback As Long, i As Long, j As Long, error As Double
    Dim concrete As CCoupledConcreteResponse, response As ICoupledPointResponse, code As ESolverFailureCode
    On Error GoTo Failed
    source = FreeFile: Open folder & "/History.csv" For Input As #source
    output = FreeFile: Open folder & "/" & name & ".Tangent.csv" For Output As #output
    Do While Not EOF(source)
        Line Input #source, line: row = Split(line, ";")
        If path <> row(0) Then
            path = row(0): oldT = 0#: oldC = 0#: Erase priorStrain: Erase priorEffective
            Set concrete = New CCoupledConcreteResponse: concrete.InitializeASD 1, flow: Set response = concrete
        End If
        For i = 1 To 6: strain(i) = Val(row(i + 1)): Next i
        If Not flow.TryStep(strain, priorStrain, priorEffective, oldT, oldC, newT, newC, _
                effective, nominal, plastic, dt, dc) Then Err.Raise 5, , "Не замкнулась контрольная история касательной."
        analytic = flow.TryTangent(strain, priorStrain, priorEffective, oldT, oldC, nominal, effective, newT, newC, tangent)
        Print #output, "STATE;" & name & ";" & path & ";" & row(1) & ";" & CStr(analytic) & TestVector(strain) & _
            TestVector(priorStrain) & TestVector(priorEffective) & ";" & Trim$(Str$(oldT)) & ";" & Trim$(Str$(oldC)) & _
            TestVector(nominal) & TestVector(effective) & ";" & Trim$(Str$(newT)) & ";" & Trim$(Str$(newC))
        If analytic Then
            available = available + 1
            For i = 1 To 6
                line = "TANGENT;" & name & ";" & path & ";" & row(1) & ";" & CStr(i)
                For j = 1 To 6: line = line & ";" & Trim$(Str$(tangent(i, j))): Next j
                Print #output, line
            Next i
        Else
            fallback = fallback + 1
        End If
        code = response.EvaluatePoint(1, strain, stress, actual, ctmStressOnly)
        If code = sfcNone Then code = response.EvaluatePoint(1, strain, stress, actual, ctmAxialGradient)
        If code = sfcNone Then code = response.EvaluatePoint(1, strain, stress, actual, ctmFull)
        If code <> sfcNone Then Err.Raise 5, , "Не рассчитана оболочка при проверке касательной."
        error = 0#
        For i = 1 To 6
            error = GeomMax(error, Abs(stress(i) - nominal(i)))
            If analytic Then
                For j = 1 To 6: error = GeomMax(error, Abs(actual(i, j) - tangent(i, j))): Next j
            End If
        Next i
        TestRecord error < 0.00000001, name & "/" & path & "/" & row(1) & " tangent cache and unchanged stress", passed, failed, report
        response.AcceptTrial
        For i = 1 To 6: priorStrain(i) = strain(i): priorEffective(i) = effective(i): Next i
        oldT = newT: oldC = newC
    Loop
    Close #source: source = 0: Close #output: output = 0
    TestRecord available >= 25 And fallback >= 10, name & " analytic and fallback coverage " & available & "/" & fallback, passed, failed, report
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If source <> 0 Then Close #source
    If output <> 0 Then Close #output
    On Error GoTo 0
    Err.Raise number, "modTestCoupledASDTangent.TestTangentHistory", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выводит инженерный вектор с инвариантным десятичным разделителем без потери текущей точности Str$.
Private Function TestVector(ByRef values() As Double) As String
    Dim i As Long
    For i = 1 To 6: TestVector = TestVector & ";" & Trim$(Str$(values(i))): Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Добавляет именованный результат без изменения ожидаемых допусков.
Private Sub TestRecord(ByVal ok As Boolean, ByVal name As String, ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If ok Then passed = passed + 1 Else failed = failed + 1
    report = report & IIf(ok, "PASS ", "FAIL ") & name & vbCrLf
End Sub
