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
        TestPrincipalBoundary folder, CStr(name), passed, failed, report
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
' Проверяет знаковые координаты, их градиенты, принятую историю и допустимую кратность двух других корней.
Private Sub TestPrincipalBoundary(ByVal folder As String, ByVal name As String, ByRef passed As Long, _
        ByRef failed As Long, ByRef report As String)
    On Error GoTo Failed
    Dim flow As CCoupledASDFlow: Set flow = modTestCoupledASDMaterial.TestReadFlow(folder, name)
    Dim concrete As New CCoupledConcreteResponse, response As ICoupledPointResponse
    concrete.InitializeASD 1, flow: Set response = concrete
    Dim base(1 To 6) As Double, requested(1 To 6) As Double, prior(1 To 6) As Double, oldEffective(1 To 6) As Double
    Dim zero(1 To 6) As Double, effective(1 To 6) As Double, nominal(1 To 6) As Double, plastic(1 To 6) As Double
    Dim gradient(1 To 6) As Double, actualGradient(1 To 6) As Double, unused(1 To 6) As Double, perturbed(1 To 6) As Double
    Dim tangent(1 To 6, 1 To 6) As Double, stress(1 To 6) As Double, t As Double, c As Double, dt As Double, dc As Double
    Dim value As Double, actual As Double, plus As Double, minus As Double, error As Double, i As Long, j As Long, rank As Long, epoch As Long
    Dim handle As Integer: handle = FreeFile
    Open folder & "/" & name & ".Boundary.csv" For Output As #handle
    base(1) = 0.0001: base(2) = -0.0002: base(3) = 0.00015
    base(4) = 0.00017: base(5) = -0.00013: base(6) = 0.00019
    For epoch = 0 To 1
        If epoch = 1 Then
            If Not flow.TryStep(base, zero, zero, 0#, 0#, t, c, effective, nominal, plastic, dt, dc) Then Err.Raise 5, , "Нет принятой истории знаковой координаты."
            If response.EvaluatePoint(1, base, stress, tangent, ctmStressOnly) <> sfcNone Then Err.Raise 5, , "Нет пробы принятой истории."
            response.AcceptTrial
            For i = 1 To 6: prior(i) = base(i): oldEffective(i) = effective(i): Next i
        End If
        For i = 1 To 6: requested(i) = base(i) + epoch * 0.00003 * i: Next i
        For rank = 1 To 3
            TestRecord flow.TryReadPrincipalBoundary(requested, prior, oldEffective, rank, value, gradient), _
                name & " boundary available " & epoch & "/" & rank, passed, failed, report
            Dim row As String
            row = "BOUNDARY;" & name & ";" & epoch & ";" & rank & TestVector(requested) & TestVector(prior) & _
                TestVector(oldEffective) & ";" & Trim$(Str$(value)) & TestVector(gradient)
            Print #handle, row
            error = 0#
            For i = 1 To 6
                For j = 1 To 6: perturbed(j) = requested(j): Next j
                perturbed(i) = requested(i) + 0.0000000001
                If Not flow.TryReadPrincipalBoundary(perturbed, prior, oldEffective, rank, plus, unused, False) Then Err.Raise 5, , "Нет положительной пробы границы."
                perturbed(i) = requested(i) - 0.0000000001
                If Not flow.TryReadPrincipalBoundary(perturbed, prior, oldEffective, rank, minus, unused, False) Then Err.Raise 5, , "Нет отрицательной пробы границы."
                error = GeomMax(error, Abs(gradient(i) - (plus - minus) / 0.0000000002))
            Next i
            TestRecord error < 0.01, name & " boundary finite difference " & epoch & "/" & rank, passed, failed, report
            If Not concrete.TryReadPrincipalBoundary(1, requested, rank, actual, actualGradient) Then Err.Raise 5, , "Нет координаты через оболочку."
            error = Abs(value - actual)
            For i = 1 To 6: error = GeomMax(error, Abs(gradient(i) - actualGradient(i))): Next i
            TestRecord error < 0.000000001, name & " boundary committed history " & epoch & "/" & rank, passed, failed, report
        Next rank
    Next epoch
    Close #handle: handle = 0
    Erase requested
    requested(1) = (3# - flow.Poisson) / flow.Modulus
    requested(2) = (1# - 3# * flow.Poisson) / flow.Modulus
    requested(3) = -4# * flow.Poisson / flow.Modulus
    TestRecord flow.TryReadPrincipalBoundary(requested, zero, zero, 3, value, gradient) And Abs(value) < 0.000000000001, _
        name & " zero principal has boundary gradient", passed, failed, report
    requested(1) = 3# / flow.Modulus: requested(2) = -3# * flow.Poisson / flow.Modulus: requested(3) = requested(2)
    TestRecord flow.TryReadPrincipalBoundary(requested, zero, zero, 1, value, gradient), _
        name & " repeated other roots permit simple principal", passed, failed, report
    TestRecord Not flow.TryReadPrincipalBoundary(requested, zero, zero, 3, value, gradient), _
        name & " repeated selected root rejects gradient", passed, failed, report
    TestRecord Not flow.TryReadPrincipalBoundary(requested, zero, zero, 0, value, gradient), _
        name & " invalid principal rejected", passed, failed, report
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    Err.Raise number, "modTestCoupledASDTangent.TestPrincipalBoundary", reason
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
