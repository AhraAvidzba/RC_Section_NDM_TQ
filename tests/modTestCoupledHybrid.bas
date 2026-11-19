Attribute VB_Name = "modTestCoupledHybrid"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: известные корни, отсутствие решения и бюджеты гибридного метода
' ==========================================================================
' Тесты выполняются в частной книге; рабочие листы не читаются и не меняются.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет полный алгоритм на аналитических системах и независимых ограничениях остановки.
Public Function RunTestCoupledHybrid() As String
    On Error GoTo Failed
    Dim report As String, passed As Long, failedCount As Long
    Dim adapter As New CHybridEquationsTest, equations As ICoupledEquationSystem, solver As New CCoupledHybridSolver
    Set equations = adapter
    Dim initial() As Double, solution() As Double, solved As Boolean, i As Long
    ReDim initial(1 To 2): adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck solved, "linear equilibrium", report, passed, failedCount
    TestCheck Abs(solution(1) - 1#) < 0.0000000001 And Abs(solution(2) + 2#) < 0.0000000001, "nonsymmetric known root", report, passed, failedCount
    TestCheck initial(1) = 0# And initial(2) = 0#, "input coordinates unchanged", report, passed, failedCount
    TestCheck solver.Steps > 0 And solver.AcceptedSteps > 0, "dogleg steps used", report, passed, failedCount
    adapter.InitializeTest "Rosenbrock": initial(1) = -1.2: initial(2) = 1#
    solved = solver.TrySolve(equations, initial, 300, 300, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck solved, "Rosenbrock equilibrium", report, passed, failedCount
    TestCheck Abs(solution(1) - 1#) < 0.00000001 And Abs(solution(2) - 1#) < 0.00000001, "Rosenbrock known root", report, passed, failedCount
    ReDim initial(1 To 9): adapter.InitializeTest "Tridiagonal"
    For i = 1 To 9: initial(i) = -1#: Next i
    solved = solver.TrySolve(equations, initial, 300, 300, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck solved And solver.BestResidual <= 0.0000000001, "MINPACK tridiagonal equilibrium", report, passed, failedCount
    For i = 1 To 9
        TestCheck solution(i) < 0# And solution(i) > -2#, "negative root coordinate " & CStr(i), report, passed, failedCount
    Next i
    report = report & "HYBRID_TRIDIAGONAL_TEST"
    For i = 1 To 9: report = report & ";" & Replace$(CStr(solution(i)), ",", "."): Next i
    report = report & vbCrLf
    ReDim initial(1 To 2): adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 1, 200, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck Not solved And solver.FunctionCalls = 1 And solver.JacobianCalls = 0 And solver.Steps = 0 And solver.StopCode = 2, _
        "one evaluation budget", report, passed, failedCount
    solved = solver.TrySolve(equations, initial, 200, 1, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck Not solved And solver.Steps = 1 And solver.StopCode = 2, "one correction budget", report, passed, failedCount
    ReDim initial(1 To 1): initial(1) = 0.5: adapter.InitializeTest "NoRoot"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck Not solved And solver.BestResidual >= 1#, "no real root is not success", report, passed, failedCount
    adapter.InitializeTest "NoRoot"
    solved = solver.TrySolve(equations, initial, 20, 20, 0.0000000001, 0.00000000001, 0.2, solution, maxRestarts:=3)
    TestCheck Not solved And solver.FunctionCalls <= 20 And solver.Steps <= 20, "shared restart budget", report, passed, failedCount
    TestCheck solver.FunctionCalls + solver.JacobianCalls = adapter.EvaluationCallsTest, "actual callback count across restarts", report, passed, failedCount
    adapter.InitializeTest "NoRoot"
    solved = solver.TrySolve(equations, initial, 1, 20, 0.0000000001, 0.00000000001, 0.2, solution, maxRestarts:=3)
    TestCheck Not solved And solver.FunctionCalls = 1 And solver.Restarts = 0, "restart cannot bypass exhausted budget", report, passed, failedCount
    adapter.InitializeTest "NoResponse"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck Not solved And solver.StopCode = -1 And Not solver.HasBestPoint, "absent response is not physical failure", report, passed, failedCount
    adapter.InitializeTest "Cancel"
    On Error Resume Next
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 0.2, solution)
    Dim cancelNumber As Long: cancelNumber = Err.Number
    Err.Clear
    On Error GoTo Failed
    TestCheck cancelNumber = 18, "cancellation propagates to caller", report, passed, failedCount
    adapter.InitializeTest "BadDimensions"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck Not solved And solver.StopCode = -1, "wrong dimensions rejected", report, passed, failedCount
    ReDim initial(1 To 2): adapter.InitializeTest "Degenerate"
    solved = solver.TrySolve(equations, initial, 300, 300, 0.0000000001, 0.00000000001, 0.2, solution)
    TestCheck solved And Abs(solution(1) + solution(2) - 3#) < 0.0000000001, "rank deficient system with existing roots", report, passed, failedCount
    ReDim initial(1 To 2): adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 0.2, solution, relativeDifferenceStep:=0.000001)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001 And Abs(solution(2) + 2#) < 0.00000001, _
        "global difference nonsymmetric known root", report, passed, failedCount
    TestCheck solver.FunctionCalls = adapter.EvaluationCallsTest, "all internal difference probes counted", report, passed, failedCount
    adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 4, 20, 0.0000000001, 0.00000000001, 0.2, solution, relativeDifferenceStep:=0.000001)
    TestCheck Not solved And solver.FunctionCalls <= 4 And solver.FunctionCalls = adapter.EvaluationCallsTest And solver.Steps = 0 And solver.StopCode = 2, _
        "partial difference matrix respects budget", report, passed, failedCount
    ReDim initial(1 To 2): initial(1) = -0.0012: initial(2) = 1000#: adapter.InitializeTest "ScaledRosenbrock"
    solved = solver.TrySolve(equations, initial, 300, 300, 0.0000000001, 0.00000000001, 0.2, solution, automaticScaling:=True)
    TestCheck solved And Abs(1000# * solution(1) - 1#) < 0.00000001 And Abs(0.001 * solution(2) - 1#) < 0.00000001, _
        "automatic scale anisotropic Rosenbrock known root", report, passed, failedCount
    adapter.InitializeTest "ScaledRosenbrock"
    solved = solver.TrySolve(equations, initial, 10000, 100, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, automaticScaling:=True, freshJacobian:=True)
    TestCheck solved And Abs(1000# * solution(1) - 1#) < 0.00000001 And Abs(0.001 * solution(2) - 1#) < 0.00000001, _
        "fresh full matrix anisotropic Rosenbrock known root", report, passed, failedCount
    TestCheck solver.FunctionCalls = adapter.EvaluationCallsTest And solver.Steps <= 100, _
        "automatic scale all full matrix probes and steps counted", report, passed, failedCount
    adapter.InitializeTest "ScaledRosenbrock"
    solved = solver.TrySolve(equations, initial, 4, 20, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, automaticScaling:=True, freshJacobian:=True)
    TestCheck Not solved And solver.FunctionCalls <= 4 And solver.FunctionCalls = adapter.EvaluationCallsTest And solver.Steps = 0, _
        "automatic scale partial matrix respects budget", report, passed, failedCount
    ReDim initial(1 To 1): initial(1) = 0.5: adapter.InitializeTest "NoRoot"
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, automaticScaling:=True, freshJacobian:=True)
    TestCheck Not solved And solver.BestResidual >= 1#, "automatic scale does not create nonexistent root", report, passed, failedCount
    Dim algebra As New CCoupledHybridFactorization, values(1 To 2) As Double
    values(1) = 1E+200: values(2) = 1E+200
    TestCheck Abs(algebra.VectorNorm(values, 1, 2) / 1E+200 - Sqr(2#)) < 0.000000000001, "large finite norm", report, passed, failedCount
    values(1) = 1E-200: values(2) = 1E-200
    TestCheck Abs(algebra.VectorNorm(values, 1, 2) / 1E-200 - Sqr(2#)) < 0.000000000001, "small finite norm", report, passed, failedCount
    RunTestCoupledHybrid = report & "TOTAL_COUPLED_HYBRID_TEST: passed=" & CStr(passed) & "; failed=" & CStr(failedCount)
    Exit Function
Failed:
    RunTestCoupledHybrid = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет каждое независимое ожидание, не прерывая остальные доступные проверки.
Private Sub TestCheck(ByVal condition As Boolean, ByVal name As String, ByRef report As String, ByRef passed As Long, ByRef failed As Long)
    If condition Then passed = passed + 1 Else failed = failed + 1
    report = report & IIf(condition, "PASS: ", "FAIL: ") & name & vbCrLf
End Sub
