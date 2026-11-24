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
    Dim bracketBefore() As Double, bracketAfter() As Double, bracketScore As Double, bracketColumn As Long
    Dim plain() As Double, plainCalls As Long, plainSteps As Long, plainSolved As Boolean
    ReDim initial(1 To 2): initial(1) = 0#: initial(2) = 0#: adapter.InitializeTest "Hinge"
    plainSolved = solver.TrySolve(equations, initial, 500, 50, 0.0000000001, 0.00000000001, 0.2, plain, relativeDifferenceStep:=0.000001)
    plainCalls = solver.FunctionCalls: plainSteps = solver.Steps
    adapter.InitializeTest "Hinge"
    solved = solver.TrySolve(equations, initial, 500, 50, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, captureDifferenceBracket:=True)
    TestCheck solved = plainSolved And solver.FunctionCalls = plainCalls And solver.Steps = plainSteps, _
        "difference observer preserves convergence and all evaluation budgets", report, passed, failedCount
    TestCheck solution(1) = plain(1) And solution(2) = plain(2), "difference observer preserves exact solution", report, passed, failedCount
    TestCheck solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn), _
        "fully evaluated nonsmooth stencil retained", report, passed, failedCount
    TestCheck bracketColumn = 2 And bracketBefore(2) < 0# And bracketAfter(2) > 0# And bracketBefore(1) = bracketAfter(1), _
        "observed bracket crosses actual hinge in selected column", report, passed, failedCount
    bracketBefore(1) = 123#
    TestCheck solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn) And bracketBefore(1) <> 123#, _
        "returned bracket is an independent copy", report, passed, failedCount
    adapter.InitializeTest "Hinge"
    solved = solver.TrySolve(equations, initial, 2, 50, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, captureDifferenceBracket:=True)
    TestCheck Not solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn) And solver.FunctionCalls <= 2, _
        "incomplete matrix exposes no stale bracket", report, passed, failedCount
    adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 500, 50, 0.0000000001, 0.00000000001, 0.2, solution, relativeDifferenceStep:=0.000001)
    TestCheck Not solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn), _
        "disabled observer clears prior solve bracket", report, passed, failedCount
    Dim bracketObserver As New CHybridDifferenceObserverTest, observer As ICoupledDifferenceObserver
    Set observer = bracketObserver
    initial(1) = 0.5: initial(2) = 0#: adapter.InitializeTest "HingeCurve"
    plainSolved = solver.TrySolve(equations, initial, 1000, 50, 0.0000000001, 0.00000000001, 0.2, plain, _
        relativeDifferenceStep:=0.000001, freshJacobian:=True)
    plainCalls = solver.FunctionCalls: plainSteps = solver.Steps
    bracketObserver.InitializeTest True, True: adapter.InitializeTest "HingeCurve"
    solved = solver.TrySolve(equations, initial, 1000, 50, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, freshJacobian:=True, captureDifferenceBracket:=True, differenceObserver:=observer)
    TestCheck solved And plainSolved And Abs(solution(1) - 1#) < 0.000000001 And solution(2) = 0#, _
        "confirmed observer retains known nonlinear root", report, passed, failedCount
    TestCheck solution(1) = plain(1) And solution(2) = plain(2) And solver.FunctionCalls = plainCalls And solver.Steps = plainSteps, _
        "confirmation preserves exact solution and original function budget", report, passed, failedCount
    TestCheck bracketObserver.AcceptedTest = 1 And bracketObserver.CallsTest > 1 And solver.JacobianCalls > 1, _
        "confirmation rejects later complete matrix probes", report, passed, failedCount
    TestCheck solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn), _
        "confirmed bracket survives later matrices without accepted crossings", report, passed, failedCount
    TestCheck bracketColumn = 2 And bracketObserver.MatchesFirstTest(bracketBefore, bracketAfter), _
        "retained confirmed bracket equals first observed Double coordinates", report, passed, failedCount
    bracketObserver.InitializeTest False: adapter.InitializeTest "HingeCurve"
    solved = solver.TrySolve(equations, initial, 1000, 50, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, freshJacobian:=True, captureDifferenceBracket:=True, differenceObserver:=observer)
    TestCheck solved And bracketObserver.CallsTest > 0 And bracketObserver.AcceptedTest = 0 And _
        Not solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn), _
        "rejected crossings expose no previous solve bracket", report, passed, failedCount
    TestCheck solution(1) = plain(1) And solution(2) = plain(2) And solver.FunctionCalls = plainCalls And solver.Steps = plainSteps, _
        "rejecting observer preserves exact solution and budgets", report, passed, failedCount
    bracketObserver.InitializeTest True: adapter.InitializeTest "HingeCurve"
    solved = solver.TrySolve(equations, initial, 2, 50, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, captureDifferenceBracket:=True, differenceObserver:=observer)
    TestCheck Not solved And bracketObserver.CallsTest = 0 And _
        Not solver.CopyDifferenceBracket(bracketBefore, bracketAfter, bracketScore, bracketColumn), _
        "partial matrix never calls confirmation observer", report, passed, failedCount
    ReDim initial(1 To 2): adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 200, 20, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, completeNewtonLineSearch:=True)
    TestCheck solved And Abs(solution(1) - 1#) < 0.000000001 And Abs(solution(2) + 2#) < 0.000000001, _
        "complete Newton nonsymmetric known root", report, passed, failedCount
    TestCheck solver.FunctionCalls = adapter.EvaluationCallsTest And solver.AcceptedSteps > 0, _
        "complete Newton actual full function calls counted", report, passed, failedCount
    initial(1) = -1.2: initial(2) = 1#: adapter.InitializeTest "Rosenbrock"
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, completeNewtonLineSearch:=True)
    TestCheck solved And Abs(solution(1) - 1#) < 0.000000001 And Abs(solution(2) - 1#) < 0.000000001, _
        "complete Newton Armijo Rosenbrock root", report, passed, failedCount
    TestCheck solver.FunctionCalls > 1 + 4 * solver.JacobianCalls + solver.Steps And solver.FunctionCalls = adapter.EvaluationCallsTest, _
        "complete Newton backtracking uses counted function probes", report, passed, failedCount
    adapter.InitializeTest "Rosenbrock"
    solved = solver.TrySolve(equations, initial, 4, 100, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.FunctionCalls <= 4 And solver.FunctionCalls = adapter.EvaluationCallsTest And solver.Steps = 0, _
        "complete Newton partial full matrix preserves function budget", report, passed, failedCount
    adapter.InitializeTest "Rosenbrock"
    solved = solver.TrySolve(equations, initial, 1000, 1, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.Steps = 1 And solver.StopCode = 2, _
        "complete Newton all line probes share one correction budget", report, passed, failedCount
    ReDim initial(1 To 1): initial(1) = 0.5: adapter.InitializeTest "NoRoot"
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.BestResidual >= 1# And solver.FunctionCalls <= 1000 And solver.Steps <= 100, _
        "complete Newton does not create nonexistent real root", report, passed, failedCount
    adapter.InitializeTest "Cancel"
    On Error Resume Next
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, _
        relativeDifferenceStep:=0.000001, completeNewtonLineSearch:=True)
    cancelNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck cancelNumber = 18, "complete Newton propagates cancellation", report, passed, failedCount
    ReDim initial(1 To 2): adapter.InitializeTest "Linear"
    solved = solver.TrySolve(equations, initial, 200, 20, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    TestCheck solved And Abs(solution(1) - 1#) < 0.000000001 And Abs(solution(2) + 2#) < 0.000000001, _
        "supplied matrix Newton nonsymmetric known root", report, passed, failedCount
    TestCheck solver.FunctionCalls + solver.JacobianCalls = adapter.EvaluationCallsTest, _
        "supplied matrix Newton actual function and matrix calls counted", report, passed, failedCount
    initial(1) = -1.2: initial(2) = 1#: adapter.InitializeTest "Rosenbrock"
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    TestCheck solved And Abs(solution(1) - 1#) < 0.000000001 And Abs(solution(2) - 1#) < 0.000000001, _
        "supplied matrix Newton Armijo Rosenbrock root", report, passed, failedCount
    TestCheck solver.FunctionCalls > 1 + solver.Steps And solver.FunctionCalls + solver.JacobianCalls = adapter.EvaluationCallsTest, _
        "supplied matrix Newton backtracking and callback totals", report, passed, failedCount
    adapter.InitializeTest "Rosenbrock"
    solved = solver.TrySolve(equations, initial, 1, 100, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.FunctionCalls = 1 And solver.JacobianCalls = 0 And solver.Steps = 0 And solver.StopCode = 2, _
        "supplied matrix Newton exhausted function budget prevents matrix request", report, passed, failedCount
    adapter.InitializeTest "Rosenbrock"
    solved = solver.TrySolve(equations, initial, 1000, 1, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.Steps = 1 And solver.StopCode = 2, _
        "supplied matrix Newton all line probes share correction budget", report, passed, failedCount
    ReDim initial(1 To 1): initial(1) = 0.5: adapter.InitializeTest "NoRoot"
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.BestResidual >= 1# And solver.FunctionCalls <= 1000 And solver.Steps <= 100, _
        "supplied matrix Newton cannot create nonexistent root", report, passed, failedCount
    adapter.InitializeTest "Cancel"
    On Error Resume Next
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    cancelNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck cancelNumber = 18, "supplied matrix Newton propagates cancellation", report, passed, failedCount
    adapter.InitializeTest "BadDimensions"
    solved = solver.TrySolve(equations, initial, 1000, 100, 0.0000000001, 0.00000000001, 0.2, solution, completeNewtonLineSearch:=True)
    TestCheck Not solved And solver.StopCode = -1, "supplied matrix Newton rejects invalid dimensions", report, passed, failedCount
    Dim difference() As Double
    ReDim initial(1 To 2): initial(1) = 0.4: initial(2) = -0.7: adapter.InitializeTest "Linear"
    solved = solver.TryLinearizeByDifferences(equations, initial, 0.000001, 5, difference)
    TestCheck solved And Abs(difference(1, 1) - 3#) < 0.000001 And Abs(difference(2, 1) - 2#) < 0.000001 And _
        Abs(difference(1, 2) - 1#) < 0.000001 And Abs(difference(2, 2) - 4#) < 0.000001, _
        "read-only full difference linearization matches nonsymmetric analytic matrix", report, passed, failedCount
    TestCheck solver.FunctionCalls = 5 And solver.JacobianCalls = 1 And solver.Steps = 0 And adapter.EvaluationCallsTest = 5, _
        "difference linearization counts every actual callback without nonlinear corrections", report, passed, failedCount
    TestCheck initial(1) = 0.4 And initial(2) = -0.7, "difference linearization preserves input coordinates", report, passed, failedCount
    adapter.InitializeTest "Linear"
    solved = solver.TryLinearizeByDifferences(equations, initial, 0.000001, 4, difference)
    TestCheck Not solved And solver.FunctionCalls = 0 And solver.StopCode = 2, "insufficient complete matrix budget prevents callbacks", report, passed, failedCount
    adapter.InitializeTest "BadDimensions"
    solved = solver.TryLinearizeByDifferences(equations, initial, 0.000001, 5, difference)
    TestCheck Not solved And solver.FunctionCalls = 1 And solver.StopCode = -1, "difference linearization rejects invalid residual dimensions", report, passed, failedCount
    adapter.InitializeTest "NoResponse"
    solved = solver.TryLinearizeByDifferences(equations, initial, 0.000001, 5, difference)
    TestCheck Not solved And solver.FunctionCalls = 1 And solver.StopCode = -1, "difference linearization rejects missing response", report, passed, failedCount
    adapter.InitializeTest "Cancel"
    On Error Resume Next
    solved = solver.TryLinearizeByDifferences(equations, initial, 0.000001, 5, difference)
    cancelNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck cancelNumber = 18, "difference linearization propagates cancellation", report, passed, failedCount
    ReDim initial(1 To 1): initial(1) = 9#: adapter.InitializeTest "DomainSqrt"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001, "existing root after unavailable dogleg trial", report, passed, failedCount
    TestCheck adapter.UnavailableCallsTest > 0 And initial(1) = 9#, "unavailable trial exercised without changing input", report, passed, failedCount
    TestCheck solver.FunctionCalls + solver.JacobianCalls = adapter.EvaluationCallsTest, "unavailable trials counted", report, passed, failedCount
    adapter.InitializeTest "DomainSqrt"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution, relativeDifferenceStep:=0.000001)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001, "difference Jacobian recovers unavailable dogleg trial", report, passed, failedCount
    TestCheck solver.FunctionCalls = adapter.EvaluationCallsTest And adapter.UnavailableCallsTest > 0, "difference and unavailable callbacks counted", report, passed, failedCount
    adapter.InitializeTest "DomainSqrt"
    solved = solver.TrySolve(equations, initial, 200, 1, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck Not solved And solver.Steps = 1 And solver.FunctionCalls = 2, "unavailable trial cannot bypass correction budget", report, passed, failedCount
    adapter.InitializeTest "DomainSqrt"
    solved = solver.TrySolve(equations, initial, 2, 200, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck Not solved And solver.FunctionCalls = 2 And solver.Steps = 1, "unavailable trial cannot bypass evaluation budget", report, passed, failedCount
    adapter.InitializeTest "TrialDimensions"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck Not solved And solver.StopCode = -1 And solver.Steps = 1 And solver.FunctionCalls = 2, _
        "wrong trial dimensions remain a contract error", report, passed, failedCount
    adapter.InitializeTest "OverflowTrial": initial(1) = 9#
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001, "trial overflow contracts trust region", report, passed, failedCount
    adapter.InitializeTest "DivisionTrial"
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001, "trial division by zero contracts trust region", report, passed, failedCount
    adapter.InitializeTest "CancelTrial"
    On Error Resume Next
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution)
    cancelNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck cancelNumber = 18, "trial cancellation propagates", report, passed, failedCount
    adapter.InitializeTest "OverflowTrial": initial(1) = -1#
    solved = solver.TrySolve(equations, initial, 200, 200, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck Not solved And solver.StopCode = -2 And Not solver.HasBestPoint, "initial arithmetic failure remains fatal", report, passed, failedCount
    adapter.InitializeTest "OverflowTrial": initial(1) = 9#
    solved = solver.TrySolve(equations, initial, 200, 1, 0.0000000001, 0.00000000001, 20#, solution)
    TestCheck Not solved And solver.Steps = 1 And solver.FunctionCalls = 2, "arithmetic rejection shares correction budget", report, passed, failedCount
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


' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет полную регуляризованную поправку, существующие корни, область определения и общие бюджеты.
Public Function RunTestCoupledLeastSquaresTrust() As String
    On Error GoTo Failed
    Dim report As String, passed As Long, failedCount As Long
    Dim adapter As New CHybridEquationsTest, equations As ICoupledEquationSystem, solver As New CCoupledHybridSolver
    Set equations = adapter
    Dim initial() As Double, solution() As Double, solved As Boolean, mode As Variant, cancelNumber As Long
    ReDim initial(1 To 2): adapter.InitializeTest "Linear"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 200, 100, 0.0000000001, solution, 0#)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001 And Abs(solution(2) + 2#) < 0.00000001, "LM nonsymmetric known root", report, passed, failedCount
    TestCheck initial(1) = 0# And initial(2) = 0#, "LM input unchanged", report, passed, failedCount
    TestCheck solver.FunctionCalls = adapter.EvaluationCallsTest, "LM analytic callbacks counted", report, passed, failedCount
    adapter.InitializeTest "Degenerate"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 300, 100, 0.0000000001, solution, 0#)
    TestCheck solved And Abs(solution(1) + solution(2) - 3#) < 0.00000001, "LM rank deficient existing roots", report, passed, failedCount
    ReDim initial(1 To 1): initial(1) = 0.5: adapter.InitializeTest "NoRoot"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 200, 100, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.BestResidual >= 1# And solver.Steps <= 100 And solver.FunctionCalls <= 200, "LM no real root remains failure", report, passed, failedCount
    For Each mode In Array("LMUnavailable", "LMOverflow", "LMDivision")
        adapter.InitializeTest CStr(mode): initial(1) = 9#
        solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution, 0#)
        TestCheck solved And Abs(solution(1) - 1#) < 0.00000001, "LM existing root after " & CStr(mode), report, passed, failedCount
        TestCheck adapter.UnavailableCallsTest > 0 And solver.FunctionCalls = adapter.EvaluationCallsTest And initial(1) = 9#, "LM unavailable callbacks counted " & CStr(mode), report, passed, failedCount
    Next mode
    adapter.InitializeTest "LMUnavailable": initial(1) = 9#
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution)
    TestCheck solved And Abs(solution(1) - 1#) < 0.00000001 And adapter.UnavailableCallsTest > 0, "LM difference Jacobian domain recovery", report, passed, failedCount
    TestCheck solver.FunctionCalls = adapter.EvaluationCallsTest, "LM difference callbacks counted", report, passed, failedCount
    adapter.InitializeTest "LMOverflow"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 1, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.Steps = 1 And solver.FunctionCalls = 3 And adapter.UnavailableCallsTest = 1, "LM invalid trial shares step budget", report, passed, failedCount
    adapter.InitializeTest "LMUnavailable"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 3, 100, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.FunctionCalls = 3 And solver.Steps = 1 And solver.StopCode = 2, "LM invalid trial shares function budget", report, passed, failedCount
    adapter.InitializeTest "LMWrongDimensions"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.StopCode = -1 And solver.Steps = 1 And solver.FunctionCalls = 3, "LM wrong trial dimensions fatal", report, passed, failedCount
    adapter.InitializeTest "LMCancel"
    On Error Resume Next
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution, 0#)
    cancelNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck cancelNumber = 18 And adapter.UnavailableCallsTest = 1, "LM trial cancellation propagates", report, passed, failedCount
    adapter.InitializeTest "LMOverflow": initial(1) = 0#
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.StopCode = -2 And Not solver.HasBestPoint And solver.FunctionCalls = 1, "LM initial arithmetic fatal", report, passed, failedCount
    adapter.InitializeTest "NoResponse": initial(1) = 9#
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.StopCode = -1 And Not solver.HasBestPoint, "LM initial unavailable fatal", report, passed, failedCount
    adapter.InitializeTest "BadDimensions"
    solved = solver.TrySolveLeastSquaresTrust(equations, initial, 400, 100, 0.0000000001, solution, 0#)
    TestCheck Not solved And solver.StopCode = -1 And solver.FunctionCalls = 1, "LM initial dimensions rejected", report, passed, failedCount
    RunTestCoupledLeastSquaresTrust = report & "TOTAL_COUPLED_LM_TEST: passed=" & CStr(passed) & "; failed=" & CStr(failedCount)
    Exit Function
Failed:
    RunTestCoupledLeastSquaresTrust = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
