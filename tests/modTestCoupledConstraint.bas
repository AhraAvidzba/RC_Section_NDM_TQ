Attribute VB_Name = "modTestCoupledConstraint"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: нелинейное исключение, производная и исходная силовая система
' ==========================================================================

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет известный корень, правило цепочки, неизменность входа, недопустимые пробы и использование общего гибридного решателя.
Public Function RunTestCoupledConstraint() As String
    On Error GoTo Failed
    Dim report As String, passed As Long, failedCount As Long
    Dim model As New CConstraintEquationsTest, provider As ICoupledEquationSystem, constraint As ICoupledScalarConstraint
    model.InitializeTest "Circle": Set provider = model: Set constraint = model
    Dim origin(1 To 3) As Double: origin(2) = 0.55: origin(3) = 0.55
    Dim adapter As New CCoupledConstraintEquations, equations As ICoupledEquationSystem
    adapter.Initialize provider, constraint, origin, 0.0000000000001, 16
    Set equations = adapter
    TestCheck adapter.Pivot = 1, "pivot selected from actual nonzero derivative", report, passed, failedCount
    Dim x() As Double, mapped() As Double, residual() As Double, matrix() As Double
    ReDim x(1 To 3): x(1) = 9#: x(2) = 0.6: x(3) = 0.6
    TestCheck equations.TryEvaluate(x, True, residual, matrix), "nonlinear projected response exists", report, passed, failedCount
    TestCheck Abs(residual(1)) < 0.000000000001 And residual(2) = 0# And residual(3) = 9#, _
        "original force rows and dummy coordinate are distinct", report, passed, failedCount
    TestCheck x(1) = 9# And x(2) = 0.6 And x(3) = 0.6, "input coordinates unchanged", report, passed, failedCount
    TestCheck Abs(matrix(2, 1) - 0.25) < 0.000000000001 And matrix(1, 1) = 0# And matrix(1, 3) = 1# And _
        matrix(2, 2) = -1# And matrix(3, 2) = 1# And matrix(3, 3) = 0#, "column-first chain rule and load column", report, passed, failedCount
    TestCheck adapter.TryMap(x, mapped) And Abs(mapped(1) + 0.2) < 0.000000000001, "actual field is recovered by mapping", report, passed, failedCount
    Dim positive() As Double, negative() As Double, rp() As Double, rm() As Double, unused() As Double
    Dim row As Long, column As Long, maximum As Double, difference As Double
    For column = 1 To 3
        positive = x: negative = x: positive(column) = positive(column) + 0.000001: negative(column) = negative(column) - 0.000001
        If Not equations.TryEvaluate(positive, False, rp, unused) Or Not equations.TryEvaluate(negative, False, rm, unused) Then _
            Err.Raise 5, , "Нет аналитической разностной пробы."
        For row = 1 To 3
            difference = Abs((rp(row) - rm(row)) / 0.000002 - matrix(column, row))
            If difference > maximum Then maximum = difference
        Next row
    Next column
    TestCheck maximum < 0.00000001, "all projected derivatives match central differences", report, passed, failedCount
    Dim solver As New CCoupledHybridSolver, solution() As Double, before As Long, solved As Boolean
    before = model.EquationCallsTest
    solved = solver.TrySolve(equations, origin, 500, 100, 0.00000000001, 0.00000000001, 0.2, solution, relativeDifferenceStep:=0.000001)
    TestCheck solved And Abs(solution(1)) < 0.00000000001 And Abs(solution(2) - 0.6) < 0.000000001 And _
        Abs(solution(3) - 0.6) < 0.000000001, "existing hybrid solver finds known constrained root", report, passed, failedCount
    TestCheck model.EquationCallsTest - before = solver.FunctionCalls, "full residual budget counts actual provider calls", report, passed, failedCount
    TestCheck adapter.TryMap(solution, mapped) And Abs(mapped(1) + 0.2) < 0.000000001, _
        "solver raw solution is mapped before physical use", report, passed, failedCount
    Dim absent As New CCoupledConstraintEquations, noRoot As New CConstraintEquationsTest
    noRoot.InitializeTest "NoRoot": Set provider = noRoot: Set constraint = noRoot
    absent.Initialize provider, constraint, origin, 0.0000000000001, 16
    Set equations = absent
    TestCheck Not equations.TryEvaluate(origin, False, residual, unused) And noRoot.EquationCallsTest = 0 And absent.ProjectionFailures = 1, _
        "missing scalar root does not evaluate or accept force state", report, passed, failedCount
    TestCheck origin(1) = 0# And origin(2) = 0.55 And origin(3) = 0.55, "rejected mapping preserves input", report, passed, failedCount
    RunTestCoupledConstraint = report & "TOTAL_COUPLED_CONSTRAINT_TEST: passed=" & CStr(passed) & "; failed=" & CStr(failedCount)
    Exit Function
Failed:
    RunTestCoupledConstraint = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет каждую независимую проверку без остановки остальных доступных утверждений.
Private Sub TestCheck(ByVal condition As Boolean, ByVal name As String, ByRef report As String, ByRef passed As Long, ByRef failed As Long)
    If condition Then passed = passed + 1 Else failed = failed + 1
    report = report & IIf(condition, "PASS: ", "FAIL: ") & name & vbCrLf
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет неизменность линейного контекста, отказы его контракта и решение известной силовой системы на наклонной плоскости.
Public Function RunTestCoupledLinearConstraint() As String
    On Error GoTo Failed
    Dim report As String, passed As Long, failedCount As Long
    Dim linear As New CCoupledLinearConstraint, constraint As ICoupledScalarConstraint
    Dim coefficients(1 To 3) As Double, x(1 To 3) As Double, value As Double, gradient() As Double
    coefficients(1) = 2#: coefficients(2) = -3#: linear.Initialize coefficients, 1.2
    Set constraint = linear
    coefficients(1) = 999#: x(1) = 2#: x(2) = 1#: x(3) = 0.25
    TestCheck constraint.TryEvaluate(x, True, value, gradient), "linear control returns response", report, passed, failedCount
    TestCheck Abs(value - 2.2) < 0.000000000001 And gradient(1) = 2# And gradient(2) = -3# And gradient(3) = 0#, _
        "input coefficient mutation does not alter prepared control", report, passed, failedCount
    gradient(2) = 999#
    TestCheck constraint.TryEvaluate(x, True, value, gradient) And gradient(2) = -3#, _
        "returned gradient is an independent copy", report, passed, failedCount
    TestCheck x(1) = 2# And x(2) = 1# And x(3) = 0.25, "scalar response preserves input field", report, passed, failedCount

    Dim model As New CConstraintEquationsTest, provider As ICoupledEquationSystem
    model.InitializeTest "Circle": Set provider = model
    Dim origin(1 To 3) As Double: origin(1) = 0.2: origin(2) = 0.1: origin(3) = 0.1
    Dim adapter As New CCoupledConstraintEquations, equations As ICoupledEquationSystem
    adapter.Initialize provider, constraint, origin, 0.0000000000001, 2: Set equations = adapter
    TestCheck adapter.Pivot = 2, "linear mapping selects actual strongest field derivative", report, passed, failedCount
    Dim solver As New CCoupledHybridSolver, solution() As Double, mapped() As Double
    Dim solved As Boolean, residual() As Double, unused() As Double
    solved = solver.TrySolveLeastSquaresTrust(equations, origin, 500, 30, 0.00000000001, solution, initialRadius:=0.2)
    TestCheck solved, "existing QR trust solver finds constrained linear root", report, passed, failedCount
    If Not solved Then Err.Raise 5, , "Не найден известный корень линейного условия управления."
    TestCheck adapter.TryMap(solution, mapped), "raw solution maps to physical coordinates", report, passed, failedCount
    TestCheck Abs(mapped(1)) < 0.000000001 And Abs(mapped(2) - 0.4) < 0.000000001 And Abs(mapped(3) - 0.4) < 0.000000001, _
        "physical field matches independent analytic equilibrium", report, passed, failedCount
    TestCheck constraint.TryEvaluate(mapped, False, value, unused) And Abs(value) < 0.000000000001, _
        "mapped root satisfies actual scalar plane", report, passed, failedCount
    TestCheck provider.TryEvaluate(mapped, False, residual, unused) And Abs(residual(1)) < 0.000000001 And Abs(residual(2)) < 0.000000001, _
        "physical force rows satisfy original provider", report, passed, failedCount

    Dim errorNumber As Long, available As Boolean, absent As New CCoupledLinearConstraint, missing As ICoupledScalarConstraint
    Set missing = absent
    On Error Resume Next
    available = missing.TryEvaluate(x, False, value, unused): errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck errorNumber = 5, "uninitialized control gives explicit input error", report, passed, failedCount
    On Error Resume Next
    linear.Initialize coefficients, 0#: errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck errorNumber = 5, "prepared control rejects reinitialization", report, passed, failedCount
    Dim shortField(1 To 2) As Double
    On Error Resume Next
    available = constraint.TryEvaluate(shortField, False, value, unused): errorNumber = Err.Number: Err.Clear
    On Error GoTo Failed
    TestCheck errorNumber = 5, "wrong field size gives explicit input error", report, passed, failedCount
    TestCheck constraint.TryEvaluate(x, False, value, unused) And Abs(value - 2.2) < 0.000000000001, _
        "rejected calls preserve prepared scalar context", report, passed, failedCount
    RunTestCoupledLinearConstraint = report & "TOTAL_COUPLED_LINEAR_CONSTRAINT_TEST: passed=" & CStr(passed) & "; failed=" & CStr(failedCount)
    Exit Function
Failed:
    RunTestCoupledLinearConstraint = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
