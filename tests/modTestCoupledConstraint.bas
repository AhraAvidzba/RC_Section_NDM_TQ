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
