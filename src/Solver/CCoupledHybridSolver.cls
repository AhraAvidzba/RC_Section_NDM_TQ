VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledHybridSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Гибридный метод Пауэлла для квадратной системы нелинейных уравнений
' ==========================================================================
' Перенос HYBRJ MINPACK/CMinpack с внешним масштабированием координат
' (mode=2, diag=1). Полная касательная чередуется с секущим обновлением
' QR ранга один; радиус dogleg меняется по фактическому снижению невязки.
' Источник и полная лицензия: literature/Coupled6D/NumericalControl/Minpack,
' ревизия 64c5635addd72fb54f44db8546c191feeb70e338, CopyrightMINPACK.txt.
' This product includes software developed by the University of Chicago,
' as Operator of Argonne National Laboratory.
' Класс не принимает историю материалов и не определяет разрушение.

Private Const MACHINE_EPSILON As Double = 2# ^ (-52)
Private mFunctionCalls As Long, mJacobianCalls As Long, mSteps As Long, mAccepted As Long, mStopCode As Long
Private mBestNorm As Double, mBest() As Double, mBestResidual() As Double
Private mHasBest As Boolean
Private mRestarts As Long

' Ищет равновесие в масштабированных координатах и возвращает лучшую фактически вычисленную точку.
' Успех требует заданной максимальной невязки; малый шаг HYBRJ сам по себе успехом не считается.
' maxEvaluations ограничивает вызовы функции, maxSteps — все dogleg-поправки, включая отвергнутые.
' При relativeDifferenceStep>0 касательная строится центральной разностью полной системы;
' каждая внутренняя разностная проба входит в общий maxEvaluations и счётчик FunctionCalls.
Friend Function TrySolve(ByVal equations As ICoupledEquationSystem, ByRef initial() As Double, _
        ByVal maxEvaluations As Long, ByVal maxSteps As Long, ByVal residualTolerance As Double, _
        ByVal stepTolerance As Double, ByVal initialRadius As Double, ByRef solution() As Double, _
        Optional ByVal maxRestarts As Long = 0, Optional ByVal relativeDifferenceStep As Double = 0#) As Boolean
    mRestarts = 0
    If maxRestarts < 0 Then Err.Raise 5, "CCoupledHybridSolver", "Число повторных стартов не может быть отрицательным."
    If relativeDifferenceStep < 0# Then Err.Raise 5, "CCoupledHybridSolver", "Шаг численного дифференцирования не может быть отрицательным."
    Dim start() As Double, usedFunctions As Long, usedJacobians As Long, usedSteps As Long, usedAccepted As Long
    Dim solved As Boolean, attempt As Long
    start = initial
    For attempt = 0 To maxRestarts
        solved = SolveFrom(equations, start, maxEvaluations - usedFunctions, maxSteps - usedSteps, _
            residualTolerance, stepTolerance, initialRadius, relativeDifferenceStep, solution)
        usedFunctions = usedFunctions + mFunctionCalls: usedJacobians = usedJacobians + mJacobianCalls
        usedSteps = usedSteps + mSteps: usedAccepted = usedAccepted + mAccepted
        If solved Or Not mHasBest Or attempt = maxRestarts Then Exit For
        If mStopCode < 3 Or mStopCode > 5 Then Exit For
        If usedFunctions >= maxEvaluations Or usedSteps >= maxSteps Then mStopCode = 2: Exit For
        start = solution: mRestarts = mRestarts + 1 ' Новый HYBRJ сбрасывает накопленные секущие изменения в лучшей фактической точке.
    Next attempt
    mFunctionCalls = usedFunctions: mJacobianCalls = usedJacobians: mSteps = usedSteps: mAccepted = usedAccepted
    TrySolve = solved
End Function

' Выполняет один HYBRJ без принятия истории материалов. Повторный старт разрешает только внешний ограниченный цикл.
Private Function SolveFrom(ByVal equations As ICoupledEquationSystem, ByRef initial() As Double, _
        ByVal maxEvaluations As Long, ByVal maxSteps As Long, ByVal residualTolerance As Double, _
        ByVal stepTolerance As Double, ByVal initialRadius As Double, ByVal relativeDifferenceStep As Double, _
        ByRef solution() As Double) As Boolean
    On Error GoTo ArithmeticFailure
    mFunctionCalls = 0: mJacobianCalls = 0: mSteps = 0: mAccepted = 0: mStopCode = 0
    mBestNorm = 1E+300: mHasBest = False: Erase mBest: Erase mBestResidual: Erase solution
    If equations Is Nothing Then Exit Function
    Dim count As Long: count = UBound(initial)
    If LBound(initial) <> 1 Or count < 1 Or maxEvaluations < 1 Or maxSteps < 1 Or _
            residualTolerance < 0# Or stepTolerance < 0# Or initialRadius <= 0# Then Exit Function
    Dim x() As Double, residual() As Double, matrix() As Double, movement() As Double, trial() As Double
    Dim nextResidual() As Double, model() As Double
    x = initial
    If Not ReadResidual(equations, x, residual, matrix, count) Then GoTo Finished
    Dim algebra As New CCoupledHybridFactorization
    Dim norm As Double, nextNorm As Double, xnorm As Double, radius As Double, movementNorm As Double
    norm = algebra.VectorNorm(residual, 1, count)
    Dim goodRadius As Long, badRadius As Long, slowSteps As Long, slowMatrices As Long
    Dim actualReduction As Double, predictedReduction As Double, ratio As Double, estimate As Double
    Dim fresh As Boolean, first As Boolean, accepted As Boolean, i As Long
    first = True
    Do
        If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
        If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
        mJacobianCalls = mJacobianCalls + 1
        If relativeDifferenceStep > 0# Then
            If Not BuildDifferenceJacobian(equations, x, relativeDifferenceStep, maxEvaluations, matrix, count) Then Exit Do
        Else
            If Not equations.TryEvaluate(x, True, residual, matrix) Then mStopCode = -1: Exit Do
        End If
        If Not ValidDimensions(residual, matrix, count) Then mStopCode = -1: Exit Do
        algebra.Factor matrix, residual, count
        fresh = True
        If first Then
            xnorm = algebra.VectorNorm(x, 1, count): radius = initialRadius * xnorm
            If radius = 0# Then radius = initialRadius
        End If
        Do
            If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
            algebra.Dogleg radius, movement
            ReDim trial(1 To count)
            For i = 1 To count: movement(i) = -movement(i): trial(i) = x(i) + movement(i): Next i
            movementNorm = algebra.VectorNorm(movement, 1, count)
            If first Then radius = GeomMin(radius, movementNorm)
            mSteps = mSteps + 1
            If Not ReadResidual(equations, trial, nextResidual, matrix, count) Then Exit Do
            nextNorm = algebra.VectorNorm(nextResidual, 1, count)
            actualReduction = -1#
            If nextNorm < norm Then actualReduction = 1# - (nextNorm / norm) ^ 2
            algebra.ModelResidual movement, model
            estimate = algebra.VectorNorm(model, 1, count): predictedReduction = 0#
            If estimate < norm Then predictedReduction = 1# - (estimate / norm) ^ 2
            ratio = 0#
            If predictedReduction > 0# Then ratio = actualReduction / predictedReduction
            If ratio < 0.1 Then
                goodRadius = 0: badRadius = badRadius + 1: radius = 0.5 * radius
            Else
                badRadius = 0: goodRadius = goodRadius + 1
                If ratio >= 0.5 Or goodRadius > 1 Then radius = GeomMax(radius, movementNorm / 0.5)
                If Abs(ratio - 1#) <= 0.1 Then radius = movementNorm / 0.5
            End If
            accepted = (ratio >= 0.0001)
            If accepted Then
                x = trial: residual = nextResidual: norm = nextNorm
                xnorm = algebra.VectorNorm(x, 1, count): mAccepted = mAccepted + 1: first = False
            End If
            slowSteps = slowSteps + 1
            If actualReduction >= 0.001 Then slowSteps = 0
            If fresh Then slowMatrices = slowMatrices + 1
            If actualReduction >= 0.1 Then slowMatrices = 0
            If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
            If radius <= stepTolerance * xnorm Or norm = 0# Then mStopCode = 1: Exit Do
            If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2
            If 0.1 * GeomMax(0.1 * radius, movementNorm) <= MACHINE_EPSILON * xnorm Then mStopCode = 3
            If slowMatrices = 5 Then mStopCode = 4
            If slowSteps = 10 Then mStopCode = 5
            If mStopCode <> 0 Then Exit Do
            If badRadius = 2 Then Exit Do
            If movementNorm = 0# Then mStopCode = 3: Exit Do
            algebra.Update movement, nextResidual, model, movementNorm, accepted
            fresh = False
        Loop
        If mStopCode <> 0 Then Exit Do
    Loop
Finished:
    If mHasBest Then
        solution = mBest
        SolveFrom = (mBestNorm <= residualTolerance)
    End If
    Exit Function
ArithmeticFailure:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    If errorNumber <> 6 And errorNumber <> 11 Then Err.Raise errorNumber, errorSource, errorDescription
    mStopCode = -2
    Resume Finished
End Function

' Строит касательную неизменённой полной невязки; пробные точки не принимают материальную историю.
' Недостающий бюджет останавливает построение до следующей пары, не публикуя частичную матрицу.
Private Function BuildDifferenceJacobian(ByVal equations As ICoupledEquationSystem, ByRef x() As Double, _
        ByVal relativeStep As Double, ByVal maximumCalls As Long, ByRef matrix() As Double, ByVal count As Long) As Boolean
    ReDim matrix(1 To count, 1 To count)
    Dim point() As Double, plus() As Double, minus() As Double, unusedMatrix() As Double
    Dim column As Long, row As Long, delta As Double
    For column = 1 To count
        If mFunctionCalls + 2 > maximumCalls Then mStopCode = 2: Exit Function
        delta = relativeStep * (1# + Abs(x(column)))
        If x(column) + delta = x(column) Or x(column) - delta = x(column) Then mStopCode = 3: Exit Function
        point = x: point(column) = x(column) + delta
        If Not ReadResidual(equations, point, plus, unusedMatrix, count) Then Exit Function
        point(column) = x(column) - delta
        If Not ReadResidual(equations, point, minus, unusedMatrix, count) Then Exit Function
        For row = 1 To count: matrix(column, row) = (plus(row) - minus(row)) / (2# * delta): Next row
    Next column
    BuildDifferenceJacobian = True
End Function

' Вычисляет одну невязку и сохраняет точную лучшую точку без округления ключей или координат.
Private Function ReadResidual(ByVal equations As ICoupledEquationSystem, ByRef x() As Double, _
        ByRef residual() As Double, ByRef matrix() As Double, ByVal count As Long) As Boolean
    mFunctionCalls = mFunctionCalls + 1
    If Not equations.TryEvaluate(x, False, residual, matrix) Then mStopCode = -1: Exit Function
    If LBound(residual) <> 1 Or UBound(residual) <> count Then mStopCode = -1: Exit Function
    Dim i As Long, norm As Double
    For i = 1 To count: norm = GeomMax(norm, Abs(residual(i))): Next i
    If Not mHasBest Or norm < mBestNorm Then mBestNorm = norm: mBest = x: mBestResidual = residual: mHasBest = True
    ReadResidual = True
End Function

' Проверяет контракт размеров полной касательной, не исправляя ошибочный отклик молча.
Private Function ValidDimensions(ByRef residual() As Double, ByRef matrix() As Double, ByVal count As Long) As Boolean
    ValidDimensions = (LBound(residual) = 1 And UBound(residual) = count And _
        LBound(matrix, 1) = 1 And UBound(matrix, 1) = count And LBound(matrix, 2) = 1 And UBound(matrix, 2) = count)
End Function

' Возвращает число чтений невязки, включая все внутренние пробы численной касательной.
Friend Property Get FunctionCalls() As Long
    FunctionCalls = mFunctionCalls
End Property

' Возвращает число запросов получения или построения полной касательной, включая прерванное построение.
Friend Property Get JacobianCalls() As Long
    JacobianCalls = mJacobianCalls
End Property

' Возвращает число всех предложенных dogleg-поправок, включая отвергнутые.
Friend Property Get Steps() As Long
    Steps = mSteps
End Property

' Возвращает число принятых численным методом поправок; это не принятие истории материалов.
Friend Property Get AcceptedSteps() As Long
    AcceptedSteps = mAccepted
End Property

' Возвращает причину остановки HYBRJ: 1 — малый шаг/невязка, 2 — бюджет,
' 3 — машинная точность, 4/5 — слабый прогресс, -1 — нет отклика, -2 — арифметика.
' Код 1 без проверки BestResidual не означает найденного равновесия.
Friend Property Get StopCode() As Long
    StopCode = mStopCode
End Property

' Возвращает максимальную компоненту лучшей фактически вычисленной невязки.
Friend Property Get BestResidual() As Double
    BestResidual = mBestNorm
End Property

' Сообщает о наличии хотя бы одной полностью вычисленной невязки и соответствующей лучшей точки.
' Сам факт обращения к интерфейсу не означает, что тот вернул пригодный отклик.
Friend Property Get HasBestPoint() As Boolean
    HasBestPoint = mHasBest
End Property

' Возвращает число фактически выполненных повторных стартов после остановки по слабому прогрессу или точности.
Friend Property Get Restarts() As Long
    Restarts = mRestarts
End Property
