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
' Гибридный метод Пауэлла и доверительное решение квадратных уравнений
' ==========================================================================
' Перенос HYBRJ MINPACK/CMinpack с внешним масштабированием координат.
' По умолчанию mode=2, diag=1; automaticScaling включает норму столбцов
' mode=1 через эквивалентную замену координат QR. freshJacobian обновляет
' полную касательную после принятой поправки вместо обновления ранга один.
' Радиус dogleg меняется по фактическому снижению невязки.
' Источник и полная лицензия: literature/Coupled6D/NumericalControl/Minpack,
' ревизия 64c5635addd72fb54f44db8546c191feeb70e338, CopyrightMINPACK.txt.
' This product includes software developed by the University of Chicago,
' as Operator of Argonne National Laboratory.
' Класс не принимает историю материалов и не определяет разрушение.

Private Const MACHINE_EPSILON As Double = 2# ^ (-52)
Private mFunctionCalls As Long, mJacobianCalls As Long, mSteps As Long, mAccepted As Long, mStopCode As Long
Private mArithmeticDiagnostic As String
Private mBestNorm As Double, mBest() As Double, mBestResidual() As Double
Private mHasBest As Boolean
Private mRestarts As Long
Private mHasDifferenceBracket As Boolean, mDifferenceBefore() As Double, mDifferenceAfter() As Double
Private mDifferenceScore As Double, mDifferenceColumn As Long

' Строит только разностную касательную полной системы от неизменной точки, переиспользуя тот же контроль парных проб и размеров.
' Неполная матрица не публикуется; число вызовов ограничено maximumCalls, входные координаты не меняются.
Friend Function TryLinearizeByDifferences(ByVal equations As ICoupledEquationSystem, ByRef coordinates() As Double, _
        ByVal relativeStep As Double, ByVal maximumCalls As Long, ByRef matrix() As Double) As Boolean
    Erase matrix
    mFunctionCalls = 0: mJacobianCalls = 0: mSteps = 0: mAccepted = 0: mStopCode = 0
    mHasBest = False: mHasDifferenceBracket = False: Erase mBest: Erase mBestResidual
    If LBound(coordinates) <> 1 Or relativeStep <= 0# Then Err.Raise 5, "CCoupledHybridSolver", "Для разностной касательной нужны координаты с индекса 1 и положительный шаг."
    Dim count As Long: count = UBound(coordinates)
    If count < 1 Or maximumCalls < 2# * count + 1# Then mStopCode = 2: Exit Function
    Dim residual() As Double, unused() As Double, completed() As Double
    If Not ReadResidual(equations, coordinates, residual, unused, count) Then Exit Function
    mJacobianCalls = 1
    If Not BuildDifferenceJacobian(equations, coordinates, residual, relativeStep, maximumCalls, completed, count, False, Nothing) Then Exit Function
    matrix = completed: mStopCode = 1: TryLinearizeByDifferences = True
End Function

' Ищет равновесие в масштабированных координатах и возвращает лучшую фактически вычисленную точку.
' Успех требует заданной максимальной невязки; малый шаг HYBRJ сам по себе успехом не считается.
' maxEvaluations ограничивает вызовы функции, maxSteps — все dogleg-поправки, включая отвергнутые.
' При relativeDifferenceStep>0 касательная строится центральной разностью полной системы;
' каждая внутренняя разностная проба входит в общий maxEvaluations и счётчик FunctionCalls.
' automaticScaling и freshJacobian меняют только численную стратегию; критерий невязки и общие бюджеты сохраняются.
' captureDifferenceBracket сохраняет две вычисленные разностные пробы без дополнительных обращений к отклику.
' completeNewtonLineSearch выбирает Newton с поиском убывания по Армихо в тех же бюджетах;
' при нулевом шаге используется матрица уравнений, при положительном — глобальная центральная разность.
Friend Function TrySolve(ByVal equations As ICoupledEquationSystem, ByRef initial() As Double, _
        ByVal maxEvaluations As Long, ByVal maxSteps As Long, ByVal residualTolerance As Double, _
        ByVal stepTolerance As Double, ByVal initialRadius As Double, ByRef solution() As Double, _
        Optional ByVal maxRestarts As Long = 0, Optional ByVal relativeDifferenceStep As Double = 0#, _
        Optional ByVal automaticScaling As Boolean = False, Optional ByVal freshJacobian As Boolean = False, _
        Optional ByVal captureDifferenceBracket As Boolean = False, Optional ByVal differenceObserver As ICoupledDifferenceObserver = Nothing, _
        Optional ByVal completeNewtonLineSearch As Boolean = False) As Boolean
    mRestarts = 0
    mArithmeticDiagnostic = vbNullString
    mHasDifferenceBracket = False: mDifferenceScore = 0#: mDifferenceColumn = 0
    Erase mDifferenceBefore: Erase mDifferenceAfter
    If maxRestarts < 0 Then Err.Raise 5, "CCoupledHybridSolver", "Число повторных стартов не может быть отрицательным."
    If relativeDifferenceStep < 0# Then Err.Raise 5, "CCoupledHybridSolver", "Шаг численного дифференцирования не может быть отрицательным."
    Dim start() As Double, usedFunctions As Long, usedJacobians As Long, usedSteps As Long, usedAccepted As Long
    Dim solved As Boolean, attempt As Long
    start = initial
    For attempt = 0 To maxRestarts
        If completeNewtonLineSearch Then
            solved = SolveNewtonFrom(equations, start, maxEvaluations - usedFunctions, maxSteps - usedSteps, _
                residualTolerance, relativeDifferenceStep, solution, captureDifferenceBracket, differenceObserver)
        Else
            solved = SolveFrom(equations, start, maxEvaluations - usedFunctions, maxSteps - usedSteps, _
                residualTolerance, stepTolerance, initialRadius, relativeDifferenceStep, solution, automaticScaling, freshJacobian, captureDifferenceBracket, differenceObserver)
        End If
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

' Восстанавливает полную касательную на каждом принятом поле; до 21 доли одной поправки проверяются по фактической норме.
' Ни уменьшение шага, ни отказ линейной системы не объявляются физическим пределом; каждая проба расходует общий бюджет.
Private Function SolveNewtonFrom(ByVal equations As ICoupledEquationSystem, ByRef initial() As Double, _
        ByVal maxEvaluations As Long, ByVal maxSteps As Long, ByVal residualTolerance As Double, ByVal relativeStep As Double, _
        ByRef solution() As Double, ByVal captureBracket As Boolean, ByVal observer As ICoupledDifferenceObserver) As Boolean
    On Error GoTo ArithmeticFailure
    mFunctionCalls = 0: mJacobianCalls = 0: mSteps = 0: mAccepted = 0: mStopCode = 0
    mBestNorm = 1E+300: mHasBest = False: Erase mBest: Erase mBestResidual: Erase solution
    If equations Is Nothing Then Exit Function
    Dim count As Long: count = UBound(initial)
    If LBound(initial) <> 1 Or count < 1 Or maxEvaluations < 1 Or maxSteps < 1 Or residualTolerance < 0# Then Exit Function
    Dim x() As Double, residual() As Double, matrix() As Double, correction() As Double, trial() As Double, nextResidual() As Double
    x = initial
    If Not ReadResidual(equations, x, residual, matrix, count) Then GoTo Finished
    Dim algebra As New CCoupledHybridFactorization, linear As New CCoupledLinearSystem
    Dim baseNorm As Double, norm As Double, alpha As Double, division As Long, i As Long, accepted As Boolean
    Do
        If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
        If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
        mJacobianCalls = mJacobianCalls + 1
        If relativeStep > 0# Then
            If Not BuildDifferenceJacobian(equations, x, residual, relativeStep, maxEvaluations, matrix, count, captureBracket, observer) Then Exit Do
        Else
            If Not equations.TryEvaluate(x, True, residual, matrix) Then mStopCode = -1: Exit Do
        End If
        If Not ValidDimensions(residual, matrix, count) Then mStopCode = -1: Exit Do
        ReDim correction(1 To count)
        For i = 1 To count: correction(i) = -residual(i): Next i
        If Not linear.Solve(matrix, correction, count) Then mStopCode = 3: Exit Do
        baseNorm = algebra.VectorNorm(residual, 1, count): alpha = 1#: accepted = False
        mSteps = mSteps + 1
        For division = 0 To 20
            If mFunctionCalls >= maxEvaluations Then mStopCode = 2: Exit For
            trial = x
            For i = 1 To count: trial(i) = x(i) + alpha * correction(i): Next i
            If ReadTrialResidual(equations, trial, nextResidual, matrix, count) Then
                norm = algebra.VectorNorm(nextResidual, 1, count)
                If norm <= (1# - 0.0001 * alpha) * baseNorm Then
                    x = trial: residual = nextResidual: mAccepted = mAccepted + 1: accepted = True
                    Exit For
                End If
            ElseIf mStopCode <> -1 Then
                Exit For
            Else
                mStopCode = 0 ' Отсутствие отклика одной доли не исключает меньшую долю той же поправки.
            End If
            alpha = alpha * 0.5
        Next division
        If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
        If mStopCode <> 0 Then Exit Do
        If Not accepted Then mStopCode = 5: Exit Do
    Loop
Finished:
    If mHasBest Then solution = mBest: SolveNewtonFrom = (mBestNorm <= residualTolerance)
    Exit Function
ArithmeticFailure:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    If errorNumber <> 6 And errorNumber <> 11 Then Err.Raise errorNumber, errorSource, errorDescription
    mArithmeticDiagnostic = "Арифметический отказ полного Newton: " & CStr(errorNumber) & "; " & errorSource & "; " & errorDescription
    mStopCode = -2
    Resume Finished
End Function

' Выполняет один HYBRJ без принятия истории материалов. Повторный старт разрешает только внешний ограниченный цикл.
' Недоступная пробная невязка уменьшает радиус; отсутствие исходного отклика и нарушение размеров завершают поиск.
Private Function SolveFrom(ByVal equations As ICoupledEquationSystem, ByRef initial() As Double, _
        ByVal maxEvaluations As Long, ByVal maxSteps As Long, ByVal residualTolerance As Double, _
        ByVal stepTolerance As Double, ByVal initialRadius As Double, ByVal relativeDifferenceStep As Double, _
        ByRef solution() As Double, ByVal automaticScaling As Boolean, ByVal freshJacobian As Boolean, _
        ByVal captureDifferenceBracket As Boolean, ByVal differenceObserver As ICoupledDifferenceObserver) As Boolean
    On Error GoTo ArithmeticFailure
    mFunctionCalls = 0: mJacobianCalls = 0: mSteps = 0: mAccepted = 0: mStopCode = 0
    mBestNorm = 1E+300: mHasBest = False: Erase mBest: Erase mBestResidual: Erase solution
    If equations Is Nothing Then Exit Function
    Dim count As Long: count = UBound(initial)
    If LBound(initial) <> 1 Or count < 1 Or maxEvaluations < 1 Or maxSteps < 1 Or _
            residualTolerance < 0# Or stepTolerance < 0# Or initialRadius <= 0# Then Exit Function
    Dim x() As Double, residual() As Double, matrix() As Double, movement() As Double, trial() As Double
    Dim nextResidual() As Double, model() As Double
    Dim phase As String: phase = "исходный отклик"
    x = initial
    If Not ReadResidual(equations, x, residual, matrix, count) Then GoTo Finished
    Dim algebra As New CCoupledHybridFactorization
    Dim norm As Double, nextNorm As Double, xnorm As Double, radius As Double, movementNorm As Double
    norm = algebra.VectorNorm(residual, 1, count)
    Dim goodRadius As Long, badRadius As Long, slowSteps As Long, slowMatrices As Long
    Dim actualReduction As Double, predictedReduction As Double, ratio As Double, estimate As Double
    Dim fresh As Boolean, first As Boolean, accepted As Boolean, i As Long, trialUnavailable As Boolean
    Dim diagonal() As Double, workMetric() As Double, diagonalReady As Boolean
    ReDim diagonal(1 To count): ReDim workMetric(1 To count)
    For i = 1 To count: diagonal(i) = 1#: Next i
    first = True
    Do
        If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
        If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
        mJacobianCalls = mJacobianCalls + 1
        phase = "полная касательная"
        If relativeDifferenceStep > 0# Then
            If Not BuildDifferenceJacobian(equations, x, residual, relativeDifferenceStep, maxEvaluations, matrix, count, captureDifferenceBracket, differenceObserver) Then Exit Do
        Else
            If Not equations.TryEvaluate(x, True, residual, matrix) Then mStopCode = -1: Exit Do
        End If
        If Not ValidDimensions(residual, matrix, count) Then mStopCode = -1: Exit Do
        phase = "масштабирование столбцов"
        If automaticScaling Then ScaleJacobianColumns matrix, diagonal, workMetric, count, diagonalReady, algebra
        phase = "QR-разложение"
        algebra.Factor matrix, residual, count
        fresh = True
        If first Then
            xnorm = ScaledCoordinateNorm(x, diagonal, workMetric, count, algebra): radius = initialRadius * xnorm
            If radius = 0# Then radius = initialRadius
        End If
        Do
RetryDoglegTrial:
            If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
            phase = "dogleg-поправка"
            algebra.Dogleg radius, movement
            ReDim trial(1 To count)
            phase = "координаты пробной точки"
            For i = 1 To count: movement(i) = -movement(i): trial(i) = x(i) + movement(i) / diagonal(i): Next i
            movementNorm = algebra.VectorNorm(movement, 1, count)
            If first Then radius = GeomMin(radius, movementNorm)
            mSteps = mSteps + 1
            phase = "пробный отклик"
            If Not ReadTrialResidual(equations, trial, nextResidual, matrix, count, trialUnavailable) Then
                If Not trialUnavailable Then Exit Do ' Ошибочный размер отклика не исправляется уменьшением радиуса.
                mStopCode = 0: goodRadius = 0
                radius = 0.5 * radius ' Неизвестная невязка отклоняет только пробу, сохраняя исходную точку и её QR.
                If radius = 0# Or radius <= stepTolerance * xnorm Then mStopCode = 3: Exit Do
                GoTo RetryDoglegTrial ' Неудачная проба уже учтена в mSteps и mFunctionCalls.
            End If
            nextNorm = algebra.VectorNorm(nextResidual, 1, count)
            actualReduction = -1#
            If nextNorm < norm Then actualReduction = 1# - (nextNorm / norm) ^ 2
            phase = "невязка линейной модели"
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
                xnorm = ScaledCoordinateNorm(x, diagonal, workMetric, count, algebra): mAccepted = mAccepted + 1: first = False
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
            If freshJacobian And accepted Then Exit Do ' Следующая линеаризация учитывает фактические активные ветви материала.
            phase = "секущий поворот QR"
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
    mArithmeticDiagnostic = "Арифметический отказ HYBRJ; этап " & phase & "; ошибка " & CStr(errorNumber) & "; " & errorSource & "; " & errorDescription
    mStopCode = -2
    Resume Finished
End Function

' Нормирует столбцы по монотонной диагонали HYBRJ mode=1. QR и его секущее
' обновление работают в координатах z=D*x; исходные x и невязка не меняются.
Private Sub ScaleJacobianColumns(ByRef matrix() As Double, ByRef diagonal() As Double, ByRef work() As Double, _
        ByVal count As Long, ByRef ready As Boolean, ByVal algebra As CCoupledHybridFactorization)
    Dim column As Long, row As Long, magnitude As Double
    For column = 1 To count
        For row = 1 To count: work(row) = matrix(column, row): Next row
        magnitude = algebra.VectorNorm(work, 1, count)
        If magnitude = 0# Then magnitude = 1# ' Нулевая производная не задаёт нулевую метрику координаты.
        If ready Then diagonal(column) = GeomMax(diagonal(column), magnitude) Else diagonal(column) = magnitude
        For row = 1 To count: matrix(column, row) = matrix(column, row) / diagonal(column): Next row
    Next column
    ready = True
End Sub

' Считает длину координаты в той же диагональной метрике, что ограничивает dogleg-поправку.
Private Function ScaledCoordinateNorm(ByRef values() As Double, ByRef diagonal() As Double, ByRef work() As Double, _
        ByVal count As Long, ByVal algebra As CCoupledHybridFactorization) As Double
    Dim i As Long
    For i = 1 To count: work(i) = diagonal(i) * values(i): Next i
    ScaledCoordinateNorm = algebra.VectorNorm(work, 1, count)
End Function

' Строит касательную неизменённой полной невязки; пробные точки не принимают материальную историю.
' Недостающий бюджет останавливает построение до следующей пары, не публикуя частичную матрицу.
Private Function BuildDifferenceJacobian(ByVal equations As ICoupledEquationSystem, ByRef x() As Double, _
        ByRef baseResidual() As Double, ByVal relativeStep As Double, ByVal maximumCalls As Long, _
        ByRef matrix() As Double, ByVal count As Long, ByVal captureBracket As Boolean, ByVal observer As ICoupledDifferenceObserver) As Boolean
    ReDim matrix(1 To count, 1 To count)
    Dim point() As Double, plus() As Double, minus() As Double, unusedMatrix() As Double
    Dim column As Long, row As Long, delta As Double
    Dim asymmetry() As Double, steps() As Double
    If captureBracket Then ReDim asymmetry(1 To count, 1 To count): ReDim steps(1 To count)
    For column = 1 To count
        If mFunctionCalls + 2 > maximumCalls Then mStopCode = 2: Exit Function
        delta = relativeStep * (1# + Abs(x(column)))
        If x(column) + delta = x(column) Or x(column) - delta = x(column) Then mStopCode = 3: Exit Function
        point = x: point(column) = x(column) + delta
        If Not ReadResidual(equations, point, plus, unusedMatrix, count) Then Exit Function
        point(column) = x(column) - delta
        If Not ReadResidual(equations, point, minus, unusedMatrix, count) Then Exit Function
        For row = 1 To count: matrix(column, row) = (plus(row) - minus(row)) / (2# * delta): Next row
        If captureBracket Then
            steps(column) = delta
            For row = 1 To count: asymmetry(column, row) = ((plus(row) - baseResidual(row)) + (minus(row) - baseResidual(row))) / delta: Next row
        End If
    Next column
    If captureBracket Then RecordDifferenceBracket x, steps, asymmetry, matrix, count, observer
    BuildDifferenceJacobian = True
End Function

' Сохраняет наиболее асимметричный столбец только полностью построенной матрицы, без новых вызовов функции или изменения решения.
Private Sub RecordDifferenceBracket(ByRef x() As Double, ByRef steps() As Double, ByRef asymmetry() As Double, _
        ByRef matrix() As Double, ByVal count As Long, ByVal observer As ICoupledDifferenceObserver)
    If observer Is Nothing Then
        mHasDifferenceBracket = False: mDifferenceScore = 0#: mDifferenceColumn = 0
        Erase mDifferenceBefore: Erase mDifferenceAfter
    End If
    Dim norms As New CCoupledHybridFactorization, rowNorm() As Double, values() As Double, i As Long, j As Long
    ReDim rowNorm(1 To count): ReDim values(1 To count)
    For i = 1 To count
        For j = 1 To count: values(j) = matrix(j, i): Next j
        rowNorm(i) = norms.VectorNorm(values, 1, count)
        If rowNorm(i) = 0# Then rowNorm(i) = 1# ' Нулевая линейная строка не скрывает наблюдавшуюся нелинейность.
    Next i
    Dim score As Double, best As Double, before() As Double, after() As Double, accepted As Boolean
    For j = 1 To count
        For i = 1 To count: values(i) = asymmetry(j, i) / rowNorm(i): Next i
        score = norms.VectorNorm(values, 1, count)
        If score > best Then
            before = x: after = x
            before(j) = x(j) - steps(j): after(j) = x(j) + steps(j)
            accepted = True
            If Not observer Is Nothing Then accepted = observer.AcceptBracket(before, after)
            If accepted Then
                mDifferenceBefore = before: mDifferenceAfter = after
                mDifferenceScore = score: mDifferenceColumn = j: mHasDifferenceBracket = True: best = score
            End If
        End If
    Next j
End Sub

' Копирует наиболее асимметричную подтверждённую пробу последней подходящей полной матрицы текущего TrySolve.
' Это диагностическое пересечение; наличие пробы само по себе не подтверждает равновесие или смену физической ветви.
Friend Function CopyDifferenceBracket(ByRef before() As Double, ByRef after() As Double, ByRef score As Double, ByRef column As Long) As Boolean
    Erase before: Erase after: score = 0#: column = 0
    If Not mHasDifferenceBracket Then Exit Function
    before = mDifferenceBefore: after = mDifferenceAfter: score = mDifferenceScore: column = mDifferenceColumn
    CopyDifferenceBracket = True
End Function

' Вычисляет одну невязку и сохраняет точную лучшую точку без округления ключей или координат.
' unavailable отличает отсутствие отклика интерфейса от нарушения размеров; исключения не перехватываются.
Private Function ReadResidual(ByVal equations As ICoupledEquationSystem, ByRef x() As Double, _
        ByRef residual() As Double, ByRef matrix() As Double, ByVal count As Long, Optional ByRef unavailable As Boolean = False) As Boolean
    unavailable = False
    mFunctionCalls = mFunctionCalls + 1
    If Not equations.TryEvaluate(x, False, residual, matrix) Then unavailable = True: mStopCode = -1: Exit Function
    If LBound(residual) <> 1 Or UBound(residual) <> count Then mStopCode = -1: Exit Function
    Dim i As Long, norm As Double
    For i = 1 To count: norm = GeomMax(norm, Abs(residual(i))): Next i
    If Not mHasBest Or norm < mBestNorm Then mBestNorm = norm: mBest = x: mBestResidual = residual: mHasBest = True
    ReadResidual = True
End Function

' Читает только пробную точку; переполнение или деление на ноль отклоняет её, не меняя исходное поле поиска.
' Вызов ReadResidual уже учитывает чтение. Отмена и остальные исключения передаются вызывающему коду.
Private Function ReadTrialResidual(ByVal equations As ICoupledEquationSystem, ByRef x() As Double, _
        ByRef residual() As Double, ByRef matrix() As Double, ByVal count As Long, Optional ByRef unavailable As Boolean = False) As Boolean
    On Error GoTo TrialArithmetic
    ReadTrialResidual = ReadResidual(equations, x, residual, matrix, count, unavailable)
    Exit Function
TrialArithmetic:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    If errorNumber <> 6 And errorNumber <> 11 Then Err.Raise errorNumber, errorSource, errorDescription
    mArithmeticDiagnostic = "Арифметический отказ пробного отклика: " & CStr(errorNumber) & "; " & errorSource & "; " & errorDescription
    unavailable = True: mStopCode = -1
End Function

' Возвращает последний арифметический отказ текущего поиска; успешный последующий шаг не превращается из-за него в отказ.
Friend Property Get ArithmeticDiagnostic() As String
    ArithmeticDiagnostic = mArithmeticDiagnostic
End Property

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

' Решает исходные квадратные уравнения свежей касательной и полной регуляризованной поправкой в доверительной области.
' Использует регуляризацию Морэ(1977), §3,5; радиус меняется по фактическому и прогнозному снижению суммы квадратов.
' Все пробные поправки и чтения функции входят в заданные бюджеты; малый градиент и малый шаг не заменяют конечного допуска невязки.
Friend Function TrySolveLeastSquaresTrust(ByVal equations As ICoupledEquationSystem, ByRef initial() As Double, _
        ByVal maxEvaluations As Long, ByVal maxSteps As Long, ByVal residualTolerance As Double, _
        ByRef solution() As Double, Optional ByVal relativeDifferenceStep As Double = 0.000001, _
        Optional ByVal initialRadius As Double = 1#) As Boolean
    On Error GoTo ArithmeticFailure
    mFunctionCalls = 0: mJacobianCalls = 0: mSteps = 0: mAccepted = 0: mRestarts = 0: mStopCode = 0
    mArithmeticDiagnostic = vbNullString
    mBestNorm = 1E+300: mHasBest = False: Erase mBest: Erase mBestResidual: Erase solution
    If equations Is Nothing Then Exit Function
    Dim count As Long: count = UBound(initial)
    If LBound(initial) <> 1 Or count < 1 Or maxEvaluations < 1 Or maxSteps < 1 Or _
            residualTolerance < 0# Or relativeDifferenceStep < 0# Or initialRadius <= 0# Then Err.Raise 5, _
            "CCoupledHybridSolver", "Проверьте размеры, бюджеты, допуск и радиус доверительной численной пробы."
    Dim x() As Double, residual() As Double, matrix() As Double, unused() As Double, trial() As Double, nextResidual() As Double
    Dim diagonal() As Double, work() As Double, gradient() As Double, movement() As Double, model() As Double
    ReDim diagonal(1 To count): ReDim work(1 To count): ReDim gradient(1 To count): ReDim model(1 To count)
    x = initial
    If Not ReadResidual(equations, x, residual, unused, count) Then GoTo Finished
    Dim algebra As New CCoupledHybridFactorization
    Dim radius As Double, nextRadius As Double, parameter As Double, parameterIterations As Long
    Dim scalingReady As Boolean, radiusReady As Boolean, i As Long, j As Long, valid As Boolean, accepted As Boolean
    Dim norm As Double, nextNorm As Double, stepNorm As Double, cost As Double, nextCost As Double
    Dim actualReduction As Double, predictedReduction As Double, ratio As Double, total As Double
    Dim originalResidual() As Double, trialUnavailable As Boolean
    Do
        If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
        If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
        mJacobianCalls = mJacobianCalls + 1
        If relativeDifferenceStep > 0# Then
            If Not BuildDifferenceJacobian(equations, x, residual, relativeDifferenceStep, maxEvaluations, matrix, count, False, Nothing) Then Exit Do
        Else
            originalResidual = residual
            mFunctionCalls = mFunctionCalls + 1
            If Not equations.TryEvaluate(x, True, residual, matrix) Then mStopCode = -1: Exit Do
            For i = 1 To count
                If residual(i) <> originalResidual(i) Then mStopCode = -1: Exit Do
            Next i
            If mStopCode <> 0 Then Exit Do
        End If
        If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
        If Not ValidDimensions(residual, matrix, count) Then mStopCode = -1: Exit Do
        ScaleJacobianColumns matrix, diagonal, work, count, scalingReady, algebra
        If Not radiusReady Then
            radius = ScaledCoordinateNorm(x, diagonal, work, count, algebra)
            If radius = 0# Then radius = initialRadius
            radiusReady = True
        End If
        For j = 1 To count
            gradient(j) = 0#
            For i = 1 To count: gradient(j) = gradient(j) + matrix(j, i) * residual(i): Next i
        Next j
        algebra.Factor matrix, residual, count, True
        norm = algebra.VectorNorm(residual, 1, count): cost = 0.5 * norm * norm
        accepted = False
        Do
            If mFunctionCalls >= maxEvaluations Or mSteps >= maxSteps Then mStopCode = 2: Exit Do
            If Not algebra.TryLeastSquaresRadius(radius, parameter, movement, parameterIterations) Then mStopCode = 3: Exit Do
            stepNorm = algebra.VectorNorm(movement, 1, count)
            If stepNorm = 0# Then mStopCode = 3: Exit Do
            predictedReduction = 0#
            For i = 1 To count
                total = 0#
                For j = 1 To count: total = total + matrix(j, i) * movement(j): Next j
                model(i) = total
                predictedReduction = predictedReduction - gradient(i) * movement(i)
            Next i
            total = algebra.VectorNorm(model, 1, count)
            predictedReduction = predictedReduction - 0.5 * total * total
            trial = x
            For i = 1 To count: trial(i) = x(i) + movement(i) / diagonal(i): Next i
            mSteps = mSteps + 1
            valid = ReadTrialResidual(equations, trial, nextResidual, unused, count, trialUnavailable)
            actualReduction = 0#: ratio = 0#: nextNorm = 0#: nextCost = 0#
            If valid Then
                nextNorm = algebra.VectorNorm(nextResidual, 1, count): nextCost = 0.5 * nextNorm * nextNorm
                actualReduction = cost - nextCost
                If predictedReduction > 0# Then ratio = actualReduction / predictedReduction
                If predictedReduction = 0# And actualReduction = 0# Then ratio = 1#
            ElseIf trialUnavailable Then
                mStopCode = 0 ' Отсутствие отклика одной пробы разрешает меньшую область от той же принятой истории.
            Else
                Exit Do
            End If
            nextRadius = radius
            If Not valid Or ratio < 0.25 Then
                nextRadius = 0.25 * stepNorm
            ElseIf ratio > 0.75 And stepNorm > 0.95 * radius Then
                nextRadius = 2# * radius
            End If
            accepted = (valid And actualReduction > 0#)
             If accepted Then
                x = trial: residual = nextResidual: mAccepted = mAccepted + 1
            End If
            If mBestNorm <= residualTolerance Then mStopCode = 1: Exit Do
            If nextRadius <= 0# Then mStopCode = 3: Exit Do
            parameter = parameter * (radius / nextRadius): radius = nextRadius
            If accepted Then Exit Do
        Loop
        If mStopCode <> 0 Then Exit Do
    Loop
Finished:
    If mHasBest Then solution = mBest: TrySolveLeastSquaresTrust = (mBestNorm <= residualTolerance)
    Exit Function
ArithmeticFailure:
    Dim number As Long, source As String, description As String
    number = Err.Number: source = Err.Source: description = Err.Description
    If number <> 6 And number <> 11 Then Err.Raise number, source, description
    mStopCode = -2
    Resume Finished
End Function
