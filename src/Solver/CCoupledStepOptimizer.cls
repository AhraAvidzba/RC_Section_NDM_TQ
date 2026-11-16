VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledStepOptimizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Линейная поправка связанного равновесия с малым изменением деформаций
' ==========================================================================
' Столбцы уже масштабированы. Силовые строки заданы в МПа. Линейные
' кинематические условия сохраняются при решении системы Ньютона.
' Линейная поправка требует отдельной проверки фактического НДС.

Private Const MAX_STRAIN_ROWS As Long = 256
Private Const STRAIN_BATCH As Long = 24
Private Const MAX_LINEAR_SOLVES As Long = 280
Private Const MAX_BACKTRACK As Long = 30
Private Const DECREMENT_TOLERANCE As Double = 0.000000000002
Private Const INITIAL_BARRIER As Double = 0.02
Private Const MIN_FORCE_BARRIER As Double = 0.0000002
Private Const MIN_STRAIN_BARRIER As Double = 0.00002
Private Const ROUNDING_GAP As Double = 0.0000000001

Private mSize As Long, mForceCount As Long, mConstraintCount As Long, mRowCount As Long
Private mRadius As Double, mForceScale As Double, mStrainScale As Double
Private mRows() As Double, mOffsets() As Double, mBounds() As Double, mFreeBound() As Boolean
Private mEquality() As Double, mEqualityOffset() As Double
Private mForceNorm As Double, mStrainNorm As Double, mLinearSolveCount As Long
Private mSystem As CCoupledLinearSystem

' Подбирает поправку с ограниченной нормой силовых строк и малой
' наибольшей деформацией. При отказе выходной массив очищен.
Friend Function TryMinimizeStrain(ByRef force() As Double, ByRef forceResidual() As Double, _
        ByRef equality() As Double, ByRef equalityResidual() As Double, ByRef strain() As Double, _
        ByVal radius As Double, ByVal forceLimit As Double, ByRef correction() As Double) As Boolean
    Dim state() As Double, cap As Double, maximum As Double, i As Long
    On Error GoTo ArithmeticFailure
    Erase correction
    mForceNorm = 0#: mStrainNorm = 0#: mLinearSolveCount = 0
    If radius <= 0# Or forceLimit <= 0# Then Exit Function
    PrepareWork force, forceResidual, equality, equalityResidual, radius, forceLimit
    If LBound(strain, 1) <> 1 Or LBound(strain, 2) <> 1 Or UBound(strain, 1) <> mSize Then _
        Err.Raise 5, "CCoupledStepOptimizer", "Размер оператора деформаций не совпадает с касательной."
    cap = forceLimit / mForceScale
    If Not FindForceInterior(state, cap) Then Exit Function

    maximum = MaximumStrain(strain, state)
    If maximum > 0# Then
        mStrainScale = maximum
        If Not MinimizeMaterialStrain(strain, state, cap) Then Exit Function
    End If
    mForceNorm = ForceMaximum(state) * mForceScale
    mStrainNorm = MaximumStrain(strain, state)
    If mForceNorm > forceLimit * (1# + ROUNDING_GAP) Then Exit Function
    If Not VerifyConstraints(state) Then Exit Function
    ReDim correction(1 To mSize)
    For i = 1 To mSize
        correction(i) = state(i) * mRadius
    Next i
    TryMinimizeStrain = True
    Exit Function

ArithmeticFailure:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
    Erase correction
End Function

' Собирает парные неравенства в координатах единичного доверительного
' интервала. Состояние имеет добавочную переменную для максимума.
Private Sub PrepareWork(ByRef force() As Double, ByRef forceResidual() As Double, _
        ByRef equality() As Double, ByRef equalityResidual() As Double, _
        ByVal radius As Double, ByVal forceLimit As Double)
    Dim i As Long, j As Long
    mSize = UBound(force, 1): mForceCount = UBound(force, 2): mConstraintCount = UBound(equality, 2)
    If LBound(force, 1) <> 1 Or LBound(force, 2) <> 1 Or LBound(equality, 1) <> 1 Or LBound(equality, 2) <> 1 Or _
            LBound(forceResidual) <> 1 Or LBound(equalityResidual) <> 1 Or UBound(equality, 1) <> mSize Or _
            UBound(forceResidual) <> mForceCount Or UBound(equalityResidual) <> mConstraintCount Then _
        Err.Raise 5, "CCoupledStepOptimizer", "Операторы и невязки поправки имеют несовместимые размеры."
    mRadius = radius: mForceScale = forceLimit
    For i = 1 To mForceCount
        If Abs(forceResidual(i)) > mForceScale Then mForceScale = Abs(forceResidual(i))
    Next i
    mEquality = equality
    ReDim mEqualityOffset(1 To mConstraintCount)
    ReDim mRows(1 To mSize, 1 To mForceCount + MAX_STRAIN_ROWS)
    ReDim mOffsets(1 To mForceCount + MAX_STRAIN_ROWS), mBounds(1 To mForceCount + MAX_STRAIN_ROWS)
    ReDim mFreeBound(1 To mForceCount + MAX_STRAIN_ROWS)
    mRowCount = mForceCount
    For i = 1 To mConstraintCount
        mEqualityOffset(i) = equalityResidual(i) / radius
    Next i
    For i = 1 To mForceCount
        mOffsets(i) = forceResidual(i) / mForceScale
        mFreeBound(i) = True
        For j = 1 To mSize
            mRows(j, i) = force(j, i) * (radius / mForceScale)
        Next j
    Next i
    Set mSystem = New CCoupledLinearSystem
End Sub

' Получает допустимое по связям начальное поле малой нормы и ищет
' внутреннюю точку заданного ограничения по силам.
Private Function FindForceInterior(ByRef state() As Double, ByVal cap As Double) As Boolean
    Dim gram() As Double, weights() As Double, i As Long, j As Long, k As Long
    Dim barrier As Double, value As Double
    ReDim gram(1 To mConstraintCount, 1 To mConstraintCount), weights(1 To mConstraintCount)
    ReDim state(1 To mSize + 1)
    For i = 1 To mConstraintCount
        weights(i) = -mEqualityOffset(i)
        For j = 1 To i
            value = 0#
            For k = 1 To mSize
                value = value + mEquality(k, i) * mEquality(k, j)
            Next k
            gram(j, i) = value: gram(i, j) = value
        Next j
    Next i
    If Not mSystem.Solve(gram, weights, mConstraintCount) Then Exit Function
    mLinearSolveCount = mLinearSolveCount + 1
    For j = 1 To mSize
        value = 0#
        For i = 1 To mConstraintCount
            value = value + mEquality(j, i) * weights(i)
        Next i
        If Abs(value) >= 1# Then Exit Function
        state(j) = value
    Next j
    value = ForceMaximum(state)
    If value < cap Then
        FindForceInterior = True
        Exit Function
    End If
    state(mSize + 1) = 1# + value
    barrier = INITIAL_BARRIER
    Do
        If Not Center(state, barrier) Then Exit Function
        If state(mSize + 1) < cap Then
            FindForceInterior = True
            Exit Function
        End If
        barrier = barrier / 10#
    Loop While barrier >= MIN_FORCE_BARRIER * 0.99
End Function

' Добавляет ограничивающие компоненты деформаций порциями. После
' каждого решения проверяются все материальные строки, включая
' отсутствующие в текущей линейной системе.
Private Function MinimizeMaterialStrain(ByRef strain() As Double, ByRef state() As Double, ByVal cap As Double) As Boolean
    Dim selected() As Boolean, values() As Double, barrier As Double, peak As Double, i As Long
    ReDim selected(1 To UBound(strain, 2)), values(1 To UBound(strain, 2))
    For i = 1 To mForceCount
        mFreeBound(i) = False: mBounds(i) = cap
    Next i
    peak = ScanStrain(strain, state, values)
    state(mSize + 1) = peak + 0.05
    AddStrainRows strain, values, selected, -1#
    barrier = INITIAL_BARRIER
    Do
        Do
            If Not Center(state, barrier) Then Exit Function
            peak = ScanStrain(strain, state, values)
            If peak <= state(mSize + 1) * (1# + ROUNDING_GAP) Then Exit Do
            If Not AddStrainRows(strain, values, selected, state(mSize + 1)) Then Exit Function
            state(mSize + 1) = peak + 0.005
        Loop
        barrier = barrier / 10#
    Loop While barrier >= MIN_STRAIN_BARRIER * 0.99
    MinimizeMaterialStrain = True
End Function

' Сохраняет наиболее нарушенные строки ограничения деформаций.
Private Function AddStrainRows(ByRef strain() As Double, ByRef values() As Double, _
        ByRef selected() As Boolean, ByVal threshold As Double) As Boolean
    Dim entry As Long, index As Long, j As Long, position As Long, best As Double, factor As Double
    factor = mRadius / mStrainScale
    For entry = 1 To STRAIN_BATCH
        If mRowCount >= mForceCount + MAX_STRAIN_ROWS Then Exit Function
        best = threshold: position = 0
        For index = 1 To UBound(values)
            If Not selected(index) And values(index) > best Then
                best = values(index): position = index
            End If
        Next index
        If position = 0 Then Exit For
        mRowCount = mRowCount + 1
        selected(position) = True: mFreeBound(mRowCount) = True
        For j = 1 To mSize
            mRows(j, mRowCount) = strain(j, position) * factor
        Next j
        AddStrainRows = True
    Next entry
End Function

' Центрирует текущую систему неравенств при фиксированном барьере.
' Только уменьшение барьерной функции внутри области даёт новую точку.
Private Function Center(ByRef state() As Double, ByVal barrier As Double) As Boolean
    Dim epigraph As Long, count As Long, attempt As Long, i As Long
    Dim matrix() As Double, correction() As Double, trial() As Double, gradient() As Double
    Dim before As Double, after As Double, derivative As Double, fraction As Double, accepted As Boolean
    epigraph = mSize + 1: count = epigraph + mConstraintCount
    ReDim trial(1 To epigraph)
    Do While mLinearSolveCount < MAX_LINEAR_SOLVES
        If Not NewtonSystem(state, barrier, matrix, gradient, correction, before) Then Exit Function
        mLinearSolveCount = mLinearSolveCount + 1
        If Not mSystem.Solve(matrix, correction, count) Then Exit Function
        derivative = 0#
        For i = 1 To epigraph
            derivative = derivative + gradient(i) * correction(i)
        Next i
        If -derivative <= DECREMENT_TOLERANCE Then
            Center = True
            Exit Function
        End If
        fraction = 1#: accepted = False
        For attempt = 1 To MAX_BACKTRACK
            For i = 1 To epigraph
                trial(i) = state(i) + fraction * correction(i)
            Next i
            If BarrierValue(trial, barrier, after) Then
                If after <= before + 0.01 * fraction * derivative Then
                    accepted = True
                    Exit For
                End If
            End If
            fraction = fraction * 0.5
        Next attempt
        If Not accepted Then Exit Function
        state = trial
    Loop
End Function

' Формирует систему Ньютона с кинематическими равенствами. Положительные
' запасы двух знаков строки дают общий множитель для её внешнего произведения.
Private Function NewtonSystem(ByRef state() As Double, ByVal barrier As Double, _
        ByRef matrix() As Double, ByRef gradient() As Double, ByRef rightSide() As Double, _
        ByRef merit As Double) As Boolean
    Dim epigraph As Long, count As Long, i As Long, j As Long, k As Long, row As Long
    Dim bound As Double, ordinate As Double, left As Double, right As Double, p As Double, q As Double
    Dim diagonal As Double, interaction As Double, multiplier As Double
    epigraph = mSize + 1: count = epigraph + mConstraintCount
    ReDim matrix(1 To count, 1 To count), gradient(1 To epigraph), rightSide(1 To count)
    merit = state(epigraph): gradient(epigraph) = 1#
    For j = 1 To mSize
        left = 1# - state(j): right = 1# + state(j)
        If left <= 0# Or right <= 0# Then Exit Function
        merit = merit - barrier * (Log(left) + Log(right))
        p = 1# / left: q = 1# / right
        gradient(j) = barrier * (p - q)
        matrix(j, j) = barrier * (p * p + q * q)
    Next j
    For row = 1 To mRowCount
        bound = mBounds(row)
        If mFreeBound(row) Then bound = state(epigraph)
        ordinate = RowValue(row, state)
        left = bound - ordinate: right = bound + ordinate
        If left <= 0# Or right <= 0# Then Exit Function
        merit = merit - barrier * (Log(left) + Log(right))
        p = 1# / left: q = 1# / right
        diagonal = barrier * (p * p + q * q)
        interaction = -barrier * (p * p - q * q)
        multiplier = barrier * (p - q)
        For j = 1 To mSize
            ordinate = mRows(j, row)
            If ordinate <> 0# Then
                gradient(j) = gradient(j) + multiplier * ordinate
                bound = diagonal * ordinate
                For k = 1 To j
                    matrix(k, j) = matrix(k, j) + bound * mRows(k, row)
                Next k
                If mFreeBound(row) Then matrix(j, epigraph) = matrix(j, epigraph) + interaction * ordinate
            End If
        Next j
        If mFreeBound(row) Then
            gradient(epigraph) = gradient(epigraph) - barrier * (p + q)
            matrix(epigraph, epigraph) = matrix(epigraph, epigraph) + diagonal
        End If
    Next row
    For j = 1 To epigraph
        rightSide(j) = -gradient(j)
        For k = 1 To j - 1
            matrix(j, k) = matrix(k, j)
        Next k
    Next j
    For i = 1 To mConstraintCount
        row = epigraph + i: ordinate = mEqualityOffset(i)
        For j = 1 To mSize
            multiplier = mEquality(j, i)
            ordinate = ordinate + multiplier * state(j)
            matrix(j, row) = multiplier: matrix(row, j) = multiplier
        Next j
        rightSide(row) = -ordinate
    Next i
    NewtonSystem = True
End Function

' Проверяет область барьера без факторизации для демпфирования поправки.
Private Function BarrierValue(ByRef state() As Double, ByVal barrier As Double, ByRef merit As Double) As Boolean
    Dim i As Long, bound As Double, value As Double
    merit = state(mSize + 1)
    For i = 1 To mSize
        value = (1# - state(i)) * (1# + state(i))
        If value <= 0# Then Exit Function
        merit = merit - barrier * Log(value)
    Next i
    For i = 1 To mRowCount
        bound = mBounds(i)
        If mFreeBound(i) Then bound = state(mSize + 1)
        value = RowValue(i, state)
        If Abs(value) >= bound Then Exit Function
        merit = merit - barrier * (Log(bound - value) + Log(bound + value))
    Next i
    BarrierValue = True
End Function

' Возвращает подписанное значение одной активной строки.
Private Function RowValue(ByVal row As Long, ByRef state() As Double) As Double
    Dim j As Long, value As Double
    value = mOffsets(row)
    For j = 1 To mSize
        value = value + mRows(j, row) * state(j)
    Next j
    RowValue = value
End Function

' Возвращает максимум по слабым уравнениям и нормированным усилиям.
Private Function ForceMaximum(ByRef state() As Double) As Double
    Dim row As Long, value As Double
    For row = 1 To mForceCount
        value = Abs(RowValue(row, state))
        If value > ForceMaximum Then ForceMaximum = value
    Next row
End Function

' Проверяет доверительный интервал и независимо замыкает равенства.
Private Function VerifyConstraints(ByRef state() As Double) As Boolean
    Dim i As Long, j As Long, value As Double
    For j = 1 To mSize
        If Abs(state(j)) > 1# Then Exit Function
    Next j
    For i = 1 To mConstraintCount
        value = mEqualityOffset(i)
        For j = 1 To mSize
            value = value + mEquality(j, i) * state(j)
        Next j
        If Abs(value) > ROUNDING_GAP Then Exit Function
    Next i
    VerifyConstraints = True
End Function

' Вычисляет максимум приращения деформации по всем станциям.
Private Function MaximumStrain(ByRef strain() As Double, ByRef state() As Double) As Double
    Dim values() As Double
    ReDim values(1 To UBound(strain, 2))
    MaximumStrain = MaterialValues(strain, state, values) * mRadius
End Function

' Проверяет все строки в масштабе текущей материальной цели.
Private Function ScanStrain(ByRef strain() As Double, ByRef state() As Double, ByRef values() As Double) As Double
    Dim i As Long, factor As Double
    factor = mRadius / mStrainScale
    ScanStrain = MaterialValues(strain, state, values) * factor
    For i = 1 To UBound(values)
        values(i) = values(i) * factor
    Next i
End Function

' Умножает немасштабированные строки деформаций на безразмерное поле.
Private Function MaterialValues(ByRef strain() As Double, ByRef state() As Double, ByRef values() As Double) As Double
    Dim i As Long, j As Long, value As Double
    For i = 1 To UBound(values)
        value = 0#
        For j = 1 To mSize
            value = value + strain(j, i) * state(j)
        Next j
        values(i) = Abs(value)
        If values(i) > MaterialValues Then MaterialValues = values(i)
    Next i
End Function

' Возвращает прогноз максимальной силовой невязки в МПа.
Friend Property Get PredictedForceNorm() As Double
    PredictedForceNorm = mForceNorm
End Property

' Возвращает максимум приращения деформации материальной точки.
Friend Property Get PredictedStrainNorm() As Double
    PredictedStrainNorm = mStrainNorm
End Property

' Возвращает число линейных решений последней попытки.
Friend Property Get LinearSolveCount() As Long
    LinearSolveCount = mLinearSolveCount
End Property
