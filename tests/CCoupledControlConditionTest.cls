VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledControlConditionTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: сравнение координат по линейной обусловленности
' ==========================================================================
' Первые n-1 строк общие; последняя строка задаёт текущую координату.
' Нормы строк приводятся к единице. Одна LU-факторизация обслуживает все
' кандидаты через замену последней строки обратной матрицы. Меньшая
' норма обратной матрицы соответствует лучшей оценке обусловленности.
' Оценка не подтверждает нелинейную сходимость или физический предел.

Private mCount As Long, mInverse() As Double, mNull() As Double, mNullLength As Double
Private mNorms As New CCoupledHybridFactorization

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Подготавливает обратную матрицу существующим LU; отказ означает только непригодную линейную опору для сравнения координат.
Friend Function TryInitializeTest(ByRef matrix() As Double, ByVal count As Long) As Boolean
    mCount = 0: Erase mInverse: Erase mNull
    If count < 2 Then Err.Raise 5, "CCoupledControlConditionTest", "Для сравнения координат нужны уравнения равновесия и строка управления."
    If LBound(matrix, 1) <> 1 Or UBound(matrix, 1) <> count Or _
            LBound(matrix, 2) <> 1 Or UBound(matrix, 2) <> count Then _
        Err.Raise 5, "CCoupledControlConditionTest", "Размер исходной матрицы не совпадает с числом уравнений."
    Dim normalized() As Double, row() As Double, length As Double, i As Long, j As Long
    ReDim normalized(1 To count, 1 To count): ReDim row(1 To count)
    For i = 1 To count
        For j = 1 To count: row(j) = matrix(j, i): Next j
        length = mNorms.VectorNorm(row, 1, count)
        If length = 0# Then Exit Function
        For j = 1 To count: normalized(j, i) = matrix(j, i) / length: Next j
    Next i
    Dim system As New CCoupledLinearSystem, right() As Double
    If Not system.Factor(normalized, count) Then Exit Function
    ReDim mInverse(1 To count, 1 To count): ReDim mNull(1 To count)
    For j = 1 To count
        ReDim right(1 To count): right(j) = 1#
        If Not system.SolveFactored(right) Then Exit Function
        For i = 1 To count: mInverse(j, i) = right(i): Next i ' Первый индекс хранит столбец обратной матрицы.
    Next j
    For i = 1 To count: mNull(i) = mInverse(count, i): Next i ' Последний столбец сохраняет первые n-1 однородных уравнений.
    mNullLength = mNorms.VectorNorm(mNull, 1, count)
    If mNullLength = 0# Then Exit Function
    mCount = count: TryInitializeTest = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает норму обратной матрицы с единичной строкой кандидата и её косинус с нулевым направлением равновесия.
' Все строки сравниваемых матриц имеют единичную норму; их общая норма Фробениуса равна sqrt(n).
Friend Function TryRankTest(ByRef gradient() As Double, ByRef inverseNorm As Double, ByRef alignment As Double) As Boolean
    inverseNorm = 0#: alignment = 0#
    If mCount = 0 Then Err.Raise 5, "CCoupledControlConditionTest", "Не подготовлена линейная опора для сравнения координат."
    If LBound(gradient) <> 1 Or UBound(gradient) <> mCount Then _
        Err.Raise 5, "CCoupledControlConditionTest", "Размер градиента не совпадает с подготовленной системой."
    Dim length As Double: length = mNorms.VectorNorm(gradient, 1, mCount)
    If length = 0# Then Exit Function
    Dim transformed() As Double, shift() As Double, i As Long, j As Long, denominator As Double
    ReDim transformed(1 To mCount): ReDim shift(1 To mCount)
    For j = 1 To mCount
        For i = 1 To mCount: transformed(j) = transformed(j) + (gradient(i) / length) * mInverse(j, i): Next i
    Next j
    denominator = transformed(mCount)
    If denominator = 0# Then Exit Function ' Кандидат не устраняет нулевое направление уравнений равновесия.
    alignment = Abs(denominator) / mNullLength
    For j = 1 To mCount: shift(j) = -transformed(j) / denominator: Next j
    shift(mCount) = (1# - transformed(mCount)) / denominator
    Dim entryScale As Double, sumSquares As Double, value As Double, magnitude As Double
    For j = 1 To mCount
        For i = 1 To mCount
            value = mInverse(j, i) + mNull(i) * shift(j) ' Замена одной строки по формуле Шермана — Моррисона.
            magnitude = Abs(value)
            If magnitude > entryScale Then
                sumSquares = 1# + sumSquares * (entryScale / magnitude) ^ 2: entryScale = magnitude
            ElseIf magnitude > 0# Then
                sumSquares = sumSquares + (magnitude / entryScale) ^ 2
            End If
        Next i
    Next j
    inverseNorm = entryScale * Sqr(sumSquares)
    TryRankTest = (inverseNorm > 0#)
End Function
