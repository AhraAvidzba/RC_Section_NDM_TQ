VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledLinearSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Линейная система для связанной постановки сечения
' ==========================================================================
' LU с выбором ведущей строки допускает неспряжённую касательную.
' Готовая факторизация обслуживает несколько правых частей без повторного
' исключения. Столбец является первым индексом матриц VBA.

Private Const MIN_SCALED_PIVOT As Double = 0.000000000001
Private mLu() As Double
Private mColumnScale() As Double
Private mRowScale() As Double
Private mPivot() As Long
Private mCount As Long
Private mFailureCode As ESolverFailureCode
Private mMinimumPivot As Double

' Решает новую систему, сразу выполняя факторизацию и обратную подстановку.
Friend Function Solve(ByRef matrix() As Double, ByRef rightHandSide() As Double, ByVal count As Long) As Boolean
    If Not Factor(matrix, count) Then Exit Function
    Solve = SolveFactored(rightHandSide)
End Function

' Факторизует копию исходной матрицы после масштабирования её столбцов
' и строк. Предыдущая факторизация теряет силу в начале нового вызова.
Friend Function Factor(ByRef matrix() As Double, ByVal count As Long) As Boolean
    On Error GoTo ArithmeticFailure
    mCount = 0
    mFailureCode = sfcNone
    mMinimumPivot = 1#
    If count < 1 Then GoTo Singular

    mLu = matrix
    ReDim mColumnScale(1 To count): ReDim mRowScale(1 To count)
    ReDim mPivot(1 To count)
    Dim row As Long, col As Long, k As Long, pivotRow As Long
    Dim magnitude As Double, candidate As Double, pivot As Double, multiplier As Double

    For col = 1 To count
        magnitude = 0#
        For row = 1 To count
            candidate = Abs(mLu(col, row))
            If candidate > magnitude Then magnitude = candidate
        Next row
        If magnitude = 0# Then GoTo Singular
        mColumnScale(col) = magnitude
        For row = 1 To count
            mLu(col, row) = mLu(col, row) / magnitude
        Next row
    Next col
    For row = 1 To count
        magnitude = 0#
        For col = 1 To count
            candidate = Abs(mLu(col, row))
            If candidate > magnitude Then magnitude = candidate
        Next col
        If magnitude = 0# Then GoTo Singular
        mRowScale(row) = magnitude
        For col = 1 To count
            mLu(col, row) = mLu(col, row) / magnitude
        Next col
    Next row

    For k = 1 To count
        pivotRow = k
        magnitude = Abs(mLu(k, k))
        For row = k + 1 To count
            candidate = Abs(mLu(k, row))
            If candidate > magnitude Then
                magnitude = candidate
                pivotRow = row
            End If
        Next row
        If magnitude < mMinimumPivot Then mMinimumPivot = magnitude
        If magnitude <= MIN_SCALED_PIVOT Then GoTo Singular

        mPivot(k) = pivotRow
        If pivotRow <> k Then
            For col = 1 To count
                candidate = mLu(col, k)
                mLu(col, k) = mLu(col, pivotRow)
                mLu(col, pivotRow) = candidate
            Next col
        End If
        pivot = mLu(k, k)
        For row = k + 1 To count
            If mLu(k, row) <> 0# Then
                multiplier = mLu(k, row) / pivot
                mLu(k, row) = multiplier ' Строго нижняя часть хранит множители прямой подстановки.
                For col = k + 1 To count
                    mLu(col, row) = mLu(col, row) - multiplier * mLu(col, k)
                Next col
            End If
        Next row
    Next k
    mCount = count
    Factor = True
    Exit Function

Singular:
    mFailureCode = sfcSingularTangent
    Exit Function
ArithmeticFailure:
    RecordArithmeticFailure Err.Number
End Function

' Решает одну правую часть подготовленной системы и восстанавливает
' физический масштаб неизвестных. При отказе массив результата непригоден.
Friend Function SolveFactored(ByRef rightHandSide() As Double) As Boolean
    If mCount = 0 Then Exit Function
    On Error GoTo ArithmeticFailure
    mFailureCode = sfcNone
    Dim row As Long, col As Long, pivot As Long, candidate As Double
    For row = 1 To mCount
        rightHandSide(row) = rightHandSide(row) / mRowScale(row)
    Next row
    For row = 1 To mCount
        pivot = mPivot(row)
        If pivot <> row Then
            candidate = rightHandSide(row)
            rightHandSide(row) = rightHandSide(pivot)
            rightHandSide(pivot) = candidate
        End If
    Next row
    For row = 2 To mCount
        candidate = rightHandSide(row)
        For col = 1 To row - 1
            candidate = candidate - mLu(col, row) * rightHandSide(col)
        Next col
        rightHandSide(row) = candidate
    Next row
    For row = mCount To 1 Step -1
        candidate = rightHandSide(row)
        For col = row + 1 To mCount
            candidate = candidate - mLu(col, row) * rightHandSide(col)
        Next col
        rightHandSide(row) = candidate / mLu(row, row)
    Next row
    For col = 1 To mCount
        rightHandSide(col) = rightHandSide(col) / mColumnScale(col)
    Next col
    SolveFactored = True
    Exit Function

ArithmeticFailure:
    RecordArithmeticFailure Err.Number
End Function

' Сохраняет машинную причину численной или внутренней ошибки арифметики.
Private Sub RecordArithmeticFailure(ByVal errorNumber As Long)
    If errorNumber = 6 Or errorNumber = 11 Then
        mFailureCode = sfcNumericalFailure
    Else
        mFailureCode = sfcInternalError
    End If
End Sub

' Возвращает код отказа вычисления или sfcNone после решения.
Friend Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property

' Возвращает минимум модулей масштабированных ведущих элементов.
Friend Property Get MinimumPivot() As Double
    MinimumPivot = mMinimumPivot
End Property
