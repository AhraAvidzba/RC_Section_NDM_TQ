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
' Метод LU с выбором ведущей строки допускает неспряжённую материальную
' касательную и множители ограничений. Переданные матрица и правая часть
' служат рабочими буферами. Столбец является первым индексом массива VBA.

Private Const MIN_SCALED_PIVOT As Double = 0.000000000001
Private mFailureCode As ESolverFailureCode
Private mMinimumPivot As Double

' Решает матрицу размера count после уравновешивания строк и столбцов.
' Не образует обратную матрицу. При отказе значения rightHandSide непригодны.
Friend Function Solve(ByRef matrix() As Double, ByRef rightHandSide() As Double, ByVal count As Long) As Boolean
    On Error GoTo ArithmeticFailure
    mFailureCode = sfcNone
    mMinimumPivot = 1#
    Dim row As Long, col As Long, k As Long, pivotRow As Long
    Dim rowMagnitude As Double, candidate As Double, pivot As Double, factor As Double
    Dim columnScale() As Double
    ReDim columnScale(1 To count)

    For col = 1 To count
        rowMagnitude = 0#
        For row = 1 To count
            candidate = Abs(matrix(col, row))
            If candidate > rowMagnitude Then rowMagnitude = candidate
        Next row
        If rowMagnitude = 0# Then GoTo Singular
        columnScale(col) = rowMagnitude
        For row = 1 To count
            matrix(col, row) = matrix(col, row) / rowMagnitude
        Next row
    Next col
    For row = 1 To count
        rowMagnitude = 0#
        For col = 1 To count
            candidate = Abs(matrix(col, row))
            If candidate > rowMagnitude Then rowMagnitude = candidate
        Next col
        If rowMagnitude = 0# Then GoTo Singular
        For col = 1 To count
            matrix(col, row) = matrix(col, row) / rowMagnitude
        Next col
        rightHandSide(row) = rightHandSide(row) / rowMagnitude
    Next row

    For k = 1 To count
        pivotRow = k
        rowMagnitude = Abs(matrix(k, k))
        For row = k + 1 To count
            candidate = Abs(matrix(k, row))
            If candidate > rowMagnitude Then
                rowMagnitude = candidate
                pivotRow = row
            End If
        Next row
        If rowMagnitude < mMinimumPivot Then mMinimumPivot = rowMagnitude
        If rowMagnitude <= MIN_SCALED_PIVOT Then GoTo Singular
        If pivotRow <> k Then
            For col = k To count
                candidate = matrix(col, k)
                matrix(col, k) = matrix(col, pivotRow)
                matrix(col, pivotRow) = candidate
            Next col
            candidate = rightHandSide(k)
            rightHandSide(k) = rightHandSide(pivotRow)
            rightHandSide(pivotRow) = candidate
        End If
        pivot = matrix(k, k)
        For row = k + 1 To count
            If matrix(k, row) <> 0# Then
                factor = matrix(k, row) / pivot
                For col = k + 1 To count
                    matrix(col, row) = matrix(col, row) - factor * matrix(col, k)
                Next col
                rightHandSide(row) = rightHandSide(row) - factor * rightHandSide(k)
            End If
        Next row
    Next k

    For row = count To 1 Step -1
        candidate = rightHandSide(row)
        For col = row + 1 To count
            candidate = candidate - matrix(col, row) * rightHandSide(col)
        Next col
        rightHandSide(row) = candidate / matrix(row, row)
    Next row
    For col = 1 To count
        rightHandSide(col) = rightHandSide(col) / columnScale(col)
    Next col
    Solve = True
    Exit Function

Singular:
    mFailureCode = sfcSingularTangent
    Exit Function
ArithmeticFailure:
    If Err.Number = 6 Or Err.Number = 11 Then
        mFailureCode = sfcNumericalFailure
    Else
        mFailureCode = sfcInternalError
    End If
    Err.Clear
End Function

' Возвращает код отказа последней факторизации или sfcNone после решения.
Friend Property Get FailureCode() As ESolverFailureCode
    FailureCode = mFailureCode
End Property

' Возвращает диагностический минимум модулей масштабированных ведущих элементов.
Friend Property Get MinimumPivot() As Double
    MinimumPivot = mMinimumPivot
End Property
