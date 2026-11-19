VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHybridEquationsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: аналитические квадратные системы для гибридного метода
' ==========================================================================
' Не обращается к Excel и не использует модель материалов сечения.
Implements ICoupledEquationSystem
Private mMode As String
Private mEvaluationCallsTest As Long

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает систему с известным корнем, отсутствием корня либо отсутствием отклика.
Friend Sub InitializeTest(ByVal mode As String)
    mMode = mode: mEvaluationCallsTest = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число реально полученных интерфейсом запросов независимо от счётчиков решателя.
Friend Property Get EvaluationCallsTest() As Long
    EvaluationCallsTest = mEvaluationCallsTest
End Property

' Реализация тестового интерфейса: точные производные, без обновления исходных координат.
Private Function ICoupledEquationSystem_TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
    mEvaluationCallsTest = mEvaluationCallsTest + 1
    If mMode = "Cancel" Then Err.Raise 18, "CHybridEquationsTest", "Тестовая отмена расчёта."
    If mMode = "NoResponse" Then Exit Function
    Dim count As Long: count = UBound(coordinates)
    ReDim residual(1 To count)
    If withJacobian Then ReDim jacobian(1 To count, 1 To count)
    Dim i As Long
    Select Case mMode
        Case "Linear"
            residual(1) = 3# * coordinates(1) + 2# * coordinates(2) + 1#
            residual(2) = coordinates(1) + 4# * coordinates(2) + 7#
            If withJacobian Then
                jacobian(1, 1) = 3#: jacobian(2, 1) = 2#: jacobian(1, 2) = 1#: jacobian(2, 2) = 4#
            End If
        Case "Rosenbrock"
            residual(1) = 10# * (coordinates(2) - coordinates(1) ^ 2): residual(2) = 1# - coordinates(1)
            If withJacobian Then
                jacobian(1, 1) = -20# * coordinates(1): jacobian(2, 1) = 10#: jacobian(1, 2) = -1#
            End If
        Case "ScaledRosenbrock"
            Dim first As Double, second As Double
            first = 1000# * coordinates(1): second = 0.001 * coordinates(2)
            residual(1) = 10# * (second - first ^ 2): residual(2) = 1# - first
            If withJacobian Then
                jacobian(1, 1) = -20000# * first: jacobian(2, 1) = 0.01: jacobian(1, 2) = -1000#
            End If
        Case "Tridiagonal"
            For i = 1 To count
                residual(i) = (3# - 2# * coordinates(i)) * coordinates(i) + 1#
                If i > 1 Then residual(i) = residual(i) - coordinates(i - 1)
                If i < count Then residual(i) = residual(i) - 2# * coordinates(i + 1)
                If withJacobian Then
                    jacobian(i, i) = 3# - 4# * coordinates(i)
                    If i > 1 Then jacobian(i - 1, i) = -1#
                    If i < count Then jacobian(i + 1, i) = -2#
                End If
            Next i
        Case "NoRoot"
            residual(1) = coordinates(1) ^ 2 + 1#
            If withJacobian Then jacobian(1, 1) = 2# * coordinates(1)
        Case "Hinge"
            residual(1) = coordinates(1) - 1#: residual(2) = Abs(coordinates(2))
            If withJacobian Then jacobian(1, 1) = 1#: jacobian(2, 2) = Sgn(coordinates(2))
        Case "HingeCurve"
            residual(1) = coordinates(1) ^ 2 - 1#: residual(2) = coordinates(2) + 0.5 * Abs(coordinates(2))
            If withJacobian Then jacobian(1, 1) = 2# * coordinates(1): jacobian(2, 2) = 1# + 0.5 * Sgn(coordinates(2))
        Case "Degenerate"
            residual(1) = coordinates(1) + coordinates(2) - 3#: residual(2) = 2# * residual(1)
            If withJacobian Then
                jacobian(1, 1) = 1#: jacobian(2, 1) = 1#: jacobian(1, 2) = 2#: jacobian(2, 2) = 2#
            End If
        Case "BadDimensions"
            ReDim residual(1 To count + 1)
        Case Else
            Err.Raise 5, , "Неизвестная тестовая система: " & mMode
    End Select
    ICoupledEquationSystem_TryEvaluate = True
End Function
