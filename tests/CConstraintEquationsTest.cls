VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CConstraintEquationsTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: аналитический отклик и круговая координата исключения
' ==========================================================================
' Корень около x=(0;0,6;0,6) отображается в (-0,2;0,6;0,6).
Implements ICoupledEquationSystem
Implements ICoupledScalarConstraint
Private mMode As String, mCalls As Long

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Выбирает аналитическую поверхность или случай отсутствия корня.
Friend Sub InitializeTest(ByVal mode As String)
    mMode = mode: mCalls = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число запросов именно силового отклика, отдельно от скалярной доводки.
Friend Property Get EquationCallsTest() As Long
    EquationCallsTest = mCalls
End Property

' Аналитический силовой отклик; третья строка намеренно не совпадает с условием однозначности адаптера.
Private Function ICoupledEquationSystem_TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
    mCalls = mCalls + 1
    ReDim residual(1 To 3)
    residual(1) = 1# + coordinates(1) + coordinates(2) - 1.4
    residual(2) = coordinates(3) - coordinates(2): residual(3) = 777#
    If withJacobian Then
        ReDim jacobian(1 To 3, 1 To 3)
        jacobian(1, 1) = 1#: jacobian(2, 1) = 1#: jacobian(2, 2) = -1#: jacobian(3, 2) = 1#
        jacobian(1, 3) = 123#: jacobian(3, 3) = 456#
    End If
    ICoupledEquationSystem_TryEvaluate = True
End Function

' Скалярная поверхность с точной производной и известными недопустимыми режимами.
Private Function ICoupledScalarConstraint_TryEvaluate(ByRef coordinates() As Double, ByVal withGradient As Boolean, _
        ByRef value As Double, ByRef gradient() As Double) As Boolean
    If mMode = "NoResponse" Then Exit Function
    If mMode = "NoRoot" Then
        value = (1# + coordinates(1)) ^ 2 + 1#
    Else
        value = (1# + coordinates(1)) ^ 2 + coordinates(2) ^ 2 - 1#
    End If
    If withGradient Then
        ReDim gradient(1 To 3)
        gradient(1) = 2# * (1# + coordinates(1))
        If mMode <> "NoRoot" Then gradient(2) = 2# * coordinates(2)
    End If
    ICoupledScalarConstraint_TryEvaluate = True
End Function
