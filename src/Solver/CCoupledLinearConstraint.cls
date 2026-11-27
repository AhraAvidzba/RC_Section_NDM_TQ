VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledLinearConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемое линейное условие управления в численных координатах
' ==========================================================================
' Коэффициенты и смещение принадлежат одному подготовленному контексту.
' Условие не читает материалы, не принимает историю и не меняет точность равновесия.
Implements ICoupledScalarConstraint

Private mGradient() As Double, mOffset As Double, mCount As Long

' Сохраняет независимую копию коэффициентов единственного линейного условия.
Friend Sub Initialize(ByRef gradient() As Double, ByVal offset As Double)
    If mCount <> 0 Then Err.Raise 5, "CCoupledLinearConstraint", "Линейное условие управления уже подготовлено."
    If LBound(gradient) <> 1 Or UBound(gradient) < 2 Then Err.Raise 5, "CCoupledLinearConstraint", "Неверный размер линейного условия управления."
    mGradient = gradient: mCount = UBound(gradient): mOffset = offset
End Sub

' Возвращает отклонение от линейного условия и при запросе независимую копию постоянного градиента.
Private Function ICoupledScalarConstraint_TryEvaluate(ByRef coordinates() As Double, ByVal withGradient As Boolean, _
        ByRef value As Double, ByRef gradient() As Double) As Boolean
    If mCount = 0 Then Err.Raise 5, "CCoupledLinearConstraint", "Линейное условие управления не подготовлено."
    If LBound(coordinates) <> 1 Or UBound(coordinates) <> mCount Then Err.Raise 5, "CCoupledLinearConstraint", "Размер численной пробы не совпадает с линейным условием."
    value = mOffset
    Dim i As Long
    For i = 1 To mCount: value = value + mGradient(i) * coordinates(i): Next i
    If withGradient Then gradient = mGradient
    ICoupledScalarConstraint_TryEvaluate = True
End Function
