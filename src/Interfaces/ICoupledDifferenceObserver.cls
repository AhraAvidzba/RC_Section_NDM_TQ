VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoupledDifferenceObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Подтверждение применимости действительно вычисленного разностного отрезка
' ==========================================================================
' Координаты относятся к действующему неизменному контексту исходного отклика.
' Наблюдатель не изменяет массивы, уравнения или принятую историю.

' Разрешает сохранить пару разностных проб как наблюдённый переход для последующего численного восстановления.
Public Function AcceptBracket(ByRef before() As Double, ByRef after() As Double) As Boolean
End Function
