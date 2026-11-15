VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledLoadTrace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Параметры принятых шагов связанного расчёта
' ==========================================================================
' Порядок совпадает с принятой историей материалов. Каждая точка хранит
' долю постоянной нагрузки и множитель переменной части.

Private mPoints As Collection

' Подготавливает пустую последовательность одной ветви.
Private Sub Class_Initialize()
    Set mPoints = New Collection
End Sub

' Дописывает подтверждённое состояние при известном параметре нагружения.
Friend Sub Add(ByVal offsetFraction As Double, ByVal lambda As Double)
    mPoints.Add Array(offsetFraction, lambda)
End Sub

' Возвращает число подтверждённых точек текущей последовательности.
Friend Property Get Count() As Long
    Count = mPoints.Count
End Property

' Читает параметры одной подтверждённой точки по порядку нагружения.
Friend Sub ReadPoint(ByVal index As Long, ByRef offsetFraction As Double, ByRef lambda As Double)
    Dim data As Variant: data = mPoints.Item(index)
    offsetFraction = data(0): lambda = data(1)
End Sub

' Создаёт независимую ветвь из принятого префикса истории.
Friend Function CopyPrefix(ByVal count As Long) As CCoupledLoadTrace
    Dim result As New CCoupledLoadTrace, i As Long, offsetFraction As Double, lambda As Double
    For i = 1 To count
        ReadPoint i, offsetFraction, lambda
        result.Add offsetFraction, lambda
    Next i
    Set CopyPrefix = result
End Function
