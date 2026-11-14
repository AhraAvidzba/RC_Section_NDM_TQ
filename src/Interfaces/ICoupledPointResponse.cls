VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoupledPointResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Конститутивный отклик материальных точек связанного сечения
' ==========================================================================
' Реализация хранит пробную и принятую истории отдельно для всех точек.
' Обновление одной точки читает только её принятую историю. Для точек
' замещения возвращается разность откликов стали и исходного бетона.

' Возвращает пробные напряжения и касательную для заданной деформации.
' Напряжения в МПа, вектор и матрица в порядке xx, yy, zz, xy, xz, yz.
' Неуспешная материальная проба не меняет принятое состояние.
' При ограниченном режиме вызывающий использует только запрошенные столбцы.
Public Function EvaluatePoint(ByVal point As Long, ByRef strain() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double, _
        Optional ByVal tangentMode As ECoupledTangentMode = ctmFull) As ESolverFailureCode
End Function

' Принимает полностью вычисленный пробный массив. Реализация публикует
' новую историю за одну операцию, без материальных вычислений внутри.
Public Sub AcceptTrial()
End Sub

' Возвращает все точки к ненагруженному состоянию нового сочетания.
Public Sub ClearHistory()
End Sub
