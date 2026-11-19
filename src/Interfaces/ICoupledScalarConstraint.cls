VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoupledScalarConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Читающая скалярная координата нелинейного равновесия
' ==========================================================================
' Значение равно нулю на заданном уровне. Градиент относится к тем же
' масштабированным координатам, что ICoupledEquationSystem. Чтение не
' принимает историю материала; False не означает физическое разрушение.

' Возвращает отклонение от заданного уровня и, при запросе, его градиент.
Public Function TryEvaluate(ByRef coordinates() As Double, ByVal withGradient As Boolean, _
        ByRef value As Double, ByRef gradient() As Double) As Boolean
End Function
