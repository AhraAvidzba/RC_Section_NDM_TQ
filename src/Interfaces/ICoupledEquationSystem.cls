VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ICoupledEquationSystem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Читающий численный отклик полной системы нелинейного равновесия
' ==========================================================================
' Координаты заранее масштабированы вызывающей стороной. Невязка и её
' производная имеют одну согласованную нормировку; первый индекс матрицы
' является столбцом. Проба не принимает и не изменяет историю материалов.

' Возвращает невязку и, при запросе, квадратную матрицу её производных.
' Массив координат доступен только для чтения. False означает отсутствие
' допустимого численного отклика, а не физическое разрушение материала.
Public Function TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
End Function
