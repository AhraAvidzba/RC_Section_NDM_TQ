VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledFieldSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Фактические поля связанного состояния в элементах сечения
' ==========================================================================
' Для центров бетонных элементов воспроизводит историю материала по полям
' принятых шагов. Эти точки вывода не входят в квадратуру равновесия.
' При локальном отказе принимает фактическое напряжение ближайшей бетонной
' точки квадратуры и сообщает об аппроксимации в комментарии состояния.
' Напряжения продольных стержней берутся из принятого состояния решателя.

Private mFields() As Double ' Строка элемента: εz, σz, τxz, τyz, γxz, γyz, w.
Private mCount As Long

' Заполняет снимок наблюдений в порядке CSectionModel. Пробные шаги
' итерационного решателя сюда не попадают.
Friend Sub Capture(ByVal section As CSectionModel, ByVal operators As CCoupledElasticState, _
        ByVal history As Collection, ByVal observer As CCoupledConcreteResponse, _
        ByVal concrete As CCoupledConcreteResponse, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef warning As String)
    If mCount > 0 Then Err.Raise 5, "CCoupledFieldSnapshot", "Поля принятого снимка уже заполнены."
    If history Is Nothing Then Err.Raise 5, "CCoupledFieldSnapshot", "Для вывода отсутствует принятая история."
    If history.Count = 0 Then Err.Raise 5, "CCoupledFieldSnapshot", "Для вывода отсутствуют принятые поля траектории."
    ReDim mFields(1 To 7, 1 To section.ConcreteCount + section.RebarCount)
    Dim field() As Double, strain(1 To 6) As Double, stress(1 To 6) As Double, warping As Double
    Dim stepIndex As Long, i As Long, lastStep As Long, affected As Long, firstElement As String
    Dim useQuadrature() As Boolean, x As Double, y As Double, pointX() As Double, pointY() As Double
    ReDim useQuadrature(1 To section.ConcreteCount)
    lastStep = history.Count
    For stepIndex = 1 To lastStep
        field = history.Item(stepIndex)
        For i = 1 To section.ConcreteCount
            If Not useQuadrature(i) Or stepIndex = lastStep Then
                x = section.ConcreteX(i): y = section.ConcreteY(i)
                operators.KinematicsAt x, y, field, strain, warping
                If Not useQuadrature(i) Then
                    If Not observer.EvaluateOutputPoint(i, strain, stress) Then
                        useQuadrature(i) = True: affected = affected + 1
                        If affected = 1 Then firstElement = section.ConcreteID(i)
                    End If
                End If
                If stepIndex = lastStep Then
                    If useQuadrature(i) Then concrete.CopyAcceptedStress NearestConcretePoint(operators, x, y, pointX, pointY), stress
                    SavePoint i, strain, stress(3), stress(5), stress(6), warping
                End If
            End If
        Next i
        observer.AcceptOutputPoints
    Next stepIndex
    For i = 1 To section.RebarCount
        operators.KinematicsAt section.RebarX(i), section.RebarY(i), field, strain, warping
        SavePoint section.ConcreteCount + i, strain, reinforcement.AcceptedSteelStress(operators.ConcretePointCount + i), _
            0#, 0#, warping ' Модель продольного стержня работает только на осевое напряжение.
    Next i
    If affected > 0 Then warning = "Для " & CStr(affected) & " бетонных элементов (первый: " & firstElement & _
        " в rngNDMSectionGeometry) напряжения приближённо представлены значениями ближайших точек интегрирования."
    mCount = section.ConcreteCount + section.RebarCount
End Sub

' Находит ближайшую точку исходного бетона по расстоянию в плоскости.
' Координаты вычисляет один раз, только если понадобится такой вывод.
Private Function NearestConcretePoint(ByVal operators As CCoupledElasticState, ByVal x As Double, ByVal y As Double, _
        ByRef pointX() As Double, ByRef pointY() As Double) As Long
    Dim point As Long, area As Double, dx As Double, dy As Double, distance As Double, nearest As Double
    If NearestConcretePointCount(pointX) = 0 Then
        ReDim pointX(1 To operators.ConcretePointCount): ReDim pointY(1 To operators.ConcretePointCount)
        For point = 1 To operators.ConcretePointCount: operators.Point point, pointX(point), pointY(point), area: Next point
    End If
    NearestConcretePoint = 1: nearest = (pointX(1) - x) ^ 2 + (pointY(1) - y) ^ 2
    For point = 2 To operators.ConcretePointCount
        dx = pointX(point) - x: dy = pointY(point) - y
        distance = dx * dx + dy * dy
        If distance < nearest Then nearest = distance: NearestConcretePoint = point
    Next point
End Function

' Различает ещё не подготовленные и заполненные координаты квадратуры.
Private Function NearestConcretePointCount(ByRef x() As Double) As Long
    On Error Resume Next
    NearestConcretePointCount = UBound(x)
End Function

' Сохраняет вектор одной точки в канонических единицах МПа и мм.
Private Sub SavePoint(ByVal index As Long, ByRef strain() As Double, ByVal axialStress As Double, _
        ByVal tauXz As Double, ByVal tauYz As Double, ByVal warping As Double)
    mFields(1, index) = strain(3): mFields(2, index) = axialStress
    mFields(3, index) = tauXz: mFields(4, index) = tauYz
    mFields(5, index) = strain(5): mFields(6, index) = strain(6): mFields(7, index) = warping
End Sub

' Возвращает осевую деформацию и осевое напряжение сохранённого элемента.
Friend Sub Axial(ByVal index As Long, ByRef strain As Double, ByRef stress As Double)
    strain = mFields(1, index): stress = mFields(2, index)
End Sub

' Возвращает два компонента сдвига и продольное перемещение в центре элемента.
Friend Sub Spatial(ByVal index As Long, ByRef tauXz As Double, ByRef tauYz As Double, _
        ByRef gammaXz As Double, ByRef gammaYz As Double, ByRef warping As Double)
    tauXz = mFields(3, index): tauYz = mFields(4, index)
    gammaXz = mFields(5, index): gammaYz = mFields(6, index): warping = mFields(7, index)
End Sub
