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
' Напряжения продольных стержней берутся из принятого состояния решателя.

Private mFields() As Double ' Строка элемента: εz, σz, τxz, τyz, γxz, γyz, w.
Private mCount As Long

' Заполняет снимок наблюдений в порядке CSectionModel, при отказе оставляет
' его пустым. Пробные шаги итерационного решателя сюда не попадают.
Friend Function Capture(ByVal section As CSectionModel, ByVal operators As CCoupledElasticState, _
        ByVal continuation As CCoupledLoadContinuation, ByVal observer As CCoupledConcreteResponse, _
        ByVal reinforcement As CCoupledReinforcedResponse, ByRef failure As String) As Boolean
    If mCount > 0 Then Err.Raise 5, "CCoupledFieldSnapshot", "Поля принятого снимка уже заполнены."
    If continuation.SavedFieldCount = 0 Then Err.Raise 5, "CCoupledFieldSnapshot", "Для вывода отсутствуют принятые поля траектории."
    ReDim mFields(1 To 7, 1 To section.ConcreteCount + section.RebarCount)
    Dim field() As Double, strain(1 To 6) As Double, stress(1 To 6) As Double, warping As Double
    Dim stepIndex As Long, i As Long, lastStep As Long
    lastStep = continuation.SavedFieldCount
    For stepIndex = 1 To lastStep
        continuation.CopySavedField stepIndex, field
        For i = 1 To section.ConcreteCount
            operators.KinematicsAt section.ConcreteX(i), section.ConcreteY(i), field, strain, warping
            If Not observer.EvaluateOutputPoint(i, strain, stress) Then
                failure = "Не удалось восстановить напряжения бетона для " & section.ConcreteID(i) & _
                    " в rngNDMSectionGeometry на принятом шаге " & CStr(stepIndex) & "."
                Erase mFields
                Exit Function
            End If
            If stepIndex = lastStep Then SavePoint i, strain, stress(3), stress(5), stress(6), warping
        Next i
        observer.AcceptOutputPoints
    Next stepIndex
    For i = 1 To section.RebarCount
        operators.KinematicsAt section.RebarX(i), section.RebarY(i), field, strain, warping
        SavePoint section.ConcreteCount + i, strain, reinforcement.AcceptedSteelStress(operators.ConcretePointCount + i), _
            0#, 0#, warping ' Модель продольного стержня работает только на осевое напряжение.
    Next i
    mCount = section.ConcreteCount + section.RebarCount
    Capture = True
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
