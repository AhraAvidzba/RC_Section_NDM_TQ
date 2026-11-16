VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionPlanarPreparation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =====================================================================
' Подготовка одноконтурной послетрещинной модели для сочетаний с T.
' Угол раскосов 45 градусов, площадь потока равна площади внутри оси
' хомута. Оба условия являются допущениями упрощённого расчёта.
' =====================================================================

Private mDemand As CTorsionTrussDemand
Private mHoop As CTorsionHoopGeometry
Private mConcreteStress() As Double, mSteelStress() As Double ' Смещения напряжений на 1 Н продольной тяги.
Private mFieldError As String

' Готовит геометрию и коэффициенты один раз на сечение. Отказ
' распределения осевых напряжений не мешает отдельно проверить хомут.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal diameter As Double, ByVal spacing As Double, _
        ByVal cover As Double, ByVal hoopResistance As Double)
    Dim hoop As New CTorsionHoopGeometry
    hoop.InitializeAutomatic section, cover, diameter
    Dim demand As New CTorsionTrussDemand
    demand.Initialize hoop.EnclosedArea, diameter, spacing, hoopResistance, 1# ' Котангенс принятого угла 45 градусов.
    Set mHoop = hoop: Set mDemand = demand
    PrepareStressField section, cover, diameter
End Sub

' Подготавливает продольную тягу и противодействующее бетонное сжатие.
' Ширина полосы отсчитывается от оси хомута в плоскости сечения.
Private Sub PrepareStressField(ByVal section As CSectionModel, ByVal cover As Double, ByVal diameter As Double)
    On Error GoTo Unavailable
    Dim bandWidth As Double: bandWidth = 2# * cover + diameter ' Удвоенное расстояние от внешней границы до оси хомута.
    Dim longitudinal As New CTorsionLongitudinalDemand, compression As New CTorsionConcreteBalance
    longitudinal.Initialize section, mHoop, bandWidth
    compression.Initialize section, mHoop, bandWidth, longitudinal.ResultantX, longitudinal.ResultantY
    longitudinal.CopyStressPerUnitForce mSteelStress
    compression.CopyStressPerUnitForce mConcreteStress
    Exit Sub
Unavailable:
    If Err.Number = 18 Then Err.Raise Err.Number, Err.Source, Err.Description
    mFieldError = Err.Description
End Sub

' Возвращает поток, напряжение и запас хомута, а также суммарную
' продольную тягу. При T=0 эту проверку не запускают.
Friend Sub Evaluate(ByVal torque As Double, ByRef flow As Double, ByRef hoopStress As Double, _
        ByRef hoopReserve As Double, ByRef longitudinalForce As Double)
    Dim traction As Double
    mDemand.Evaluate torque, flow, traction, hoopStress
    hoopReserve = mDemand.HoopMomentLimit / Abs(torque)
    longitudinalForce = traction * mHoop.Length
End Sub

' Переводит самоуравновешенные напряжения в исходные деформации по
' физическим ветвям диаграмм. Отказ материала отмечается отдельно от
' неприменимости распределения или последующей численной несходимости.
Friend Function BuildInitialStrains(ByVal section As CSectionModel, ByVal longitudinalForce As Double, _
        ByVal concrete As CMaterialDiagram, ByVal steel As CMaterialDiagram, ByRef limitDetails As String) As CSectionInitialStrains
    limitDetails = vbNullString
    Dim concreteStrains() As Double, steelStrains() As Double, i As Long
    ReDim concreteStrains(1 To section.ConcreteCount), steelStrains(1 To section.RebarCount)
    For i = 1 To section.ConcreteCount
        If mConcreteStress(i) <> 0# Then
            If Not concrete.TryInitialStrainForStress(mConcreteStress(i) * longitudinalForce, concreteStrains(i)) Then
                limitDetails = "Бетонный элемент " & section.ConcreteID(i) & " не воспринимает начальное сжатие от T по физической диаграмме."
                Exit Function
            End If
        End If
    Next i
    For i = 1 To section.RebarCount
        If mSteelStress(i) <> 0# Then
            If Not steel.TryInitialStrainForStress(mSteelStress(i) * longitudinalForce, steelStrains(i)) Then
                limitDetails = "Продольный стержень " & section.RebarID(i) & " не воспринимает начальное растяжение от T по физической диаграмме."
                Exit Function
            End If
        End If
    Next i
    Dim state As New CSectionInitialStrains
    state.Initialize section, concreteStrains, steelStrains
    Set BuildInitialStrains = state
End Function

' Возвращает причину отсутствия добавок при доступной проверке хомута.
Friend Property Get FieldError() As String
    FieldError = mFieldError
End Property
