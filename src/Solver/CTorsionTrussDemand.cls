VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionTrussDemand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------------
' Усилия в одноконтурной пространственной ферме от заданного кручения.
' Геометрия и наклон бетонных раскосов задаются до расчёта. Проверка
' ограничена спросом на арматуру в послетрещинной схеме. Предел бетона,
' деформации кручения и образование трещин определяются отдельно.
' Внутренние единицы: Н, мм, Н*мм и МПа.
' ----------------------------------------------------------------------

Private mFlowPerMoment As Double ' Поток касательных усилий на единицу момента.
Private mLongitudinalForcePerMomentLength As Double ' Тяга на единицу момента и длины контура.
Private mHoopStressPerMoment As Double ' Напряжение в одном сечении замкнутого хомута на единицу момента.
Private mHoopMomentLimit As Double ' Момент, отвечающий достижению расчётного сопротивления хомута.

' Подготавливает коэффициенты для всех сочетаний одного сечения.
' shearFlowArea — площадь внутри оси расчётного потока. strutCotangent —
' котангенс угла бетонного раскоса к продольной оси стержня.
Friend Sub Initialize(ByVal shearFlowArea As Double, ByVal hoopDiameter As Double, ByVal hoopSpacing As Double, _
        ByVal hoopResistance As Double, ByVal strutCotangent As Double)
    If shearFlowArea <= 0# Then Err.Raise vbObjectError + 4500, TypeName(Me), "Площадь контура потока должна быть больше нуля."
    If hoopDiameter <= 0# Or hoopSpacing <= 0# Then _
        Err.Raise vbObjectError + 4501, TypeName(Me), "Диаметр и шаг замкнутого хомута должны быть больше нуля."
    If hoopResistance <= 0# Then Err.Raise vbObjectError + 4502, TypeName(Me), "Расчётное сопротивление хомута должно быть больше нуля."
    If strutCotangent <= 0# Then Err.Raise vbObjectError + 4503, TypeName(Me), "Угол бетонного раскоса должен лежать между 0° и 90°."

    Dim hoopLegArea As Double
    hoopLegArea = Atn(1#) * hoopDiameter * hoopDiameter ' Площадь одной ветви замкнутого стержня равна pi*d^2/4.
    mFlowPerMoment = 0.5 / shearFlowArea ' Равновесие потока по замкнутому контуру T = 2 Ao q.
    mLongitudinalForcePerMomentLength = mFlowPerMoment * strutCotangent
    mHoopStressPerMoment = mFlowPerMoment * hoopSpacing / (hoopLegArea * strutCotangent)
    mHoopMomentLimit = hoopResistance / mHoopStressPerMoment
End Sub

' Возвращает модуль потока, продольную тягу на единицу длины контура
' и напряжение хомута. При смене знака T меняется направление потока,
' а растягивающие усилия в продольных стержнях и хомуте сохраняют знак.
Friend Sub Evaluate(ByVal torque As Double, ByRef shearFlow As Double, ByRef longitudinalForcePerLength As Double, _
        ByRef hoopStress As Double)
    If mHoopMomentLimit <= 0# Then Err.Raise vbObjectError + 4504, TypeName(Me), "Сначала задайте геометрию и свойства замкнутого хомута."
    Dim magnitude As Double
    magnitude = Abs(torque)
    shearFlow = magnitude * mFlowPerMoment
    longitudinalForcePerLength = magnitude * mLongitudinalForcePerMomentLength
    hoopStress = magnitude * mHoopStressPerMoment
End Sub

' Возвращает предельный модуль T только по сопротивлению хомута.
Friend Property Get HoopMomentLimit() As Double
    HoopMomentLimit = mHoopMomentLimit
End Property
