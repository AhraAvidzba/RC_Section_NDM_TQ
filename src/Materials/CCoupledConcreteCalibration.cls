VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledConcreteCalibration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Общая калибровка одноосных огибающих пространственного бетона
' ==========================================================================
' CDP и ASD используют одни параметры сопротивлений, пика, энергий и деградации.
' Это инженерная калибровка проекта; пространственные законы при этом различаются.
' Объект готовится один раз и не содержит истории деформирования материала.

Private mTension As CCoupledConcreteEnvelope, mCompression As CCoupledConcreteEnvelope
Private mModulus As Double, mReady As Boolean

' Принимает модуль и прочности в МПа, энергии в Н/мм, характерную длину в мм.
' Формулы пика и деградации совпадают с калибровкой пространственного CDP.
Friend Sub Initialize(ByVal modulus As Double, ByVal fc0 As Double, ByVal fc As Double, ByVal ft As Double, _
        ByVal peakCompression As Double, ByVal gt As Double, ByVal gc As Double, ByVal localizationLength As Double)
    If mReady Then Err.Raise 5, "CCoupledConcreteCalibration", "Калибровка неизменяема. Создайте новый объект."
    If modulus <= 0# Or ft <= 0# Or fc0 <= ft Or fc <= fc0 Or gt <= 0# Or gc <= 0# Or localizationLength <= 0# Then _
        Err.Raise 5, "CCoupledConcreteCalibration", "Укажите положительные модуль, энергии, длину и прочности ft < fc0 < fc."
    If peakCompression <= fc / modulus Then Err.Raise 5, "CCoupledConcreteCalibration", "Деформация пика сжатия должна превышать упругую деформацию при fc."
    Dim halfTension As Double, compressDamage As Double, plasticPeak As Double
    halfTension = Log(2#) * gt / (ft * localizationLength)
    compressDamage = 0.423 * (1# - fc / (peakCompression * modulus))
    plasticPeak = peakCompression - fc / ((1# - compressDamage) * modulus)
    Set mTension = New CCoupledConcreteEnvelope: Set mCompression = New CCoupledConcreteEnvelope
    mTension.Initialize ft, ft, 0#, gt, localizationLength, 0.5, halfTension
    mCompression.Initialize fc0, fc, plasticPeak, gc, localizationLength, compressDamage, plasticPeak
    mModulus = modulus: mReady = True
End Sub

' Передаёт подготовленную огибающую растяжения для чтения материальным законом.
Friend Property Get TensionEnvelope() As CCoupledConcreteEnvelope
    Set TensionEnvelope = mTension
End Property

' Передаёт подготовленную огибающую сжатия для чтения материальным законом.
Friend Property Get CompressionEnvelope() As CCoupledConcreteEnvelope
    Set CompressionEnvelope = mCompression
End Property

' Создаёт закон ASD по тем же огибающим; однократное обновление соответствует ASDConcrete3D без IMPL-EX.
' Замкнутая интеграция доступна отдельным явным аргументом и имеет собственный LawIdentifier.
Friend Function CreateASDFlow(ByVal poisson As Double, Optional ByVal nativeSingleUpdate As Boolean = True) As CCoupledASDFlow
    If Not mReady Then Err.Raise 5, "CCoupledConcreteCalibration", "Сначала подготовьте параметры бетона."
    Dim tension As CCoupledConcreteBackbone, compression As CCoupledConcreteBackbone
    Set tension = mTension.PrepareASDBackbone(mModulus): Set compression = mCompression.PrepareASDBackbone(mModulus)
    Dim flow As New CCoupledASDFlow
    flow.Initialize tension, compression, mModulus, poisson, nativeSingleUpdate
    Set CreateASDFlow = flow
End Function
