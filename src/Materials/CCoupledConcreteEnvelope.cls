VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledConcreteEnvelope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Одноосная огибающая в пластической координате для пространственного бетона
' ==========================================================================
' Калибровка программы: линейное упрочнение до заданного пика, после
' пика экспоненциальное разупрочнение. Площадь всей пластической диаграммы
' составляет G / l. Характерная длина вводится как физический параметр.
' Отдельная экспоненциальная кривая задаёт деградацию жёсткости.

Private mYield As Double, mPeak As Double, mPlasticPeak As Double
Private mHardening As Double, mAscendingEnergy As Double, mEnergy As Double
Private mSoftening As Double, mDegradation As Double

' Калибрует огибающую из сопротивлений, энергии и пластической деформации.
' DamageAtReference и PlasticDamageReference задают одну точку ветви D.
' При отсутствии упрочнения пластическая координата пика равна нулю.
Friend Sub Initialize(ByVal yieldStress As Double, ByVal peakStress As Double, _
        ByVal plasticPeak As Double, ByVal fractureEnergy As Double, ByVal localizationLength As Double, _
        ByVal damageAtReference As Double, ByVal plasticDamageReference As Double)
    If yieldStress <= 0# Or peakStress < yieldStress Or plasticPeak < 0# Or _
            fractureEnergy <= 0# Or localizationLength <= 0# Or damageAtReference < 0# Or damageAtReference >= 1# Or _
            plasticDamageReference <= 0# Then InvalidEnvelope "Недопустимые параметры одноосной огибающей."
    If (peakStress > yieldStress) <> (plasticPeak > 0#) Then _
        InvalidEnvelope "Ветвь упрочнения требует разных сопротивлений и положительной пластической деформации пика."
    Dim energy As Double, prePeakEnergy As Double
    energy = fractureEnergy / localizationLength
    prePeakEnergy = (yieldStress + peakStress) * plasticPeak * 0.5
    If energy <= prePeakEnergy Then InvalidEnvelope _
        "Энергия разрушения на характерную длину должна превышать работу пластического упрочнения до пика."
    mYield = yieldStress: mPeak = peakStress: mPlasticPeak = plasticPeak
    mEnergy = energy: mAscendingEnergy = prePeakEnergy
    mSoftening = peakStress / (energy - prePeakEnergy)
    mDegradation = -Log(1# - damageAtReference) / plasticDamageReference
    mHardening = 0#
    If plasticPeak > 0# Then mHardening = (peakStress - yieldStress) / plasticPeak
End Sub

' Возвращает требуемую прочность, сцепление и деградацию материала.
' Эффективное сцепление определяется до применения коэффициента жёсткости.
Friend Sub Evaluate(ByVal plastic As Double, ByRef stress As Double, _
        ByRef cohesion As Double, ByRef degraded As Double, ByRef energyRatio As Double)
    If mPeak <= 0# Or plastic < 0# Then InvalidEnvelope "Одноосная огибающая не подготовлена или заданы отрицательные пластические деформации."
    Dim exponent As Double, consumed As Double
    If plastic < mPlasticPeak Then
        stress = mYield + mHardening * plastic
        cohesion = stress * Exp(mDegradation * plastic)
        consumed = (mYield + stress) * plastic * 0.5
    Else
        exponent = -mSoftening * (plastic - mPlasticPeak)
        stress = mPeak * Exp(exponent)
        cohesion = mPeak * Exp(exponent + mDegradation * plastic)
        consumed = mAscendingEnergy + mPeak * (1# - Exp(exponent)) / mSoftening
    End If
    degraded = 1# - Exp(-mDegradation * plastic)
    energyRatio = consumed / mEnergy
End Sub

' Возвращает пластическую деформацию при половине напряжения на нисходящей ветви.
' Значение используется для воспроизводимой калибровки растягивающей ветви.
Friend Property Get HalfStrengthPlasticStrain() As Double
    HalfStrengthPlasticStrain = mPlasticPeak + Log(2#) / mSoftening
End Property

' Выдаёт проблему физической калибровки с источником закона.
Private Sub InvalidEnvelope(ByVal reason As String)
    Err.Raise vbObjectError + 1782, "CCoupledConcreteEnvelope", reason
End Sub
