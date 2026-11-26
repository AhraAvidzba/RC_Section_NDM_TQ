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

' Готовит неизменяемую табличную ветвь ASD из этой же огибающей и деградации.
' Упрочнение разбито на 40 отрезков, спад — с шагом 0,05 по показателю экспоненты до 13,8.
' Ограничение таблицы при полной деформации 0,1 либо почти единичном повреждении не является критерием прочности.
Friend Function PrepareASDBackbone(ByVal modulus As Double) As CCoupledConcreteBackbone
    If mPeak <= 0# Or modulus <= 0# Then InvalidEnvelope "Подготовьте огибающую и положительный модуль бетона для ASD."
    Dim strain(1 To 318) As Double, nominal(1 To 318) As Double, effective(1 To 318) As Double
    Dim count As Long: count = 1
    Dim index As Long, plastic As Double, stress As Double, cohesion As Double, damage As Double, energy As Double, epsilon As Double
    If mPlasticPeak > 0# Then
        For index = 0 To 39
            plastic = mPlasticPeak * CDbl(index) / 40#
            Evaluate plastic, stress, cohesion, damage, energy
            epsilon = plastic + cohesion / modulus
            If epsilon > 0.1 Or damage > 1# - 0.000000000001 Then Exit For
            count = count + 1: strain(count) = epsilon: nominal(count) = stress: effective(count) = cohesion
        Next index
    End If
    For index = 0 To 276
        plastic = mPlasticPeak + CDbl(index) / 20# / mSoftening
        Evaluate plastic, stress, cohesion, damage, energy
        epsilon = plastic + cohesion / modulus
        If epsilon > 0.1 Or damage > 1# - 0.000000000001 Then Exit For
        count = count + 1: strain(count) = epsilon: nominal(count) = stress: effective(count) = cohesion
    Next index
    Dim minimumJump As Double, jump As Double
    For index = 2 To count
        jump = Abs(nominal(index) - nominal(index - 1))
        If jump > 0# Then
            If minimumJump = 0# Or jump < minimumJump Then minimumJump = jump
        End If
    Next index
    If count < 2 Or minimumJump <= 0# Then InvalidEnvelope "Огибающая не даёт пригодной табличной ветви ASD."
    Dim branch As New CCoupledConcreteBackbone
    branch.Initialize strain, nominal, effective, count, 0.000001 * minimumJump ' Тот же нижний уровень интерполяции, что у ASDConcrete3D: минимум ненулевого соседнего скачка напряжения.
    Set PrepareASDBackbone = branch
End Function

' Снимает точки от начала до нисходящей ветви. Полная деформация включает
' пластическое смещение и упругую долю от эффективного напряжения.
Friend Sub CopyCalibrationPoints(ByVal elasticModulus As Double, ByVal direction As Double, _
        ByRef strain() As Double, ByRef stress() As Double)
    If mPeak <= 0# Or elasticModulus <= 0# Or Abs(direction) <> 1# Then InvalidEnvelope "Недопустимые параметры вывода калибровки."
    Dim hardeningSteps As Long
    If mPlasticPeak > 0# Then hardeningSteps = 4
    Dim ratios As Variant: ratios = Array(0.9, 0.75, 0.5, 0.25, 0.1, 0.03, 0.01)
    Dim total As Long: total = 2 + hardeningSteps + UBound(ratios) + 1
    ReDim strain(1 To total): ReDim stress(1 To total)

    Dim point As Long, i As Long, plastic As Double
    point = 2
    PutCalibrationPoint point, 0#, elasticModulus, direction, strain, stress
    For i = 1 To hardeningSteps
        point = point + 1
        PutCalibrationPoint point, mPlasticPeak * i / hardeningSteps, elasticModulus, direction, strain, stress
    Next i
    For i = LBound(ratios) To UBound(ratios)
        plastic = mPlasticPeak - Log(CDbl(ratios(i))) / mSoftening
        If mDegradation * plastic + Log(CDbl(ratios(i))) > 700# Then Exit For ' Дальше полная деформация графика может выйти за диапазон Double.
        point = point + 1
        PutCalibrationPoint point, plastic, elasticModulus, direction, strain, stress
    Next i
    If point < total Then ReDim Preserve strain(1 To point): ReDim Preserve stress(1 To point)
End Sub

' Координата полной деформации использует эффективное напряжение напрямую.
Private Sub PutCalibrationPoint(ByVal point As Long, ByVal plastic As Double, ByVal modulus As Double, _
        ByVal direction As Double, ByRef strain() As Double, ByRef stress() As Double)
    Dim nominal As Double, effective As Double, degraded As Double, energy As Double
    Evaluate plastic, nominal, effective, degraded, energy
    strain(point) = direction * (plastic + effective / modulus)
    stress(point) = direction * nominal
End Sub

' Выдаёт проблему физической калибровки с источником закона.
Private Sub InvalidEnvelope(ByVal reason As String)
    Err.Raise vbObjectError + 1782, "CCoupledConcreteEnvelope", reason
End Sub
