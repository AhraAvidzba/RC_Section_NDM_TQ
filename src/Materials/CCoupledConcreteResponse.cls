VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledConcreteResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Пространственный бетон с пластичностью и раздельным повреждением
' ==========================================================================
' Эффективные напряжения подчиняются поверхности Барселоны и гиперболическому
' потенциалу. Две пластические координаты хранят историю энергий растяжения
' и сжатия. Поправка повреждения учитывает частичное закрытие трещин.
' Истории исходного и замещённого бетона хранятся одинаково.

Implements ICoupledPointResponse

Private Type TConcreteHistory
    Plastic(1 To 6) As Double
    Tension As Double
    Compression As Double
End Type

Private Const N_COMPONENTS As Long = 6
Private Const MATERIAL_ITERATIONS As Long = 70
Private Const STATE_TOLERANCE As Double = 0.0000000005
Private Const STRAIN_PERTURBATION As Double = 0.000000002
Private Const MIXED_DAMAGE_WEIGHT As Double = 20#
Private Const MAXIMUM_DAMPING As Double = 0.5

Private mHistory() As TConcreteHistory
Private mCommitted As Long, mTrial As Long, mPointCount As Long
Private mElastic(1 To 6, 1 To 6) As Double
Private mReturn As CCoupledConcreteReturn
Private mTension As CCoupledConcreteEnvelope, mCompression As CCoupledConcreteEnvelope
Private mFt As Double, mFc As Double, mAlpha As Double, mGamma As Double
Private mFrictionMax As Double, mDilationMax As Double

' Готовит единственный набор параметров бетона для текущей секции.
' Прочности в МПа, G в Н/мм, характерная длина в мм, углы в градусах.
' Деформация пика сжатия задаётся положительной величиной.
Friend Sub Initialize(ByVal pointCount As Long, ByVal modulus As Double, ByVal poisson As Double, _
        ByVal fc0 As Double, ByVal fc As Double, ByVal ft As Double, ByVal peakCompression As Double, _
        ByVal gt As Double, ByVal gc As Double, ByVal localizationLength As Double, _
        Optional ByVal alpha As Double = 0.12, Optional ByVal gamma As Double = 3#, _
        Optional ByVal friction As Double = 65#, Optional ByVal dilation As Double = 32#, _
        Optional ByVal eccentricity As Double = 0.1)
    If pointCount < 1 Or modulus <= 0# Or ft <= 0# Or fc0 <= ft Or fc <= fc0 Or _
            gt <= 0# Or gc <= 0# Or localizationLength <= 0# Then _
        InvalidInput "Укажите положительные модули, энергии, характерную длину и прочности ft < fc0 < fc."
    If peakCompression <= fc / modulus Or InvalidDilationAngles(friction, dilation) Then _
        InvalidInput "Проверьте деформацию пика сжатия и углы дилатансии и внутреннего трения."
    If alpha < 0# Or alpha >= 1# Or gamma < 0# Or eccentricity <= 0# Or poisson <= -1# Or poisson >= 0.5 Then _
        InvalidInput "Проверьте коэффициент Пуассона и параметры пластического потенциала."
    If fc0 / ft * (1# - alpha) < 1# + alpha Then _
        InvalidInput "Соотношение начальных прочностей несовместимо с поверхностью текучести."

    Set mReturn = New CCoupledConcreteReturn
    mReturn.Initialize modulus, poisson, alpha, gamma, eccentricity
    CoupledElasticTensor modulus, poisson, mElastic
    Set mTension = New CCoupledConcreteEnvelope: Set mCompression = New CCoupledConcreteEnvelope
    Dim halfTension As Double, compressDamage As Double, plasticPeak As Double
    halfTension = Log(2#) * gt / (ft * localizationLength)
    compressDamage = 0.423 * (1# - fc / (peakCompression * modulus))
    plasticPeak = peakCompression - fc / ((1# - compressDamage) * modulus)
    mTension.Initialize ft, ft, 0#, gt, localizationLength, 0.5, halfTension
    mCompression.Initialize fc0, fc, plasticPeak, gc, localizationLength, compressDamage, plasticPeak
    mAlpha = alpha: mGamma = gamma
    mFt = ft: mFc = fc
    mFrictionMax = friction * GEOM_PI / 180#: mDilationMax = dilation * GEOM_PI / 180#
    mPointCount = pointCount
    ICoupledPointResponse_ClearHistory
End Sub

' Вычисляет одну пробу и её касательную. Возмущения вычисляются от
' того же принятого состояния, без изменения пробной истории.
Private Function ICoupledPointResponse_EvaluatePoint(ByVal point As Long, ByRef strain() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double) As ESolverFailureCode
    Dim updated As TConcreteHistory, scratch As TConcreteHistory, virginElastic As Boolean
    If Not UpdatePoint(strain, mHistory(point, mCommitted), updated, stress, virginElastic) Then GoTo Failure
    Dim i As Long, j As Long
    If virginElastic Then
        For j = 1 To N_COMPONENTS
            For i = 1 To N_COMPONENTS: tangent(i, j) = mElastic(i, j): Next i
        Next j
    Else
        Dim perturbed(1 To N_COMPONENTS) As Double, plus(1 To N_COMPONENTS) As Double, minus(1 To N_COMPONENTS) As Double
        Dim delta As Double, plusOk As Boolean, minusOk As Boolean, unused As Boolean
        For i = 1 To N_COMPONENTS: perturbed(i) = strain(i): Next i
        For j = 1 To N_COMPONENTS
            delta = STRAIN_PERTURBATION * (1# + Abs(strain(j)))
            perturbed(j) = strain(j) + delta
            plusOk = UpdatePoint(perturbed, mHistory(point, mCommitted), scratch, plus, unused)
            perturbed(j) = strain(j) - delta
            minusOk = UpdatePoint(perturbed, mHistory(point, mCommitted), scratch, minus, unused)
            perturbed(j) = strain(j)
            If Not plusOk And Not minusOk Then GoTo Failure
            For i = 1 To N_COMPONENTS
                If plusOk And minusOk Then
                    tangent(i, j) = (plus(i) - minus(i)) / (2# * delta)
                ElseIf plusOk Then
                    tangent(i, j) = (plus(i) - stress(i)) / delta
                Else
                    tangent(i, j) = (stress(i) - minus(i)) / delta
                End If
            Next i
        Next j
    End If
    mHistory(point, mTrial) = updated
    Exit Function
Failure:
    ICoupledPointResponse_EvaluatePoint = sfcNumericalFailure
End Function

' Восстанавливает напряжения и историю точки вывода на известном принятом
' поле. Точка не участвует в равновесии, её касательная не нужна.
Friend Function EvaluateOutputPoint(ByVal point As Long, ByRef strain() As Double, ByRef stress() As Double) As Boolean
    Dim updated As TConcreteHistory, virginElastic As Boolean
    If Not UpdatePoint(strain, mHistory(point, mCommitted), updated, stress, virginElastic) Then Exit Function
    mHistory(point, mTrial) = updated
    EvaluateOutputPoint = True
End Function

' Подтверждает вычисленный шаг точек вывода перед переходом к следующему полю.
Friend Sub AcceptOutputPoints()
    ICoupledPointResponse_AcceptTrial
End Sub

' Меняет владение готовыми массивами без покомпонентной записи истории.
Private Sub ICoupledPointResponse_AcceptTrial()
    mCommitted = mTrial: mTrial = 3 - mCommitted
End Sub

' Выделяет пустые истории всех бетонных, включая замещённые, точек.
Private Sub ICoupledPointResponse_ClearHistory()
    ReDim mHistory(1 To mPointCount, 1 To 2)
    mCommitted = 1: mTrial = 2
End Sub

' Снимает одноосную огибающую из фактического закона данного бетона.
' Отрицательная ветвь передаётся от начала к хвосту сжатия.
Friend Sub CopyCalibrationPoints(ByVal compression As Boolean, ByVal modulus As Double, _
        ByRef strain() As Double, ByRef stress() As Double)
    If compression Then
        mCompression.CopyCalibrationPoints modulus, -1#, strain, stress
    Else
        mTension.CopyCalibrationPoints modulus, 1#, strain, stress
    End If
End Sub

' Возвращает физические меры повреждения принятой точки для диагностики.
Friend Sub AcceptedPointDamage(ByVal point As Long, ByRef tensionEnergy As Double, ByRef compressionEnergy As Double)
    Dim unusedStress As Double, unusedCohesion As Double, unusedDamage As Double
    mTension.Evaluate mHistory(point, mCommitted).Tension, unusedStress, unusedCohesion, unusedDamage, tensionEnergy
    mCompression.Evaluate mHistory(point, mCommitted).Compression, unusedStress, unusedCohesion, unusedDamage, compressionEnergy
End Sub

' Находит совместную пробу пластического возврата, энергий и дилатансии.
' Фиксированная точка чувствительных ветвей проверяется до выхода.
Private Function UpdatePoint(ByRef strain() As Double, ByRef committed As TConcreteHistory, _
        ByRef updated As TConcreteHistory, ByRef apparent() As Double, ByRef virginElastic As Boolean) As Boolean
    Dim elastic(1 To N_COMPONENTS) As Double, effective(1 To N_COMPONENTS) As Double
    Dim increment(1 To N_COMPONENTS) As Double, tensorIncrement(1 To N_COMPONENTS) As Double
    Dim i As Long, j As Long, elasticComponent As Double
    For j = 1 To N_COMPONENTS
        elasticComponent = strain(j) - committed.Plastic(j)
        For i = 1 To N_COMPONENTS
            elastic(i) = elastic(i) + mElastic(i, j) * elasticComponent
        Next i
    Next j
    Dim tension As Double, compression As Double, dilationRate As Double
    Dim nextTension As Double, nextCompression As Double, nextDilation As Double
    Dim duration As Double, trace As Double, q As Double, high As Double, mid As Double, low As Double
    Dim ct As Double, cc As Double, dt As Double, dc As Double, kt As Double, kc As Double
    Dim unusedStress As Double, r As Double, residual As Double
    Dim deltaTension As Double, deltaCompression As Double, deltaDilation As Double
    tension = committed.Tension: compression = committed.Compression
    virginElastic = False
    updated = committed
    mTension.Evaluate tension, unusedStress, ct, dt, kt
    mCompression.Evaluate compression, unusedStress, cc, dc, kc
    CoupledStressInvariants elastic, trace, q, high, mid, low
    r = PositiveStressFraction(high, mid, low)
    residual = CoupledBarcelonaYield(trace, q, high, cc, ct, mAlpha, mGamma)
    If residual <= 1E-10 * cc Then
        Dim fraction As Double
        fraction = (1# - r * dt) * (1# - dc)
        For i = 1 To N_COMPONENTS: apparent(i) = fraction * elastic(i): Next i
        virginElastic = (tension = 0# And compression = 0# And residual < -0.0002 * cc)
        UpdatePoint = True
        Exit Function
    End If

    dilationRate = CurrentDilation(kt, kc, r, trace, q)
    For j = 1 To MATERIAL_ITERATIONS
        If Not mReturn.Correct(elastic, cc, ct, mFt, dilationRate, effective, increment, duration) Then Exit Function
        CoupledStressInvariants effective, trace, q, high, mid, low
        r = PositiveStressFraction(high, mid, low)
        For i = 1 To 3
            tensorIncrement(i) = increment(i)
            tensorIncrement(i + 3) = increment(i + 3) * 0.5 ' Главные пластические деформации используют тензорные сдвиги.
        Next i
        Dim eTrace As Double, eQ As Double, eHigh As Double, eMid As Double, eLow As Double
        CoupledStressInvariants tensorIncrement, eTrace, eQ, eHigh, eMid, eLow
        nextTension = committed.Tension
        nextCompression = committed.Compression
        If eHigh > 0# Then nextTension = nextTension + r * eHigh
        If eLow < 0# Then nextCompression = nextCompression - (1# - r) * eLow
        mTension.Evaluate nextTension, unusedStress, ct, dt, kt
        mCompression.Evaluate nextCompression, unusedStress, cc, dc, kc
        nextDilation = CurrentDilation(kt, kc, r, trace, q)
        deltaTension = nextTension - tension
        deltaCompression = nextCompression - compression
        deltaDilation = nextDilation - dilationRate
        residual = CoupledBarcelonaYield(trace, q, high, cc, ct, mAlpha, mGamma)
        If Abs(deltaTension) <= STATE_TOLERANCE * (mFt / mElastic(3, 3) + nextTension) And _
                Abs(deltaCompression) <= STATE_TOLERANCE * (mFc / mElastic(3, 3) + nextCompression) And _
                Abs(deltaDilation) <= STATE_TOLERANCE And Abs(residual) <= 1E-8 * cc Then Exit For
        tension = tension + MAXIMUM_DAMPING * deltaTension
        compression = compression + MAXIMUM_DAMPING * deltaCompression
        dilationRate = dilationRate + MAXIMUM_DAMPING * deltaDilation
        mTension.Evaluate tension, unusedStress, ct, dt, kt
        mCompression.Evaluate compression, unusedStress, cc, dc, kc
    Next j
    If j > MATERIAL_ITERATIONS Then Exit Function
    fraction = (1# - r * dt) * (1# - dc)
    For i = 1 To N_COMPONENTS
        updated.Plastic(i) = committed.Plastic(i) + increment(i)
        apparent(i) = fraction * effective(i)
    Next i
    updated.Tension = nextTension: updated.Compression = nextCompression
    UpdatePoint = True
End Function

' Возвращает долю положительных главных напряжений по абсолютной сумме.
Private Function PositiveStressFraction(ByVal high As Double, ByVal mid As Double, ByVal low As Double) As Double
    Dim denominator As Double
    denominator = Abs(high) + Abs(mid) + Abs(low)
    If denominator = 0# Then Exit Function
    If high > 0# Then PositiveStressFraction = high
    If mid > 0# Then PositiveStressFraction = PositiveStressFraction + mid
    If low > 0# Then PositiveStressFraction = PositiveStressFraction + low
    If PositiveStressFraction >= denominator Then
        PositiveStressFraction = 1# ' Округление суммы не создаёт отрицательную долю сжатия.
    Else
        PositiveStressFraction = PositiveStressFraction / denominator
    End If
End Function

' Считает дилатансию по накопленной энергии и напорному состоянию.
' Верхняя граница углов действует при сдвиге и при гидростатическом
' растяжении. После полного повреждения углы достигают своего предела.
Private Function CurrentDilation(ByVal kt As Double, ByVal kc As Double, ByVal r As Double, _
        ByVal trace As Double, ByVal q As Double) As Double
    Dim mixed As Double, confine As Double, exponent As Double
    mixed = 1# - (1# - (1# + MIXED_DAMAGE_WEIGHT * r) * kt) * (1# - (1# - r) * kc)
    If mixed <= 0# Then Exit Function
    If mixed > 1# Then mixed = 1#
    confine = 1#
    If q > 1E-12 * mFc Then
        exponent = -(1# + trace / q) * trace / mFc
        If exponent < 0# Then confine = Exp(exponent)
    End If
    Dim sinFriction As Double, sinPeakDilation As Double, sinConstantVolume As Double, sinCurrentFriction As Double
    Dim sinDilation As Double
    sinFriction = Sin(mFrictionMax * confine)
    sinPeakDilation = Sin(mDilationMax * confine)
    sinConstantVolume = (sinFriction - sinPeakDilation) / (1# - sinFriction * sinPeakDilation)
    sinCurrentFriction = 2# * Sqr(mixed) / (1# + mixed) * sinFriction
    If sinCurrentFriction <= sinConstantVolume Then Exit Function
    sinDilation = (sinCurrentFriction - sinConstantVolume) / (1# - sinCurrentFriction * sinConstantVolume)
    CurrentDilation = sinDilation / Sqr((1# - sinDilation) * (1# + sinDilation))
End Function

' Проверяет диапазон углов до перевода в радианы.
Private Function InvalidDilationAngles(ByVal friction As Double, ByVal dilation As Double) As Boolean
    InvalidDilationAngles = (friction <= 0# Or friction >= 90# Or dilation < 0# Or dilation >= friction)
End Function

' Передаёт причину неверных материальных характеристик вызывающему слою.
Private Sub InvalidInput(ByVal reason As String)
    Err.Raise vbObjectError + 1783, "CCoupledConcreteResponse", reason
End Sub
