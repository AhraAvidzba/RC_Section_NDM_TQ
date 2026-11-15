VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledConcreteReturn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Спектральный возврат эффективных напряжений бетона
' ==========================================================================
' Один локальный возврат при фиксированных сцеплениях и дилатансии.
' При изотропной упругости поток сохраняет оси пробного девиатора.
' Поэтому обратный Эйлер сводится к скаляру — отношению q/qElastic.

Private Const DEVIATORIC_NORM As Double = 0.816496580927726
Private Const MAX_ITERATIONS As Long = 64
Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000001
Private mTwiceShear As Double
Private mBulk As Double
Private mAlpha As Double, mGamma As Double, mEccentricity As Double

' Задаёт упругость и форму поверхности; все модули в МПа.
Friend Sub Initialize(ByVal modulus As Double, ByVal poisson As Double, _
        ByVal alpha As Double, ByVal gamma As Double, ByVal eccentricity As Double)
    If modulus <= 0# Or poisson <= -1# Or poisson >= 0.5 Or _
            alpha < 0# Or alpha >= 1# Or gamma < 0# Or eccentricity <= 0# Then _
        Err.Raise 5, "CCoupledConcreteReturn", "Неверные параметры упругости или пластического возврата бетона."
    mTwiceShear = modulus / (1# + poisson)
    mBulk = modulus / (3# * (1# - 2# * poisson))
    mAlpha = alpha: mGamma = gamma: mEccentricity = eccentricity
End Sub

' Возвращает эффективные напряжения и приращение пластической деформации.
' Вход elastic уже учитывает ранее принятую пластическую деформацию.
' Неудача оставляет вызывающему право уменьшить материальный шаг.
Friend Function Correct(ByRef elastic() As Double, ByVal cc As Double, ByVal ct As Double, _
        ByVal ft0 As Double, ByVal dilation As Double, ByRef corrected() As Double, _
        ByRef plasticIncrement() As Double, ByRef stepMultiplier As Double) As Boolean
    If mBulk <= 0# Or cc <= 0# Or ct <= 0# Or ft0 <= 0# Or dilation < 0# Then _
        Err.Raise 5, "CCoupledConcreteReturn", "Не заданы параметры текущей поверхности бетона."
    Dim beta As Double
    beta = cc / ct * (1# - mAlpha) - (1# + mAlpha)
    If beta < 0# Then Exit Function ' Монотонная скобка требует допустимой геометрии поверхности.
    Dim trace As Double, q As Double, principalMax As Double, principalMid As Double, principalMin As Double
    CoupledStressInvariants elastic, trace, q, principalMax, principalMid, principalMin
    Dim tolerance As Double, value As Double, k As Long
    tolerance = RELATIVE_YIELD_TOLERANCE * cc
    value = CoupledBarcelonaYield(trace, q, principalMax, cc, ct, mAlpha, mGamma)
    stepMultiplier = 0#
    If value <= tolerance Then
        For k = 1 To 6: corrected(k) = elastic(k): plasticIncrement(k) = 0#: Next k
        Correct = True
        Exit Function
    End If

    Dim factor As Double, newTrace As Double, devLength As Double, offset As Double
    Dim trialMean As Double: trialMean = trace / 3#
    If q <= 1E-14 * cc Then
        If Not TensileApex(trace, dilation, beta, cc, newTrace, stepMultiplier) Then Exit Function
    Else
        devLength = DEVIATORIC_NORM * q
        offset = mEccentricity * dilation * ft0
        Dim lower As Double, upper As Double, i As Long, apexPressure As Double
        Dim largestDev As Double: largestDev = principalMax - trialMean
        upper = 1#
        If offset = 0# Then
            apexPressure = trace
            If CoupledBarcelonaYield(apexPressure, 0#, apexPressure / 3#, cc, ct, mAlpha, mGamma) > tolerance Then Exit Function
        End If
        For i = 1 To MAX_ITERATIONS
            factor = (lower + upper) * 0.5
            Dim radius As Double
            radius = Sqr(offset * offset + (devLength * factor) ^ 2)
            stepMultiplier = (1# - factor) / factor * radius / mTwiceShear
            newTrace = trace - 9# * mBulk * dilation * stepMultiplier
            value = CoupledBarcelonaYield(newTrace, q * factor, newTrace / 3# + largestDev * factor, cc, ct, mAlpha, mGamma)
            If Abs(value) <= tolerance Then Exit For
            If value > 0# Then upper = factor Else lower = factor
        Next i
        If Abs(value) > tolerance Then
            If Not SolveScaledApex(trace, q, largestDev, cc, ct, beta, ft0, dilation, tolerance, _
                    factor, newTrace, stepMultiplier) Then Exit Function
        End If
    End If

    Dim correctedMean As Double, volumetric As Double
    correctedMean = newTrace / 3#
    volumetric = (trace - newTrace) / (9# * mBulk) ' Приращение каждой из трёх нормальных компонент от дилатансии.
    For k = 1 To 3
        corrected(k) = correctedMean + factor * (elastic(k) - trialMean)
        corrected(k + 3) = factor * elastic(k + 3)
        plasticIncrement(k) = volumetric + (1# - factor) * (elastic(k) - trialMean) / mTwiceShear
        plasticIncrement(k + 3) = 2# * (1# - factor) * elastic(k + 3) / mTwiceShear
    Next k
    Correct = True
End Function

' У растягивающей вершины q/qElastic пропорционально квадрату дилатансии.
' Поиск в масштабе этого квадрата сохраняет точность даже при f ниже диапазона Double.
Private Function SolveScaledApex(ByVal trace As Double, ByVal q As Double, ByVal largestDev As Double, _
        ByVal cc As Double, ByVal ct As Double, ByVal beta As Double, ByVal ft0 As Double, _
        ByVal dilation As Double, ByVal tolerance As Double, ByRef factor As Double, _
        ByRef newTrace As Double, ByRef stepMultiplier As Double) As Boolean
    If dilation <= 0# Then Exit Function
    Dim slope As Double: slope = mAlpha + beta / 3#
    If slope <= 0# Then Exit Function
    Dim apexDistance As Double: apexDistance = trace - cc * (1# - mAlpha) / slope
    If apexDistance <= 0# Then Exit Function

    Dim offsetPerDilation As Double: offsetPerDilation = mEccentricity * ft0
    Dim bulkToShear As Double: bulkToShear = 9# * mBulk / mTwiceShear
    Dim factorScale As Double: factorScale = bulkToShear * offsetPerDilation / apexDistance
    Dim lower As Double, upper As Double, ratio As Double, bracketed As Boolean, i As Long
    Dim scaledFactor As Double, radius As Double, value As Double, scaledVolume As Double
    ratio = 1#
    For i = 1 To MAX_ITERATIONS
        scaledFactor = (ratio * factorScale) * dilation ' f/dilation конечен и после потери самого f при округлении.
        factor = scaledFactor * dilation
        If factor >= 1# Then
            factor = 1#: scaledVolume = 0#: newTrace = trace
        Else
            radius = Sqr(offsetPerDilation * offsetPerDilation + (DEVIATORIC_NORM * q * scaledFactor) ^ 2)
            scaledVolume = (1# - factor) * radius / (mTwiceShear * (ratio * factorScale)) ' Произведение λ·dilation в деформациях.
            newTrace = trace - 9# * mBulk * scaledVolume
        End If
        value = CoupledBarcelonaYield(newTrace, q * factor, newTrace / 3# + largestDev * factor, cc, ct, mAlpha, mGamma)
        If Abs(value) <= tolerance Then
            stepMultiplier = scaledVolume / dilation
            SolveScaledApex = True
            Exit Function
        End If
        If value > 0# Then
            upper = ratio: bracketed = True
        Else
            lower = ratio
        End If
        If bracketed Then
            ratio = lower + (upper - lower) * 0.5
            If ratio <= lower Or ratio >= upper Then Exit Function
        Else
            ratio = 2# * ratio
        End If
    Next i
End Function

' Возвращает изотропную растягивающую пробу непосредственно на вершину.
' При нулевом объёмном потоке вершина недостижима из внешней точки.
Private Function TensileApex(ByVal trace As Double, ByVal dilation As Double, ByVal beta As Double, ByVal cc As Double, _
        ByRef newTrace As Double, ByRef stepMultiplier As Double) As Boolean
    Dim slope As Double: slope = mAlpha + beta / 3#
    If dilation <= 0# Or slope <= 0# Then Exit Function
    newTrace = cc * (1# - mAlpha) / slope
    stepMultiplier = (trace - newTrace) / (9# * mBulk * dilation)
    TensileApex = (stepMultiplier > 0#)
End Function
