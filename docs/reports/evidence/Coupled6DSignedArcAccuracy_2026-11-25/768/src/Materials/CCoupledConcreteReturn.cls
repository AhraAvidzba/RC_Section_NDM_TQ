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
Private Const NEWTON_ITERATIONS As Long = 10
Private Const SCALED_APEX_FACTOR As Double = 0.00000001
Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000001
Private Const NEAR_YIELD_RANGE_RATIO As Double = 0.00000004
Private Const NEAR_YIELD_MAX_DEVIATOR_LOSS As Double = 0.0001
Private Const NEAR_YIELD_ITERS As Long = 8
Private Const ROUNDING_LIMIT As Double = 0.000000000000007105427357601 ' 32 относительных шага Double.
Private Const MAX_DIRECT_MULTIPLIER As Double = 1E+308 ' Запас до переполнения при сравнении перед делением.
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
' Вход elastic учитывает ранее принятую пластическую деформацию.
' Инварианты этого тензора одинаковы на всех итерациях локальной истории.
' Множитель равен stepMultiplier / multiplierDivisor. Обычно делитель 1.
' У вершины конечный объёмный поток может соответствовать множителю вне Double.
' Неудача оставляет вызывающему право уменьшить материальный шаг.
Friend Function Correct(ByRef elastic() As Double, ByVal trace As Double, ByVal q As Double, ByVal principalMax As Double, _
        ByVal cc As Double, ByVal ct As Double, _
        ByVal ft0 As Double, ByVal dilation As Double, ByRef corrected() As Double, _
        ByRef plasticIncrement() As Double, ByRef stepMultiplier As Double, ByRef deviatorFactor As Double, _
        ByRef multiplierDivisor As Double) As Boolean
    If mBulk <= 0# Or cc <= 0# Or ct <= 0# Or ft0 <= 0# Or dilation < 0# Then _
        Err.Raise 5, "CCoupledConcreteReturn", "Не заданы параметры текущей поверхности бетона."
    multiplierDivisor = 1#
    Dim beta As Double
    beta = cc / ct * (1# - mAlpha) - (1# + mAlpha)
    If beta < 0# Then Exit Function ' Монотонная скобка требует допустимой геометрии поверхности.
    Dim tolerance As Double, value As Double, k As Long
    tolerance = RELATIVE_YIELD_TOLERANCE * cc
    value = CoupledBarcelonaYield(trace, q, principalMax, cc, ct, mAlpha, mGamma)
    If value > 0# And value <= NEAR_YIELD_RANGE_RATIO * cc Then
        If TryRegularNearYield(elastic, trace, q, principalMax, cc, ct, ft0, dilation, value, _
                corrected, plasticIncrement, stepMultiplier, deviatorFactor) Then
            Correct = True
            Exit Function
        End If
    End If
    stepMultiplier = 0#
    If value <= tolerance Then
        For k = 1 To 6: corrected(k) = elastic(k): plasticIncrement(k) = 0#: Next k
        deviatorFactor = 1#
        Correct = True
        Exit Function
    End If

    Dim factor As Double, newTrace As Double, devLength As Double, offset As Double
    Dim trialMean As Double: trialMean = trace / 3#
    If q <= 1E-14 * cc Then
        If Not TensileApex(trace, dilation, beta, cc, newTrace, stepMultiplier, multiplierDivisor) Then Exit Function
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
        Dim radius As Double, offsetSquared As Double, devSquared As Double, traceScale As Double
        Dim maxPrincipal As Double, branchSlope As Double, traceSlope As Double, yieldSlope As Double, nextFactor As Double
        Dim scaledOffset As Double, proposed As Double, triedScaledApex As Boolean, solvedScaledApex As Boolean
        offsetSquared = offset * offset: devSquared = devLength * devLength
        traceScale = 9# * mBulk * dilation
        factor = 0.5
        For i = 1 To NEWTON_ITERATIONS + MAX_ITERATIONS
            If factor < SCALED_APEX_FACTOR And Not triedScaledApex And offset > 0# Then
                triedScaledApex = True
                solvedScaledApex = SolveScaledApex(trace, q, largestDev, cc, ct, beta, ft0, dilation, tolerance, _
                    factor, newTrace, stepMultiplier, multiplierDivisor)
                If solvedScaledApex Then Exit For
            End If
            radius = Sqr(offsetSquared + devSquared * factor * factor)
            stepMultiplier = (1# - factor) / factor * radius / mTwiceShear
            newTrace = trace - traceScale * stepMultiplier
            maxPrincipal = newTrace / 3# + largestDev * factor
            value = CoupledBarcelonaYield(newTrace, q * factor, maxPrincipal, cc, ct, mAlpha, mGamma)
            If Abs(value) <= tolerance Then Exit For
            If value > 0# Then upper = factor Else lower = factor
            nextFactor = (lower + upper) * 0.5
            If i <= NEWTON_ITERATIONS Then
                branchSlope = mGamma
                If maxPrincipal > 0# Then branchSlope = beta
                traceSlope = 0#
                If traceScale > 0# And radius > 0# Then
                    scaledOffset = offset / factor
                    traceSlope = (traceScale / mTwiceShear) * (scaledOffset * scaledOffset + devSquared * factor) / radius
                End If
                ' Производная числителя F вдоль f = q/qtrial; знак главного напряжения задаёт ветвь.
                yieldSlope = q + branchSlope * largestDev + (mAlpha + branchSlope / 3#) * traceSlope
                If yieldSlope > 0# Then
                    proposed = factor - value * (1# - mAlpha) / yieldSlope
                    If proposed > lower And proposed < upper Then nextFactor = proposed
                End If
            End If
            If nextFactor <= lower Or nextFactor >= upper Then Exit For
            factor = nextFactor
        Next i
        If Not solvedScaledApex And Abs(value) > tolerance Then
            If triedScaledApex Then Exit Function
            If Not SolveScaledApex(trace, q, largestDev, cc, ct, beta, ft0, dilation, tolerance, _
                    factor, newTrace, stepMultiplier, multiplierDivisor) Then Exit Function
        End If
    End If

    AssembleCorrection elastic, trace, newTrace, factor, corrected, plasticIncrement
    deviatorFactor = factor
    Correct = True
End Function

' Восстанавливает тензор поправки из следа и доли пробного девиатора.
' Деформации сдвига инженерные, модуль сдвига задан через 2G.
Friend Sub AssembleCorrection(ByRef elastic() As Double, ByVal trace As Double, ByVal newTrace As Double, _
        ByVal factor As Double, ByRef corrected() As Double, ByRef plasticIncrement() As Double)
    Dim trialMean As Double: trialMean = trace / 3#
    Dim k As Long, correctedMean As Double, volumetric As Double
    correctedMean = newTrace / 3#
    volumetric = (trace - newTrace) / (9# * mBulk) ' Приращение каждой из трёх нормальных компонент от дилатансии.
    For k = 1 To 3
        corrected(k) = correctedMean + factor * (elastic(k) - trialMean)
        corrected(k + 3) = factor * elastic(k + 3)
        plasticIncrement(k) = volumetric + (1# - factor) * (elastic(k) - trialMean) / mTwiceShear
        plasticIncrement(k + 3) = 2# * (1# - factor) * elastic(k + 3) / mTwiceShear
    Next k
End Sub

' Возвращает малую положительную пробу точно на F=0 на одной гладкой
' меридиональной ветви. При вершине, смене знака sigma1 или неопределённом
' знаке F по точности Double решение оставляет общему пути.
Friend Function TryRegularNearYield(ByRef elastic() As Double, ByVal trace As Double, ByVal q As Double, _
        ByVal principalMax As Double, ByVal cc As Double, ByVal ct As Double, ByVal ft0 As Double, _
        ByVal dilation As Double, ByVal trialYield As Double, ByRef corrected() As Double, _
        ByRef plasticIncrement() As Double, ByRef stepMultiplier As Double, ByRef deviatorFactor As Double) As Boolean
    If q <= 1E-14 * cc Then Exit Function
    If trialYield <= 0# Or trialYield > NEAR_YIELD_RANGE_RATIO * cc Then Exit Function

    Dim largestDev As Double, trialMean As Double, branchSlope As Double, elasticPart As Double, beta As Double
    Dim pressurePart As Double, initialNumerator As Double, inputScale As Double
    trialMean = trace / 3#
    largestDev = principalMax - trialMean
    If largestDev < 0# Then Exit Function ' Для верхнего главного напряжения требуется неотрицательный девиатор.
    beta = cc / ct * (1# - mAlpha) - (1# + mAlpha)
    If beta < 0# Then Exit Function
    branchSlope = mGamma
    If principalMax > 0# Then branchSlope = beta
    initialNumerator = trialYield * (1# - mAlpha)
    inputScale = Abs(mAlpha * trace) + q + Abs(branchSlope * principalMax) + cc * (1# - mAlpha)
    If initialNumerator <= ROUNDING_LIMIT * inputScale Then Exit Function

    elasticPart = q + branchSlope * largestDev
    pressurePart = mAlpha + branchSlope / 3#
    Dim dev As Double, offset As Double, volumeScale As Double, firstRadius As Double
    dev = DEVIATORIC_NORM * q
    offset = mEccentricity * dilation * ft0
    volumeScale = 9# * mBulk * dilation / mTwiceShear
    firstRadius = Sqr(dev * dev + offset * offset)
    If firstRadius <= 0# Then Exit Function

    ' При u=1-f приращение давления пропорционально u*r(u), где r
    ' не убывает. Линеаризация в u=0 ограничивает физический корень сверху.
    Dim lower As Double, upper As Double, loss As Double, offsetScaled As Double
    Dim pressureDrop As Double, radius As Double, rest As Double, residualTolerance As Double
    upper = initialNumerator / (elasticPart + pressurePart * volumeScale * firstRadius)
    If upper > NEAR_YIELD_MAX_DEVIATOR_LOSS Then Exit Function
    offsetScaled = offset / (1# - upper)
    radius = Sqr(dev * dev + offsetScaled * offsetScaled)
    pressureDrop = upper * volumeScale * radius
    If principalMax > 0# Then
        If principalMax <= pressureDrop / 3# + upper * largestDev Then Exit Function
    End If

    loss = upper
    residualTolerance = ROUNDING_LIMIT * initialNumerator
    Dim iteration As Long, denominator As Double, proposal As Double
    For iteration = 1 To NEAR_YIELD_ITERS
        offsetScaled = offset / (1# - loss)
        radius = Sqr(dev * dev + offsetScaled * offsetScaled)
        pressureDrop = loss * volumeScale * radius
        rest = initialNumerator - loss * elasticPart - pressurePart * pressureDrop
        If Abs(rest) <= residualTolerance Then Exit For
        If rest > 0# Then lower = loss Else upper = loss
        denominator = elasticPart + pressurePart * volumeScale * _
            (radius + loss * offsetScaled * offsetScaled / ((1# - loss) * radius))
        proposal = loss + rest / denominator
        If proposal <= lower Or proposal >= upper Then proposal = lower + (upper - lower) * 0.5
        loss = proposal
    Next iteration
    If iteration > NEAR_YIELD_ITERS Then Exit Function

    ' Собираем приращение через малое u. Вычитание (1-f) теряет
    ' эту величину при округлении f к единице.
    stepMultiplier = loss * radius / mTwiceShear
    deviatorFactor = 1# - loss
    Dim k As Long, deviator As Double, volumetric As Double
    volumetric = dilation * stepMultiplier
    For k = 1 To 3
        deviator = loss * (elastic(k) - trialMean)
        corrected(k) = elastic(k) - deviator - pressureDrop / 3#
        corrected(k + 3) = elastic(k + 3) - loss * elastic(k + 3)
        plasticIncrement(k) = volumetric + deviator / mTwiceShear
        plasticIncrement(k + 3) = 2# * loss * elastic(k + 3) / mTwiceShear
    Next k
    TryRegularNearYield = True
End Function

' У растягивающей вершины q/qElastic пропорционально квадрату дилатансии.
' Поиск в масштабе этого квадрата сохраняет точность даже при f ниже диапазона Double.
Private Function SolveScaledApex(ByVal trace As Double, ByVal q As Double, ByVal largestDev As Double, _
        ByVal cc As Double, ByVal ct As Double, ByVal beta As Double, ByVal ft0 As Double, _
        ByVal dilation As Double, ByVal tolerance As Double, ByRef factor As Double, _
        ByRef newTrace As Double, ByRef stepMultiplier As Double, ByRef multiplierDivisor As Double) As Boolean
    If dilation <= 0# Then Exit Function
    Dim slope As Double: slope = mAlpha + beta / 3#
    If slope <= 0# Then Exit Function
    Dim apexDistance As Double: apexDistance = trace - cc * (1# - mAlpha) / slope
    If apexDistance <= 0# Then Exit Function

    Dim offsetPerDilation As Double: offsetPerDilation = mEccentricity * ft0
    Dim bulkToShear As Double: bulkToShear = 9# * mBulk / mTwiceShear
    Dim factorScale As Double: factorScale = bulkToShear * offsetPerDilation / apexDistance
    Dim lower As Double, upper As Double, ratio As Double, bracketed As Boolean, i As Long
    Dim scaledFactor As Double, trialFactor As Double, trialTrace As Double, radius As Double
    Dim value As Double, scaledVolume As Double
    ratio = 1#
    For i = 1 To MAX_ITERATIONS
        scaledFactor = (ratio * factorScale) * dilation ' f/dilation конечен и после потери самого f при округлении.
        trialFactor = scaledFactor * dilation
        If trialFactor >= 1# Then
            trialFactor = 1#: scaledVolume = 0#: trialTrace = trace
        Else
            radius = Sqr(offsetPerDilation * offsetPerDilation + (DEVIATORIC_NORM * q * scaledFactor) ^ 2)
            scaledVolume = (1# - trialFactor) * radius / (mTwiceShear * (ratio * factorScale)) ' Произведение λ·dilation в деформациях.
            trialTrace = trace - 9# * mBulk * scaledVolume
        End If
        value = CoupledBarcelonaYield(trialTrace, q * trialFactor, trialTrace / 3# + largestDev * trialFactor, cc, ct, mAlpha, mGamma)
        If Abs(value) <= tolerance Then
            factor = trialFactor: newTrace = trialTrace
            If MultiplierNeedsScaling(scaledVolume, dilation) Then
                stepMultiplier = scaledVolume: multiplierDivisor = dilation
            Else
                stepMultiplier = scaledVolume / dilation
            End If
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
        ByRef newTrace As Double, ByRef stepMultiplier As Double, ByRef multiplierDivisor As Double) As Boolean
    Dim slope As Double: slope = mAlpha + beta / 3#
    If dilation <= 0# Or slope <= 0# Then Exit Function
    newTrace = cc * (1# - mAlpha) / slope
    Dim traceDrop As Double: traceDrop = trace - newTrace
    If traceDrop <= 0# Then Exit Function
    Dim bulkScale As Double: bulkScale = 9# * mBulk
    Dim traceScale As Double: traceScale = bulkScale * dilation
    If MultiplierNeedsScaling(traceDrop, traceScale) Then
        stepMultiplier = traceDrop / bulkScale: multiplierDivisor = dilation
    Else
        stepMultiplier = traceDrop / traceScale
    End If
    TensileApex = (stepMultiplier > 0#)
End Function

' Проверяет диапазон частного до деления, включая субнормальный знаменатель.
' При знаменателе от единицы частное не превышает конечный числитель.
Private Function MultiplierNeedsScaling(ByVal numerator As Double, ByVal denominator As Double) As Boolean
    If denominator < 1# Then MultiplierNeedsScaling = (numerator > MAX_DIRECT_MULTIPLIER * denominator)
End Function
