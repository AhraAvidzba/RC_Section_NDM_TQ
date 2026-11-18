VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledASDTangent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Алгоритмическая производная замкнутого спектрального закона ASD
' ==========================================================================
' Дифференцирует все фактически выполненные обновления двух необратимых координат.
' Спектральные проекторы вычисляются один раз из упругого прогноза; вращение
' главных направлений входит через разделённые разности. Напряжения и история
' вызывающего материала не изменяются. При близких корнях, знаковой границе
' или несогласовании с фактическим шагом возвращает False для разностного пути.

Private Type TBranchGradient
    Coordinate As Double
    Gradient(1 To 3) As Double
    NominalFactor As Double
    EffectiveFactor As Double
    NominalGradient(1 To 3) As Double
    EffectiveGradient(1 To 3) As Double
End Type

Private Const PRINCIPAL_GAP_RATIO As Double = 0.00001 ' Ограничивает обусловленность полиномиальных проекторов.
Private Const SIGN_GAP_RATIO As Double = 0.0000000001 ' На знаковой границе производная выбирается разностным алгоритмом.
Private Const PROJECTOR_TOLERANCE As Double = 0.0000001
Private mTension As CCoupledConcreteBackbone, mCompression As CCoupledConcreteBackbone
Private mModulus As Double, mG As Double, mLame As Double, mAlpha As Double, mGamma As Double, mRatio As Double, mBeta As Double
Private mClosureTolerance As Double, mClosureIterations As Long, mReady As Boolean
Private mElastic(1 To 3, 1 To 3, 1 To 6) As Double

' Подключает те же неизменяемые ветви и численные параметры замыкания, что использует основной закон.
Friend Sub Initialize(ByVal tension As CCoupledConcreteBackbone, ByVal compression As CCoupledConcreteBackbone, _
        ByVal modulus As Double, ByVal poisson As Double, ByVal alpha As Double, ByVal gamma As Double, _
        ByVal strengthRatio As Double, ByVal closureTolerance As Double, ByVal closureIterations As Long)
    If mReady Then Err.Raise 5, "CCoupledASDTangent", "Подготовленная касательная неизменяема. Создайте новый объект."
    If tension Is Nothing Or compression Is Nothing Or modulus <= 0# Or closureTolerance <= 0# Or closureIterations < 1 Then _
        Err.Raise 5, "CCoupledASDTangent", "Передайте готовые ветви, положительный модуль и параметры замыкания."
    Set mTension = tension: Set mCompression = compression
    mModulus = modulus: mG = modulus / (2# * (1# + poisson)): mLame = 2# * mG * poisson / (1# - 2# * poisson)
    mAlpha = alpha: mGamma = gamma: mRatio = strengthRatio: mBeta = (1# - alpha) * strengthRatio - (1# + alpha)
    mClosureTolerance = closureTolerance: mClosureIterations = closureIterations
    Dim i As Long, j As Long
    For j = 1 To 3
        For i = 1 To 3: mElastic(i, i, j) = mLame: Next i
        mElastic(j, j, j) = mLame + 2# * mG
    Next j
    mElastic(1, 2, 4) = mG: mElastic(2, 1, 4) = mG
    mElastic(1, 3, 5) = mG: mElastic(3, 1, 5) = mG
    mElastic(2, 3, 6) = mG: mElastic(3, 2, 6) = mG
    mReady = True
End Sub

' Строит производную по инженерным деформациям только при согласовании с действительным замкнутым шагом.
' Отказ не означает отказ материала: вызывающий слой может собрать его разностную касательную.
Friend Function TryBuild(ByRef strain() As Double, ByRef oldStrain() As Double, ByRef oldEffective() As Double, _
        ByVal oldT As Double, ByVal oldC As Double, ByRef expectedStress() As Double, ByRef expectedEffective() As Double, _
        ByVal expectedT As Double, ByVal expectedC As Double, ByRef tangent() As Double) As Boolean
    If Not mReady Then Exit Function
    On Error GoTo ArithmeticFailed
    Dim trial(1 To 6) As Double, increment(1 To 6) As Double, volume As Double, moved As Boolean, i As Long, j As Long
    For i = 1 To 6
        increment(i) = strain(i) - oldStrain(i)
        If increment(i) <> 0# Then moved = True
    Next i
    If Not moved Then Exit Function ' Точная принятая точка может иметь разные односторонние касательные.
    volume = mLame * (increment(1) + increment(2) + increment(3))
    For i = 1 To 3: trial(i) = oldEffective(i) + 2# * mG * increment(i) + volume: Next i
    For i = 4 To 6: trial(i) = oldEffective(i) + mG * increment(i): Next i
    Dim trace As Double, q As Double, initial(1 To 3) As Double, values(1 To 3) As Double, jacobian(1 To 3, 1 To 3) As Double
    CoupledStressInvariants trial, trace, q, initial(1), initial(2), initial(3)
    Dim projectors(1 To 3, 1 To 3, 1 To 3) As Double
    If Not PrepareProjectors(trial, initial, projectors) Then Exit Function
    For i = 1 To 3: values(i) = initial(i): jacobian(i, i) = 1#: Next i
    Dim priorT As TBranchGradient, priorC As TBranchGradient, nextT As TBranchGradient, nextC As TBranchGradient, selected As TBranchGradient
    priorT.Coordinate = oldT: priorC.Coordinate = oldC
    Dim equivalent As Double, equivalentGradient(1 To 3) As Double
    Dim nominal(1 To 3) As Double, effective(1 To 3) As Double
    Dim nominalJacobian(1 To 3, 1 To 3) As Double, effectiveJacobian(1 To 3, 1 To 3) As Double
    Dim stress(1 To 6) As Double, elasticStress(1 To 6) As Double, lastStress(1 To 6) As Double, lastEffective(1 To 6) As Double
    Dim iteration As Long, maximum As Double, error As Double
    For iteration = 0 To mClosureIterations
        EquivalentCoordinate values, jacobian, False, equivalent, equivalentGradient
        AdvanceBranch mTension, priorT, equivalent, equivalentGradient, nextT
        EquivalentCoordinate values, jacobian, True, equivalent, equivalentGradient
        AdvanceBranch mCompression, priorC, equivalent, equivalentGradient, nextC
        For i = 1 To 3
            If values(i) < 0# Then selected = nextC Else selected = nextT
            nominal(i) = values(i) * selected.NominalFactor
            effective(i) = values(i) * selected.EffectiveFactor
            For j = 1 To 3
                nominalJacobian(i, j) = jacobian(i, j) * selected.NominalFactor + values(i) * selected.NominalGradient(j)
                effectiveJacobian(i, j) = jacobian(i, j) * selected.EffectiveFactor + values(i) * selected.EffectiveGradient(j)
            Next j
        Next i
        Reconstruct projectors, nominal, stress: Reconstruct projectors, effective, elasticStress
        If iteration > 0 Then
            maximum = mModulus * GeomMax(Abs(nextT.Coordinate - priorT.Coordinate), Abs(nextC.Coordinate - priorC.Coordinate))
            For i = 1 To 6
                maximum = GeomMax(maximum, Abs(stress(i) - lastStress(i)))
                maximum = GeomMax(maximum, Abs(elasticStress(i) - lastEffective(i)))
            Next i
            If maximum <= mClosureTolerance Then
                error = mModulus * GeomMax(Abs(nextT.Coordinate - expectedT), Abs(nextC.Coordinate - expectedC))
                For i = 1 To 6
                    error = GeomMax(error, Abs(stress(i) - expectedStress(i)))
                    error = GeomMax(error, Abs(elasticStress(i) - expectedEffective(i)))
                Next i
                If error > 8# * mClosureTolerance Then Exit Function
                AssembleTangent initial, nominal, nominalJacobian, projectors, tangent
                TryBuild = True: Exit Function
            End If
        End If
        priorT = nextT: priorC = nextC
        For i = 1 To 3
            values(i) = effective(i)
            For j = 1 To 3: jacobian(i, j) = effectiveJacobian(i, j): Next j
        Next i
        For i = 1 To 6: lastStress(i) = stress(i): lastEffective(i) = elasticStress(i): Next i
    Next iteration
    Exit Function
ArithmeticFailed:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Вычисляет полиномиальные проекторы простых корней и проверяет их полноту и идемпотентность.
' Близкие корни и знаковая граница требуют другого вычисления производной.
Private Function PrepareProjectors(ByRef stress() As Double, ByRef values() As Double, ByRef projectors() As Double) As Boolean
    Dim magnitude As Double, i As Long, j As Long, k As Long, a As Long, b As Long
    magnitude = 1#
    For i = 1 To 3: magnitude = GeomMax(magnitude, Abs(values(i))): Next i
    For i = 1 To 3
        If Abs(values(i)) <= SIGN_GAP_RATIO * magnitude Then Exit Function
        For j = i + 1 To 3
            If Abs(values(i) - values(j)) <= PRINCIPAL_GAP_RATIO * magnitude Then Exit Function
        Next j
    Next i
    Dim tensor(1 To 3, 1 To 3) As Double, squared(1 To 3, 1 To 3) As Double, denominator As Double, value As Double, error As Double
    For i = 1 To 3: tensor(i, i) = stress(i): Next i
    tensor(1, 2) = stress(4): tensor(2, 1) = stress(4)
    tensor(1, 3) = stress(5): tensor(3, 1) = stress(5)
    tensor(2, 3) = stress(6): tensor(3, 2) = stress(6)
    For a = 1 To 3
        For b = 1 To 3
            For k = 1 To 3: squared(a, b) = squared(a, b) + tensor(a, k) * tensor(k, b): Next k
        Next b
    Next a
    For i = 1 To 3
        j = i Mod 3 + 1: k = j Mod 3 + 1
        denominator = (values(i) - values(j)) * (values(i) - values(k))
        For a = 1 To 3
            For b = 1 To 3
                value = squared(a, b) - (values(j) + values(k)) * tensor(a, b)
                If a = b Then value = value + values(j) * values(k)
                projectors(a, b, i) = value / denominator
            Next b
        Next a
    Next i
    For a = 1 To 3
        For b = 1 To 3
            value = 0#
            For i = 1 To 3: value = value + projectors(a, b, i): Next i
            If a = b Then value = value - 1#
            error = GeomMax(error, Abs(value))
            For i = 1 To 3
                value = -projectors(a, b, i)
                For k = 1 To 3: value = value + projectors(a, k, i) * projectors(k, b, i): Next k
                error = GeomMax(error, Abs(value))
            Next i
        Next b
    Next a
    PrepareProjectors = (error <= PROJECTOR_TOLERANCE)
End Function

' Дифференцирует эквивалентную координату выбранной знаковой части по трём исходным главным напряжениям.
Private Sub EquivalentCoordinate(ByRef values() As Double, ByRef jacobian() As Double, ByVal compression As Boolean, _
        ByRef equivalent As Double, ByRef gradient() As Double)
    Dim selected(1 To 3) As Double, active(1 To 3) As Boolean, trace As Double, mean As Double, q As Double
    Dim i As Long, j As Long, factor As Double, derivative As Double, high As Double
    equivalent = 0#: For j = 1 To 3: gradient(j) = 0#: Next j
    For i = 1 To 3
        active(i) = (Not compression Or values(i) < 0#)
        If active(i) Then selected(i) = values(i)
        trace = trace + selected(i)
    Next i
    high = selected(1): mean = trace / 3#
    q = Sqr(0.5 * ((selected(1) - selected(2)) ^ 2 + (selected(2) - selected(3)) ^ 2 + (selected(3) - selected(1)) ^ 2))
    factor = (1# - mAlpha) * mModulus
    If compression Then
        equivalent = (mAlpha * trace + q + mGamma * high) / factor
    Else
        If high < mTension.StressFloor Then Exit Sub
        factor = factor * mRatio: equivalent = (mAlpha * trace + q + mBeta * high) / factor
    End If
    For i = 1 To 3
        If active(i) Then
            derivative = mAlpha
            If q > 0# Then derivative = derivative + 1.5 * (selected(i) - mean) / q
            If i = 1 Then
                If compression Then derivative = derivative + mGamma Else derivative = derivative + mBeta
            End If
            For j = 1 To 3: gradient(j) = gradient(j) + derivative * jacobian(i, j) / factor: Next j
        End If
    Next i
End Sub

' Применяет правило цепочки к необратимой координате, обеим таблицам и знаменателю масштабирования.
Private Sub AdvanceBranch(ByVal branch As CCoupledConcreteBackbone, ByRef prior As TBranchGradient, _
        ByVal equivalent As Double, ByRef equivalentGradient() As Double, ByRef updated As TBranchGradient)
    Dim nominal As Double, effective As Double, nominalSlope As Double, effectiveSlope As Double
    Dim plastic As Double, plasticGradient(1 To 3) As Double, reference As Double, referenceGradient As Double, j As Long
    branch.StressesAndSlopes prior.Coordinate, nominal, effective, nominalSlope, effectiveSlope
    plastic = prior.Coordinate - effective / mModulus
    For j = 1 To 3: plasticGradient(j) = (1# - effectiveSlope / mModulus) * prior.Gradient(j): Next j
    updated.Coordinate = GeomMax(prior.Coordinate, equivalent + plastic)
    For j = 1 To 3
        If equivalent + plastic > prior.Coordinate Then
            updated.Gradient(j) = equivalentGradient(j) + plasticGradient(j)
        Else
            updated.Gradient(j) = prior.Gradient(j)
        End If
    Next j
    branch.StressesAndSlopes updated.Coordinate, nominal, effective, nominalSlope, effectiveSlope
    reference = mModulus * (updated.Coordinate - plastic)
    updated.NominalFactor = nominal / reference: updated.EffectiveFactor = effective / reference
    For j = 1 To 3
        referenceGradient = mModulus * (updated.Gradient(j) - plasticGradient(j))
        updated.NominalGradient(j) = (nominalSlope * updated.Gradient(j) - updated.NominalFactor * referenceGradient) / reference
        updated.EffectiveGradient(j) = (effectiveSlope * updated.Gradient(j) - updated.EffectiveFactor * referenceGradient) / reference
    Next j
End Sub

' Восстанавливает шесть компонент тензора из неизменных проекторов и текущих главных значений.
Private Sub Reconstruct(ByRef projectors() As Double, ByRef principal() As Double, ByRef tensor() As Double)
    Dim i As Long, j As Long
    For i = 1 To 6: tensor(i) = 0#: Next i
    For j = 1 To 3
        For i = 1 To 3: tensor(i) = tensor(i) + principal(j) * projectors(i, i, j): Next i
        tensor(4) = tensor(4) + principal(j) * projectors(1, 2, j)
        tensor(5) = tensor(5) + principal(j) * projectors(1, 3, j)
        tensor(6) = tensor(6) + principal(j) * projectors(2, 3, j)
    Next j
End Sub

' Собирает полную двойную сумму главных производных и изменение направлений; симметрия касательной не предполагается.
Private Sub AssembleTangent(ByRef initial() As Double, ByRef nominal() As Double, ByRef jacobian() As Double, _
        ByRef projectors() As Double, ByRef tangent() As Double)
    Dim column As Long, a As Long, b As Long, i As Long, j As Long, k As Long
    Dim principalIncrement(1 To 3) As Double, derivative(1 To 3, 1 To 3) As Double, product(1 To 3, 1 To 3) As Double
    Dim mixed(1 To 3, 1 To 3) As Double, value As Double, ratio As Double
    For column = 1 To 6
        For i = 1 To 3
            principalIncrement(i) = 0#
            For a = 1 To 3
                For b = 1 To 3: principalIncrement(i) = principalIncrement(i) + projectors(a, b, i) * mElastic(a, b, column): Next b
            Next a
        Next i
        Erase derivative
        For i = 1 To 3
            value = 0#
            For j = 1 To 3: value = value + jacobian(i, j) * principalIncrement(j): Next j
            For a = 1 To 3
                For b = 1 To 3: derivative(a, b) = derivative(a, b) + value * projectors(a, b, i): Next b
            Next a
        Next i
        For i = 1 To 2
            For j = i + 1 To 3
                ratio = (nominal(i) - nominal(j)) / (initial(i) - initial(j))
                Erase product: Erase mixed
                For a = 1 To 3
                    For b = 1 To 3
                        For k = 1 To 3: product(a, b) = product(a, b) + projectors(a, k, i) * mElastic(k, b, column): Next k
                    Next b
                Next a
                For a = 1 To 3
                    For b = 1 To 3
                        For k = 1 To 3: mixed(a, b) = mixed(a, b) + product(a, k) * projectors(k, b, j): Next k
                    Next b
                Next a
                For a = 1 To 3
                    For b = 1 To 3: derivative(a, b) = derivative(a, b) + ratio * (mixed(a, b) + mixed(b, a)): Next b
                Next a
            Next j
        Next i
        For i = 1 To 3: tangent(i, column) = derivative(i, i): Next i
        tangent(4, column) = derivative(1, 2): tangent(5, column) = derivative(1, 3): tangent(6, column) = derivative(2, 3)
    Next column
End Sub
