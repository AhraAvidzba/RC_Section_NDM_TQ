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
' CDP использует поверхность Барселоны, гиперболический потенциал и историю
' пластических энергий. ASD использует спектральное разложение и полные
' эквивалентные координаты подготовленных одноосных ветвей. Значения этих
' координат не взаимозаменяемы. Истории исходного и замещённого бетона едины.

Implements ICoupledPointResponse

Private Type TConcreteHistory
    Strain(1 To 6) As Double ' Полная деформация принятой точки для локального деления шага.
    Stress(1 To 6) As Double ' Напряжения той же принятой точки при точно нулевом приращении.
    Effective(1 To 6) As Double ' Эффективный тензор спектрального закона при той же принятой деформации.
    Plastic(1 To 6) As Double
    Tension As Double ' Пластическая координата CDP либо полная эквивалентная деформация ASD.
    Compression As Double ' Положительная мера сжатия в том же контракте выбранного закона.
    ElasticTangentExact As Boolean
    TensionCohesion As Double ' Значения огибающих при тех же Tension и Compression.
    CompressionCohesion As Double
    TensionDamage As Double
    CompressionDamage As Double
    TensionEnergy As Double
    CompressionEnergy As Double
End Type

Private Const N_COMPONENTS As Long = 6
Private Const PLASTIC_ROUNDOFF As Double = 32# / (2# ^ 52) ' Относительный масштаб округления при восстановлении пластической деформации ASD.
Private Const MATERIAL_ITERATIONS As Long = 70
Private Const STATE_TOLERANCE As Double = 0.0000000005
Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000001 ' Единый допуск для возврата и повторной оценки принятой истории.
Private Const STRAIN_PERTURBATION As Double = 0.000000002
Private Const MIXED_DAMAGE_WEIGHT As Double = 20#
Private Const MAXIMUM_DAMPING As Double = 0.5
Private Const MAX_LOCAL_DIVISIONS As Long = 16
Private Const TENSILE_AXIS_BAND_RATIO As Double = 1E-12 ' Непрерывное сопряжение внутри численной полосы q/fc.

Private mHistory() As TConcreteHistory
Private mCommitted As Long, mTrial As Long, mPointCount As Long
Private mTrialEpoch As Long ' Меняется при каждом принятии истории.
Private mCachedEpoch() As Long, mCachedMode() As ECoupledTangentMode
Private mCachedTangent() As Double ' Столбцы касательной последней нелинейной пробы каждой точки.
Private mElastic(1 To 6, 1 To 6) As Double
Private mReturn As CCoupledConcreteReturn
Private mTension As CCoupledConcreteEnvelope, mCompression As CCoupledConcreteEnvelope
Private mASDFlow As CCoupledASDFlow
Private mLastFailure As CCoupledMaterialFailure
Private mFt As Double, mFc As Double, mAlpha As Double, mGamma As Double
Private mTensionEnergyDensity As Double, mCompressionEnergyDensity As Double
Private mFrictionMax As Double, mDilationMax As Double
Private mPressureModulus As Double ' 3K: изменение давления от одинакового приращения трёх нормальных деформаций.
Private mApexPotentialOffset As Double ' Эксцентриситет потенциала, умноженный на ft, МПа.
Private mDeviatorCompliance As Double ' 1/(2G): перевод девиатора напряжений в тензорные деформации.

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

    Set mASDFlow = Nothing
    Set mReturn = New CCoupledConcreteReturn
    mReturn.Initialize modulus, poisson, alpha, gamma, eccentricity
    CoupledElasticTensor modulus, poisson, mElastic
    mPressureModulus = modulus / (1# - 2# * poisson)
    mDeviatorCompliance = 1# / (2# * mElastic(4, 4))
    Dim calibration As New CCoupledConcreteCalibration
    calibration.Initialize modulus, fc0, fc, ft, peakCompression, gt, gc, localizationLength
    Set mTension = calibration.TensionEnvelope: Set mCompression = calibration.CompressionEnvelope
    mTensionEnergyDensity = gt / localizationLength: mCompressionEnergyDensity = gc / localizationLength
    mAlpha = alpha: mGamma = gamma
    mFt = ft: mFc = fc
    mApexPotentialOffset = eccentricity * ft
    mFrictionMax = friction * GEOM_PI / 180#: mDilationMax = dilation * GEOM_PI / 180#
    mPointCount = pointCount
    ICoupledPointResponse_ClearHistory
End Sub

' Подключает неизменяемый спектральный закон с собственными таблицами.
' Касательные, пробные буферы и восстановление полей используют общий контракт материала.
Friend Sub InitializeASD(ByVal pointCount As Long, ByVal flow As CCoupledASDFlow)
    If flow Is Nothing Then InvalidInput "Подготовьте спектральный закон бетона."
    If pointCount < 1 Or flow.Modulus <= 0# Then InvalidInput "Проверьте число точек и готовность спектрального бетона."
    Set mASDFlow = flow
    Set mReturn = Nothing: Set mTension = Nothing: Set mCompression = Nothing
    CoupledElasticTensor flow.Modulus, flow.Poisson, mElastic
    mFt = flow.PeakStress(False): mFc = flow.PeakStress(True)
    mPointCount = pointCount
    ICoupledPointResponse_ClearHistory
End Sub

' Возвращает обозначение применённого пространственного закона, отдельно от одноосных диаграмм.
Friend Property Get LawIdentifier() As String
    If mASDFlow Is Nothing Then LawIdentifier = "Barcelona-CDP" Else LawIdentifier = mASDFlow.LawIdentifier
End Property

' Возвращает масштаб CDP-координаты огибающих; ноль означает неприменимость к текущему закону.
Friend Property Get EnvelopeReferenceStress() As Double
    If mASDFlow Is Nothing And mPointCount > 0 Then EnvelopeReferenceStress = mFt
End Property

' Читает принятые полные эквивалентные деформации ASD без обновления материала.
' Обе меры неотрицательны и необратимы; они не являются компонентами тензора.
' Для CDP возвращает False и нули: его пластические координаты имеют иной смысл.
Friend Function TryReadAcceptedASDStrains(ByVal point As Long, ByRef tension As Double, ByRef compression As Double) As Boolean
    tension = 0#: compression = 0#
    If mASDFlow Is Nothing Then Exit Function
    If point < 1 Or point > mPointCount Then InvalidInput "Точка эквивалентных деформаций ASD находится вне подготовленного материала."
    tension = mHistory(point, mCommitted).Tension
    compression = mHistory(point, mCommitted).Compression
    TryReadAcceptedASDStrains = True
End Function

' Читает координаты принятого либо уже вычисленного пробного буфера без повторного обновления.
' Пробное значение читает владелец сразу после успешной оценки данной точки с recordTrial=True.
Friend Function TryReadBufferedASDStrains(ByVal point As Long, ByVal accepted As Boolean, _
        ByRef tension As Double, ByRef compression As Double) As Boolean
    tension = 0#: compression = 0#
    If mASDFlow Is Nothing Then Exit Function
    If point < 1 Or point > mPointCount Then InvalidInput "Точка эквивалентных деформаций ASD находится вне подготовленного материала."
    Dim buffer As Long: buffer = mTrial
    If accepted Then buffer = mCommitted
    tension = mHistory(point, buffer).Tension: compression = mHistory(point, buffer).Compression
    TryReadBufferedASDStrains = True
End Function

' Вычисляет полные эквивалентные деформации кандидата ASD от принятой истории.
' Не меняет принятый или пробный буфер, касательную и исходный массив деформаций.
' Меры относятся к тому же обновлению, что и напряжения; порог разрушения здесь не задаётся.
Friend Function TryReadASDStrains(ByVal point As Long, ByRef strain() As Double, _
        ByRef tension As Double, ByRef compression As Double) As Boolean
    tension = 0#: compression = 0#
    If mASDFlow Is Nothing Then Exit Function
    If point < 1 Or point > mPointCount Then InvalidInput "Точка эквивалентных деформаций ASD находится вне подготовленного материала."
    Dim updated As TConcreteHistory, stress(1 To N_COMPONENTS) As Double, elastic As Boolean, substeps As Long
    If Not UpdateWithSubsteps(strain, mHistory(point, mCommitted), updated, stress, elastic, substeps) Then Exit Function
    tension = updated.Tension: compression = updated.Compression
    TryReadASDStrains = True
End Function

' Читает координату огибающих в МПа из принятого либо полностью подготовленного пробного буфера.
' Пробный буфер допустим только после успешного расчёта всего кандидата; координата не является критерием разрушения.
Friend Function TryReadBufferedEnvelope(ByVal point As Long, ByVal accepted As Boolean, ByRef value As Double) As Boolean
    If EnvelopeReferenceStress <= 0# Then Exit Function
    If point < 1 Or point > mPointCount Then InvalidInput "Точка координаты огибающих находится вне подготовленного материала."
    Dim buffer As Long: buffer = mTrial
    If accepted Then buffer = mCommitted
    value = mTensionEnergyDensity * mHistory(point, buffer).TensionEnergy + _
        mCompressionEnergyDensity * mHistory(point, buffer).CompressionEnergy
    TryReadBufferedEnvelope = True
End Function

' Вычисляет ту же пробную координату от принятой истории без изменения буферов, включая внутреннее деление шага.
Friend Function TryReadEnvelope(ByVal point As Long, ByRef strain() As Double, ByRef value As Double) As Boolean
    If EnvelopeReferenceStress <= 0# Then Exit Function
    If point < 1 Or point > mPointCount Then InvalidInput "Точка координаты огибающих находится вне подготовленного материала."
    Dim updated As TConcreteHistory, stress(1 To N_COMPONENTS) As Double, elastic As Boolean, substeps As Long
    If Not UpdateWithSubsteps(strain, mHistory(point, mCommitted), updated, stress, elastic, substeps) Then Exit Function
    value = mTensionEnergyDensity * updated.TensionEnergy + mCompressionEnergyDensity * updated.CompressionEnergy
    TryReadEnvelope = True
End Function

' Передаёт неизменяемый контекст последнего отказа; успешные пробы его не подменяют.
Friend Property Get LastFailure() As CCoupledMaterialFailure
    Set LastFailure = mLastFailure
End Property

' Вычисляет одну пробу и её касательную. Возмущения вычисляются от
' того же принятого состояния, без изменения пробной истории.
Private Function ICoupledPointResponse_EvaluatePoint(ByVal point As Long, ByRef strain() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double, _
        Optional ByVal tangentMode As ECoupledTangentMode = ctmFull, _
        Optional ByVal recordTrial As Boolean = True) As ESolverFailureCode
    Dim updated As TConcreteHistory, virginElastic As Boolean, cachedMode As ECoupledTangentMode
    Dim sameTrial As Boolean, i As Long, j As Long, baseSubsteps As Long, failedComponent As Long
    Dim failureStage As Long: failureStage = 1 ' Название этапа формируется только при фактическом отказе.
    Dim branchConsistent As Boolean
    branchConsistent = (tangentMode = ctmFullBranchConsistent Or tangentMode = ctmRefinedBranchConsistent)
    If branchConsistent And recordTrial Then
        ICoupledPointResponse_EvaluatePoint = sfcInvalidConfiguration
        Exit Function
    End If
    cachedMode = ctmStressOnly
    If Not branchConsistent Then sameTrial = MatchesNonlinearTrial(point, strain)
    If sameTrial Then
        updated = mHistory(point, mTrial)
        For i = 1 To N_COMPONENTS: stress(i) = updated.Stress(i): Next i
        cachedMode = mCachedMode(point)
    Else
        If recordTrial Then mCachedEpoch(point) = 0 ' Касательная базового поля остаётся пригодной после возмущения.
        If Not UpdateWithSubsteps(strain, mHistory(point, mCommitted), updated, stress, virginElastic, baseSubsteps) Then GoTo Failure
    End If
    If tangentMode = ctmStressOnly Then GoTo StoreTrial
    If virginElastic Then
        For j = 1 To N_COMPONENTS
            If tangentMode <> ctmAxialGradient Or j = 3 Or j >= 5 Then
                For i = 1 To N_COMPONENTS: tangent(i, j) = mElastic(i, j): Next i
            End If
        Next j
    Else
        failureStage = 2
        If Not NonlinearTangent(point, strain, stress, tangent, tangentMode, cachedMode, recordTrial, _
                updated, baseSubsteps, failedComponent) Then GoTo Failure
    End If
StoreTrial:
    If Not recordTrial Then Exit Function
    mHistory(point, mTrial) = updated
    If Not virginElastic Then
        mCachedEpoch(point) = mTrialEpoch
        If tangentMode < cachedMode Then cachedMode = tangentMode
        mCachedMode(point) = cachedMode
    End If
    Exit Function
Failure:
    RememberFailure point, failureStage, failedComponent, tangentMode, strain
    ICoupledPointResponse_EvaluatePoint = sfcNumericalFailure
End Function

' Снимает вход и принятую историю, не меняя ни материальных буферов, ни статуса разрушения.
Private Sub RememberFailure(ByVal point As Long, ByVal stage As Long, ByVal component As Long, _
        ByVal mode As ECoupledTangentMode, ByRef requested() As Double)
    Dim prior(1 To 6) As Double, effective(1 To 6) As Double, i As Long, j As Long
    For i = 1 To 6: prior(i) = mHistory(point, mCommitted).Strain(i): Next i
    If mASDFlow Is Nothing Then
        For i = 1 To 6
            For j = 1 To 6
                effective(i) = effective(i) + mElastic(i, j) * (prior(j) - mHistory(point, mCommitted).Plastic(j))
            Next j
        Next i
    Else
        For i = 1 To 6: effective(i) = mHistory(point, mCommitted).Effective(i): Next i
    End If
    Dim snapshot As New CCoupledMaterialFailure
    snapshot.Initialize point, mTrialEpoch, stage, component, mode, LawIdentifier, requested, prior, effective, _
        mHistory(point, mCommitted).Tension, mHistory(point, mCommitted).Compression
    Set mLastFailure = snapshot
End Sub

' Сравнивает все шесть деформаций точно. Номер принятой истории исключает
' повторное использование после перестановки буферов или восстановления.
Private Function MatchesNonlinearTrial(ByVal point As Long, ByRef strain() As Double) As Boolean
    If mCachedEpoch(point) <> mTrialEpoch Then Exit Function
    Dim i As Long
    For i = 1 To N_COMPONENTS
        If strain(i) <> mHistory(point, mTrial).Strain(i) Then Exit Function
    Next i
    MatchesNonlinearTrial = True
End Function

' Сохраняет готовые столбцы при переходе от невязки к полной касательной.
' Все недостающие разности вычисляются от той же принятой истории. При
' возмущении найденные столбцы не заменяют столбцы базовой деформации.
' При заданной материальной ветви предпочтительны те же признаки
' приращения и число подшагов. У точно принятого поля на границе может
' существовать только односторонняя производная соседней ветви.
Private Function NonlinearTangent(ByVal point As Long, ByRef strain() As Double, ByRef stress() As Double, _
        ByRef tangent() As Double, ByVal mode As ECoupledTangentMode, ByVal cachedMode As ECoupledTangentMode, _
        ByVal recordTrial As Boolean, ByRef baseHistory As TConcreteHistory, ByVal baseSubsteps As Long, _
        ByRef failedComponent As Long) As Boolean
    Dim perturbed(1 To N_COMPONENTS) As Double, plus(1 To N_COMPONENTS) As Double, minus(1 To N_COMPONENTS) As Double
    Dim delta As Double, plusOk As Boolean, minusOk As Boolean, unused As Boolean, attempt As Long
    Dim scratch As TConcreteHistory, i As Long, j As Long, gradientColumn As Boolean
    Dim branchConsistent As Boolean, basePlastic As Boolean, basePrincipalSign As Long, sideSubsteps As Long
    Dim atAcceptedPosition As Boolean, plusAdjacent As Boolean, minusAdjacent As Boolean
    branchConsistent = (mode = ctmFullBranchConsistent Or mode = ctmRefinedBranchConsistent)
    If Not branchConsistent And Not mASDFlow Is Nothing And cachedMode <> ctmFull Then
        If mode <> ctmAxialGradient Or cachedMode <> ctmAxialGradient Then
            If TryASDTangent(point, strain, stress, tangent, baseHistory) Then
                If recordTrial Then
                    For j = 1 To N_COMPONENTS
                        If mode <> ctmAxialGradient Or j = 3 Or j >= 5 Then
                            For i = 1 To N_COMPONENTS: mCachedTangent(i, j, point) = tangent(i, j): Next i
                        End If
                    Next j
                End If
                NonlinearTangent = True: Exit Function
            End If
        End If
    End If
    If branchConsistent Then
        ReadBranchSigns baseHistory, mHistory(point, mCommitted), basePlastic, basePrincipalSign
        atAcceptedPosition = (baseSubsteps = 0)
        If atAcceptedPosition Then baseSubsteps = 1 ' Нулевая проба возвращает принятое поле; соседняя проба входит через один возврат.
    End If
    For i = 1 To N_COMPONENTS: perturbed(i) = strain(i): Next i
    For j = 1 To N_COMPONENTS
        gradientColumn = (j = 3 Or j >= 5) ' Продольный градиент имеет компоненты zz, xz, yz.
        If mode <> ctmAxialGradient Or gradientColumn Then
            If cachedMode = ctmFull Or (cachedMode = ctmAxialGradient And gradientColumn) Then
                For i = 1 To N_COMPONENTS: tangent(i, j) = mCachedTangent(i, j, point): Next i
            Else
                delta = STRAIN_PERTURBATION * (1# + Abs(strain(j)))
                If mode = ctmRefinedBranchConsistent Then delta = delta / 64#
                For attempt = 1 To 3
                    perturbed(j) = strain(j) + delta
                    plusOk = UpdateWithSubsteps(perturbed, mHistory(point, mCommitted), scratch, plus, unused, sideSubsteps)
                    If plusOk And branchConsistent Then
                        plusOk = (sideSubsteps = baseSubsteps)
                        plusAdjacent = plusOk
                        If plusOk Then _
                            plusOk = MatchesBranchSigns(scratch, mHistory(point, mCommitted), basePlastic, basePrincipalSign)
                    Else
                        plusAdjacent = False
                    End If
                    perturbed(j) = strain(j) - delta
                    minusOk = UpdateWithSubsteps(perturbed, mHistory(point, mCommitted), scratch, minus, unused, sideSubsteps)
                    If minusOk And branchConsistent Then
                        minusOk = (sideSubsteps = baseSubsteps)
                        minusAdjacent = minusOk
                        If minusOk Then _
                            minusOk = MatchesBranchSigns(scratch, mHistory(point, mCommitted), basePlastic, basePrincipalSign)
                    Else
                        minusAdjacent = False
                    End If
                    If plusOk Or minusOk Then Exit For
                    If attempt = 3 And atAcceptedPosition Then
                        plusOk = plusAdjacent: minusOk = (minusAdjacent And Not plusAdjacent)
                        If plusOk Or minusOk Then Exit For
                    End If
                    delta = delta * 0.25 ' На границе ветвей уменьшаем только численное возмущение.
                Next attempt
                perturbed(j) = strain(j)
                If Not plusOk And Not minusOk Then failedComponent = j: Exit Function
                For i = 1 To N_COMPONENTS
                    If plusOk And minusOk Then
                        tangent(i, j) = (plus(i) - minus(i)) / (2# * delta)
                    ElseIf plusOk Then
                        tangent(i, j) = (plus(i) - stress(i)) / delta
                    Else
                        tangent(i, j) = (stress(i) - minus(i)) / delta
                    End If
                    If recordTrial Then mCachedTangent(i, j, point) = tangent(i, j)
                Next i
            End If
        End If
    Next j
    NonlinearTangent = True
End Function

' Передаёт неизменные копии принятого состояния и сверяемой пробы в спектральную производную.
' У ASD нет деления материального приращения; отказ производной оставляет общий разностный путь.
Private Function TryASDTangent(ByVal point As Long, ByRef strain() As Double, ByRef stress() As Double, _
        ByRef tangent() As Double, ByRef updated As TConcreteHistory) As Boolean
    Dim priorStrain(1 To 6) As Double, priorEffective(1 To 6) As Double, effective(1 To 6) As Double, i As Long
    For i = 1 To 6
        priorStrain(i) = mHistory(point, mCommitted).Strain(i)
        priorEffective(i) = mHistory(point, mCommitted).Effective(i)
        effective(i) = updated.Effective(i)
    Next i
    TryASDTangent = mASDFlow.TryTangent(strain, priorStrain, priorEffective, _
        mHistory(point, mCommitted).Tension, mHistory(point, mCommitted).Compression, stress, effective, _
        updated.Tension, updated.Compression, tangent)
End Function

' Читает значение пробной поверхности Барселоны в МПа от принятой
' истории. Производная в инженерных деформациях определена при ненулевом
' девиаторе и простом старшем главном напряжении вне границы знака.
Friend Function TryReadTrialYield(ByVal point As Long, ByRef strain() As Double, ByRef value As Double, _
        ByRef gradient() As Double, Optional ByVal needGradient As Boolean = True) As Boolean
    If Not mASDFlow Is Nothing Then Exit Function ' Скалярная поверхность возврата CDP не является границей спектрального обновления.
    If point < 1 Or point > mPointCount Then Exit Function
    Dim c As TConcreteHistory: c = mHistory(point, mCommitted)
    Dim effective(1 To N_COMPONENTS) As Double, i As Long, j As Long
    Dim ex As Double, ey As Double, ez As Double
    ex = strain(1) - c.Plastic(1): ey = strain(2) - c.Plastic(2): ez = strain(3) - c.Plastic(3)
    For i = 1 To 3
        effective(i) = mElastic(i, 1) * ex + mElastic(i, 2) * ey + mElastic(i, 3) * ez
        effective(i + 3) = mElastic(i + 3, i + 3) * (strain(i + 3) - c.Plastic(i + 3))
    Next i
    Dim trace As Double, q As Double, high As Double, middle As Double, low As Double
    CoupledStressInvariants effective, trace, q, high, middle, low
    value = CoupledBarcelonaYield(trace, q, high, c.CompressionCohesion, c.TensionCohesion, mAlpha, mGamma)
    If Not needGradient Then TryReadTrialYield = True: Exit Function
    Dim stressScale As Double
    stressScale = GeomMax(1#, GeomMax(Abs(high), Abs(low)))
    If q <= 0.000000000001 * stressScale Or Abs(high) <= 0.000000000001 * stressScale Or _
            high - middle <= 0.000000000001 * stressScale Then Exit Function
    Dim tensor(1 To 3, 1 To 3) As Double, highProjector(1 To 3, 1 To 3) As Double
    For i = 1 To 3: tensor(i, i) = effective(i): Next i
    tensor(1, 2) = effective(4): tensor(2, 1) = effective(4)
    tensor(1, 3) = effective(5): tensor(3, 1) = effective(5)
    tensor(2, 3) = effective(6): tensor(3, 2) = effective(6)
    Dim k As Long, divisor As Double, x As Double, branchSlope As Double
    divisor = (high - middle) * (high - low)
    For i = 1 To 3
        For j = 1 To 3
            x = -(middle + low) * tensor(i, j)
            For k = 1 To 3: x = x + tensor(i, k) * tensor(k, j): Next k
            If i = j Then x = x + middle * low
            highProjector(i, j) = x / divisor
        Next j
    Next i
    branchSlope = mGamma
    If high > 0# Then _
        branchSlope = (c.CompressionCohesion / c.TensionCohesion) * (1# - mAlpha) - (1# + mAlpha)
    Dim derivative(1 To N_COMPONENTS) As Double
    For i = 1 To 3
        derivative(i) = (mAlpha + 1.5 * (effective(i) - trace / 3#) / q + branchSlope * highProjector(i, i)) / (1# - mAlpha)
    Next i
    derivative(4) = (3# * effective(4) / q + 2# * branchSlope * highProjector(1, 2)) / (1# - mAlpha)
    derivative(5) = (3# * effective(5) / q + 2# * branchSlope * highProjector(1, 3)) / (1# - mAlpha)
    derivative(6) = (3# * effective(6) / q + 2# * branchSlope * highProjector(2, 3)) / (1# - mAlpha)
    For j = 1 To N_COMPONENTS
        gradient(j) = 0#
        For i = 1 To N_COMPONENTS: gradient(j) = gradient(j) + derivative(i) * mElastic(i, j): Next i
    Next j
    TryReadTrialYield = True
End Function

' Сообщает, что подготовленный материал предоставляет знаковые границы спектрального прогноза.
Friend Property Get HasPrincipalBoundaries() As Boolean
    HasPrincipalBoundaries = Not (mASDFlow Is Nothing)
End Property

' Читает конкретную знаковую границу ASD из принятого эффективного тензора без обновления буферов.
' Номер главного напряжения идёт от наибольшего к наименьшему; это не поверхность разрушения.
Friend Function TryReadPrincipalBoundary(ByVal point As Long, ByRef strain() As Double, ByVal principalIndex As Long, _
        ByRef value As Double, ByRef gradient() As Double, Optional ByVal needGradient As Boolean = True) As Boolean
    If mASDFlow Is Nothing Or point < 1 Or point > mPointCount Then Exit Function
    Dim priorStrain(1 To 6) As Double, priorEffective(1 To 6) As Double, i As Long
    For i = 1 To 6
        priorStrain(i) = mHistory(point, mCommitted).Strain(i)
        priorEffective(i) = mHistory(point, mCommitted).Effective(i)
    Next i
    TryReadPrincipalBoundary = mASDFlow.TryReadPrincipalBoundary(strain, priorStrain, priorEffective, principalIndex, value, gradient, needGradient)
End Function

' Возвращает признаки фактически вычисленной пробы от принятой истории:
' пластическое приращение, знак старшего главного напряжения и подшаги.
' Ноль подшагов означает точное совпадение с принятой деформацией.
Friend Function TryReadConcreteBranch(ByVal point As Long, ByRef strain() As Double, _
        ByRef plasticIncrement As Boolean, ByRef principalSign As Long, ByRef substeps As Long) As Boolean
    If point < 1 Or point > mPointCount Then Exit Function
    Dim updated As TConcreteHistory, stress(1 To N_COMPONENTS) As Double, virginElastic As Boolean
    If Not UpdateWithSubsteps(strain, mHistory(point, mCommitted), updated, stress, virginElastic, substeps) Then Exit Function
    ReadBranchSigns updated, mHistory(point, mCommitted), plasticIncrement, principalSign
    TryReadConcreteBranch = True
End Function

' Сравнивает с принятой пластической деформацией. Знак старшего
' номинального напряжения совпадает со знаком эффективного при
' положительном множителе повреждения.
Private Sub ReadBranchSigns(ByRef updated As TConcreteHistory, ByRef committed As TConcreteHistory, _
        ByRef plasticIncrement As Boolean, ByRef principalSign As Long)
    Dim stress(1 To N_COMPONENTS) As Double, i As Long, strainScale As Double, threshold As Double
    If Not mASDFlow Is Nothing Then
        For i = 1 To N_COMPONENTS
            strainScale = GeomMax(strainScale, GeomMax(Abs(updated.Strain(i)), Abs(committed.Strain(i))))
            strainScale = GeomMax(strainScale, GeomMax(Abs(updated.Plastic(i)), Abs(committed.Plastic(i))))
        Next i
        threshold = PLASTIC_ROUNDOFF * strainScale ' Меняется только признак ветви, фактическая история не округляется.
    End If
    plasticIncrement = False
    For i = 1 To N_COMPONENTS
        If Abs(updated.Plastic(i) - committed.Plastic(i)) > threshold Then plasticIncrement = True
        stress(i) = updated.Stress(i)
    Next i
    Dim trace As Double, q As Double, high As Double, middle As Double, low As Double
    CoupledStressInvariants stress, trace, q, high, middle, low
    principalSign = Sgn(high)
End Sub

' Проверяет сохранение знаковых признаков материального возврата
' при одном возмущении касательной.
Private Function MatchesBranchSigns(ByRef updated As TConcreteHistory, ByRef committed As TConcreteHistory, _
        ByVal basePlastic As Boolean, ByVal basePrincipalSign As Long) As Boolean
    Dim plasticIncrement As Boolean, principalSign As Long
    ReadBranchSigns updated, committed, plasticIncrement, principalSign
    MatchesBranchSigns = (plasticIncrement = basePlastic And principalSign = basePrincipalSign)
End Function

' Восстанавливает напряжения и историю точки вывода на известном принятом
' поле. Точка не участвует в равновесии, её касательная не нужна.
Friend Function EvaluateOutputPoint(ByVal point As Long, ByRef strain() As Double, ByRef stress() As Double) As Boolean
    Dim updated As TConcreteHistory, virginElastic As Boolean
    If Not UpdateWithSubsteps(strain, mHistory(point, mCommitted), updated, stress, virginElastic) Then Exit Function
    mHistory(point, mTrial) = updated
    mCachedEpoch(point) = 0
    EvaluateOutputPoint = True
End Function

' Подтверждает вычисленный шаг точек вывода перед переходом к следующему полю.
Friend Sub AcceptOutputPoints()
    ICoupledPointResponse_AcceptTrial
End Sub

' При отказе обратного Эйлера делит локальный путь деформации. Все
' попытки начинают от того же материала. Промежуточные подшаги не
' переходят в принятое состояние до приёмки всего сечения.
Private Function UpdateWithSubsteps(ByRef strain() As Double, ByRef committed As TConcreteHistory, _
        ByRef updated As TConcreteHistory, ByRef stress() As Double, ByRef virginElastic As Boolean, _
        Optional ByRef substeps As Long = 0) As Boolean
    Dim i As Long, moved As Boolean
    substeps = 0
    For i = 1 To N_COMPONENTS
        If strain(i) <> committed.Strain(i) Then
            moved = True
            Exit For
        End If
    Next i
    If Not moved Then
        updated = committed
        For i = 1 To N_COMPONENTS: stress(i) = committed.Stress(i): Next i
        virginElastic = committed.ElasticTangentExact
        UpdateWithSubsteps = True
        Exit Function
    End If
    If UpdatePoint(strain, committed, updated, stress, virginElastic) Then
        StoreStatePosition updated, strain, stress
        substeps = 1
        updated.ElasticTangentExact = virginElastic
        UpdateWithSubsteps = True
        Exit Function
    End If
    If Not mASDFlow Is Nothing Then Exit Function ' Спектральный шаг замыкается при фиксированной деформации, без скрытой смены пути её приращения.
    Dim divisions As Long, stepIndex As Long, intermediate(1 To N_COMPONENTS) As Double
    Dim history As TConcreteHistory, fraction As Double, succeeded As Boolean
    divisions = 2
    Do
        history = committed
        For stepIndex = 1 To divisions
            fraction = CDbl(stepIndex) / divisions
            For i = 1 To N_COMPONENTS
                If stepIndex = divisions Then
                    intermediate(i) = strain(i)
                Else
                    intermediate(i) = committed.Strain(i) + fraction * (strain(i) - committed.Strain(i))
                End If
            Next i
            succeeded = UpdatePoint(intermediate, history, updated, stress, virginElastic)
            If Not succeeded Then Exit For
            history = updated
        Next stepIndex
        If succeeded Then
            substeps = divisions
            StoreStatePosition updated, strain, stress
            virginElastic = False
            updated.ElasticTangentExact = False
            UpdateWithSubsteps = True
            Exit Function
        End If
        divisions = divisions * 2
    Loop While divisions <= MAX_LOCAL_DIVISIONS
End Function

' Сохраняет сопряжённые полную деформацию и напряжения данной пробы.
Private Sub StoreStatePosition(ByRef history As TConcreteHistory, ByRef strain() As Double, ByRef stress() As Double)
    Dim i As Long
    For i = 1 To N_COMPONENTS
        history.Strain(i) = strain(i)
        history.Stress(i) = stress(i)
    Next i
End Sub

' Меняет владение готовыми массивами без покомпонентной записи истории.
Private Sub ICoupledPointResponse_AcceptTrial()
    mCommitted = mTrial: mTrial = 3 - mCommitted
    mTrialEpoch = mTrialEpoch + 1
End Sub

' Выделяет пустые истории всех бетонных, включая замещённые, точек.
Private Sub ICoupledPointResponse_ClearHistory()
    Set mLastFailure = Nothing
    ReDim mHistory(1 To mPointCount, 1 To 2)
    ReDim mCachedEpoch(1 To mPointCount): ReDim mCachedMode(1 To mPointCount)
    ReDim mCachedTangent(1 To N_COMPONENTS, 1 To N_COMPONENTS, 1 To mPointCount)
    mCommitted = 1: mTrial = 2
    mTrialEpoch = 1
    Dim point As Long, initial As TConcreteHistory, unusedStress As Double
    initial.ElasticTangentExact = (mASDFlow Is Nothing)
    If mASDFlow Is Nothing Then
        mTension.Evaluate 0#, unusedStress, initial.TensionCohesion, initial.TensionDamage, initial.TensionEnergy
        mCompression.Evaluate 0#, unusedStress, initial.CompressionCohesion, initial.CompressionDamage, initial.CompressionEnergy
    End If
    For point = 1 To mPointCount: mHistory(point, mCommitted) = initial: Next point
End Sub

' Читает напряжения принятой точки квадратуры без повторного решения материала.
Friend Sub CopyAcceptedStress(ByVal point As Long, ByRef stress() As Double)
    Dim i As Long
    For i = 1 To N_COMPONENTS: stress(i) = mHistory(point, mCommitted).Stress(i): Next i
End Sub

' Снимает одноосную огибающую из фактического закона данного бетона.
' Отрицательная ветвь передаётся от начала к хвосту сжатия.
Friend Sub CopyCalibrationPoints(ByVal compression As Boolean, ByVal modulus As Double, _
        ByRef strain() As Double, ByRef stress() As Double)
    If Not mASDFlow Is Nothing Then
        If modulus <> mASDFlow.Modulus Then InvalidInput "Модуль вывода диаграммы не совпадает с фактическим спектральным материалом."
        mASDFlow.CopyCurve compression, strain, stress
        Exit Sub
    End If
    If compression Then
        mCompression.CopyCalibrationPoints modulus, -1#, strain, stress
    Else
        mTension.CopyCalibrationPoints modulus, 1#, strain, stress
    End If
End Sub

' Возвращает меры повреждения принятой точки для диагностики: доли энергии CDP
' или факторы потери жёсткости спектрального закона. Это не критерий разрушения сечения.
Friend Sub AcceptedPointDamage(ByVal point As Long, ByRef tensionEnergy As Double, ByRef compressionEnergy As Double)
    If mASDFlow Is Nothing Then
        tensionEnergy = mHistory(point, mCommitted).TensionEnergy
        compressionEnergy = mHistory(point, mCommitted).CompressionEnergy
    Else
        tensionEnergy = mHistory(point, mCommitted).TensionDamage
        compressionEnergy = mHistory(point, mCommitted).CompressionDamage
    End If
End Sub

' Находит совместную пробу пластического возврата, энергий и дилатансии.
' Фиксированная точка чувствительных ветвей проверяется до выхода.
Private Function UpdatePoint(ByRef strain() As Double, ByRef committed As TConcreteHistory, _
        ByRef updated As TConcreteHistory, ByRef apparent() As Double, ByRef virginElastic As Boolean) As Boolean
    If Not mASDFlow Is Nothing Then
        virginElastic = False
        UpdatePoint = UpdateASDPoint(strain, committed, updated, apparent)
        Exit Function
    End If
    Dim elastic(1 To N_COMPONENTS) As Double, effective(1 To N_COMPONENTS) As Double
    Dim increment(1 To N_COMPONENTS) As Double
    Dim i As Long, j As Long, strainX As Double, strainY As Double, strainZ As Double
    strainX = strain(1) - committed.Plastic(1)
    strainY = strain(2) - committed.Plastic(2)
    strainZ = strain(3) - committed.Plastic(3)
    ' Изотропная упругость разделяет нормальный блок и независимые сдвиги. Порядок сложения нормальных вкладов сохранён.
    elastic(1) = mElastic(1, 1) * strainX + mElastic(1, 2) * strainY + mElastic(1, 3) * strainZ
    elastic(2) = mElastic(2, 1) * strainX + mElastic(2, 2) * strainY + mElastic(2, 3) * strainZ
    elastic(3) = mElastic(3, 1) * strainX + mElastic(3, 2) * strainY + mElastic(3, 3) * strainZ
    elastic(4) = mElastic(4, 4) * (strain(4) - committed.Plastic(4))
    elastic(5) = mElastic(5, 5) * (strain(5) - committed.Plastic(5))
    elastic(6) = mElastic(6, 6) * (strain(6) - committed.Plastic(6))
    Dim tension As Double, compression As Double, dilationRate As Double
    Dim nextTension As Double, nextCompression As Double, nextDilation As Double
    Dim plasticMultiplier As Double, multiplierDivisor As Double, trace As Double, q As Double
    Dim high As Double, mid As Double, low As Double
    Dim trialTrace As Double, trialQ As Double, trialHigh As Double, trialMid As Double, trialLow As Double
    Dim ct As Double, cc As Double, dt As Double, dc As Double, kt As Double, kc As Double
    Dim unusedStress As Double, r As Double, residual As Double
    Dim deltaTension As Double, deltaCompression As Double, deltaDilation As Double
    tension = committed.Tension: compression = committed.Compression
    virginElastic = False
    updated = committed
    ct = committed.TensionCohesion: cc = committed.CompressionCohesion
    dt = committed.TensionDamage: dc = committed.CompressionDamage
    kt = committed.TensionEnergy: kc = committed.CompressionEnergy
    CoupledStressInvariants elastic, trialTrace, trialQ, trialHigh, trialMid, trialLow
    r = PositiveStressFraction(trialHigh, trialMid, trialLow)
    residual = CoupledBarcelonaYield(trialTrace, trialQ, trialHigh, cc, ct, mAlpha, mGamma)
    Dim firstReturnReady As Boolean, deviatorFactor As Double
    multiplierDivisor = 1# ' Регулярный возврат отдаёт непосредственно множитель.
    If residual > 0# And residual <= YIELD_TOLERANCE_RATIO * cc Then
        dilationRate = CurrentDilation(kt, kc, r, trialTrace, trialQ)
        firstReturnReady = mReturn.TryRegularNearYield(elastic, trialTrace, trialQ, trialHigh, cc, ct, mFt, _
            dilationRate, residual, effective, increment, plasticMultiplier, deviatorFactor)
    End If
    If residual <= YIELD_TOLERANCE_RATIO * cc And Not firstReturnReady Then
        Dim fraction As Double
        fraction = (1# - r * dt) * (1# - dc)
        For i = 1 To N_COMPONENTS: apparent(i) = fraction * elastic(i): Next i
        virginElastic = (tension = 0# And compression = 0# And residual < -0.0002 * cc)
        UpdatePoint = True
        Exit Function
    End If

    If Not firstReturnReady Then dilationRate = CurrentDilation(kt, kc, r, trialTrace, trialQ)
    Dim trialMean As Double, highDev As Double, midDev As Double, lowDev As Double
    trialMean = trialTrace / 3#
    highDev = trialHigh - trialMean: midDev = trialMid - trialMean: lowDev = trialLow - trialMean
    Dim mean As Double, plasticMean As Double, plasticFactor As Double
    Dim eHigh As Double, eLow As Double
    Dim apexTransition As Boolean
    For j = 1 To MATERIAL_ITERATIONS + 1 ' После штатного цикла допускается один скалярный возврат.
        apexTransition = (j > MATERIAL_ITERATIONS)
        If apexTransition Then
            If Not TryTensileApexTransition(elastic, committed, trialTrace, trialQ, highDev, lowDev, _
                    effective, increment, dilationRate, plasticMultiplier, deviatorFactor) Then Exit Function
            multiplierDivisor = 1#
        ElseIf j > 1 Or Not firstReturnReady Then
            If Not mReturn.Correct(elastic, trialTrace, trialQ, trialHigh, cc, ct, mFt, dilationRate, _
                    effective, increment, plasticMultiplier, deviatorFactor, multiplierDivisor) Then
                If j <> 1 Or dilationRate <> 0# Or mDilationMax = 0# Then Exit Function
                ' Нулевая начальная дилатансия не позволяет уменьшить давление в растягивающей пробе.
                ' Начальное значение можно взять положительным, но принять только корень всех исходных условий ниже.
                dilationRate = Tan(mDilationMax)
                If Not mReturn.Correct(elastic, trialTrace, trialQ, trialHigh, cc, ct, mFt, dilationRate, _
                        effective, increment, plasticMultiplier, deviatorFactor, multiplierDivisor) Then Exit Function
            End If
        End If
        ' Изотропный возврат меняет среднее и масштаб девиатора, сохраняя его главные оси.
        trace = effective(1) + effective(2) + effective(3)
        mean = trace / 3#: q = deviatorFactor * trialQ
        high = mean + deviatorFactor * highDev
        mid = mean + deviatorFactor * midDev
        low = mean + deviatorFactor * lowDev
        r = PositiveStressFraction(high, mid, low)
        plasticMean = (increment(1) + increment(2) + increment(3)) / 3#
        plasticFactor = (1# - deviatorFactor) * mDeviatorCompliance
        eHigh = plasticMean + plasticFactor * highDev: eLow = plasticMean + plasticFactor * lowDev
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
        If apexTransition Then
            If r <> 1# Or nextCompression <> committed.Compression Then Exit Function
            If (1# + MIXED_DAMAGE_WEIGHT) * kt < 1# Then Exit Function
            ' Пластические координаты выведены непосредственно из найденного поля.
            ' Проверяются заново поверхность и закон с обновлённой историей.
            If Abs(residual) > YIELD_TOLERANCE_RATIO * cc Or Abs(deltaDilation) > STATE_TOLERANCE Then Exit Function
            If Not DilationConsistent(Abs(deltaDilation), plasticMultiplier, multiplierDivisor, _
                    YIELD_TOLERANCE_RATIO * cc / mPressureModulus) Then Exit Function
            Exit For
        End If
        If Abs(deltaTension) <= STATE_TOLERANCE * (mFt / mElastic(3, 3) + nextTension) And _
                Abs(deltaCompression) <= STATE_TOLERANCE * (mFc / mElastic(3, 3) + nextCompression) And _
                Abs(deltaDilation) <= STATE_TOLERANCE And Abs(residual) <= YIELD_TOLERANCE_RATIO * cc Then
            ' У вершины даже малое изменение угла может вывести давление за допуск поверхности.
            If DilationConsistent(Abs(deltaDilation), plasticMultiplier, multiplierDivisor, _
                    YIELD_TOLERANCE_RATIO * cc / mPressureModulus) Then Exit For
        End If
        tension = tension + MAXIMUM_DAMPING * deltaTension
        compression = compression + MAXIMUM_DAMPING * deltaCompression
        dilationRate = dilationRate + MAXIMUM_DAMPING * deltaDilation
        mTension.Evaluate tension, unusedStress, ct, dt, kt
        mCompression.Evaluate compression, unusedStress, cc, dc, kc
    Next j

    fraction = (1# - r * dt) * (1# - dc)
    For i = 1 To N_COMPONENTS
        updated.Plastic(i) = committed.Plastic(i) + increment(i)
        apparent(i) = fraction * effective(i)
    Next i
    updated.Tension = nextTension: updated.Compression = nextCompression
    updated.TensionCohesion = ct: updated.CompressionCohesion = cc
    updated.TensionDamage = dt: updated.CompressionDamage = dc
    updated.TensionEnergy = kt: updated.CompressionEnergy = kc
    UpdatePoint = True
End Function

' Передаёт спектральному закону копии принятого тензора и координат истории.
' Неудачная проба не изменяет ни принятый, ни подготовленный общий материал.
Private Function UpdateASDPoint(ByRef strain() As Double, ByRef committed As TConcreteHistory, _
        ByRef updated As TConcreteHistory, ByRef apparent() As Double) As Boolean
    Dim priorStrain(1 To 6) As Double, priorEffective(1 To 6) As Double
    Dim effective(1 To 6) As Double, plastic(1 To 6) As Double, i As Long
    updated = committed
    For i = 1 To 6: priorStrain(i) = committed.Strain(i): priorEffective(i) = committed.Effective(i): Next i
    If Not mASDFlow.TryStep(strain, priorStrain, priorEffective, committed.Tension, committed.Compression, _
            updated.Tension, updated.Compression, effective, apparent, plastic, updated.TensionDamage, updated.CompressionDamage) Then Exit Function
    For i = 1 To 6: updated.Effective(i) = effective(i): updated.Plastic(i) = plastic(i): Next i
    UpdateASDPoint = True
End Function

' Сверяет изменение объёмной деформации с тем же допуском поверхности.
' Масштабирует обе части неравенства, если множитель хранится дробью.
Private Function DilationConsistent(ByVal change As Double, ByVal multiplier As Double, _
        ByVal divisor As Double, ByVal tolerance As Double) As Boolean
    If divisor = 1# Then
        DilationConsistent = (change * multiplier <= tolerance)
    Else
        Dim largest As Double: largest = GeomMax(change, divisor)
        DilationConsistent = ((change / largest) * multiplier <= (divisor / largest) * tolerance)
    End If
End Function

' Возвращает долю положительных главных напряжений по абсолютной сумме.
Private Function PositiveStressFraction(ByVal high As Double, ByVal mid As Double, ByVal low As Double) As Double
    If high <= 0# Then Exit Function ' Упорядоченный спектр без растягивающих напряжений.
    If low >= 0# Then
        PositiveStressFraction = 1# ' При полностью растянутом спектре доля задаётся точно.
        Exit Function
    End If
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
    If q <= 0# Then
        confine = 1#
    ElseIf trace > 0# And q < TENSILE_AXIS_BAND_RATIO * mFc Then
        confine = TensileApexConfinement(trace, q)
    ElseIf q > TENSILE_AXIS_BAND_RATIO * mFc Or (trace > 0# And q >= TENSILE_AXIS_BAND_RATIO * mFc) Then
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

' Сопрягает гидростатическое значение и направленный закон в узкой полосе q.
' Вес (1-s)^3(1+3s+6s²) имеет нулевые первую и вторую производные на концах.
' Непрерывное сопряжение — численное допущение модели около оси растяжения.
Private Function TensileApexConfinement(ByVal trace As Double, ByVal q As Double) As Double
    Dim s As Double: s = q / (TENSILE_AXIS_BAND_RATIO * mFc)
    Dim remainder As Double: remainder = 1# - s
    Dim weight As Double: weight = GeomMin(1#, remainder * remainder * remainder * (1# + s * (3# + 6# * s))) ' Ограничивает ошибку округления веса около s = 0.
    Dim pressureRatio As Double: pressureRatio = trace / mFc
    Dim directional As Double
    If pressureRatio < 1# Then
        Dim intensityRatio As Double: intensityRatio = q / mFc
        Dim pressureSquared As Double: pressureSquared = pressureRatio * pressureRatio
        If pressureSquared < (745# - pressureRatio) * intensityRatio Then _
            directional = Exp(-pressureRatio - pressureSquared / intensityRatio)
    End If
    TensileApexConfinement = weight + (1# - weight) * directional
End Function

' Ищет локальный возврат внутри полосы после исчерпания обычных итераций.
' Использует только полностью растянутую ветвь с насыщенной энергией.
' Защита нижней оценки следа исключает ненулевой направленный член
' в полосе и сохраняет однозначность скобки для этой ветви.
Private Function TryTensileApexTransition(ByRef elastic() As Double, ByRef committed As TConcreteHistory, _
        ByVal trialTrace As Double, ByVal trialQ As Double, ByVal highDeviator As Double, _
        ByVal lowDeviator As Double, ByRef effective() As Double, ByRef increment() As Double, _
        ByRef dilation As Double, ByRef multiplier As Double, ByRef factor As Double) As Boolean
    If trialTrace <= 0# Or trialQ <= TENSILE_AXIS_BAND_RATIO * mFc Or mDilationMax = 0# Then Exit Function
    If (1# + MIXED_DAMAGE_WEIGHT) * committed.TensionEnergy < 1# Then Exit Function
    Dim cc As Double: cc = committed.CompressionCohesion
    Dim beta As Double: beta = cc / committed.TensionCohesion * (1# - mAlpha) - (1# + mAlpha)
    If beta < 0# Then Exit Function
    Dim pressureSlope As Double: pressureSlope = mAlpha + beta / 3#
    If pressureSlope <= 0# Then Exit Function
    Dim apexTrace As Double: apexTrace = cc * (1# - mAlpha) / pressureSlope
    Dim traceExcess As Double: traceExcess = trialTrace - apexTrace
    Dim deviatorSlope As Double: deviatorSlope = (trialQ + beta * highDeviator) / pressureSlope
    If deviatorSlope <= 0# Then Exit Function
    Dim bandWidth As Double: bandWidth = TENSILE_AXIS_BAND_RATIO * mFc
    Dim bandFactor As Double: bandFactor = bandWidth / trialQ
    If bandFactor <= 0# Or bandFactor >= 1# Then Exit Function
    Dim minTrace As Double: minTrace = apexTrace - deviatorSlope * bandFactor
    If minTrace / 3# + bandFactor * lowDeviator <= 7.105427357601E-15 * apexTrace Then Exit Function
    If minTrace <= mFc * Sqr(746# * TENSILE_AXIS_BAND_RATIO) Then Exit Function ' Направленный множитель должен округляться к нулю во всей скобке.
    Dim lower As Double
    If traceExcess < 0# Then lower = -traceExcess / deviatorSlope ' След возврата не должен превысить пробный след.
    If lower >= bandFactor Then Exit Function
    Dim trace As Double, volume As Double, requiredFlow As Double, potentialFlow As Double
    ApexTransitionFlow committed, trialQ, apexTrace, traceExcess, deviatorSlope, lower, trace, volume, dilation, requiredFlow, potentialFlow
    If requiredFlow >= potentialFlow Then Exit Function
    Dim upper As Double: upper = bandFactor
    ApexTransitionFlow committed, trialQ, apexTrace, traceExcess, deviatorSlope, upper, trace, volume, dilation, requiredFlow, potentialFlow
    If requiredFlow <= potentialFlow Then Exit Function
    Dim iteration As Long
    For iteration = 1 To 52
        factor = lower + (upper - lower) * 0.5
        If factor <= lower Or factor >= upper Then Exit For
        ApexTransitionFlow committed, trialQ, apexTrace, traceExcess, deviatorSlope, factor, trace, volume, dilation, requiredFlow, potentialFlow
        If requiredFlow > potentialFlow Then upper = factor Else lower = factor
    Next iteration
    ApexTransitionFlow committed, trialQ, apexTrace, traceExcess, deviatorSlope, factor, trace, volume, dilation, requiredFlow, potentialFlow
    If factor <= 0# Or factor >= bandFactor Or volume <= 0# Or dilation <= 0# Then Exit Function
    Dim returnedQ As Double: returnedQ = factor * trialQ
    If returnedQ <= 0# Or returnedQ >= bandWidth Then Exit Function
    If dilation < 1# Then
        If volume > 1E+308 * dilation Then Exit Function
    End If
    multiplier = volume / dilation
    ' Контроль объёмного приращения по решённому гиперболическому потоку.
    If Abs(requiredFlow - potentialFlow) * mDeviatorCompliance / factor > YIELD_TOLERANCE_RATIO * cc / mPressureModulus Then Exit Function
    mReturn.AssembleCorrection elastic, trialTrace, trace, factor, effective, increment
    TryTensileApexTransition = True
End Function

' Вычисляет обе стороны уравнения гиперболического потока в масштабе f=q/qtrial.
' След задаётся растянутой ветвью поверхности, объём — потерей следа.
Private Sub ApexTransitionFlow(ByRef committed As TConcreteHistory, ByVal trialQ As Double, ByVal apexTrace As Double, _
        ByVal traceExcess As Double, ByVal deviatorSlope As Double, ByVal factor As Double, ByRef trace As Double, _
        ByRef volume As Double, ByRef dilation As Double, ByRef requiredFlow As Double, ByRef potentialFlow As Double)
    Dim traceLoss As Double: traceLoss = deviatorSlope * factor
    trace = apexTrace - traceLoss
    volume = (traceExcess + traceLoss) / (3# * mPressureModulus)
    Dim q As Double: q = trialQ * factor
    dilation = CurrentDilation(committed.TensionEnergy, committed.CompressionEnergy, 1#, trace, q)
    Dim offset As Double: offset = dilation * mApexPotentialOffset
    Dim dev As Double: dev = 0.816496580927726 * q ' Норма девиатора из второго инварианта.
    requiredFlow = volume * factor / mDeviatorCompliance
    potentialFlow = (1# - factor) * dilation * Sqr(offset * offset + dev * dev)
End Sub

' Проверяет диапазон углов до перевода в радианы.
Private Function InvalidDilationAngles(ByVal friction As Double, ByVal dilation As Double) As Boolean
    InvalidDilationAngles = (friction <= 0# Or friction >= 90# Or dilation < 0# Or dilation >= friction)
End Function

' Передаёт причину неверных материальных характеристик вызывающему слою.
Private Sub InvalidInput(ByVal reason As String)
    Err.Raise vbObjectError + 1783, "CCoupledConcreteResponse", reason
End Sub
