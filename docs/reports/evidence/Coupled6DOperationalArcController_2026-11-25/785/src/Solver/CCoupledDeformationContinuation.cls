VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledDeformationContinuation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Поиск предела связанной ветви при управлении деформацией
' ==========================================================================
' Приращения меры деформации положительны. Плечи предполагаемого максимума
' проверяются повторно с меньшим шагом от восстановленной истории левого
' плеча. Численный отказ, плато и предел диапазона не считаются максимумом.
' В железобетоне локальный максимум может предшествовать набору нагрузки
' на трещиноватой ветви. Он сохраняется в диагностике, поиск идёт дальше.

Private Type TBranchPoint
    Coordinate As Double
    Lambda As Double
    Uncertainty As Double
    HistoryIndex As Long
End Type

Private Type TSteelCrossing
    Admissible As TBranchPoint
    Exceeded As TBranchPoint
    AdmissibleRatio As Double
    ExceededRatio As Double
    CriticalPoint As Long
End Type

Private Const COMPONENT_COUNT As Long = 6
Private Const STEP_LIMIT As Long = 32
Private Const PROBE_LIMIT As Long = 96 ' Общий предел проб продолжения, деления и повторного прохождения плеч.
Private Const FIXED_YIELD_RECOVERY_LIMIT As Long = 4 ' Проверки перехода без смены меры за одно продолжение.
Private Const YIELD_RECOVERY_LIMIT As Long = 4 ' Для CDP на весь участок, для знаковых границ ASD на одну ревизию принятой истории; общий бюджет проб сохраняется.
Private Const MIN_STEP_RATIO As Double = 0.015625
Private Const NEAR_ONE_TOLERANCE As Double = 0.0001
Private Const STEP_GROWTH As Double = 1.25
Private Const MATERIAL_STEP_GROWTH As Double = 1.25
Private Const ENVELOPE_INITIAL_STEP_RATIO As Double = 0.0001 ' Начальный численный масштаб H относительно уже накопленного уровня, не допуск материала.
Private Const ENVELOPE_MAXIMUM_STEP_RATIO As Double = 0.02 ' Верхний относительный шаг H ограничивает рост после лёгких решений, не меняя материальные огибающие.
Private Const POST_PEAK_STEP_RATIO As Double = 2# ' Верхний шаг относительно шага, выбранного при смене меры.
Private Const POST_PEAK_EXTENDED_RATIO As Double = 4# ' Верхний шаг после шести гладких приращений подряд.
Private Const TRANSVERSE_STEP_RATIO As Double = 32# ' Ограничивает рост сдвигового шага относительно различимой начальной секущей.
Private Const POST_PEAK_SECANT_TOLERANCE As Double = 0.02 ' Допустимое отклонение приращений поля и множителя от секущей.

Private mSolver As CCoupledNewtonSolver, mReinforcement As CCoupledReinforcedResponse
Private mOffset() As Double, mDirection() As Double, mProjection() As Double
Private mAbsTol As Double, mRelTol As Double, mLambdaTolerance As Double
Private mMaxIterations As Long, mMaximumLambda As Double, mBestLambda As Double
Private mBestBeforeCurrent As Double ' Наибольший множитель до последнего принятого поля.
Private mControlName As String, mProbeDetails As String, mMeta As CResultMeta
Private mPeak As TBranchPoint, mPeakAccuracy As Double
Private mHasPeak As Boolean
Private mVerifiedPeak As TBranchPoint, mVerifiedPeakAccuracy As Double
Private mVerifiedPeakFields As Collection, mVerifiedPeakTrace As CCoupledLoadTrace
Private mVerifiedPeakDesignIndex As Long
Private mDesignHistoryIndex As Long ' Индекс правого принятого узла, подтверждающего сочетание.
Private mLocalPeak As Double, mLocalPeakAccuracy As Double
Private mFollowingLocalPeak As Boolean, mRecoveredPeaks As Long
Private mReusedRightPoints As Long
Private mShoulderProbes As Long
Private mSteelLow As Double, mSteelHigh As Double, mSteelPoint As Long
Private mSteelEventPoint As Long, mSteelEventDetails As String
Private mProbeCount As Long
Private mYieldRecoveryProbes As Long, mYieldRecoveredPoints As Long
Private mSpatialReversalsRejected As Long
Private mPrincipalCornerProbes As Long, mPrincipalCorners As Long
Private mCornerAttemptRevision As Long, mCornerAttemptsAtRevision As Long
Private mHasCornerReference As Boolean, mCornerReferenceRevision As Long
Private mCornerDirection() As Double, mCornerLambdaDirection As Double, mCornerReferenceIncrement As Double
Private mHasCornerProbe As Boolean, mCornerProbe() As Double, mCornerProbeLambda As Double, mCornerProbeCoordinate As Double
Private mCornerProbeSeeds As Long
Private mMaterialGrowthActive As Boolean, mMaterialGrowthProjection() As Double
Private mMaterialStepExpansions As Long, mMaterialMaximumStep As Double
Private mMaterialRecoveryRevision As Long, mMaterialRecoveryAtRevision As Long
Private mFixedYieldRecoveryProbes As Long
Private mArcRecoveryTried As Boolean, mArcRecoveryProbes As Long
Private mUsesYieldControl As Boolean
Private mBranchTangentControl As Boolean ' Согласованные с ветвью разности нужны для восстановления CDP; спектральный ASD использует свою касательную.
Private mPostPeakControlName As String, mPostPeakControlPoints As Long
Private mPostPeakWidePoints As Long
Private mLimitFields As Collection
Private mLimitResultants() As Double
Private mCurrentResultants() As Double
Private mTrace As CCoupledLoadTrace, mLimitTrace As CCoupledLoadTrace
Private mCurrentFields As Collection ' Принятый путь для воспроизводимой диагностики незавершённого поиска.
Private mUsesTransverseControl As Boolean, mUnverifiedDecline As Boolean
Private mUsesFieldControl As Boolean, mFieldControlChanges As Long
Private mLastFieldControlRevision As Long
Private mLastHybridRevision As Long, mLastHybridProjection() As Double
Private mLastHybridTarget As Double
Private mHybridRecoveryProbes As Long, mHybridRecoveredPoints As Long
Private mFailedProbeRevision As Long, mFailedProbeProjection() As Double
Private mFailedProbeTarget As Double, mFailedProbeCode As EResultCode
Private mUnresolvedLoadRetries As Long
Private mUnverifiedReference As Double, mUnverifiedDips As Long, mRecoveredDips As Long
Private mObserver As ICoupledContinuationProgress
Private mEnvelopePrevious As Double, mEnvelopeCurrent As Double, mEnvelopeLastIncrement As Double
Private mEnvelopeHistoryIndex As Long, mEnvelopeAttemptRevision As Long, mEnvelopeAttemptsAtRevision As Long
Private mEnvelopeProbes As Long, mEnvelopeRecoveries As Long
Private mUsesEnvelopeControl As Boolean
Private mEnvelopeResolution As Double
Private mEnvelopeDepartures As Long, mEnvelopeDepartureRevision As Long
Private mDepartureRecoveryRevision As Long, mDepartureRecoveryProbes As Long, mDepartureRecoveredPoints As Long ' Отдельный резерв не повторяется на той же принятой истории и расходует общий бюджет.
Private mGuideRecoveryProbes As Long, mGuideRecoverySolutions As Long ' Численные направляющие пробы и успешные возвраты; материал принимает общий путь контроллера.
Private mSignedArcRevision As Long, mSignedArcAttempts As Long, mSignedArcRatio As Double
Private mSignedArcProbes As Long, mSignedArcRecovered As Long

' Начинает деформационный участок после двух принятых силовых точек.
' Выбор меры производится по их фактической разности. История первых
' точек принадлежит той же траектории, что и текущий решатель.
Friend Sub Trace(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef offset() As Double, ByRef direction() As Double, ByVal history As Collection, _
        ByVal previousLambda As Double, ByVal currentLambda As Double, ByRef initialResultants() As Double, _
        ByVal maximumLambda As Double, ByVal lambdaTolerance As Double, ByVal maximumIterations As Long, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, Optional ByVal loadTrace As CCoupledLoadTrace = Nothing, _
        Optional ByVal observer As ICoupledContinuationProgress = Nothing)
    Set mSolver = solver: Set mReinforcement = reinforcement
    Set mObserver = observer
    mOffset = offset: mDirection = direction
    mMaxIterations = maximumIterations: mMaximumLambda = maximumLambda
    mLambdaTolerance = lambdaTolerance: mAbsTol = absoluteTolerance: mRelTol = relativeTolerance
    mBestBeforeCurrent = previousLambda: mBestLambda = currentLambda ' Начальный силовой участок монотонно возрастает.
    mDesignHistoryIndex = 0
    mHasPeak = False: mSteelPoint = 0: mProbeCount = 0
    Set mVerifiedPeakFields = Nothing: Set mVerifiedPeakTrace = Nothing
    mVerifiedPeakAccuracy = 0#: mVerifiedPeakDesignIndex = 0
    mSteelEventPoint = 0: mSteelEventDetails = vbNullString
    mYieldRecoveryProbes = 0: mYieldRecoveredPoints = 0: mFixedYieldRecoveryProbes = 0
    mSpatialReversalsRejected = 0
    mPrincipalCornerProbes = 0: mPrincipalCorners = 0: mHasCornerReference = False
    mHasCornerProbe = False: Erase mCornerProbe: mCornerProbeSeeds = 0
    mMaterialGrowthActive = False: Erase mMaterialGrowthProjection
    mMaterialStepExpansions = 0: mMaterialMaximumStep = 0#
    mEnvelopePrevious = 0#: mEnvelopeCurrent = 0#: mEnvelopeLastIncrement = 0#: mEnvelopeHistoryIndex = 0
    mEnvelopeAttemptRevision = -1: mEnvelopeAttemptsAtRevision = 0: mEnvelopeProbes = 0: mEnvelopeRecoveries = 0
    mUsesEnvelopeControl = False: mEnvelopeResolution = 0#
    mEnvelopeDepartures = 0: mEnvelopeDepartureRevision = -1
    mDepartureRecoveryRevision = -1: mDepartureRecoveryProbes = 0: mDepartureRecoveredPoints = 0
    mCornerAttemptRevision = -1: mCornerAttemptsAtRevision = 0
    mMaterialRecoveryRevision = -1: mMaterialRecoveryAtRevision = 0
    mArcRecoveryTried = False: mArcRecoveryProbes = 0
    mUsesYieldControl = False
    mBranchTangentControl = False
    mUsesTransverseControl = False: mUnverifiedDecline = False
    mUsesFieldControl = False: mFieldControlChanges = 0
    mLastFieldControlRevision = -1
    mLastHybridRevision = -1: Erase mLastHybridProjection
    mLastHybridTarget = 0#
    mHybridRecoveryProbes = 0: mHybridRecoveredPoints = 0
    mFailedProbeRevision = -1: Erase mFailedProbeProjection
    mFailedProbeTarget = 0#: mFailedProbeCode = rcNone: mUnresolvedLoadRetries = 0
    mUnverifiedReference = 0#: mUnverifiedDips = 0: mRecoveredDips = 0
    mPostPeakControlName = vbNullString: mPostPeakControlPoints = 0
    mPostPeakWidePoints = 0
    mLocalPeak = 0#: mLocalPeakAccuracy = 0#
    mFollowingLocalPeak = False: mRecoveredPeaks = 0
    mGuideRecoveryProbes = 0: mGuideRecoverySolutions = 0
    mSignedArcRevision = -1: mSignedArcAttempts = 0: mSignedArcRatio = 1#: mSignedArcProbes = 0: mSignedArcRecovered = 0
    mReusedRightPoints = 0
    mShoulderProbes = 0
    Set mLimitFields = Nothing: Erase mLimitResultants
    Set mTrace = Nothing: Set mLimitTrace = Nothing
    Set mCurrentFields = Nothing
    If Not loadTrace Is Nothing Then Set mTrace = loadTrace.CopyPrefix(loadTrace.Count)
    mCurrentResultants = initialResultants
    mProbeDetails = vbNullString
    Set mMeta = New CResultMeta
    If history Is Nothing Then
        StopSearch "Для продолжения не сохранены принятые поля."
        Exit Sub
    End If
    If history.Count < 2 Or previousLambda >= currentLambda Then
        StopSearch "Для продолжения нужны две последовательные точки возрастающей силовой ветви."
        Exit Sub
    End If
    If Not mTrace Is Nothing Then
        If mTrace.Count <> history.Count Then
            StopSearch "Количество записей нагрузочной траектории не совпадает с принятой историей полей."
            Exit Sub
        End If
        mBestBeforeCurrent = PrefixMaximumLambda(history.Count - 1)
        mBestLambda = GeomMax(mBestBeforeCurrent, currentLambda) ' Сохраняемый префикс может содержать ранее пройденный локальный максимум.
    End If

    Set mCurrentFields = history
    mDesignHistoryIndex = InitialDesignHistoryIndex(history.Count, previousLambda, currentLambda)

    Dim left As TBranchPoint, current As TBranchPoint, trial As TBranchPoint
    Dim previousField() As Double, field() As Double, increment As Double, minimum As Double, nominal As Double
    left.HistoryIndex = history.Count - 1: left.Lambda = previousLambda
    current.HistoryIndex = history.Count: current.Lambda = currentLambda
    SelectInitialAnchor history, current, left
    previousField = history(left.HistoryIndex): field = history(current.HistoryIndex)
    Dim selected As Boolean
    If mSolver.UsesSpatialHistory Then selected = mSolver.SelectTransverseDeformationControl(previousField, field, _
        mDirection, mProjection, current.Coordinate, nominal, mControlName)
    mUsesTransverseControl = selected
    If Not selected Then selected = mSolver.SelectDeformationControl(previousField, field, mProjection, current.Coordinate, nominal, mControlName)
    If Not selected Then
        StopSearch "Не удалось выбрать меняющуюся меру деформации."
        Exit Sub
    End If
    left.Coordinate = CoordinateOf(previousField)
    nominal = current.Coordinate - left.Coordinate
    If nominal <= 0# Then
        StopSearch "Приращение начальной меры деформации не различается в принятом поле."
        Exit Sub
    End If
    increment = nominal * 0.5
    minimum = increment * MIN_STEP_RATIO
    SelectInitialEnvelopeControl left, current, previousLambda, nominal, increment, minimum

    Dim attempt As Long, reinforced As Boolean, probed As Boolean, changedControl As Boolean
    Dim lastProbeFailed As Boolean, lastNoProgress As Boolean, lastLoadUnresolved As Boolean
    Dim unresolvedSteel As Boolean, hybridRecovery As Boolean, envelopeRecovered As Boolean, departureRecovered As Boolean
    Dim arcRecovered As Boolean
    Dim targetCoordinate As Double
    Dim smoothTrial As Boolean, smoothSteps As Long, stepCeiling As Double, designAtTrial As Boolean, allowMaterialRecovery As Boolean
    Dim recoveryProjection() As Double
    Dim cumulativeLeft As TBranchPoint, cumulativeCurrent As TBranchPoint
    Dim arcFields As Collection
    If Not reinforcement Is Nothing Then reinforced = reinforcement.HasSteel
    For attempt = 1 To PROBE_LIMIT
        hybridRecovery = False: envelopeRecovered = False: departureRecovered = False: arcRecovered = False
        If mProbeCount >= PROBE_LIMIT Then Exit For
        If mUsesEnvelopeControl And mProbeCount + 2 > PROBE_LIMIT Then
            StopSearch "Остатка общего бюджета недостаточно для следующего материального шага с обязательной проверкой поворота."
            Exit Sub
        End If
        If attempt > STEP_LIMIT And mLocalPeak = 0# And mUnverifiedDips = 0 And Not mUsesEnvelopeControl And _
                mEnvelopeDepartures = 0 And mDepartureRecoveredPoints = 0 Then ' Принятый материальный переход завершает начальный участок; общий бюджет проб продолжает действовать.
            StopSearch "Достигнут лимит начального участка деформационного продолжения до обнаружения спада нагрузки."
            Exit Sub
        End If
        If increment < minimum Then
            If Not unresolvedSteel Then arcRecovered = TryArcContinuationProbe(history, current, trial, nominal, increment, minimum)
            If arcRecovered Then
                lastProbeFailed = False: lastNoProgress = False: lastLoadUnresolved = False: smoothSteps = 0
            ElseIf mUsesEnvelopeControl Then
                If unresolvedSteel Then Exit For ' Наблюдавшееся пересечение предела стали нельзя обходить сменой численной координаты.
                If Not TryLeaveEnvelopeControl(history, left, current, nominal, increment, minimum) Then
                    StopSearch "Материальное продолжение потребовало шаг ниже различимого масштаба до подтверждения предела."
                    Exit Sub
                End If
                lastProbeFailed = False: lastNoProgress = False: lastLoadUnresolved = False: smoothSteps = 0
            Else
                If Not (lastProbeFailed Or lastNoProgress Or lastLoadUnresolved) Or unresolvedSteel Then Exit For
                Dim enteredCorner As Boolean
                If PrepareHybridRecovery(lastProbeFailed, lastLoadUnresolved, current.Coordinate, increment) Then
                    hybridRecovery = True
                ElseIf TryCrossPrincipalCorner(history, left, current, nominal, increment, minimum, enteredCorner) Then
                    lastProbeFailed = False: lastNoProgress = False: smoothSteps = 0
                ElseIf enteredCorner Then
                    StopSearch "Материальный переход не удалось завершить с проверенным направлением продолжения."
                    Exit Sub
                ElseIf Not TryChangePostPeakControl(history, left, current, nominal, increment, minimum, lastNoProgress) Then
                    If Not TryChangeStalledFieldControl(history, left, current, nominal, increment, minimum) Then
                        If mSolver.HasPrincipalBoundaries Then Exit For
                        Dim enteredMaterialReserve As Boolean
                        If TryCrossPrincipalCorner(history, left, current, nominal, increment, minimum, enteredMaterialReserve, True) Then
                            lastProbeFailed = False: lastNoProgress = False: smoothSteps = 0
                        ElseIf enteredMaterialReserve Then
                            StopSearch "Материальный переход не удалось завершить с проверенным направлением продолжения."
                            Exit Sub
                        Else
                            envelopeRecovered = TryEnvelopeRecovery(history, left, current, trial, nominal, increment, minimum)
                            If Not envelopeRecovered Then
                                If reinforced Then departureRecovered = TryRecoverEnvelopeDeparture(history, current, trial)
                                If Not departureRecovered Then Exit For
                            End If
                        End If
                    End If
                End If
            End If
        End If
        Dim seed As Variant, seedLambda As Double
        If envelopeRecovered Or departureRecovered Or arcRecovered Then
            targetCoordinate = trial.Coordinate: probed = True
        Else
            targetCoordinate = current.Coordinate + increment
            If targetCoordinate <= current.Coordinate Then Exit For
            trial.Coordinate = targetCoordinate
            Set arcFields = Nothing
            If Not mUsesEnvelopeControl And Not mArcRecoveryTried And Not mUsesYieldControl And Not mFollowingLocalPeak And mDesignHistoryIndex > 0 Then _
                Set arcFields = history ' Повтор точного lambda=1 уже принят, поэтому новая цель не перерешается.
            allowMaterialRecovery = MayRecoverMaterialBoundary()
            If PrepareSecantSeed(history, left, current, trial.Coordinate, seed, seedLambda) Then
                probed = Probe(trial, seedLambda, allowMaterialRecovery, _
                    arcFields:=arcFields, initialField:=seed, allowHybridRecovery:=hybridRecovery, acceptedLambda:=current.Lambda)
            Else
                probed = Probe(trial, current.Lambda, allowMaterialRecovery, _
                    arcFields:=arcFields, allowHybridRecovery:=hybridRecovery, acceptedLambda:=current.Lambda)
            End If
        End If
        lastProbeFailed = Not probed
        lastNoProgress = False: lastLoadUnresolved = False
        changedControl = envelopeRecovered Or arcRecovered: smoothTrial = False
        If probed And Not envelopeRecovered And Not mUsesEnvelopeControl Then
            If mSolver.CopyRecoveredControl(recoveryProjection) Then
                probed = TryAdoptRecoveredControl(history, current, trial, recoveryProjection, nominal, increment, minimum)
                changedControl = probed
                If probed Then targetCoordinate = trial.Coordinate Else mSolver.DiscardStep
            End If
        End If
        If probed And reinforced And Not envelopeRecovered And Not arcRecovered And Not mUsesEnvelopeControl Then
            If Not BeyondSteelLimit() And IsEnvelopeUnloading(current, trial) Then
                mSolver.DiscardStep
                envelopeRecovered = TryEnvelopeRecovery(history, left, current, trial, nominal, increment, minimum)
                probed = envelopeRecovered: changedControl = envelopeRecovered
                If probed Then targetCoordinate = trial.Coordinate Else lastProbeFailed = True
            End If
        End If
        If probed Then
            probed = UpdateTrialCoordinate(trial, current.Coordinate)
            If probed And Not envelopeRecovered And Not mUsesEnvelopeControl And _
                    (mSolver.HasPrincipalBoundaries Or mFieldControlChanges > 0 Or hybridRecovery) Then
                probed = SpatialTrialAdvances(history, current)
                If Not probed Then mSolver.DiscardStep
            End If
            lastNoProgress = Not probed
            If probed And Not envelopeRecovered Then
                If mUsesEnvelopeControl Then
                    smoothTrial = (mSolver.IterationCount <= GeomMax(1, mMaxIterations \ 2)) ' Как в адаптивном управлении OpenSees, затраты предыдущего решения ограничивают рост следующего шага.
                ElseIf IsMaterialGrowthControl() Then
                    smoothTrial = FitsMaterialGrowthSecant(history, left, current, trial)
                ElseIf mLocalPeak > 0# Then
                    smoothTrial = FitsMaterialGrowthSecant(history, left, current, trial)
                End If
            End If
        End If
        If Not probed Then
            smoothSteps = 0
            Dim enteredAfterFailure As Boolean
            If Not unresolvedSteel Then
                If TryCrossPrincipalCorner(history, left, current, nominal, increment, minimum, enteredAfterFailure) Then
                    lastProbeFailed = False: lastNoProgress = False
                ElseIf enteredAfterFailure Then
                    StopSearch "Материальный переход не удалось завершить с проверенным направлением продолжения."
                    Exit Sub
                Else
                    increment = increment * 0.5
                End If
            Else
                increment = increment * 0.5
            End If
        ElseIf Not mUsesEnvelopeControl And (IsMaterialGrowthControl() Or mLocalPeak > 0#) And _
                increment > nominal And Not smoothTrial Then
            If BeyondSteelLimit() Then unresolvedSteel = True ' Уточнение требует меньшей пробы на прослеженной ветви.
            mSolver.DiscardStep ' На прежнем размере можно пройти изменение наклона без экстраполяции ветви.
            smoothSteps = 0
            increment = increment * 0.5
        ElseIf trial.Lambda > maximumLambda Then
            StopAtSearchBound
            Exit Sub
        ElseIf BeyondSteelLimit() Then
            unresolvedSteel = True ' Смена меры не может пропустить наблюдавшееся пересечение предела арматуры.
            If ResolveSteelCrossing(history, current, trial) Then Exit Sub
            smoothSteps = 0
            increment = increment * 0.5
        ElseIf SelectCumulativePeak(history, trial, cumulativeLeft, cumulativeCurrent) Then
            Dim cumulativeTrial As TBranchPoint, cumulativeField() As Double, cumulativeResultants() As Double
            cumulativeTrial = trial ' Уточнение плеч может заменить переданные узлы; исходный сошедшийся спад сохраняется отдельно.
            mSolver.CopyTrial cumulativeField, cumulativeResultants
            mSolver.DiscardStep
            If VerifyPeak(history, cumulativeLeft, cumulativeCurrent, trial, reinforced) Then
                If Not reinforced Then Exit Sub
                left = cumulativeLeft: current = cumulativeCurrent
                ResumeAfterLocalPeak history, current, trial
                increment = nominal * 0.25: minimum = increment * MIN_STEP_RATIO
            ElseIf reinforced And mSolver.UsesSpatialHistory And (mUsesTransverseControl Or mUsesFieldControl) Then
                trial = cumulativeTrial
                Dim cumulativeResumed As Boolean, cumulativeSeed As Variant
                cumulativeSeed = cumulativeField
                If Not ReplayUnverifiedDecline(history, left, current, trial, cumulativeSeed, cumulativeResumed) Then Exit Sub
                smoothSteps = 0
                If cumulativeResumed Then
                    lastProbeFailed = False: lastNoProgress = False: lastLoadUnresolved = False
                Else
                    lastProbeFailed = True: increment = increment * 0.5
                End If
            Else
                StopSearch "Суммарный спад найден, но его максимум не подтверждён повтором плеч." & " Причина: " & mProbeDetails
                Exit Sub
            End If
        ElseIf ((changedControl Or mUnverifiedDecline) And Not mUsesEnvelopeControl) Or mFollowingLocalPeak Or _
                trial.Lambda > current.Lambda + LoadResolution(current.Lambda) Or _
                (mUsesEnvelopeControl And trial.Lambda >= current.Lambda - LoadResolution(current.Lambda)) Or _
                AdvancesWithSmallLoadChange(history, current, trial) Then
            If trial.Lambda <= 0# Then
                mSolver.DiscardStep
                StopSearch "После локального максимума множитель вышел из области положительной нагрузки."
                Exit Sub
            End If
            designAtTrial = False
            If mDesignHistoryIndex = 0 And current.Lambda < 1# And trial.Lambda >= 1# And mProbeCount + 2 <= PROBE_LIMIT Then
                If Not RepeatDesignTrial(current, trial, targetCoordinate, designAtTrial, changedControl And mSolver.HasPrincipalBoundaries) Then Exit Sub
            End If
            Dim acceptedIncrement As Double: acceptedIncrement = trial.Coordinate - current.Coordinate
            AcceptPoint history, trial
            If hybridRecovery Then mHybridRecoveredPoints = mHybridRecoveredPoints + 1
            If designAtTrial And trial.Lambda >= 1# Then mDesignHistoryIndex = trial.HistoryIndex
            unresolvedSteel = False
            If changedControl And Not departureRecovered Then mYieldRecoveredPoints = mYieldRecoveredPoints + 1
            If departureRecovered Then mDepartureRecoveredPoints = mDepartureRecoveredPoints + 1
            If Len(mPostPeakControlName) > 0 Then
                If mPostPeakControlPoints = 0 Then mControlName = mPostPeakControlName
                mPostPeakControlPoints = mPostPeakControlPoints + 1
                If increment > nominal Then mPostPeakWidePoints = mPostPeakWidePoints + 1
            End If
            left = current: current = trial
            If mFollowingLocalPeak Then CheckRecovery current.Lambda
            If mUnverifiedDecline And current.Lambda > mUnverifiedReference + LoadResolution(mUnverifiedReference) Then
                mUnverifiedDecline = False: mRecoveredDips = mRecoveredDips + 1
            End If
            If smoothTrial And (mLocalPeak > 0# Or IsMaterialGrowthControl()) Then smoothSteps = smoothSteps + 1 Else smoothSteps = 0
            stepCeiling = nominal
            If IsMaterialGrowthControl() Or mLocalPeak > 0# Then
                increment = acceptedIncrement ' Доверенный масштаб и следующая цель относятся к фактически принятому полю, а не к прогнозу.
                mMaterialMaximumStep = GeomMax(mMaterialMaximumStep, increment)
                If smoothSteps >= 2 Then
                    If increment > nominal Then mMaterialStepExpansions = mMaterialStepExpansions + 1
                    nominal = GeomMax(nominal, increment) ' Новый доверенный масштаб получен из принятого равновесия с малой ошибкой прогноза.
                    minimum = nominal * MIN_STEP_RATIO * MIN_STEP_RATIO
                    stepCeiling = MATERIAL_STEP_GROWTH * nominal
                End If
            Else
                If smoothSteps >= 2 Then stepCeiling = POST_PEAK_STEP_RATIO * nominal
                If smoothSteps >= 6 Then stepCeiling = POST_PEAK_EXTENDED_RATIO * nominal
            End If
            If mUsesTransverseControl Then stepCeiling = TRANSVERSE_STEP_RATIO * nominal
            If mUsesEnvelopeControl Then
                stepCeiling = GeomMin(stepCeiling, GeomMax(4# * mEnvelopeResolution, ENVELOPE_MAXIMUM_STEP_RATIO * current.Coordinate))
                minimum = 2# * mEnvelopeResolution ' Растущий номинальный шаг не запрещает последующее уточнение до действительного разрешения H.
            End If
            If increment > stepCeiling Then increment = stepCeiling
            If (IsMaterialGrowthControl() Or mLocalPeak > 0#) And smoothSteps >= 2 Then
                increment = GeomMin(stepCeiling, increment * MATERIAL_STEP_GROWTH)
            ElseIf increment < stepCeiling / STEP_GROWTH Then
                increment = increment * STEP_GROWTH
            End If
        ElseIf trial.Lambda < current.Lambda - LoadResolution(current.Lambda) Then
            Dim declinedField() As Double, declinedResultants() As Double
            If reinforced And (mUsesTransverseControl Or mUsesFieldControl) Then mSolver.CopyTrial declinedField, declinedResultants
            mSolver.DiscardStep
            If VerifyPeak(history, left, current, trial, reinforced) Then
                If Not reinforced Then Exit Sub
                ResumeAfterLocalPeak history, current, trial
                increment = nominal * 0.25
                minimum = increment * MIN_STEP_RATIO
            Else
                If reinforced And (mUsesTransverseControl Or mUsesFieldControl) Then
                    Dim repeatSeed As Variant, declineResumed As Boolean
                    repeatSeed = declinedField
                    If Not ReplayUnverifiedDecline(history, left, current, trial, repeatSeed, declineResumed) Then Exit Sub
                    If Not declineResumed Then lastProbeFailed = True: increment = increment * 0.5
                Else
                    If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
                        StopSearch "Не удалось восстановить историю возрастающей ветви."
                        Exit Sub
                    End If
                    increment = increment * 0.5
                End If
            End If
        Else
            lastLoadUnresolved = True ' Дальнейшее деление уже сошедшегося шага не делает малое изменение lambda различимым.
            mSolver.DiscardStep
            increment = increment * 0.5
        End If
    Next attempt
    If mProbeCount >= PROBE_LIMIT Then
        StopSearch "Исчерпан общий лимит проб деформационного продолжения и проверки максимума."
    Else
        StopSearch "Деформационное продолжение остановилось до подтверждения физического предела."
    End If
End Sub


' Повторяет сошедшийся спад от исходной принятой истории после неудачи проверки плеч.
' Возвращает признак восстановления истории отдельно от приёмки нового поля; неподтверждённый максимум не становится пределом.
Private Function ReplayUnverifiedDecline(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef trial As TBranchPoint, ByRef prediction As Variant, ByRef resumed As Boolean) As Boolean
    resumed = False
    If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
        StopSearch "Не удалось восстановить принятую историю после численной неудачи проверки максимума."
        Exit Function
    End If
    ReplayUnverifiedDecline = True
    If Not Probe(trial, trial.Lambda, initialField:=prediction, acceptedLambda:=current.Lambda) Then Exit Function
    If Not UpdateTrialCoordinate(trial, current.Coordinate) Or BeyondSteelLimit() Or trial.Lambda <= 0# Then
        mSolver.DiscardStep
        Exit Function
    End If
    If Not mUsesEnvelopeControl And (mSolver.HasPrincipalBoundaries Or mFieldControlChanges > 0) Then
        If Not SpatialTrialAdvances(history, current) Then mSolver.DiscardStep: Exit Function
    End If
    mUnverifiedDecline = True: mUnverifiedReference = mBestLambda
    mUnverifiedDips = mUnverifiedDips + 1
    AcceptPoint history, trial
    left = current: current = trial
    resumed = True
End Function

' Выбирает H сразу при различимом росте в последних двух принятых состояниях; без такой пары сохраняет выбранную геометрическую меру.
' Начальный масштаб и дальнейшая адаптация относятся к численному продолжению, не расширяют допустимую область материалов.
Private Sub SelectInitialEnvelopeControl(ByRef left As TBranchPoint, ByRef current As TBranchPoint, ByVal previousLambda As Double, _
        ByRef nominal As Double, ByRef increment As Double, ByRef minimum As Double)
    If Not mReinforcement Is Nothing Then
        If mReinforcement.HasSteel Then Exit Sub ' Между изменениями бетона сталь может воспринимать нагрузку при практически постоянном H; начальная геометрическая мера сохраняется.
    End If
    Dim field() As Double, before As Double, after As Double, baseline As Double, resolution As Double
    If Not mSolver.TryAcceptedEnvelopeSegment(field, before, after) Then Exit Sub
    If Not mSolver.TryEnvelopeControlState(mOffset, mDirection, current.Lambda, mAbsTol, mRelTol, baseline, resolution) Then Exit Sub
    If after <> baseline Or baseline <= 0# Or after - before <= 2# * resolution Then Exit Sub
    left.HistoryIndex = current.HistoryIndex - 1: left.Coordinate = before: left.Lambda = previousLambda
    If Not mTrace Is Nothing Then
        Dim fraction As Double
        mTrace.ReadPoint left.HistoryIndex, fraction, left.Lambda
    End If
    current.Coordinate = after
    nominal = GeomMax(after - before, GeomMax(8# * resolution, ENVELOPE_INITIAL_STEP_RATIO * baseline))
    increment = GeomMin(0.5 * nominal, GeomMax(4# * resolution, ENVELOPE_MAXIMUM_STEP_RATIO * baseline))
    minimum = 2# * resolution: mEnvelopeResolution = resolution
    mUsesEnvelopeControl = True: mUsesTransverseControl = False: mUsesFieldControl = True
    mMaterialGrowthActive = True: mEnvelopeLastIncrement = after - before
    mControlName = "материальная координата H огибающих"
End Sub

' Возвращает геометрическое управление после неразличимого роста H, не принимая отвергнутую пробу и не меняя нагрузочную траекторию.
' Направление задают последние принятые поля; масштаб берётся из ближайшего различимого участка той же истории.
Private Function TryLeaveEnvelopeControl(ByVal history As Collection, ByRef left As TBranchPoint, ByRef current As TBranchPoint, _
        ByRef nominal As Double, ByRef increment As Double, ByRef minimum As Double) As Boolean
    If Not mUsesEnvelopeControl Or current.HistoryIndex < 2 Or current.HistoryIndex <> history.Count Then Exit Function
    If mEnvelopeDepartureRevision = mSolver.AcceptedRevision Then Exit Function
    Dim previous() As Double, accepted() As Double, movement() As Double, projection() As Double, coordinate As Double, spacing As Double
    previous = history.Item(current.HistoryIndex - 1): accepted = history.Item(current.HistoryIndex)
    ReDim movement(1 To UBound(accepted))
    Dim degree As Long
    For degree = 1 To mSolver.PhysicalDegreeCount: movement(degree) = accepted(degree) - previous(degree): Next degree
    If Not mSolver.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then Exit Function
    If spacing <= 0# Then Exit Function
    Dim reference() As Double, index As Long, fraction As Double, factor As Double
    If Not mTrace Is Nothing Then
        For index = current.HistoryIndex - 1 To 1 Step -1
            mTrace.ReadPoint index, fraction, factor
            If fraction <> 1# Then Exit For
            If Abs(current.Lambda - factor) > 8# * LoadResolution(current.Lambda) Then
                reference = history.Item(index): spacing = GeomMax(spacing, mSolver.SpatialFieldDistance(reference, accepted))
                Exit For
            End If
        Next index
    End If
    mProjection = projection: current.Coordinate = coordinate
    left.HistoryIndex = current.HistoryIndex - 1: left.Coordinate = CoordinateOf(previous)
    If Not mTrace Is Nothing Then mTrace.ReadPoint left.HistoryIndex, fraction, left.Lambda
    nominal = spacing: increment = 0.5 * nominal: minimum = increment * MIN_STEP_RATIO
    mUsesEnvelopeControl = False: mUsesYieldControl = False: mUsesFieldControl = True: mUsesTransverseControl = False
    mBranchTangentControl = False: mMaterialGrowthActive = False: Erase mMaterialGrowthProjection
    mHasCornerReference = False: mHasCornerProbe = False: Erase mCornerProbe
    mFailedProbeRevision = -1: mLastHybridRevision = -1: Erase mFailedProbeProjection: Erase mLastHybridProjection
    mLastHybridTarget = 0#: mFailedProbeTarget = 0#: mFailedProbeCode = rcNone
    mPostPeakControlName = vbNullString: mControlName = "пространственная секущая после остановки H"
    mEnvelopeDepartures = mEnvelopeDepartures + 1: mEnvelopeDepartureRevision = mSolver.AcceptedRevision
    mFieldControlChanges = mFieldControlChanges + 1: mLastFieldControlRevision = mSolver.AcceptedRevision
    TryLeaveEnvelopeControl = True
End Function

' Проверяет один материальный резерв после выхода из H либо исчерпания его проб на той же принятой ревизии.
' Использует только прошлую нагрузочно различимую секущую; новую меру и поле принимает общий путь CopyRecoveredControl, а не эта проба.
Private Function TryRecoverEnvelopeDeparture(ByVal history As Collection, ByRef current As TBranchPoint, _
        ByRef trial As TBranchPoint) As Boolean
    If mUsesEnvelopeControl Then Exit Function
    If mEnvelopeDepartureRevision <> mSolver.AcceptedRevision Then
        If mEnvelopeAttemptRevision <> mSolver.AcceptedRevision Or mEnvelopeAttemptsAtRevision < YIELD_RECOVERY_LIMIT Then Exit Function ' Неудачные H-пробы также завершают этот способ продолжения, не меняя принятую историю.
    End If
    If mDepartureRecoveryRevision = mSolver.AcceptedRevision Or mProbeCount >= PROBE_LIMIT Then Exit Function
    If history Is Nothing Or mTrace Is Nothing Then Exit Function
    If current.HistoryIndex < 2 Or current.HistoryIndex <> history.Count Or mTrace.Count <> history.Count Then Exit Function
    Dim previous() As Double, accepted() As Double, movement() As Double, projection() As Double
    previous = history(current.HistoryIndex - 1): accepted = history(current.HistoryIndex)
    ReDim movement(1 To UBound(accepted))
    Dim i As Long, index As Long, referenceIndex As Long, fraction As Double, factor As Double
    For i = 1 To mSolver.PhysicalDegreeCount: movement(i) = accepted(i) - previous(i): Next i
    Dim coordinate As Double, spacing As Double, reference() As Double, referenceCoordinate As Double
    If Not mSolver.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then Exit Function
    For index = current.HistoryIndex - 1 To 1 Step -1
        mTrace.ReadPoint index, fraction, factor
        If fraction <> 1# Then Exit For
        If Abs(current.Lambda - factor) > 8# * LoadResolution(current.Lambda) Then
            reference = history(index): spacing = GeomMax(spacing, mSolver.SpatialFieldDistance(reference, accepted)) ' Масштаб совпадает с выходом из H и не задаёт физический предел.
            Exit For
        End If
    Next index
    referenceIndex = current.HistoryIndex - 1
    For index = current.HistoryIndex - 1 To 1 Step -1
        mTrace.ReadPoint index, fraction, factor
        If fraction = 1# And current.Lambda - factor > 8# * LoadResolution(current.Lambda) Then
            reference = history(index): referenceCoordinate = 0#
            For i = 1 To mSolver.PhysicalDegreeCount: referenceCoordinate = referenceCoordinate + projection(i) * reference(i): Next i
            If coordinate > referenceCoordinate Then referenceIndex = index: Exit For
        End If
    Next index
    reference = history(referenceIndex): referenceCoordinate = 0#
    mTrace.ReadPoint referenceIndex, fraction, factor
    If fraction <> 1# Then Exit Function
    For i = 1 To mSolver.PhysicalDegreeCount: referenceCoordinate = referenceCoordinate + projection(i) * reference(i): Next i
    trial.Coordinate = coordinate + spacing
    If coordinate <= referenceCoordinate Or trial.Coordinate <= coordinate Then Exit Function
    Dim ratio As Double, seed() As Double, prediction As Variant, initialLambda As Double
    ratio = (trial.Coordinate - coordinate) / (coordinate - referenceCoordinate): seed = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = accepted(i) + ratio * (accepted(i) - reference(i)): Next i
    initialLambda = current.Lambda + ratio * (current.Lambda - factor): prediction = seed ' Служебные множители связей сохраняются из принятого поля.
    Dim originalProjection() As Double, recoveredProjection() As Double, solved As Boolean
    originalProjection = mProjection
    mDepartureRecoveryRevision = mSolver.AcceptedRevision: mDepartureRecoveryProbes = mDepartureRecoveryProbes + 1
    Dim errorNumber As Long, errorSource As String, errorText As String
    On Error GoTo RestoreAfterError
    mProjection = projection
    solved = Probe(trial, initialLambda, initialField:=prediction, allowHybridRecovery:=True, _
        acceptedLambda:=current.Lambda, robustDeparture:=True)
    mProjection = originalProjection
    If Not solved Then Exit Function
    If Not mSolver.CopyRecoveredControl(recoveredProjection) Then
        mSolver.DiscardStep
        mProbeDetails = "Материальный резерв после остановки H не предоставил новую подтверждённую меру управления."
        Exit Function
    End If
    TryRecoverEnvelopeDeparture = True ' Сталь, направление, диапазон lambda и принятие новой меры проверяет общий путь контроллера.
    Exit Function
RestoreAfterError:
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    mProjection = originalProjection
    Err.Raise errorNumber, errorSource, errorText
End Function

' Выбирает ближайшую принятую секущую с различимой нагрузкой, сохраняя все промежуточные состояния материала.
' Запас восьми разрешений нагрузки не позволяет первому полушагу заведомо затеряться в её погрешности.
Private Sub SelectInitialAnchor(ByVal history As Collection, ByRef current As TBranchPoint, ByRef left As TBranchPoint)
    If mTrace Is Nothing Then Exit Sub
    If mTrace.Count <> history.Count Then Exit Sub
    If current.Lambda - left.Lambda > 8# * LoadResolution(current.Lambda) Then Exit Sub
    Dim index As Long, fraction As Double, factor As Double
    For index = current.HistoryIndex - 1 To 1 Step -1
        mTrace.ReadPoint index, fraction, factor
        If fraction <> 1# Or factor >= current.Lambda Then Exit Sub
        If current.Lambda - factor > 8# * LoadResolution(current.Lambda) Then
            left.HistoryIndex = index: left.Lambda = factor
            Exit Sub
        End If
    Next index
End Sub

' Готовит начальное приближение из собственных полей либо касательной принятого материального перехода без изменения истории.
' Секущая использует механические координаты; касательная полной системы прогнозирует также вспомогательные множители связей.
Private Function PrepareSecantSeed(ByVal history As Collection, ByRef left As TBranchPoint, ByRef current As TBranchPoint, _
        ByVal target As Double, ByRef seed As Variant, ByRef factor As Double, Optional ByVal predictorTrace As CCoupledLoadTrace = Nothing) As Boolean
    If Not mUsesEnvelopeControl And mEnvelopeDepartureRevision = mSolver.AcceptedRevision Then Exit Function
    If left.HistoryIndex < 1 Or left.HistoryIndex >= current.HistoryIndex Or current.HistoryIndex <> history.Count Then Exit Function
    Dim previous() As Double, accepted() As Double, field() As Double, oldControl As Double, control As Double, ratio As Double, i As Long
    previous = history(left.HistoryIndex): accepted = history(current.HistoryIndex)
    If mHasCornerReference And mCornerReferenceRevision = mSolver.AcceptedRevision Then
        If mHasCornerProbe And target = mCornerProbeCoordinate Then
            seed = mCornerProbe: factor = mCornerProbeLambda
            mCornerProbeSeeds = mCornerProbeSeeds + 1
            PrepareSecantSeed = True: Exit Function
        End If
        control = CoordinateOf(accepted)
        Dim rate As Double
        For i = 1 To mSolver.PhysicalDegreeCount: rate = rate + mProjection(i) * mCornerDirection(i): Next i
        If target <= control Or rate <= 0# Then Exit Function
        ratio = (target - control) / rate: field = accepted
        For i = 1 To UBound(field): field(i) = accepted(i) + ratio * mCornerDirection(i): Next i ' Вспомогательные множители прогнозируются из полной касательной, не из материальных деформаций.
        seed = field: factor = current.Lambda + ratio * mCornerLambdaDirection
        PrepareSecantSeed = True
        Exit Function
    End If
    Dim selectedTrace As CCoupledLoadTrace: Set selectedTrace = predictorTrace
    If selectedTrace Is Nothing Then Set selectedTrace = mTrace
    Dim previousLambda As Double: previousLambda = left.Lambda
    If Not selectedTrace Is Nothing Then
        Dim fraction As Double
        previous = history(current.HistoryIndex - 1)
        selectedTrace.ReadPoint current.HistoryIndex - 1, fraction, previousLambda
        If fraction <> 1# Then Exit Function
    End If
    If mUsesEnvelopeControl Then
        If Not mSolver.TryAcceptedEnvelopeSegment(previous, oldControl, control) Then Exit Function
    Else
        oldControl = CoordinateOf(previous): control = CoordinateOf(accepted)
    End If
    If control <= oldControl Or target <= control Then Exit Function
    ratio = (target - control) / (control - oldControl)
    field = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: field(i) = accepted(i) + ratio * (accepted(i) - previous(i)): Next i
    seed = field: factor = current.Lambda + ratio * (current.Lambda - previousLambda)
    PrepareSecantSeed = True
End Function

' Продолжает от уже проверенного правого плеча. Последняя уточняющая
' проходка завершилась именно в этом материальном состоянии.
Private Sub ResumeAfterLocalPeak(ByRef history As Collection, ByRef current As TBranchPoint, ByRef right As TBranchPoint)
    mLocalPeak = mPeak.Lambda: mLocalPeakAccuracy = mPeakAccuracy
    mHasPeak = False: mFollowingLocalPeak = True
    Set history = mLimitFields
    Set mCurrentFields = history
    Set mTrace = mLimitTrace: Set mLimitTrace = Nothing
    mCurrentResultants = mLimitResultants
    Set mLimitFields = Nothing: Erase mLimitResultants
    current = right
    mReusedRightPoints = mReusedRightPoints + current.HistoryIndex - mPeak.HistoryIndex
End Sub

' Возвращает независимую числовую копию принятого пути; незавершённый путь не становится предельным НДС.
Friend Property Get AcceptedFields() As Collection
    Dim source As Collection
    If mHasPeak Or mSteelPoint > 0 Then Set source = mLimitFields Else Set source = mCurrentFields
    If source Is Nothing Then Exit Property
    Dim result As New Collection, i As Long, field() As Double
    For i = 1 To source.Count: field = source(i): result.Add field: Next i
    Set AcceptedFields = result
End Property

' Один раз выбирает другую меру, если после подтверждённого спада
' численная проба потребовала шаг ниже минимального. Сохраняет принятое
' поле, историю и одностороннюю материальную касательную.
Private Function TryChangePostPeakControl(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef nominal As Double, ByRef increment As Double, ByRef minimum As Double, _
        ByVal stalledOnCoordinate As Boolean) As Boolean
    If Len(mPostPeakControlName) > 0 Or Not mFollowingLocalPeak Or Not mUsesYieldControl Then Exit Function
    If history.Count < 3 Or current.HistoryIndex <> history.Count Or left.HistoryIndex <> history.Count - 1 Then Exit Function
    Dim code As EResultCode: code = mSolver.ResultMeta.ResultCode
    If Not stalledOnCoordinate And code <> rcNumericalFailure And code <> rcSingularTangent Then Exit Function

    Dim oldest() As Double, previous() As Double, currentField() As Double, accepted() As Double
    oldest = history(history.Count - 2): previous = history(left.HistoryIndex): currentField = history(current.HistoryIndex)
    mSolver.CopyAccepted accepted
    Dim alternate() As Double, currentCoordinate As Double, currentSpacing As Double, name As String
    If Not mSolver.SelectDeformationControl(previous, currentField, alternate, currentCoordinate, currentSpacing, name) Then Exit Function
    Dim i As Long, oldestCoordinate As Double, previousCoordinate As Double
    For i = 1 To UBound(alternate)
        If currentField(i) <> accepted(i) Then Exit Function
        previousCoordinate = previousCoordinate + alternate(i) * previous(i)
        oldestCoordinate = oldestCoordinate + alternate(i) * oldest(i)
    Next i
    Dim spacing As Double: spacing = GeomMin(currentCoordinate - previousCoordinate, previousCoordinate - oldestCoordinate)
    If spacing <= 0# Or currentCoordinate + 0.5 * spacing <= currentCoordinate Then Exit Function

    mProjection = alternate: mPostPeakControlName = name
    mUsesTransverseControl = False: mUsesFieldControl = False
    left.Coordinate = previousCoordinate: current.Coordinate = currentCoordinate
    nominal = spacing: increment = 0.5 * spacing: minimum = increment * MIN_STEP_RATIO
    TryChangePostPeakControl = True
End Function

' Меняет параметризацию после численного отказа до локального максимума или на неподтверждённом спаде нагрузки.
' Последняя отвергнутая поправка не принимается; новая цель решается из полной неизменной истории.
' На одну ревизию принятой истории разрешена одна смена; все пробы входят в общий бюджет.
Private Function TryChangeStalledFieldControl(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef nominal As Double, ByRef increment As Double, ByRef minimum As Double) As Boolean
    If mLastFieldControlRevision = mSolver.AcceptedRevision Then Exit Function
    If mFollowingLocalPeak Then Exit Function
    If Not mSolver.UsesSpatialHistory Or history.Count < 2 Or current.HistoryIndex <> history.Count Then Exit Function
    Dim code As EResultCode: code = mSolver.ResultMeta.ResultCode
    If code <> rcNumericalFailure And code <> rcSingularTangent Then Exit Function
    Dim previous() As Double, accepted() As Double, alternate() As Double, coordinate As Double, spacing As Double, name As String
    previous = history(history.Count - 1): accepted = history(history.Count)
    If Not mSolver.SelectRejectedDeformationControl(previous, accepted, alternate, coordinate, spacing, name) Then Exit Function
    If spacing <= 0# Or coordinate + 0.5 * spacing <= coordinate Then Exit Function
    mProjection = alternate: mControlName = name
    mUsesTransverseControl = False: mUsesFieldControl = True: mUsesYieldControl = False
    mBranchTangentControl = False
    mFieldControlChanges = mFieldControlChanges + 1
    mLastFieldControlRevision = mSolver.AcceptedRevision
    left.HistoryIndex = history.Count - 1: left.Coordinate = coordinate - spacing
    If Not mTrace Is Nothing Then
        Dim fraction As Double
        mTrace.ReadPoint left.HistoryIndex, fraction, left.Lambda
    End If
    current.Coordinate = coordinate
    nominal = spacing: increment = 0.5 * spacing: minimum = increment * MIN_STEP_RATIO
    TryChangeStalledFieldControl = True
End Function

' Запрашивает один повтор конкретной цели для неизменной истории и проекции после исчерпания деления.
' Если сошедшийся меньший шаг не различает нагрузку, повторяет только ранее численно отвергнутую большую цель.
' Условия равновесия, направления и материала проверяются обычной приёмкой; неуточнённый предел стали не обходится.
Private Function PrepareHybridRecovery(ByVal lastProbeFailed As Boolean, ByVal lastLoadUnresolved As Boolean, _
        ByVal acceptedCoordinate As Double, ByRef increment As Double) As Boolean
    If Not mSolver.SupportsHybridRecovery Then Exit Function
    Dim code As EResultCode, target As Double, sameProjection As Boolean, i As Long
    If lastProbeFailed Then
        code = mSolver.ResultMeta.ResultCode
        target = acceptedCoordinate + 2# * increment
    ElseIf lastLoadUnresolved Then
        If mFailedProbeRevision <> mSolver.AcceptedRevision Then Exit Function
        If mFailedProbeTarget <= acceptedCoordinate + 2# * increment Then Exit Function
        For i = 1 To UBound(mProjection)
            If mFailedProbeProjection(i) <> mProjection(i) Then Exit Function
        Next i
        code = mFailedProbeCode: target = mFailedProbeTarget
    Else
        Exit Function
    End If
    If code <> rcNumericalFailure And code <> rcSingularTangent Then Exit Function
    If target <= acceptedCoordinate Then Exit Function
    If mLastHybridRevision = mSolver.AcceptedRevision Then
        sameProjection = True
        For i = 1 To UBound(mProjection)
            If mLastHybridProjection(i) <> mProjection(i) Then sameProjection = False: Exit For
        Next i
        If sameProjection And target = mLastHybridTarget Then Exit Function ' Другая правая часть задаёт другую задачу, даже при той же истории и проекции.
    End If
    mLastHybridRevision = mSolver.AcceptedRevision: mLastHybridProjection = mProjection
    mLastHybridTarget = target
    mHybridRecoveryProbes = mHybridRecoveryProbes + 1
    If lastLoadUnresolved Then mUnresolvedLoadRetries = mUnresolvedLoadRetries + 1
    increment = target - acceptedCoordinate ' Возвращается конкретная отвергнутая цель, новая экстраполяция нагрузки не вводится.
    PrepareHybridRecovery = True
End Function

' Сравнивает приращения на двух соседних шагах одной нисходящей ветви.
' Не меняет ни поля, ни историю. Рост шага возможен при устойчивых
' наклонах нагрузки и механического поля.
Private Function FitsPostPeakSecant(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef trial As TBranchPoint) As Boolean
    If left.HistoryIndex <> current.HistoryIndex - 1 Then Exit Function
    Dim spacing As Double: spacing = current.Coordinate - left.Coordinate
    If spacing <= 0# Or trial.Lambda <= 0# Then Exit Function
    Dim ratio As Double: ratio = (trial.Coordinate - current.Coordinate) / spacing
    Dim loadChange As Double: loadChange = ratio * (current.Lambda - left.Lambda)
    If loadChange >= -LoadResolution(current.Lambda) Then Exit Function
    If Abs(trial.Lambda - current.Lambda - loadChange) > Abs(loadChange) * POST_PEAK_SECANT_TOLERANCE Then Exit Function

    Dim previous() As Double, accepted() As Double, field() As Double, resultants() As Double
    previous = history(left.HistoryIndex): accepted = history(current.HistoryIndex)
    mSolver.CopyTrial field, resultants
    FitsPostPeakSecant = FitsMechanicalSecant(previous, accepted, field, mSolver.PhysicalDegreeCount, ratio)
End Function

' Разрешает адаптацию только для той же точной пространственной меры, установленной проверенным CDP-переходом.
' Смена меры автоматически прекращает этот режим; принятые состояния материала здесь не кэшируются.
Private Function IsMaterialGrowthControl() As Boolean
    If mUsesEnvelopeControl Then IsMaterialGrowthControl = True: Exit Function
    If Not mMaterialGrowthActive Or Not mUsesFieldControl Then Exit Function
    Dim i As Long
    For i = 1 To UBound(mProjection)
        If mProjection(i) <> mMaterialGrowthProjection(i) Then Exit Function
    Next i
    IsMaterialGrowthControl = True
End Function

' Проверяет прогноз принятой секущей по тензорной метрике трёх станций и множителю нагрузки.
' Неразличимое изменение lambda не препятствует росту шага; это проверка прогноза, а не ослабление силового равновесия.
Private Function FitsMaterialGrowthSecant(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef trial As TBranchPoint) As Boolean
    If left.HistoryIndex <> current.HistoryIndex - 1 Or trial.Lambda <= 0# Then Exit Function
    Dim spacing As Double: spacing = current.Coordinate - left.Coordinate
    If spacing <= 0# Then Exit Function
    Dim ratio As Double: ratio = (trial.Coordinate - current.Coordinate) / spacing
    If ratio <= 0# Then Exit Function
    Dim predictedLoad As Double: predictedLoad = ratio * (current.Lambda - left.Lambda)
    Dim growthResolutionTest As Double: growthResolutionTest = LoadResolution(current.Lambda)
    If mFollowingLocalPeak Then growthResolutionTest = GeomMax(growthResolutionTest, mLambdaTolerance / 16#) ' Численный масштаб прогноза шага; не допуск приёмки равновесия или максимума.
    If Abs(trial.Lambda - current.Lambda - predictedLoad) > _
            GeomMax(Abs(predictedLoad) * POST_PEAK_SECANT_TOLERANCE, growthResolutionTest) Then Exit Function
    Dim previous() As Double, accepted() As Double, field() As Double, predicted() As Double, resultants() As Double, i As Long
    previous = history(left.HistoryIndex): accepted = history(current.HistoryIndex)
    mSolver.CopyTrial field, resultants
    predicted = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: predicted(i) = accepted(i) + ratio * (accepted(i) - previous(i)): Next i
    Dim expectedLength As Double: expectedLength = mSolver.SpatialFieldDistance(accepted, predicted)
    If expectedLength <= 0# Then Exit Function
    FitsMaterialGrowthSecant = (mSolver.SpatialFieldDistance(field, predicted) <= POST_PEAK_SECANT_TOLERANCE * expectedLength)
End Function

' Проверяет относительное отклонение механических приращений от секущей.
' Множители кинематических связей лежат за границей mechanicalCount.
Friend Function FitsMechanicalSecant(ByRef previous() As Double, ByRef accepted() As Double, _
        ByRef field() As Double, ByVal mechanicalCount As Long, ByVal ratio As Double) As Boolean
    Dim i As Long, delta As Double, expectedDelta As Double, stepScale As Double, deviation As Double
    For i = 1 To mechanicalCount
        delta = field(i) - accepted(i)
        expectedDelta = ratio * (accepted(i) - previous(i))
        stepScale = GeomMax(stepScale, Abs(expectedDelta))
        deviation = GeomMax(deviation, Abs(delta - expectedDelta))
    Next i
    FitsMechanicalSecant = (stepScale > 0# And deviation <= stepScale * POST_PEAK_SECANT_TOLERANCE)
End Function

' Заменяет цель фактической мерой поля только при прямом движении ветви.
' Равновесный повтор на месте допустим как НДС и не образует нового шага.
Private Function UpdateTrialCoordinate(ByRef point As TBranchPoint, ByVal previousCoordinate As Double) As Boolean
    Dim field() As Double, resultants() As Double, coordinate As Double
    mSolver.CopyTrial field, resultants
    If mUsesEnvelopeControl Then
        If Not mSolver.TryReadEnvelopeCoordinate(False, coordinate) Then Exit Function
    Else
        coordinate = CoordinateOf(field)
    End If
    If coordinate <= previousCoordinate Then
        mSolver.DiscardStep
        mProbeDetails = "Проба не продвинулась по мере деформации с " & CStr(previousCoordinate) & " до " & CStr(coordinate) & "."
        Exit Function
    End If
    point.Coordinate = coordinate
    UpdateTrialCoordinate = True
End Function

' Считает восстановление после проверенного нисходящего плеча.
' Возвращённая нагрузка должна превышать максимум с учётом его погрешности.
Private Sub CheckRecovery(ByVal lambda As Double)
    If lambda <= mLocalPeak + mLocalPeakAccuracy Then Exit Sub
    mFollowingLocalPeak = False
    mRecoveredPeaks = mRecoveredPeaks + 1
End Sub

' Возвращает максимальный подтверждённый множитель на пройденной
' допустимой ветви, даже если общий предел остался неопределённым.
Friend Property Get BestAdmissibleLambda() As Double
    BestAdmissibleLambda = mBestLambda
End Property

' Возвращает историю выбранной ветви, исключая отменённые пробы плеч.
Friend Property Get SavedTrace() As CCoupledLoadTrace
    If mHasPeak Or mSteelPoint > 0 Then Set SavedTrace = mLimitTrace Else Set SavedTrace = mTrace
End Property

' Передаёт подтверждённый максимум: для железобетона только после физического завершения допустимой ветви.
Friend Function CopyPeak(ByRef lambda As Double, ByRef accuracy As Double, ByRef designVerified As Boolean) As Boolean
    If Not mHasPeak Then Exit Function
    lambda = mPeak.Lambda: accuracy = mPeakAccuracy: designVerified = (mDesignHistoryIndex > 0)
    CopyPeak = True
End Function

' Передаёт только поля до подтверждённой предельной точки и её
' фактические интегралы. Поздние точки нисходящего плеча не включаются.
Friend Function CopyLimitState(ByRef fields As Collection, ByRef resultants() As Double) As Boolean
    If mLimitFields Is Nothing Then Exit Function
    Set fields = mLimitFields
    resultants = mLimitResultants
    CopyLimitState = True
End Function

' Передаёт арматурное ограничение, определяющее достигнутую нагрузку.
Friend Function CopySteelLimit(ByRef low As Double, ByRef high As Double, ByRef criticalPoint As Long) As Boolean
    If mSteelPoint = 0 Then Exit Function
    low = mSteelLow: high = mSteelHigh: criticalPoint = mSteelPoint
    CopySteelLimit = True
End Function

' Возвращает точку установленного арматурного события для пояснения в отчёте.
Friend Property Get SteelEventPoint() As Long
    SteelEventPoint = mSteelEventPoint
End Property

' Возвращает независимый статус контролируемого участка.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Разрешает ограниченное число восстановлений на неизменной истории ASD; принятая новая история получает собственный запас попыток.
' Для CDP действует общий предел участка. Каждое решение остаётся внутри общего бюджета деформационных проб.
Private Function MayRecoverMaterialBoundary() As Boolean
    If Not mSolver.HasPrincipalBoundaries Then
        MayRecoverMaterialBoundary = (Not mFollowingLocalPeak And mYieldRecoveryProbes < YIELD_RECOVERY_LIMIT)
    ElseIf mMaterialRecoveryRevision <> mSolver.AcceptedRevision Then
        MayRecoverMaterialBoundary = True
    Else
        MayRecoverMaterialBoundary = (mMaterialRecoveryAtRevision < YIELD_RECOVERY_LIMIT)
    End If
End Function

' Решает пробу при заданной мере. Проверки продвижения и касательной
' выполняются после выбора действующей проекции.
Private Function Probe(ByRef point As TBranchPoint, ByVal initialLambda As Double, _
        Optional ByVal recoverYieldControl As Boolean = False, Optional ByVal keepYieldControl As Boolean = False, _
        Optional ByVal arcFields As Collection = Nothing, Optional ByRef initialField As Variant, _
        Optional ByVal allowHybridRecovery As Boolean = False, Optional ByVal acceptedLambda As Double = 0#, _
        Optional ByVal robustDeparture As Boolean = False, Optional ByVal accuracyScale As Double = 1#) As Boolean
    If accuracyScale <= 0# Or accuracyScale > 1# Then Err.Raise 5, "CCoupledDeformationContinuation", "Масштаб точности должен быть положительным и не превышать единицу."
    If mUsesEnvelopeControl And accuracyScale <> 1# Then Err.Raise 5, "CCoupledDeformationContinuation", "Для меры H уточнение задаётся отдельным алгоритмом."
    If mProbeCount >= PROBE_LIMIT Then Exit Function
    If mUsesEnvelopeControl And mProbeCount + 2 > PROBE_LIMIT Then
        mProbeDetails = "Оставшегося бюджета недостаточно для полного материального шага и обязательной половинной проверки при повороте."
        Exit Function
    End If
    mProbeCount = mProbeCount + 1
    NotifyProgress
    point.Uncertainty = 0#
    If mUsesEnvelopeControl Then
        Probe = ProbeEnvelope(point, acceptedLambda, initialLambda, initialField)
        Exit Function
    End If
    Dim solved As Boolean, recoverSameControl As Boolean, recoveryReference As Double
    Dim postPeakRecovery As Boolean: postPeakRecovery = (mLocalPeak > 0# And Not mSolver.HasPrincipalBoundaries)
    Dim guideProbes As Long, guideReturned As Boolean, allowGuide As Boolean
    allowGuide = mFollowingLocalPeak And Not robustDeparture And (mFieldControlChanges > 0 Or mBranchTangentControl) And _
        IsArray(initialField) And mProbeCount + 2 <= PROBE_LIMIT ' Возврат после локального максимума не расходует бюджет ещё не прослеженной восходящей ветви.
    If robustDeparture Then recoveryReference = acceptedLambda
    recoverSameControl = keepYieldControl And mBranchTangentControl And mFixedYieldRecoveryProbes < FIXED_YIELD_RECOVERY_LIMIT
    solved = mSolver.TryDeformationStep(mOffset, mDirection, mProjection, point.Coordinate, initialLambda, _
        point.Lambda, mAbsTol * accuracyScale, mRelTol * accuracyScale, mMaxIterations, recoverYieldControl, (mBranchTangentControl Or robustDeparture Or postPeakRecovery), recoverSameControl, _
        initialField:=initialField, allowSmallCorrection:=(mFieldControlChanges > 0), allowHybridRecovery:=(allowHybridRecovery Or postPeakRecovery), _
        conditionHybridRecovery:=(robustDeparture Or postPeakRecovery), referenceLambda:=recoveryReference, _
        allowRefinedRestart:=(IsMaterialGrowthControl() Or robustDeparture Or postPeakRecovery), allowGuideRecovery:=allowGuide, _
        guideReferenceLambda:=acceptedLambda, guideProbes:=guideProbes, guideReturnSucceeded:=guideReturned)
    mProbeCount = mProbeCount + guideProbes: mGuideRecoveryProbes = mGuideRecoveryProbes + guideProbes
    If guideReturned Then mGuideRecoverySolutions = mGuideRecoverySolutions + 1
    If Not robustDeparture Then ' Резерв после H имеет отдельные счётчики при неизменном общем бюджете проб.
        If mSolver.SameControlRecoveryTried Then
            mFixedYieldRecoveryProbes = mFixedYieldRecoveryProbes + 1
        ElseIf mSolver.YieldRecoveryTried Then
            mYieldRecoveryProbes = mYieldRecoveryProbes + 1
            If mSolver.HasPrincipalBoundaries Then
                If mMaterialRecoveryRevision <> mSolver.AcceptedRevision Then
                    mMaterialRecoveryRevision = mSolver.AcceptedRevision: mMaterialRecoveryAtRevision = 0
                End If
                mMaterialRecoveryAtRevision = mMaterialRecoveryAtRevision + 1
            End If
        End If
    End If
    If Not solved And Not arcFields Is Nothing And mProbeCount + 2 <= PROBE_LIMIT Then
        Dim extraProbes As Long
        solved = mSolver.TryRecoverDeformationByArc(arcFields, mTrace, initialLambda, mAbsTol * accuracyScale, mRelTol * accuracyScale, _
            mMaxIterations, point.Lambda, extraProbes)
        If extraProbes > 0 Then
            mArcRecoveryTried = True: mArcRecoveryProbes = extraProbes
            mProbeCount = mProbeCount + extraProbes
        End If
    End If
    If Not solved Then
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
        mFailedProbeRevision = mSolver.AcceptedRevision: mFailedProbeProjection = mProjection
        mFailedProbeTarget = point.Coordinate: mFailedProbeCode = mSolver.ResultMeta.ResultCode
        Exit Function
    End If
    Probe = True
End Function

' Решает следующий уровень H от пары фактически принятых состояний, в том числе после повторного прохождения плеч.
' Прогноз использует только прошлые поля; обратная пространственная секущая проверяется половинным H-шагом в решателе.
Private Function ProbeEnvelope(ByRef point As TBranchPoint, ByVal acceptedLambda As Double, ByVal initialLambda As Double, _
        Optional ByRef initialField As Variant) As Boolean
    Dim previous() As Double, accepted() As Double, seed() As Double, projection() As Double
    Dim before As Double, after As Double, ratio As Double, degree As Long, probes As Long
    Dim baseline As Double
    If Not mSolver.TryEnvelopeControlState(mOffset, mDirection, acceptedLambda, mAbsTol, mRelTol, baseline, mEnvelopeResolution) Then Exit Function
    If Not mSolver.TryAcceptedEnvelopeSegment(previous, before, after) Then
        mProbeDetails = "Не сохранена последовательная пара принятых состояний для материального продолжения."
        Exit Function
    End If
    mSolver.CopyAccepted accepted
    If IsMissing(initialField) Then
        seed = accepted
        If after > before And point.Coordinate > after Then ratio = (point.Coordinate - after) / (after - before)
        For degree = 1 To mSolver.PhysicalDegreeCount: seed(degree) = accepted(degree) + ratio * (accepted(degree) - previous(degree)): Next degree
    Else
        seed = initialField
    End If
    ProbeEnvelope = mSolver.TryEnvelopeStep(mOffset, mDirection, previous, acceptedLambda, seed, initialLambda, _
        point.Coordinate, point.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True)
    mProbeCount = mProbeCount + GeomMax(0, probes - 1): mEnvelopeProbes = mEnvelopeProbes + probes
    If ProbeEnvelope Then
        mProjection = projection
    Else
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
    End If
End Function

' Передаёт наблюдателю только независимые копии принятого пути, без пробных состояний и ввода-вывода ядра.
' Полный снимок сохраняет смысл после замены хвоста при уточнении локального максимума.
Private Sub NotifyProgress()
    modTestCoupledElasticState.TestCaptureStockContinuation Me, Details()
    If mObserver Is Nothing Then Exit Sub
    Dim fields As Collection, trace As CCoupledLoadTrace, sourceTrace As CCoupledLoadTrace, meta As CResultMeta
    Set fields = AcceptedFields: Set sourceTrace = SavedTrace
    If fields Is Nothing Or sourceTrace Is Nothing Then Exit Sub
    Set trace = sourceTrace.CopyPrefix(sourceTrace.Count)
    Set meta = mMeta.Clone: meta.DiagnosticDetails = Details()
    mObserver.Report fields, trace, meta
End Sub

' Принимает новую координату подтверждённого материального перехода, даже когда изменение lambda меньше порога распознавания пика.
' До наблюдаемого спада не допускает уменьшения нагрузки сверх этого порога; проверки плеч и стали сохраняют выбранную меру.
Private Function TryAdoptRecoveredControl(ByVal history As Collection, ByRef current As TBranchPoint, _
        ByRef trial As TBranchPoint, ByRef projection() As Double, ByRef nominal As Double, _
        ByRef increment As Double, ByRef minimum As Double) As Boolean
    If trial.Lambda <= 0# Or trial.Lambda > mMaximumLambda Then Exit Function
    If Not mSolver.HasPrincipalBoundaries Then
        If trial.Lambda <= current.Lambda + LoadResolution(current.Lambda) Then Exit Function
    ElseIf Not mFollowingLocalPeak And Not mUnverifiedDecline Then
        If trial.Lambda < current.Lambda - LoadResolution(current.Lambda) Then Exit Function
    End If
    If BeyondSteelLimit() Then Exit Function
    Dim previousField() As Double, acceptedField() As Double, field() As Double, resultants() As Double
    previousField = history(current.HistoryIndex - 1)
    acceptedField = history(current.HistoryIndex)
    mSolver.CopyTrial field, resultants
    If mSolver.HasPrincipalBoundaries Then
        If Not SpatialTrialAdvances(history, current) Then Exit Function
    End If
    Dim i As Long, previousCoordinate As Double, oldCoordinate As Double, newCoordinate As Double
    For i = 1 To UBound(projection)
        previousCoordinate = previousCoordinate + projection(i) * previousField(i)
        oldCoordinate = oldCoordinate + projection(i) * acceptedField(i)
        newCoordinate = newCoordinate + projection(i) * field(i)
    Next i
    Dim difference As Double: difference = Abs(newCoordinate - oldCoordinate)
    If difference <= 0.000000000001 Then Exit Function
    If (oldCoordinate - previousCoordinate) * (newCoordinate - oldCoordinate) <= 0# Then Exit Function
    nominal = difference
    If newCoordinate < oldCoordinate Then
        For i = 1 To UBound(projection): projection(i) = -projection(i): Next i
        oldCoordinate = -oldCoordinate: newCoordinate = -newCoordinate
    End If
    mProjection = projection: mControlName = "мера деформации у перехода бетона"
    mUsesYieldControl = True
    mBranchTangentControl = mSolver.RecoveredControlRequiresBranchTangent
    mUsesTransverseControl = False: mUsesFieldControl = False
    current.Coordinate = oldCoordinate: trial.Coordinate = newCoordinate
    increment = 0.5 * nominal: minimum = increment * MIN_STEP_RATIO
    TryAdoptRecoveredControl = True
End Function

' Проверяет направление полного пространственного пробного поля, в том числе после смены управляющей координаты.
' Отказ требует дробления шага либо другой параметризации и не устанавливает физический предел сечения.
Private Function SpatialTrialAdvances(ByVal history As Collection, ByRef current As TBranchPoint) As Boolean
    Dim previous() As Double, accepted() As Double, trial() As Double, resultants() As Double
    ReadSpatialReference history, current.HistoryIndex, previous
    accepted = history(current.HistoryIndex)
    mSolver.CopyTrial trial, resultants
    Dim signedAdvance As Double, referenceIncrement As Double
    SpatialTrialAdvances = mSolver.SpatialIncrementAdvances(previous, accepted, trial, signedAdvance, referenceIncrement)
    If SpatialTrialAdvances Then Exit Function
    mSolver.PreserveRejectedSpatialTrial previous
    mSpatialReversalsRejected = mSpatialReversalsRejected + 1
    mProbeDetails = "Пробное равновесие не продолжает направление предыдущего пространственного приращения." & _
        " Ход по секущей: " & CStr(signedAdvance) & "; предыдущее приращение: " & CStr(referenceIncrement) & "."
End Function

' Продвигает принятую материальную историю и фиксирует индекс её поля.
Private Sub AcceptPoint(ByVal history As Collection, ByRef point As TBranchPoint)
    Dim field() As Double
    mHasCornerReference = False: mHasCornerProbe = False: Erase mCornerProbe
    mSolver.CopyTrial field, mCurrentResultants
    Dim before As Double, after As Double
    If mSolver.TryReadEnvelopeCoordinate(True, before) Then
        If mSolver.TryReadEnvelopeCoordinate(False, after) Then
            mEnvelopePrevious = before: mEnvelopeCurrent = after: mEnvelopeHistoryIndex = history.Count + 1
            If after > before Then mEnvelopeLastIncrement = after - before
        End If
    End If
    mSolver.AcceptStep
    history.Add field
    If Not mTrace Is Nothing Then mTrace.Add 1#, point.Lambda
    point.HistoryIndex = history.Count
    mBestBeforeCurrent = mBestLambda
    mBestLambda = GeomMax(mBestLambda, point.Lambda)
End Sub

' Выделяет снижение нагрузки без дальнейшего развития огибающих; это повод сменить параметризацию, не признак разрушения.
Private Function IsEnvelopeUnloading(ByRef current As TBranchPoint, ByRef trial As TBranchPoint) As Boolean
    If trial.Lambda >= current.Lambda - LoadResolution(current.Lambda) Then Exit Function
    Dim accepted As Double, candidate As Double
    If Not mSolver.TryReadEnvelopeCoordinate(True, accepted) Then Exit Function
    If accepted <= 0# Then Exit Function
    If Not mSolver.TryReadEnvelopeCoordinate(False, candidate) Then Exit Function
    IsEnvelopeUnloading = (candidate <= accepted)
End Function

' После исчерпания обычного шага либо на разгрузке ищет растущую материальную координату от той же истории.
' Принимает только новый кандидат; сталь, граница поиска и точное сочетание проверяются общим циклом до записи истории.
Private Function TryEnvelopeRecovery(ByVal history As Collection, ByRef left As TBranchPoint, ByRef current As TBranchPoint, _
        ByRef trial As TBranchPoint, ByRef nominal As Double, ByRef increment As Double, ByRef minimum As Double) As Boolean
    If mEnvelopeDepartureRevision = mSolver.AcceptedRevision Then Exit Function
    If mReinforcement Is Nothing Then Exit Function
    If Not mReinforcement.HasSteel Or mSolver.HasPrincipalBoundaries Or mProbeCount + 2 > PROBE_LIMIT Then Exit Function
    Dim baseline As Double, resolution As Double
    If Not mSolver.TryEnvelopeControlState(mOffset, mDirection, current.Lambda, mAbsTol, mRelTol, baseline, resolution) Then Exit Function
    If baseline <= 0# Then Exit Function
    If mEnvelopeAttemptRevision <> mSolver.AcceptedRevision Then
        mEnvelopeAttemptRevision = mSolver.AcceptedRevision: mEnvelopeAttemptsAtRevision = 0
    End If
    Dim desired As Double: desired = GeomMax(mEnvelopeLastIncrement, GeomMax(4# * resolution, ENVELOPE_INITIAL_STEP_RATIO * baseline))
    Dim previous() As Double, accepted() As Double, seed() As Double, projection() As Double, field() As Double, actual() As Double
    ReadSpatialReference history, current.HistoryIndex, previous
    accepted = history(current.HistoryIndex)
    Dim last() As Double, pairBefore As Double, pairAfter As Double, hasPair As Boolean
    hasPair = mSolver.TryAcceptedEnvelopeSegment(last, pairBefore, pairAfter)
    If hasPair Then
        If pairAfter = baseline And pairAfter > pairBefore Then desired = GeomMax(desired, pairAfter - pairBefore)
    End If
    desired = GeomMin(desired, GeomMax(4# * resolution, ENVELOPE_MAXIMUM_STEP_RATIO * baseline))
    Dim ratio As Double, factor As Double, probes As Long, degree As Long, previousFactor As Double, fraction As Double
    previousFactor = left.Lambda
    If Not mTrace Is Nothing Then mTrace.ReadPoint current.HistoryIndex - 1, fraction, previousFactor
    Do While mEnvelopeAttemptsAtRevision < YIELD_RECOVERY_LIMIT And desired > 2# * resolution And mProbeCount + 2 <= PROBE_LIMIT
        mEnvelopeAttemptsAtRevision = mEnvelopeAttemptsAtRevision + 1
        ratio = 0#
        If hasPair Then
            If pairAfter = baseline And pairAfter > pairBefore Then ratio = GeomMin(1024#, desired / (pairAfter - pairBefore))
        End If
        seed = accepted
        last = history(current.HistoryIndex - 1)
        For degree = 1 To mSolver.PhysicalDegreeCount: seed(degree) = accepted(degree) + ratio * (accepted(degree) - last(degree)): Next degree
        factor = current.Lambda + ratio * (current.Lambda - previousFactor)
        mProbeCount = mProbeCount + 1: NotifyProgress
        trial.Uncertainty = 0#
        If mSolver.TryEnvelopeStep(mOffset, mDirection, previous, current.Lambda, seed, factor, baseline + desired, _
                trial.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True) Then
            mProbeCount = mProbeCount + GeomMax(0, probes - 1): mEnvelopeProbes = mEnvelopeProbes + probes
            mSolver.CopyTrial field, actual
            Dim oldCoordinate As Double, newCoordinate As Double, previousCoordinate As Double
            oldCoordinate = baseline
            If Not mSolver.TryReadEnvelopeCoordinate(False, newCoordinate) Then mSolver.DiscardStep: Exit Function
            If Not mSolver.TryAcceptedEnvelopeSegment(last, previousCoordinate, oldCoordinate) Then mSolver.DiscardStep: Exit Function
            If newCoordinate <= oldCoordinate Then mSolver.DiscardStep: Exit Function
            mProjection = projection: current.Coordinate = oldCoordinate: trial.Coordinate = newCoordinate
            left.HistoryIndex = current.HistoryIndex - 1: left.Coordinate = previousCoordinate
            If Not mTrace Is Nothing Then mTrace.ReadPoint left.HistoryIndex, fraction, left.Lambda
            nominal = newCoordinate - oldCoordinate: increment = 0.5 * nominal: minimum = 2# * resolution
            mEnvelopeResolution = resolution
            mUsesEnvelopeControl = True
            mUsesYieldControl = False: mUsesFieldControl = True: mUsesTransverseControl = False: mBranchTangentControl = False
            mHasCornerReference = False: mHasCornerProbe = False: Erase mCornerProbe
            mPostPeakControlName = vbNullString: mControlName = "материальная координата H огибающих"
            mMaterialGrowthActive = True: mMaterialGrowthProjection = mProjection
            mFieldControlChanges = mFieldControlChanges + 1: mLastFieldControlRevision = mSolver.AcceptedRevision
            If trial.Lambda < current.Lambda - LoadResolution(current.Lambda) And Not mFollowingLocalPeak Then
                If Not mUnverifiedDecline Then mUnverifiedDips = mUnverifiedDips + 1
                mUnverifiedDecline = True: mUnverifiedReference = GeomMax(mUnverifiedReference, mBestLambda)
            End If
            mEnvelopeRecoveries = mEnvelopeRecoveries + 1
            TryEnvelopeRecovery = True
            Exit Function
        End If
        mProbeCount = mProbeCount + GeomMax(0, probes - 1): mEnvelopeProbes = mEnvelopeProbes + probes
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
        mSolver.DiscardStep: desired = desired * 0.5
    Loop
End Function

' Повторяет правое поле после силовой пробы. Точка lambda = 1 должна
' принадлежать фактическому интервалу именно принятого повтора.
Private Function RepeatDesignTrial(ByRef current As TBranchPoint, ByRef trial As TBranchPoint, _
        ByVal targetCoordinate As Double, ByRef designVerified As Boolean, _
        Optional ByVal materialCoordinateProgress As Boolean = False) As Boolean
    Dim designCoordinate As Double
    Dim prediction As Variant
    If DesignLoadPrediction(current, trial, prediction) Then
        designVerified = VerifyDesignLoad(current.Coordinate, trial.Coordinate, designCoordinate, prediction)
    Else
        designVerified = VerifyDesignLoad(current.Coordinate, trial.Coordinate, designCoordinate)
    End If
    trial.Coordinate = targetCoordinate
    If Not Probe(trial, current.Lambda, keepYieldControl:=True, acceptedLambda:=current.Lambda) Then
        StopSearch "Не удалось повторить пробу после проверки заданных усилий."
        Exit Function
    End If
    If Not UpdateTrialCoordinate(trial, current.Coordinate) Then
        StopSearch "Повторная проба после проверки заданных усилий не продвинулась по деформационной ветви."
        Exit Function
    End If
    If Not mUsesEnvelopeControl And (mSolver.HasPrincipalBoundaries Or mFieldControlChanges > 0) Then
        If Not SpatialTrialAdvances(mCurrentFields, current) Then
            mSolver.DiscardStep
            StopSearch "Повторная проба не подтвердила направление прослеживаемой ветви."
            Exit Function
        End If
    End If
    If trial.Lambda > mMaximumLambda Then
        StopAtSearchBound
        Exit Function
    End If
    If trial.Lambda <= 0# Then
        mSolver.DiscardStep
        StopSearch "Повторная проба не подтвердила допустимое изменение множителя нагрузки."
        Exit Function
    End If
    If Not mFollowingLocalPeak Then
        Dim admissibleChange As Boolean
        If materialCoordinateProgress Or mUsesEnvelopeControl Then
            admissibleChange = (trial.Lambda >= current.Lambda - LoadResolution(current.Lambda))
        Else
            admissibleChange = (trial.Lambda > current.Lambda + LoadResolution(current.Lambda))
        End If
        If Not admissibleChange Then
            mSolver.DiscardStep
            StopSearch "Повторная проба не подтвердила допустимое изменение множителя нагрузки."
            Exit Function
        End If
    End If
    If BeyondSteelLimit() Then
        mSolver.DiscardStep
        StopSearch "Повторная проба не подтвердила допустимость арматуры."
        Exit Function
    End If
    If designVerified Then designVerified = (designCoordinate >= current.Coordinate And designCoordinate <= trial.Coordinate)
    RepeatDesignTrial = True
End Function

' Готовит начальное поле точного сочетания между принятым левым и найденным правым узлами CDP.
' Интерполяция не является решением и не переносит пробную историю материала.
Private Function DesignLoadPrediction(ByRef left As TBranchPoint, ByRef right As TBranchPoint, ByRef prediction As Variant) As Boolean
    If Not mSolver.UsesSpatialHistory Or mSolver.HasPrincipalBoundaries Then Exit Function
    If right.Lambda <= left.Lambda Then Exit Function
    Dim ratio As Double: ratio = (1# - left.Lambda) / (right.Lambda - left.Lambda)
    If ratio < 0# Or ratio > 1# Then Exit Function
    Dim first() As Double, last() As Double, actual() As Double, degree As Long
    mSolver.CopyAccepted first: mSolver.CopyTrial last, actual
    For degree = 1 To UBound(first)
        last(degree) = first(degree) + ratio * (last(degree) - first(degree))
    Next degree
    prediction = last
    DesignLoadPrediction = True
End Function

' Проверяет точное lambda = 1 от левого плеча перед продвижением истории.
' Контрольная мера должна попасть внутрь проверяемого участка.
Private Function VerifyDesignLoad(ByVal leftCoordinate As Double, ByVal rightCoordinate As Double, _
        Optional ByRef coordinate As Double = 0#, Optional ByRef prediction As Variant) As Boolean
    If mProbeCount >= PROBE_LIMIT Then Exit Function
    mProbeCount = mProbeCount + 1
    Dim load(1 To COMPONENT_COUNT) As Double, i As Long
    For i = 1 To COMPONENT_COUNT: load(i) = mOffset(i) + mDirection(i): Next i
    Dim solved As Boolean
    If IsMissing(prediction) Then
        solved = mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), mAbsTol, mRelTol, mMaxIterations)
    Else
        solved = mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), mAbsTol, mRelTol, mMaxIterations, fallbackField:=prediction)
    End If
    If Not solved Then Exit Function
    Dim field() As Double, resultants() As Double
    mSolver.CopyTrial field, resultants
    If mUsesEnvelopeControl Then
        If Not mSolver.TryReadEnvelopeCoordinate(False, coordinate) Then mSolver.DiscardStep: Exit Function
    Else
        coordinate = CoordinateOf(field)
    End If
    If Not BeyondSteelLimit() Then VerifyDesignLoad = (coordinate >= leftCoordinate And coordinate <= rightCoordinate)
    mSolver.DiscardStep
End Function

' Читает точную силовую точку из сохранённой истории с полной постоянной частью.
' Без следа можно подтвердить только один из двух переданных конечных узлов.
Private Function InitialDesignHistoryIndex(ByVal historyCount As Long, _
        ByVal previousLambda As Double, ByVal currentLambda As Double) As Long
    If Not mTrace Is Nothing Then
        If mTrace.Count = historyCount Then
            Dim i As Long, fraction As Double, factor As Double
            For i = 1 To historyCount
                mTrace.ReadPoint i, fraction, factor
                If fraction = 1# And factor = 1# Then
                    InitialDesignHistoryIndex = i
                    Exit Function
                End If
            Next i
            Exit Function
        End If
    End If
    If currentLambda = 1# Then
        InitialDesignHistoryIndex = historyCount
    ElseIf previousLambda = 1# Then
        InitialDesignHistoryIndex = historyCount - 1
    End If
End Function

' Проверяет точное сочетание от сохранённого уточнённого левого поля.
' Силовая проба отбрасывается, затем восстанавливается вся передаваемая история.
Private Function VerifyRefinedDesign(ByVal fields As Collection, ByVal leftIndex As Long, _
        ByRef designIndex As Long) As Boolean
    If Not mSolver.ReplayAcceptedFields(fields, leftIndex) Then Exit Function
    Dim leftField() As Double, rightField() As Double
    leftField = fields.Item(leftIndex): rightField = fields.Item(leftIndex + 1)
    Dim leftCoordinate As Double, rightCoordinate As Double
    If mUsesEnvelopeControl Then
        If Not mSolver.TryReadEnvelopeCoordinate(True, leftCoordinate) Then Exit Function
        If Not mSolver.TryEnvelopeCoordinateAt(rightField, rightCoordinate) Then Exit Function
    Else
        leftCoordinate = CoordinateOf(leftField): rightCoordinate = CoordinateOf(rightField)
    End If
    If VerifyDesignLoad(leftCoordinate, rightCoordinate) Then designIndex = leftIndex + 1
    VerifyRefinedDesign = mSolver.ReplayAcceptedFields(fields, fields.Count)
End Function


' Повторяет оба плеча от одного принятого левого состояния и оценивает точность.
' Железобетон продолжает также от менее точного локального пика; определяющим он становится только при выполнении допуска.
Private Function VerifyPeak(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint, ByVal keepRight As Boolean) As Boolean
    If Not SelectPeakAnchor(history, left, middle) Then Exit Function
    Dim sharedPrefix As Long: sharedPrefix = left.HistoryIndex ' При уточнении сохраняются поля по этот индекс.
    Dim shoulderStart As Double: shoulderStart = left.Coordinate
    Dim first As TBranchPoint, second As TBranchPoint, tolerance As Double
    Dim checkedLeft As TBranchPoint, checkedRight As TBranchPoint
    checkedLeft = left: checkedRight = right
    tolerance = mLambdaTolerance
    If Abs(middle.Lambda - 1#) <= 0.05 And tolerance > NEAR_ONE_TOLERANCE Then tolerance = NEAR_ONE_TOLERANCE
    Dim sampling As Double, refinedFields As Collection, actual As Variant
    Dim accuracy As Double, divisions As Long, refinedDesignLeft As Long
    Dim retainedMaximum As Double, retainedPreviousMaximum As Double
    Dim peakLeft As TBranchPoint, peakRight As TBranchPoint, checkedMiddle As TBranchPoint
    If keepRight Then
        Dim sourceFields As Collection, sourceTrace As CCoupledLoadTrace, refinement As Long
        Dim retainedLeft As TBranchPoint, retainedRight As TBranchPoint, retainedPeak As TBranchPoint
        Set sourceFields = history: Set sourceTrace = mTrace: checkedMiddle = middle: first = middle
        For refinement = 1 To STEP_LIMIT
            If Not ReplayShoulders(sourceFields, checkedLeft, checkedMiddle, checkedRight, 2, second, sampling, peakLeft, peakRight, _
                    refinedFields, actual, True, retainedMaximum, refinedDesignLeft, retainedPreviousMaximum, sourceTrace) Then
                If refinedFields Is Nothing Then Exit Function
                If Not mSolver.ReplayAcceptedFields(refinedFields, refinedFields.Count) Then Exit Function
                checkedLeft = retainedLeft: checkedRight = retainedRight: second = retainedPeak
                Exit For ' Неудача следующего уточнения не отменяет уже проверенные плечи и их фактическую погрешность.
            End If
            accuracy = Abs(first.Lambda - second.Lambda)
            If sampling > accuracy Then accuracy = sampling
            accuracy = accuracy + first.Uncertainty + second.Uncertainty
            If retainedMaximum - second.Lambda > accuracy Then accuracy = retainedMaximum - second.Lambda
            retainedLeft = checkedLeft: retainedRight = checkedRight: retainedPeak = second
            If accuracy <= tolerance Then Exit For
            If mSolver.UsesSpatialHistory And mSolver.SupportsHybridRecovery And mDesignHistoryIndex = 0 Then Exit For ' Только пространственный CDP откладывает точность раннего пика; конечная приёмка требует исходного допуска.
            If first.Uncertainty + second.Uncertainty >= tolerance Or mProbeCount + 4 > PROBE_LIMIT Then Exit For
            first = second
            Set sourceFields = refinedFields: Set sourceTrace = mLimitTrace
            checkedLeft = peakLeft: checkedMiddle = second: checkedRight = peakRight ' Вне новой окрестности сохраняются уже принятые поля, а не пересчитываются широкие плечи.
        Next refinement
        If Not mLimitTrace Is Nothing Then refinedDesignLeft = RefinedDesignLeftIndex(mLimitTrace)
    Else
        If Not ReplayShoulders(history, left, middle, right, 2, first, sampling, peakLeft, peakRight) Then Exit Function
        divisions = 4
        Do
            If Not ReplayShoulders(history, left, middle, right, divisions, second, sampling, peakLeft, peakRight, refinedFields, actual, _
                    retainedMaximum:=retainedMaximum, designLeftIndex:=refinedDesignLeft, _
                    retainedPreviousMaximum:=retainedPreviousMaximum) Then Exit Function
            accuracy = Abs(first.Lambda - second.Lambda)
            If sampling > accuracy Then accuracy = sampling
            accuracy = accuracy + first.Uncertainty + second.Uncertainty
            If mBestLambda - second.Lambda > accuracy Then accuracy = mBestLambda - second.Lambda
            If accuracy <= tolerance Then Exit Do
            first = second
            divisions = divisions * 2
        Loop While divisions <= 16
        If accuracy > tolerance Then Exit Function
    End If
    Dim designIndex As Long: designIndex = mDesignHistoryIndex
    If designIndex > sharedPrefix Then designIndex = 0
    If designIndex = 0 And refinedDesignLeft > 0 And mProbeCount < PROBE_LIMIT Then
        If Not VerifyRefinedDesign(refinedFields, refinedDesignLeft, designIndex) Then
            Set mLimitTrace = Nothing
            Exit Function
        End If
    End If
    If keepRight Then left = checkedLeft: right = checkedRight
    mDesignHistoryIndex = designIndex
    mPeak = second: mPeakAccuracy = accuracy: mHasPeak = True
    Set mLimitFields = refinedFields: mLimitResultants = actual
    If keepRight Then RetainVerifiedPeak sharedPrefix, second, accuracy, refinedFields, designIndex
    mBestBeforeCurrent = retainedPreviousMaximum
    mBestLambda = retainedMaximum ' Равновесие удалённого крупного шага остаётся только в оценке погрешности.
    mMeta.SetResult rsSuccess, rcCheckPassed, rkCapacity, _
        "Подтверждён первый максимум прослеженной ветви.", Details() & " | Оценка погрешности lambda: " & CStr(accuracy) & _
        " | Координаты плеч: " & CStr(shoulderStart) & ", " & CStr(right.Coordinate)
    VerifyPeak = True
End Function

' Выбирает различимое восходящее плечо пологого пика по фактически принятому следу, без интерполяции равновесий.
' Для H используются значения того же принятого пути; координата старого поля не пересчитывается от новой материальной истории.
Private Function SelectPeakAnchor(ByVal history As Collection, ByRef left As TBranchPoint, ByRef middle As TBranchPoint) As Boolean
    If middle.Lambda > left.Lambda + LoadResolution(middle.Lambda) Then SelectPeakAnchor = True: Exit Function
    If mTrace Is Nothing Then Exit Function
    If mTrace.Count <> history.Count Or middle.HistoryIndex <> history.Count Then Exit Function
    Dim index As Long, fraction As Double, factor As Double, coordinate As Double, field() As Double
    For index = middle.HistoryIndex - 1 To 1 Step -1
        mTrace.ReadPoint index, fraction, factor
        If fraction <> 1# Then Exit Function
        If middle.Lambda > factor + LoadResolution(middle.Lambda) Then
            If mUsesEnvelopeControl Then
                If Not mSolver.TryReadAcceptedEnvelopeAt(index, history.Count, coordinate) Then Exit Function
            Else
                field = history.Item(index): coordinate = CoordinateOf(field)
            End If
            If coordinate < middle.Coordinate Then
                left.HistoryIndex = index: left.Lambda = factor: left.Coordinate = coordinate: left.Uncertainty = 0#
                SelectPeakAnchor = True
                Exit Function
            End If
        End If
    Next index
End Function

' Воспроизводит левые принятые поля, затем рассчитывает более мелкие
' шаги только вперёд. Сравнивает пик с обоими концами с учётом невязки.
' При сохранении спада передаёт последнюю пару точек уточнённой истории.
Private Function ReplayShoulders(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint, ByVal divisions As Long, _
        ByRef peak As TBranchPoint, ByRef sampling As Double, ByRef peakLeft As TBranchPoint, ByRef peakRight As TBranchPoint, _
        Optional ByRef limitFields As Collection = Nothing, _
        Optional ByRef limitResultants As Variant, Optional ByVal keepRight As Boolean = False, _
        Optional ByRef retainedMaximum As Double = 0#, Optional ByRef designLeftIndex As Long = 0, _
        Optional ByRef retainedPreviousMaximum As Double = 0#, Optional ByVal sourceTrace As CCoupledLoadTrace = Nothing) As Boolean
    If Not mSolver.ReplayAcceptedFields(history, left.HistoryIndex) Then Exit Function
    Dim start As Double, finish As Double, previousLambda As Double, point As TBranchPoint
    Dim samples() As TBranchPoint, shoulder As Long, i As Long, index As Long
    Dim saveFields As Boolean, refined As Collection, integrated As Collection
    saveFields = Not IsMissing(limitResultants)
    If saveFields Then
        Set refined = New Collection: Set integrated = New Collection
    End If
    Dim field() As Double, resultants() As Double
    Dim seedFields As New Collection, seedTrace As CCoupledLoadTrace, seedSource As CCoupledLoadTrace
    Dim seedLeft As TBranchPoint, seedCurrent As TBranchPoint, prediction As Variant, predictedLambda As Double
    Dim seedIndex As Long, seedFraction As Double, probed As Boolean
    For seedIndex = 1 To left.HistoryIndex: field = history.Item(seedIndex): seedFields.Add field: Next seedIndex
    Set seedSource = sourceTrace: If seedSource Is Nothing Then Set seedSource = mTrace
    If Not seedSource Is Nothing Then Set seedTrace = seedSource.CopyPrefix(left.HistoryIndex)
    ReDim samples(0 To 2 * divisions)
    samples(0) = left: previousLambda = left.Lambda
    For shoulder = 1 To 2
        If shoulder = 1 Then
            start = left.Coordinate: finish = middle.Coordinate
        Else
            start = middle.Coordinate: finish = right.Coordinate
        End If
        For i = 1 To divisions
            point.Coordinate = start + (finish - start) * i / divisions
            If point.Coordinate <= samples(index).Coordinate Then
                mProbeDetails = "Уточнённая цель плеча лежит на уже пройденной части деформационной ветви."
                Exit Function
            End If
            If mProbeCount < PROBE_LIMIT Then mShoulderProbes = mShoulderProbes + 1
            seedCurrent = samples(index): seedLeft.HistoryIndex = 0
            If index > 0 Then
                seedLeft = samples(index - 1)
            ElseIf left.HistoryIndex > 1 And Not seedTrace Is Nothing Then
                seedLeft.HistoryIndex = left.HistoryIndex - 1
                seedTrace.ReadPoint seedLeft.HistoryIndex, seedFraction, seedLeft.Lambda
                If seedFraction <> 1# Then seedLeft.HistoryIndex = 0
            End If
            If PrepareSecantSeed(seedFields, seedLeft, seedCurrent, point.Coordinate, prediction, predictedLambda, seedTrace) Then
                probed = Probe(point, predictedLambda, keepYieldControl:=keepRight, initialField:=prediction, acceptedLambda:=previousLambda)
            Else
                probed = Probe(point, previousLambda, keepYieldControl:=keepRight, acceptedLambda:=previousLambda)
            End If
            If Not probed Then Exit Function
            If Not UpdateTrialCoordinate(point, samples(index).Coordinate) Then Exit Function
            If BeyondSteelLimit() Or point.Lambda > mMaximumLambda Then
                mSolver.DiscardStep
                Exit Function
            End If
            If Not ConfirmShoulderAccuracy(point, previousLambda, samples(index).Coordinate, keepRight) Then
                mSolver.DiscardStep
                Exit Function
            End If
            If keepRight And point.Lambda <= 0# Then
                mSolver.DiscardStep
                Exit Function
            End If
            mSolver.CopyTrial field, resultants
            If saveFields Then
                refined.Add field: integrated.Add resultants
            End If
            mSolver.AcceptStep
            seedFields.Add field
            If Not seedTrace Is Nothing Then seedTrace.Add 1#, point.Lambda
            index = index + 1
            point.HistoryIndex = left.HistoryIndex + index
            samples(index) = point
            previousLambda = point.Lambda
        Next i
    Next shoulder
    Dim peakIndex As Long, lower As Double
    For i = 1 To index - 1
        point = samples(i): lower = point.Lambda - point.Uncertainty
        If lower > samples(i - 1).Lambda + samples(i - 1).Uncertainty And _
                lower > samples(i + 1).Lambda + samples(i + 1).Uncertainty Then
            If lower <= left.Lambda + LoadResolution(point.Lambda) Or _
                    lower <= samples(index).Lambda + samples(index).Uncertainty Then
                If Not keepRight Then Exit Function
            ElseIf peakIndex = 0 Then
                peakIndex = i
                If Not keepRight Then Exit For
            ElseIf lower > samples(peakIndex).Lambda + samples(peakIndex).Uncertainty Then
                peakIndex = i ' Больший максимум должен различаться с учётом невязки.
            End If
        End If
    Next i
    If peakIndex = 0 Then
        mProbeDetails = "При повторе плеч не найден максимум, отделённый от соседних уровней с учётом погрешности lambda. " & _
            "Крайние множители: " & CStr(samples(0).Lambda) & " и " & CStr(samples(index).Lambda) & "; проверено промежуточных уровней: " & CStr(index - 1) & "."
        Exit Function
    End If

    peak = samples(peakIndex): peakLeft = samples(peakIndex - 1): peakRight = samples(peakIndex + 1)
    Dim dropLeft As Double, dropRight As Double
    dropLeft = peak.Lambda - samples(peakIndex - 1).Lambda + samples(peakIndex - 1).Uncertainty
    dropRight = peak.Lambda - samples(peakIndex + 1).Lambda + samples(peakIndex + 1).Uncertainty
    sampling = GeomMax(dropLeft, dropRight)
    If saveFields Then
        Set limitFields = New Collection
        Dim j As Long, fieldCount As Long
        fieldCount = peakIndex
        If keepRight Then fieldCount = index
        For j = 1 To left.HistoryIndex: limitFields.Add history.Item(j): Next j
        Dim trace As CCoupledLoadTrace: Set trace = sourceTrace
        If trace Is Nothing Then Set trace = mTrace
        retainedMaximum = PrefixMaximumLambda(left.HistoryIndex, trace): designLeftIndex = 0 ' Удалённые промежуточные поля не остаются подтверждённой границей нагрузки.
        For j = 1 To fieldCount
            limitFields.Add refined.Item(j)
            retainedPreviousMaximum = retainedMaximum
            retainedMaximum = GeomMax(retainedMaximum, samples(j).Lambda)
            If designLeftIndex = 0 And samples(j - 1).Lambda < 1# And samples(j).Lambda >= 1# Then _
                designLeftIndex = left.HistoryIndex + j - 1
        Next j
        If Not trace Is Nothing Then
            Set mLimitTrace = trace.CopyPrefix(left.HistoryIndex)
            For j = 1 To fieldCount: mLimitTrace.Add 1#, samples(j).Lambda: Next j
        End If
        limitResultants = integrated.Item(fieldCount)
        If keepRight Then
            left = samples(index - 1) ' Пара сохранённых полей нужна для выбора меры после максимума.
            right = samples(index)
        End If
    End If
    ReplayShoulders = True

End Function

' Находит восходящий интервал lambda = 1 во всём сохраняемом следе после нескольких локальных уточнений плеч.
Private Function RefinedDesignLeftIndex(ByVal trace As CCoupledLoadTrace) As Long
    Dim i As Long, fraction As Double, factor As Double, previousFraction As Double, previous As Double
    For i = 1 To trace.Count
        trace.ReadPoint i, fraction, factor
        If i > 1 And previousFraction = 1# And fraction = 1# And previous < 1# And factor >= 1# Then
            RefinedDesignLeftIndex = i - 1
            Exit Function
        End If
        previousFraction = fraction: previous = factor
    Next i
End Function

' Сохраняет точный префикс наибольшего уточнённого максимума; последующие поля не входят в его материальную историю.
' Погрешность включает верхние оценки остальных проверенных максимумов той же сохраняемой ветви.
Private Sub RetainVerifiedPeak(ByVal sharedPrefix As Long, ByRef point As TBranchPoint, ByVal accuracy As Double, _
        ByVal fields As Collection, ByVal designIndex As Long)
    If mTrace Is Nothing Or mLimitTrace Is Nothing Then Exit Sub
    If Not mVerifiedPeakFields Is Nothing Then
        If mVerifiedPeak.HistoryIndex > sharedPrefix Then Set mVerifiedPeakFields = Nothing ' Замена уточняемого хвоста отменяет его прежний кандидат.
    End If
    If Not mVerifiedPeakFields Is Nothing Then
        If mVerifiedPeak.Lambda >= point.Lambda Then
            mVerifiedPeakAccuracy = GeomMax(mVerifiedPeakAccuracy, point.Lambda + accuracy - mVerifiedPeak.Lambda)
            Exit Sub
        End If
        accuracy = GeomMax(accuracy, mVerifiedPeak.Lambda + mVerifiedPeakAccuracy - point.Lambda)
    End If
    mVerifiedPeak = point: mVerifiedPeakAccuracy = accuracy
    mVerifiedPeakDesignIndex = 0
    If designIndex > 0 And designIndex <= point.HistoryIndex Then mVerifiedPeakDesignIndex = designIndex
    Set mVerifiedPeakFields = New Collection
    Dim i As Long
    For i = 1 To point.HistoryIndex: mVerifiedPeakFields.Add fields.Item(i): Next i
    Set mVerifiedPeakTrace = mLimitTrace.CopyPrefix(point.HistoryIndex)
End Sub

' Возвращает максимум только после локализованного первого предела стали на продолженной допустимой ветви.
' Непроверенный спад, больший поздний узел, недостаточная точность или невоспроизводимое поле сохраняют NumFail.
Private Function CompleteVerifiedPeak(ByVal history As Collection, ByVal lastBeforeEvent As Long, ByVal eventUpper As Double) As Boolean
    If mVerifiedPeakFields Is Nothing Or mVerifiedPeakTrace Is Nothing Then Exit Function
    If mUnverifiedDecline Or mVerifiedPeak.HistoryIndex > lastBeforeEvent Then Exit Function
    If mVerifiedPeakAccuracy > mLambdaTolerance Then Exit Function
    If mBestLambda > mVerifiedPeak.Lambda + mVerifiedPeakAccuracy Then Exit Function
    If eventUpper >= mVerifiedPeak.Lambda - mVerifiedPeakAccuracy Then Exit Function
    If mProbeCount >= PROBE_LIMIT Then Exit Function
    If Not mSolver.ReplayAcceptedFields(mVerifiedPeakFields, mVerifiedPeakFields.Count) Then GoTo RestoreTail
    Dim point As TBranchPoint, field() As Double, checked() As Double, actual() As Double, i As Long
    field = mVerifiedPeakFields.Item(mVerifiedPeakFields.Count)
    If mUsesEnvelopeControl Then
        Dim load(1 To COMPONENT_COUNT) As Double, seed As Variant
        For i = 1 To COMPONENT_COUNT: load(i) = mOffset(i) + mVerifiedPeak.Lambda * mDirection(i): Next i
        seed = field: mProbeCount = mProbeCount + 1: point.Lambda = mVerifiedPeak.Lambda
        If Not mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), _
                mAbsTol, mRelTol, 1, initialField:=seed) Then GoTo RestoreTail ' Повтор уже принятого пика не является новым положительным H-шагом.
    Else
        point.Coordinate = CoordinateOf(field)
        If Not Probe(point, mVerifiedPeak.Lambda, keepYieldControl:=True) Then GoTo RestoreTail
    End If
    If Abs(point.Lambda - mVerifiedPeak.Lambda) > LoadResolution(mVerifiedPeak.Lambda) Or BeyondSteelLimit() Then GoTo RestoreTail
    mSolver.CopyTrial checked, actual
    For i = 1 To UBound(field)
        If checked(i) <> field(i) Then GoTo RestoreTail ' Равновесие соседнего поля не подтверждает сохранённую точку максимума.
    Next i
    mSolver.DiscardStep ' Проба проверяет уже принятое поле; нового шага истории здесь нет.
    mPeak = mVerifiedPeak: mPeakAccuracy = mVerifiedPeakAccuracy: mHasPeak = True
    Set mLimitFields = mVerifiedPeakFields: Set mLimitTrace = mVerifiedPeakTrace
    mLimitResultants = actual: mDesignHistoryIndex = mVerifiedPeakDesignIndex
    mMeta.SetResult rsSuccess, rcCheckPassed, rkCapacity, _
        "Подтверждён максимум нагрузки допустимой ветви, завершённой предельной деформацией арматуры.", Details() & _
        " | Определяющий максимум lambda: " & CStr(mPeak.Lambda) & "; оценка погрешности: " & CStr(mPeakAccuracy)
    CompleteVerifiedPeak = True
    Exit Function
RestoreTail:
    mSolver.DiscardStep
    If Not mSolver.ReplayAcceptedFields(history, history.Count) Then mProbeDetails = "Не удалось восстановить допустимый хвост после проверки кандидата максимума."
End Function

' Вычисляет подтверждённую нагрузку только по сохраняемому префиксу, включая максимум до локального спада.
' Без журнала левое поле непосредственно предшествует текущему и используется сохранённая общая граница.
Private Function PrefixMaximumLambda(ByVal lastIndex As Long, Optional ByVal sourceTrace As CCoupledLoadTrace = Nothing) As Double
    Dim trace As CCoupledLoadTrace: Set trace = sourceTrace
    If trace Is Nothing Then Set trace = mTrace
    If trace Is Nothing Then
        PrefixMaximumLambda = mBestBeforeCurrent
        Exit Function
    End If
    Dim i As Long, fraction As Double, factor As Double
    For i = 1 To lastIndex
        trace.ReadPoint i, fraction, factor
        If fraction = 1# And factor > PrefixMaximumLambda Then PrefixMaximumLambda = factor
    Next i
End Function

' Уточняет первое пересечение по возрастающей деформационной координате.
' Позднее событие сохраняет как диагностику, если был достигнут больший
' допустимый множитель или положение относительно единицы не определено.
Private Function ResolveSteelCrossing(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef right As TBranchPoint) As Boolean
    Dim eventState As TSteelCrossing, refined As New Collection, multipliers As New Collection
    Dim actual() As Double: actual = mCurrentResultants
    eventState.Admissible = left: eventState.Exceeded = right
    If Not RefineSteelCrossing(eventState, refined, multipliers, actual) Then
        mSolver.DiscardStep
        If refined.Count = 0 Then Exit Function
        If mSolver.ReplayAcceptedFields(history, left.HistoryIndex) Then Exit Function
        StopSearch "Не удалось восстановить принятую историю после уточнения предела арматуры."
        ResolveSteelCrossing = True
        Exit Function
    End If

    Dim factor As Variant
    For Each factor In multipliers
        If factor > mBestLambda Then mBestLambda = factor
    Next factor
    Dim low As Double, high As Double, lowerBound As Double, upperBound As Double
    low = eventState.Admissible.Lambda: high = eventState.Exceeded.Lambda
    mSteelEventPoint = eventState.CriticalPoint
    lowerBound = GeomMin(low - eventState.Admissible.Uncertainty, high - eventState.Exceeded.Uncertainty)
    upperBound = GeomMax(low + eventState.Admissible.Uncertainty, high + eventState.Exceeded.Uncertainty)
    mSteelEventDetails = "Арматурное событие: lambda до = " & CStr(low) & ", lambda после = " & CStr(high) & _
        ". Координата деформации: [" & CStr(eventState.Admissible.Coordinate) & ", " & CStr(eventState.Exceeded.Coordinate) & _
        "]. Точка арматуры: " & CStr(eventState.CriticalPoint) & "; отношения деформаций к пределам: " & _
        CStr(eventState.AdmissibleRatio) & " и " & CStr(eventState.ExceededRatio) & _
        "; оценка интервала lambda: [" & CStr(lowerBound) & ", " & CStr(upperBound) & "]"
    Dim reason As String
    If high <= low Or mFollowingLocalPeak Or low + LoadResolution(low) < mBestLambda Then
        reason = "Достигнута предельная деформация арматуры после максимума или снижения нагрузки. Определяющий предел требует проверки."
    ElseIf lowerBound <= 1# And upperBound >= 1# Then
        reason = "Достигнута предельная деформация арматуры. Интервал её достижения охватывает расчётную нагрузку (lambda = 1)."
    End If
    Dim i As Long
    If Len(reason) > 0 Then
        For i = 1 To refined.Count: history.Add refined.Item(i): Next i ' Эти допустимые подшаги уже приняты материалом при уточнении интервала.
        If Not mTrace Is Nothing Then
            For Each factor In multipliers: mTrace.Add 1#, factor: Next factor
        End If
        Set mCurrentFields = history: mCurrentResultants = actual
        If Not CompleteVerifiedPeak(history, left.HistoryIndex, upperBound) Then StopSearch reason
    Else
        mSteelLow = low: mSteelHigh = high: mSteelPoint = eventState.CriticalPoint
        Set mLimitFields = New Collection
        For i = 1 To left.HistoryIndex: mLimitFields.Add history.Item(i): Next i
        For i = 1 To refined.Count: mLimitFields.Add refined.Item(i): Next i
        If Not mTrace Is Nothing Then
            Set mLimitTrace = mTrace.CopyPrefix(left.HistoryIndex)
            For Each factor In multipliers: mLimitTrace.Add 1#, factor: Next factor
        End If
        mLimitResultants = actual
        mMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
            "Достигнута предельная деформация арматуры на возрастающем участке.", Details()
    End If
    ResolveSteelCrossing = True
End Function

' Сохраняет только допустимые подшаги. После их принятия правое плечо
' рассчитывается повторно, поскольку изменилось состояние материалов.
Private Function RefineSteelCrossing(ByRef eventState As TSteelCrossing, ByVal refined As Collection, _
        ByVal multipliers As Collection, ByRef actual() As Double) As Boolean
    Dim low As TBranchPoint, high As TBranchPoint, point As TBranchPoint
    low = eventState.Admissible: high = eventState.Exceeded
    If low.Lambda <= 0# Or high.Coordinate <= low.Coordinate Then Exit Function
    eventState.AdmissibleRatio = mReinforcement.AcceptedSteelUtilization
    If eventState.AdmissibleRatio > 1# Or Not ReadSteelTrialBound(high) Then Exit Function
    eventState.ExceededRatio = mReinforcement.TrialSteelUtilization
    eventState.CriticalPoint = mReinforcement.TrialSteelCriticalPoint
    If eventState.ExceededRatio <= 1# Or eventState.CriticalPoint = 0 Then Exit Function
    mSolver.DiscardStep
    low.Uncertainty = GeomMax(low.Uncertainty, LoadResolution(low.Lambda))
    Dim tolerance As Double, coordinateTolerance As Double
    tolerance = mLambdaTolerance
    If GeomMin(low.Lambda, high.Lambda) <= 1.05 And GeomMax(low.Lambda, high.Lambda) >= 0.95 Then _
        tolerance = GeomMin(tolerance, NEAR_ONE_TOLERANCE)
    coordinateTolerance = (high.Coordinate - low.Coordinate) * MIN_STEP_RATIO ' Плато lambda тоже требует локализации деформации.
    Dim attempt As Long, field() As Double, ratio As Double
    For attempt = 1 To STEP_LIMIT
        If high.Coordinate - low.Coordinate <= coordinateTolerance And Abs(high.Lambda - low.Lambda) <= tolerance Then
            eventState.Admissible = low: eventState.Exceeded = high
            RefineSteelCrossing = True
            Exit Function
        End If
        point.Coordinate = low.Coordinate + 0.5 * (high.Coordinate - low.Coordinate)
        If point.Coordinate <= low.Coordinate Or point.Coordinate >= high.Coordinate Then Exit Function
        If Not Probe(point, low.Lambda, keepYieldControl:=True, acceptedLambda:=low.Lambda) Then Exit Function
        If Not UpdateTrialCoordinate(point, low.Coordinate) Then Exit Function
        If point.Coordinate >= high.Coordinate Then
            mSolver.DiscardStep
            mProbeDetails = "Фактическая мера промежуточной пробы вышла за правое плечо предела арматуры."
            Exit Function
        End If
        If Not ReadSteelTrialBound(point) Then Exit Function
        ratio = mReinforcement.TrialSteelUtilization
        If ratio > 1# Then
            high = point: eventState.ExceededRatio = ratio
            eventState.CriticalPoint = mReinforcement.TrialSteelCriticalPoint
        Else
            low = point: eventState.AdmissibleRatio = ratio
            mSolver.CopyTrial field, actual
            mSolver.AcceptStep
            refined.Add field: multipliers.Add low.Lambda
            If Not Probe(high, low.Lambda, keepYieldControl:=True, acceptedLambda:=low.Lambda) Then Exit Function
            If Not UpdateTrialCoordinate(high, low.Coordinate) Then Exit Function
            If Not ReadSteelTrialBound(high) Then Exit Function
            eventState.ExceededRatio = mReinforcement.TrialSteelUtilization
            eventState.CriticalPoint = mReinforcement.TrialSteelCriticalPoint
            If eventState.ExceededRatio <= 1# Then Exit Function
        End If
        mSolver.DiscardStep
    Next attempt
End Function

' Проверяет область поиска и оценивает влияние невязки около единицы.
Private Function ReadSteelTrialBound(ByRef point As TBranchPoint) As Boolean
    If point.Lambda <= 0# Or point.Lambda > mMaximumLambda Then Exit Function
    point.Uncertainty = LoadResolution(point.Lambda)
    If Abs(point.Lambda - 1#) <= 0.05 Then
        Dim uncertainty As Double
        If Not mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, uncertainty) Then Exit Function
        point.Uncertainty = GeomMax(point.Uncertainty, uncertainty)
    End If
    ReadSteelTrialBound = True
End Function

' Определяет превышение по пробной точке до фиксации истории.
Private Function BeyondSteelLimit() As Boolean
    If Not mReinforcement Is Nothing Then BeyondSteelLimit = (mReinforcement.TrialSteelUtilization > 1#)
End Function

' Считывает выбранную физическую меру из поля.
Private Function CoordinateOf(ByRef field() As Double) As Double
    Dim i As Long
    For i = 1 To mSolver.PhysicalDegreeCount: CoordinateOf = CoordinateOf + mProjection(i) * field(i): Next i
End Function

' Выделяет изменение множителя для выбора пробного интервала. В итоговой
' проверке дополнительно используется чувствительность к невязке.
Private Function LoadResolution(ByVal lambda As Double) As Double
    LoadResolution = 8# * mRelTol * (1# + Abs(lambda))
End Function

' Отклоняет пробу за установленной пользователем границей множителя
' и сохраняет достигнутую допустимую нагрузку в диагностике.
Private Sub StopAtSearchBound()
    mSolver.DiscardStep
    mMeta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, _
        "Достигнута граница Capacity.MaxLambda до подтверждения предельного состояния.", Details()
End Sub

' Сохраняет численную неопределённость и достигнутую допустимую нагрузку.
Private Sub StopSearch(ByVal reason As String)
    If mLocalPeak > 0# Then reason = "Подтверждён локальный максимум при lambda = " & Format$(mLocalPeak, "0.000000") & _
        ". Предельное состояние железобетонного сечения не установлено. " & reason
    If mBestLambda > 0# Then reason = reason & " Допустимое равновесие подтверждено до lambda = " & Format$(mBestLambda, "0.000000") & "."
    mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, reason, Details()
    modTestCoupledElasticState.TestCaptureStockContinuation Me, Details()
End Sub

' Записывает выбранную управляемую величину и область проверенной ветви.
Private Function Details() As String
    If mSignedArcProbes > 0 Then Details = "; вызовов проверенной знаковой дуги " & CStr(mSignedArcProbes) & "; принятых направлений " & CStr(mSignedArcRecovered)
    Details = Details & "Мера деформации: " & mControlName & "; наибольший допустимый lambda: " & CStr(mBestLambda) & _
        "; равновесие при lambda = 1 подтверждено: " & CStr(mDesignHistoryIndex > 0) & "; проб: " & CStr(mProbeCount)
    If mUnverifiedDips > 0 Then Details = Details & "; неподтверждённых спадов: " & CStr(mUnverifiedDips) & _
        "; восстановлений после них: " & CStr(mRecoveredDips)
    If mUnverifiedDecline Then Details = Details & "; текущий спад не является подтверждённым физическим пределом"
    If mFieldControlChanges > 0 Then Details = Details & "; смен пространственной меры после отказа: " & CStr(mFieldControlChanges)
    If mHybridRecoveryProbes > 0 Then Details = Details & "; проб полного разностного резерва после деления шага: " & _
        CStr(mHybridRecoveryProbes) & "; принято после него: " & CStr(mHybridRecoveredPoints) & _
        "; последняя цель резерва: " & CStr(mLastHybridTarget)
    If mUnresolvedLoadRetries > 0 Then Details = Details & "; повторов после неразличимого изменения lambda: " & CStr(mUnresolvedLoadRetries)
    If mYieldRecoveryProbes > 0 Then Details = Details & "; проб смены меры: " & CStr(mYieldRecoveryProbes) & _
        "; принято после смены меры: " & CStr(mYieldRecoveredPoints)
    If mSpatialReversalsRejected > 0 Then Details = Details & "; обратных пространственных приращений отклонено: " & CStr(mSpatialReversalsRejected)
    If mPrincipalCornerProbes > 0 Then Details = Details & "; проб материального перехода: " & CStr(mPrincipalCornerProbes) & _
        "; подтверждённых переходов: " & CStr(mPrincipalCorners) & "; прогнозов из проверочного равновесия: " & CStr(mCornerProbeSeeds)
    If mMaterialGrowthActive Then Details = Details & "; расширений шага CDP: " & CStr(mMaterialStepExpansions) & _
        "; наибольший принятый шаг после CDP: " & CStr(mMaterialMaximumStep)
    If mEnvelopeProbes > 0 Then Details = Details & "; проб материальной координаты: " & CStr(mEnvelopeProbes) & _
        "; восстановлений по ней: " & CStr(mEnvelopeRecoveries)
    If mUsesEnvelopeControl Then Details = Details & "; последующие шаги и плечи проверяются по H"
    If mEnvelopeDepartures > 0 Then Details = Details & "; возвратов к пространственной мере после остановки H: " & CStr(mEnvelopeDepartures)
    If mDepartureRecoveryProbes > 0 Then Details = Details & "; материальных резервов после остановки H: " & _
        CStr(mDepartureRecoveryProbes) & "; принято после них: " & CStr(mDepartureRecoveredPoints)
    If mGuideRecoveryProbes > 0 Then Details = Details & "; направляющих проб и возвратов: " & CStr(mGuideRecoveryProbes) & _
        "; подтверждённых возвратов к исходной цели: " & CStr(mGuideRecoverySolutions)
    If Len(mPostPeakControlName) > 0 Then Details = Details & "; пробная мера после максимума: " & mPostPeakControlName & _
        "; принято полей после максимума в новой мере: " & CStr(mPostPeakControlPoints)
    If mFixedYieldRecoveryProbes > 0 Then Details = Details & _
        "; проверок перехода в исходной мере: " & CStr(mFixedYieldRecoveryProbes)
    If mPostPeakWidePoints > 0 Then Details = Details & "; принято увеличенных шагов: " & CStr(mPostPeakWidePoints)
    If mLocalPeak > 0# Then
        Details = Details & "; локальный максимум lambda: " & CStr(mLocalPeak) & _
            "; погрешность локального максимума: " & CStr(mLocalPeakAccuracy) & _
            "; максимумов превзойдено после спада: " & CStr(mRecoveredPeaks)
        Details = Details & "; повторно использовано точек правого плеча: " & CStr(mReusedRightPoints)
        Details = Details & "; проб при проверке плеч: " & CStr(mShoulderProbes)
    End If
    If mArcRecoveryTried Then Details = Details & "; вспомогательных проб псевдодуги и повтора: " & CStr(mArcRecoveryProbes)
    If Len(mSteelEventDetails) > 0 Then Details = Details & " | " & mSteelEventDetails
    If Len(mProbeDetails) > 0 Then Details = Details & " | Последняя неудачная проба: " & mProbeDetails
End Function

' Проходит локальную материальную границу двумя отдельными принятыми состояниями.
' Разворот допускается только внутри решённого события; следующие обычные шаги снова проверяют пространственную ориентацию.
Private Function TryCrossPrincipalCorner(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef nominal As Double, ByRef increment As Double, _
        ByRef minimum As Double, ByRef entered As Boolean, Optional ByVal afterStepExhaustion As Boolean = False) As Boolean
    entered = False
    If mUsesEnvelopeControl Then Exit Function
    If Not mSolver.HasPrincipalTransitionSegment Or mProbeCount + 2 > PROBE_LIMIT Then Exit Function
    If Not mSolver.HasPrincipalBoundaries And Not afterStepExhaustion Then Exit Function ' CDP-событие не заменяет доступные деление шага и смену пространственной меры.
    Dim previous() As Double, accepted() As Double, incomingField() As Double, field() As Double, actual() As Double
    ReadSpatialReference history, current.HistoryIndex, previous
    accepted = history(current.HistoryIndex)
    If Not mSolver.CanPreparePrincipalTransition(previous, current.Lambda, mAbsTol, mRelTol) Then Exit Function
    If mCornerAttemptRevision <> mSolver.AcceptedRevision Then
        mCornerAttemptRevision = mSolver.AcceptedRevision: mCornerAttemptsAtRevision = 0
    End If
    If mCornerAttemptsAtRevision >= YIELD_RECOVERY_LIMIT Then Exit Function
    mCornerAttemptsAtRevision = mCornerAttemptsAtRevision + 1
    Dim transition As CCoupledPrincipalTransition, projection() As Double, point As TBranchPoint
    mProbeCount = mProbeCount + 1: mPrincipalCornerProbes = mPrincipalCornerProbes + 1: NotifyProgress
    If Not mSolver.TryPreparePrincipalTransition(mOffset, mDirection, previous, current.Lambda, mAbsTol, mRelTol, _
            mMaxIterations, transition, point.Lambda, projection) Then
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
        Exit Function
    End If
    If point.Lambda > mMaximumLambda Or point.Lambda <= 0# Or BeyondSteelLimit() Then
        mProbeDetails = "Входящая сторона материального перехода вышла за допустимую нагрузку или предел арматуры."
        mSolver.DiscardStep
        Exit Function
    End If
    If mDesignHistoryIndex = 0 And (current.Lambda - 1#) * (point.Lambda - 1#) <= 0# And Not mSolver.PrincipalDesignLoadVerified Then
        mProbeDetails = "Переход через заданное сочетание требует отдельного подтверждения точного lambda=1."
        mSolver.DiscardStep
        Exit Function
    End If
    mSolver.CopyTrial incomingField, actual
    Dim oldCoordinate As Double, newCoordinate As Double, i As Long, difference As Double
    For i = 1 To mSolver.PhysicalDegreeCount
        oldCoordinate = oldCoordinate + projection(i) * accepted(i)
        newCoordinate = newCoordinate + projection(i) * incomingField(i)
    Next i
    difference = newCoordinate - oldCoordinate
    If difference = 0# Then mSolver.DiscardStep: Exit Function
    If difference < 0# Then
        For i = 1 To UBound(projection): projection(i) = -projection(i): Next i
        oldCoordinate = -oldCoordinate: newCoordinate = -newCoordinate
    End If
    mProjection = projection: current.Coordinate = oldCoordinate: point.Coordinate = newCoordinate
    mUsesYieldControl = True: mBranchTangentControl = False
    mUsesTransverseControl = False: mUsesFieldControl = False
    mControlName = "входящая мера материального перехода"
    nominal = Abs(difference): increment = 0.5 * nominal: minimum = increment * MIN_STEP_RATIO
    Dim designAtPoint As Boolean: designAtPoint = mSolver.PrincipalDesignLoadVerified
    AcceptPoint history, point
    If designAtPoint Then mDesignHistoryIndex = point.HistoryIndex
    left = current: current = point: entered = True

    mProbeCount = mProbeCount + 1: mPrincipalCornerProbes = mPrincipalCornerProbes + 1: NotifyProgress
    If Not mSolver.TryCompletePrincipalTransition(transition, mOffset, mDirection, current.Lambda, _
            mMaxIterations, point.Lambda, projection, mAbsTol, mRelTol) Then
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
        Exit Function
    End If
    If point.Lambda > mMaximumLambda Or point.Lambda <= 0# Or BeyondSteelLimit() Then
        mProbeDetails = "Исходящая сторона материального перехода вышла за допустимую нагрузку или предел арматуры."
        mSolver.DiscardStep
        Exit Function
    End If
    If mDesignHistoryIndex = 0 And (current.Lambda - 1#) * (point.Lambda - 1#) <= 0# And Not mSolver.PrincipalDesignLoadVerified Then
        mProbeDetails = "Исходящая сторона требует отдельного подтверждения точного lambda=1."
        mSolver.DiscardStep
        Exit Function
    End If
    mSolver.CopyTrial field, actual
    point.Coordinate = CoordinateOf(field)
    designAtPoint = mSolver.PrincipalDesignLoadVerified
    AcceptPoint history, point
    If designAtPoint Then mDesignHistoryIndex = point.HistoryIndex
    left = current: current = point
    Dim coordinate As Double, probe As Variant, probeLambda As Double
    If Not mSolver.TryPrincipalDepartureControl(transition, current.Lambda, mCornerDirection, mCornerLambdaDirection, projection, coordinate, _
            linearizationField:=probe, linearizationLambda:=probeLambda, maxIterations:=mMaxIterations) Then
        mProbeDetails = "Не подтверждена касательная исходящей ветви после принятия материального перехода."
        Exit Function
    End If
    mProjection = projection: current.Coordinate = coordinate
    left.Coordinate = CoordinateOf(incomingField)
    mUsesYieldControl = False: mUsesFieldControl = True: mBranchTangentControl = False
    mPostPeakControlName = vbNullString
    mControlName = "пространственная мера после материального перехода"
    mPrincipalCorners = mPrincipalCorners + 1: mFieldControlChanges = mFieldControlChanges + 1
    mLastFieldControlRevision = mSolver.AcceptedRevision
    nominal = transition.ReferenceIncrement
    increment = nominal * MIN_STEP_RATIO: minimum = increment * MIN_STEP_RATIO
    mHasCornerReference = True: mCornerReferenceRevision = mSolver.AcceptedRevision
    mCornerReferenceIncrement = nominal
    If transition.PrincipalIndex = 0 And current.Lambda < left.Lambda And Not mFollowingLocalPeak Then
        mUnverifiedDecline = True: mUnverifiedReference = mBestLambda
        mUnverifiedDips = mUnverifiedDips + 1 ' Спад после решённого события требует продолжения ветви и не подтверждает предельную нагрузку.
    End If
    Dim probeRadius As Double: probeRadius = 0.25 * nominal
    If transition.PrincipalIndex = 0 Then
        mMaterialGrowthActive = True: mMaterialGrowthProjection = mProjection
        probeRadius = nominal ' Радиус проверочной CDP-пробы совпадает с радиусом её нативного решения.
    End If
    PreparePrincipalProbeSeed history, current, probe, probeLambda, increment, probeRadius
    TryCrossPrincipalCorner = True
End Function

' Использует уже найденное непринятое равновесие как прогноз обычной пробы, если оно продвигает исходящее направление.
' Точный ключ связан с принятой ревизией через mHasCornerReference; материал, направление и пределы проверяются обычным Probe заново.
Private Sub PreparePrincipalProbeSeed(ByVal history As Collection, ByRef current As TBranchPoint, _
        ByRef probe As Variant, ByVal lambda As Double, ByRef increment As Double, ByVal maximumSpacing As Double)
    mHasCornerProbe = False
    If IsEmpty(probe) Or lambda <= 0# Then Exit Sub
    Dim field() As Double, accepted() As Double, previous() As Double
    field = probe: accepted = history(current.HistoryIndex)
    Dim coordinate As Double, spacing As Double, advance As Double, reference As Double
    coordinate = CoordinateOf(field): spacing = coordinate - current.Coordinate
    If spacing <= increment Or spacing > maximumSpacing Then Exit Sub
    ReadSpatialReference history, current.HistoryIndex, previous
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, field, advance, reference) Then Exit Sub
    mCornerProbe = field: mCornerProbeLambda = lambda: mCornerProbeCoordinate = coordinate
    mHasCornerProbe = True: increment = spacing
End Sub

' После материального угла использует касательную именно принятой исходящей истории до первого обычного принятого шага.
' Это ориентир для проверки и прогноза, а не дополнительное принятое состояние материала.
Private Sub ReadSpatialReference(ByVal history As Collection, ByVal index As Long, ByRef previous() As Double)
    previous = history(index - 1)
    If Not mHasCornerReference Or mCornerReferenceRevision <> mSolver.AcceptedRevision Then Exit Sub
    previous = history(index)
    Dim i As Long
    For i = 1 To mSolver.PhysicalDegreeCount: previous(i) = previous(i) - mCornerReferenceIncrement * mCornerDirection(i): Next i
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Запускает фактический Probe с граничным остатком общего бюджета; цель и прогноз задаёт проверяемая история сечения.
Friend Function TestProbeComputedGuide(ByVal solver As CCoupledNewtonSolver, ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef projection() As Double, ByVal target As Double, ByRef prediction As Variant, ByVal factor As Double, _
        ByVal acceptedFactor As Double, ByVal initialProbes As Long, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        Optional ByVal afterVerifiedPeak As Boolean = True) As Variant
    Set mSolver = solver: mOffset = baseLoad: mDirection = direction: mProjection = projection
    mMaxIterations = 30: mAbsTol = absoluteTolerance: mRelTol = relativeTolerance
    mProbeCount = initialProbes: mGuideRecoveryProbes = 0: mGuideRecoverySolutions = 0
    mBranchTangentControl = True: mFieldControlChanges = 1: mUsesEnvelopeControl = False
    mFollowingLocalPeak = afterVerifiedPeak
    Dim point As TBranchPoint, solved As Boolean: point.Coordinate = target
    solved = Probe(point, factor, initialField:=prediction, allowHybridRecovery:=True, acceptedLambda:=acceptedFactor)
    TestProbeComputedGuide = Array(Abs(CLng(solved)), mProbeCount, mGuideRecoveryProbes, mGuideRecoverySolutions, point.Lambda)
End Function

' При малом изменении нагрузки принимает только уже сошедшееся движение по прежней пространственной ветви.
' Разрешение максимума по lambda не является минимальным допустимым приращением нагрузки.
Private Function AdvancesWithSmallLoadChange(ByVal history As Collection, ByRef current As TBranchPoint, ByRef trial As TBranchPoint) As Boolean
    If Not mSolver.UsesSpatialHistory Or mUsesEnvelopeControl Or Not mSolver.SupportsHybridRecovery Then Exit Function
    If trial.Lambda <= 0# Or Abs(trial.Lambda - current.Lambda) > LoadResolution(current.Lambda) Then Exit Function
    AdvancesWithSmallLoadChange = SpatialTrialAdvances(history, current)
End Function

' Выделяет различимый суммарный спад от наибольшего принятого узла, даже если отдельные приращения меньше разрешения lambda.
' Левое плечо и кандидат берутся только из фактически принятого следа; будущий ответ не интерполируется.
Private Function SelectCumulativePeak(ByVal history As Collection, ByRef trial As TBranchPoint, _
        ByRef left As TBranchPoint, ByRef middle As TBranchPoint) As Boolean
    If Not mSolver.UsesSpatialHistory Or mUsesEnvelopeControl Or mFollowingLocalPeak Then Exit Function
    If mTrace Is Nothing Then Exit Function
    If mTrace.Count <> history.Count Or history.Count < 3 Then Exit Function
    Dim index As Long, bestIndex As Long, fraction As Double, factor As Double, best As Double
    For index = 1 To history.Count
        mTrace.ReadPoint index, fraction, factor
        If fraction <> 1# Then Exit Function
        If factor > best Then best = factor: bestIndex = index
    Next index
    If bestIndex < 2 Or best - trial.Lambda <= LoadResolution(best) Then Exit Function
    Dim field() As Double, selectedLeft As TBranchPoint, selectedMiddle As TBranchPoint
    field = history(bestIndex)
    selectedMiddle.HistoryIndex = bestIndex: selectedMiddle.Lambda = best: selectedMiddle.Coordinate = CoordinateOf(field): selectedMiddle.Uncertainty = 0#
    If trial.Coordinate <= selectedMiddle.Coordinate Then Exit Function
    For index = bestIndex - 1 To 1 Step -1
        mTrace.ReadPoint index, fraction, factor
        If best > factor + LoadResolution(best) Then
            field = history(index)
            selectedLeft.HistoryIndex = index: selectedLeft.Lambda = factor: selectedLeft.Coordinate = CoordinateOf(field): selectedLeft.Uncertainty = 0#
            If selectedLeft.Coordinate < selectedMiddle.Coordinate Then
                left = selectedLeft: middle = selectedMiddle: SelectCumulativePeak = True
            End If
            Exit Function
        End If
    Next index
End Function

' Подтверждает нужную точность множителя тем же сопряжённым оператором после более точного решения только критической точки.
' Формула погрешности, материал и конечный допуск максимума не меняются; повтор расходует обычную пробу из общего бюджета.
Private Function ConfirmShoulderAccuracy(ByRef point As TBranchPoint, ByVal previousLambda As Double, _
        ByVal previousCoordinate As Double, ByVal keepRight As Boolean) As Boolean
    If Not mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, point.Uncertainty) Then
        mProbeDetails = "Не удалось оценить чувствительность множителя к невязке."
        Exit Function
    End If
    If mUsesEnvelopeControl Or Not mSolver.SupportsHybridRecovery Then ConfirmShoulderAccuracy = True: Exit Function
    Dim target As Double: target = mLambdaTolerance
    If Abs(point.Lambda - 1#) <= 0.05 Then target = GeomMin(target, NEAR_ONE_TOLERANCE)
    target = target / 16# ' Отдельные погрешности плеч занимают ограниченную часть допуска максимума; ошибка дискретизации оценивается независимо.
    Dim attempt As Long, accuracyScale As Double: accuracyScale = 1#
    Dim field() As Double, resultants() As Double, prediction As Variant, initialLambda As Double
    For attempt = 1 To 3
        If point.Uncertainty <= target Then ConfirmShoulderAccuracy = True: Exit Function
        If mProbeCount >= PROBE_LIMIT Then Exit Function
        accuracyScale = accuracyScale * (0.5 * (target / point.Uncertainty))
        If accuracyScale <= 0# Then
            mProbeDetails = "Требуемая точность критического равновесия вышла за численный диапазон."
            Exit Function
        End If
        mSolver.CopyTrial field, resultants: prediction = field: initialLambda = point.Lambda
        mSolver.DiscardStep
        mShoulderProbes = mShoulderProbes + 1
        If Not Probe(point, initialLambda, keepYieldControl:=keepRight, initialField:=prediction, _
                acceptedLambda:=previousLambda, accuracyScale:=accuracyScale) Then Exit Function
        If Not UpdateTrialCoordinate(point, previousCoordinate) Then Exit Function
        If BeyondSteelLimit() Or point.Lambda > mMaximumLambda Or (keepRight And point.Lambda <= 0#) Then
            mSolver.DiscardStep
            mProbeDetails = "Уточнённая точка плеча вышла из допустимой области материалов или нагрузки."
            Exit Function
        End If
        If Not mSolver.TrialLambdaUncertainty(mAbsTol * accuracyScale, mRelTol * accuracyScale, point.Uncertainty) Then
            mProbeDetails = "Не подтверждена чувствительность множителя после уточнения критического равновесия."
            Exit Function
        End If
    Next attempt
    ConfirmShoulderAccuracy = (point.Uncertainty <= target)
    If Not ConfirmShoulderAccuracy Then mProbeDetails = "Точность критического равновесия уточнялась, но погрешность множителя осталась выше выделенной части допуска."
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет одну штатную процедуру резерва от заданной принятой истории, не изменяя её.
Friend Function TestAuditEnvelopeDeparture(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef offset() As Double, ByRef direction() As Double, ByVal history As Collection, ByVal trace As CCoupledLoadTrace, _
        ByVal options As Variant) As Variant
    Set mSolver = solver: Set mReinforcement = reinforcement
    mOffset = offset: mDirection = direction
    Set mTrace = trace.CopyPrefix(trace.Count): Set mCurrentFields = history
    Set mMeta = New CResultMeta
    mMaxIterations = CLng(options(2)): mAbsTol = options(3): mRelTol = options(4)
    mLambdaTolerance = options(5): mMaximumLambda = options(6)
    mEnvelopeDepartureRevision = solver.AcceptedRevision: mDepartureRecoveryRevision = -1
    mMaterialRecoveryRevision = -1: mFailedProbeRevision = -1: mLastHybridRevision = -1
    mUsesFieldControl = True: mFieldControlChanges = 1
    Dim accepted() As Double: solver.CopyAccepted accepted
    ReDim mProjection(1 To UBound(accepted))
    Dim current As TBranchPoint, trial As TBranchPoint, fraction As Double, factor As Double, index As Long
    current.HistoryIndex = history.Count
    For index = 1 To trace.Count
        trace.ReadPoint index, fraction, factor
        mBestLambda = GeomMax(mBestLambda, factor)
        If fraction = 1# And factor = 1# Then mDesignHistoryIndex = index
    Next index
    trace.ReadPoint trace.Count, fraction, current.Lambda
    Dim revision As Long: revision = solver.AcceptedRevision
    Dim solved As Boolean: solved = TryRecoverEnvelopeDeparture(history, current, trial)
    If solver.AcceptedRevision <> revision Then Err.Raise 5, , "Диагностическая процедура приняла материал."
    TestAuditEnvelopeDeparture = Array(Abs(CLng(solved)), mProbeCount, trial.Lambda, current.Lambda, revision, solver.AcceptedRevision, mProbeDetails)
End Function

' При остановке выбранной меры проверяет одну знаковую дугу от текущей истории; все вызовы учитываются общим владельцем бюджета.
' Новое поле остаётся пробным: предел стали, область нагрузки, принятие материала и конечный статус проверяет основной цикл.
Private Function TryArcContinuationProbe(ByVal history As Collection, ByRef current As TBranchPoint, ByRef trial As TBranchPoint, _
        ByRef nominal As Double, ByRef increment As Double, ByRef minimum As Double) As Boolean
    If Not mSolver.UsesSpatialHistory Or mSolver.HasPrincipalBoundaries Or mDesignHistoryIndex = 0 Then Exit Function
    If history.Count < 4 Or mTrace Is Nothing Or mProbeCount + 3 > PROBE_LIMIT Then Exit Function
    If mSignedArcRevision <> mSolver.AcceptedRevision Then
        mSignedArcRevision = mSolver.AcceptedRevision: mSignedArcAttempts = 0: mSignedArcRatio = 1#
    End If
    If mSignedArcAttempts >= 4 Then Exit Function
    mSignedArcAttempts = mSignedArcAttempts + 1
    mSolver.DiscardStep
    Dim arc As New CCoupledArcContinuation, projection() As Double, solved As Boolean
    arc.Initialize mSolver, mOffset, mDirection, mAbsTol, mRelTol, mMaxIterations
    solved = arc.TryStep(history, mTrace, mSignedArcRatio, PROBE_LIMIT - mProbeCount, trial.Lambda, projection)
    mProbeCount = mProbeCount + arc.Calls: mSignedArcProbes = mSignedArcProbes + arc.Calls
    mProbeDetails = arc.Details
    If Not solved Then mSignedArcRatio = 0.5 * mSignedArcRatio: Exit Function
    Dim accepted() As Double, field() As Double, actual() As Double
    accepted = history(current.HistoryIndex): mSolver.CopyTrial field, actual
    Dim previous() As Double: previous = history(current.HistoryIndex - 1)
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, field, nominal, increment) Then
        mSolver.DiscardStep: Exit Function
    End If
    mProjection = projection: current.Coordinate = CoordinateOf(accepted): trial.Coordinate = CoordinateOf(field)
    nominal = trial.Coordinate - current.Coordinate
    If nominal <= 0# Then mSolver.DiscardStep: Exit Function
    increment = 0.5 * nominal: minimum = increment * MIN_STEP_RATIO
    mUsesEnvelopeControl = False: mUsesYieldControl = False: mUsesFieldControl = True: mUsesTransverseControl = False
    mBranchTangentControl = True: mMaterialGrowthActive = False: Erase mMaterialGrowthProjection
    mFieldControlChanges = mFieldControlChanges + 1: mLastFieldControlRevision = mSolver.AcceptedRevision
    mControlName = "пространственная мера проверенной знаковой дуги"
    If trial.Lambda < current.Lambda - LoadResolution(current.Lambda) Then
        If Not mUnverifiedDecline Then mUnverifiedDips = mUnverifiedDips + 1
        mUnverifiedDecline = True: mUnverifiedReference = GeomMax(mUnverifiedReference, mBestLambda)
    End If
    mSignedArcRecovered = mSignedArcRecovered + 1
    TryArcContinuationProbe = True
End Function

