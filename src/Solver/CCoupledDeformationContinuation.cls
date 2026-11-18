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
Private Const YIELD_RECOVERY_LIMIT As Long = 4 ' Дорогие пробы с выбором другой меры при остановке.
Private Const MIN_STEP_RATIO As Double = 0.015625
Private Const NEAR_ONE_TOLERANCE As Double = 0.0001
Private Const STEP_GROWTH As Double = 1.25
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
Private mDesignHistoryIndex As Long ' Индекс правого принятого узла, подтверждающего сочетание.
Private mLocalPeak As Double, mLocalPeakAccuracy As Double
Private mFollowingLocalPeak As Boolean, mRecoveredPeaks As Long
Private mReusedRightPoints As Long
Private mShoulderProbes As Long
Private mSteelLow As Double, mSteelHigh As Double, mSteelPoint As Long
Private mSteelEventPoint As Long, mSteelEventDetails As String
Private mProbeCount As Long
Private mYieldRecoveryProbes As Long, mYieldRecoveredPoints As Long
Private mFixedYieldRecoveryProbes As Long
Private mArcRecoveryTried As Boolean, mArcRecoveryProbes As Long
Private mUsesYieldControl As Boolean
Private mPostPeakControlName As String, mPostPeakControlPoints As Long
Private mPostPeakWidePoints As Long
Private mLimitFields As Collection
Private mLimitResultants() As Double
Private mCurrentResultants() As Double
Private mTrace As CCoupledLoadTrace, mLimitTrace As CCoupledLoadTrace
Private mCurrentFields As Collection ' Принятый путь для воспроизводимой диагностики незавершённого поиска.
Private mUsesTransverseControl As Boolean, mUnverifiedDecline As Boolean
Private mUnverifiedReference As Double, mUnverifiedDips As Long, mRecoveredDips As Long
Private mObserver As ICoupledContinuationProgress

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
    mSteelEventPoint = 0: mSteelEventDetails = vbNullString
    mYieldRecoveryProbes = 0: mYieldRecoveredPoints = 0: mFixedYieldRecoveryProbes = 0
    mArcRecoveryTried = False: mArcRecoveryProbes = 0
    mUsesYieldControl = False
    mUsesTransverseControl = False: mUnverifiedDecline = False
    mUnverifiedReference = 0#: mUnverifiedDips = 0: mRecoveredDips = 0
    mPostPeakControlName = vbNullString: mPostPeakControlPoints = 0
    mPostPeakWidePoints = 0
    mLocalPeak = 0#: mLocalPeakAccuracy = 0#
    mFollowingLocalPeak = False: mRecoveredPeaks = 0
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

    Dim attempt As Long, reinforced As Boolean, probed As Boolean, changedControl As Boolean
    Dim lastProbeFailed As Boolean, lastNoProgress As Boolean, unresolvedSteel As Boolean
    Dim targetCoordinate As Double
    Dim smoothTrial As Boolean, smoothSteps As Long, stepCeiling As Double, designAtTrial As Boolean
    Dim recoveryProjection() As Double
    Dim arcFields As Collection
    If Not reinforcement Is Nothing Then reinforced = reinforcement.HasSteel
    For attempt = 1 To PROBE_LIMIT
        If mProbeCount >= PROBE_LIMIT Then Exit For
        If attempt > STEP_LIMIT And mLocalPeak = 0# And mUnverifiedDips = 0 Then
            StopSearch "Достигнут лимит начального участка деформационного продолжения до обнаружения спада нагрузки."
            Exit Sub
        End If
        If increment < minimum Then
            If Not (lastProbeFailed Or lastNoProgress) Or unresolvedSteel Then Exit For
            If Not TryChangePostPeakControl(history, left, current, nominal, increment, minimum, lastNoProgress) Then Exit For
        End If
        targetCoordinate = current.Coordinate + increment
        If targetCoordinate <= current.Coordinate Then Exit For
        trial.Coordinate = targetCoordinate
        Set arcFields = Nothing
        If Not mArcRecoveryTried And Not mUsesYieldControl And Not mFollowingLocalPeak And mDesignHistoryIndex > 0 Then _
            Set arcFields = history ' Повтор точного lambda=1 уже принят, поэтому новая цель не перерешается.
        Dim seed As Variant, seedLambda As Double
        If PrepareSecantSeed(history, left, current, trial.Coordinate, seed, seedLambda) Then
            probed = Probe(trial, seedLambda, mYieldRecoveryProbes < YIELD_RECOVERY_LIMIT And Not mFollowingLocalPeak, _
                arcFields:=arcFields, initialField:=seed)
        Else
            probed = Probe(trial, current.Lambda, mYieldRecoveryProbes < YIELD_RECOVERY_LIMIT And Not mFollowingLocalPeak, _
                arcFields:=arcFields)
        End If
        lastProbeFailed = Not probed
        lastNoProgress = False
        changedControl = False: smoothTrial = False
        If probed Then
            If mSolver.CopyRecoveredControl(recoveryProjection) Then
                probed = TryAdoptRecoveredControl(history, current, trial, recoveryProjection, nominal, increment, minimum)
                changedControl = probed
                If probed Then targetCoordinate = trial.Coordinate Else mSolver.DiscardStep
            End If
        End If
        If probed Then
            probed = UpdateTrialCoordinate(trial, current.Coordinate)
            lastNoProgress = Not probed
            If probed And mFollowingLocalPeak And Len(mPostPeakControlName) > 0 Then _
                smoothTrial = FitsPostPeakSecant(history, left, current, trial)
        End If
        If Not probed Then
            smoothSteps = 0
            increment = increment * 0.5
        ElseIf mFollowingLocalPeak And Len(mPostPeakControlName) > 0 And increment > nominal And Not smoothTrial Then
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
        ElseIf mFollowingLocalPeak Or mUnverifiedDecline Or trial.Lambda > current.Lambda + LoadResolution(current.Lambda) Then
            If trial.Lambda <= 0# Then
                mSolver.DiscardStep
                StopSearch "После локального максимума множитель вышел из области положительной нагрузки."
                Exit Sub
            End If
            designAtTrial = False
            If mDesignHistoryIndex = 0 And current.Lambda < 1# And trial.Lambda >= 1# And mProbeCount + 2 <= PROBE_LIMIT Then
                If Not RepeatDesignTrial(current, trial, targetCoordinate, designAtTrial) Then Exit Sub
            End If
            AcceptPoint history, trial
            If designAtTrial And trial.Lambda >= 1# Then mDesignHistoryIndex = trial.HistoryIndex
            unresolvedSteel = False
            If changedControl Then mYieldRecoveredPoints = mYieldRecoveredPoints + 1
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
            If smoothTrial And mFollowingLocalPeak Then smoothSteps = smoothSteps + 1 Else smoothSteps = 0
            stepCeiling = nominal
            If smoothSteps >= 2 Then stepCeiling = POST_PEAK_STEP_RATIO * nominal
            If smoothSteps >= 6 Then stepCeiling = POST_PEAK_EXTENDED_RATIO * nominal
            If mUsesTransverseControl Then stepCeiling = TRANSVERSE_STEP_RATIO * nominal
            If increment > stepCeiling Then increment = stepCeiling
            If increment < stepCeiling / STEP_GROWTH Then increment = increment * STEP_GROWTH
        ElseIf trial.Lambda < current.Lambda - LoadResolution(current.Lambda) Then
            Dim declinedField() As Double, declinedResultants() As Double
            If reinforced And mUsesTransverseControl Then mSolver.CopyTrial declinedField, declinedResultants
            mSolver.DiscardStep
            If VerifyPeak(history, left, current, trial, reinforced) Then
                If Not reinforced Then Exit Sub
                ResumeAfterLocalPeak history, current, trial
                increment = nominal * 0.25
                minimum = increment * MIN_STEP_RATIO
            Else
                If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
                    StopSearch "Не удалось восстановить историю возрастающей ветви."
                    Exit Sub
                End If
                If reinforced And mUsesTransverseControl Then
                    Dim repeatSeed As Variant: repeatSeed = declinedField
                    If Probe(trial, trial.Lambda, initialField:=repeatSeed) Then
                        If UpdateTrialCoordinate(trial, current.Coordinate) And Not BeyondSteelLimit() And trial.Lambda > 0# Then
                            mUnverifiedDecline = True: mUnverifiedReference = mBestLambda
                            mUnverifiedDips = mUnverifiedDips + 1
                            AcceptPoint history, trial
                            left = current: current = trial
                        Else
                            mSolver.DiscardStep: increment = increment * 0.5
                        End If
                    Else
                        lastProbeFailed = True
                        increment = increment * 0.5
                    End If
                Else
                    increment = increment * 0.5
                End If
            End If
        Else
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

' Строит только начальное приближение из двух собственных полей; материал остаётся в последней принятой истории.
' Координаты множителей связей не экстраполируются, нагрузка сохраняет заданные постоянную и переменную части.
Private Function PrepareSecantSeed(ByVal history As Collection, ByRef left As TBranchPoint, ByRef current As TBranchPoint, _
        ByVal target As Double, ByRef seed As Variant, ByRef factor As Double) As Boolean
    If left.HistoryIndex < 1 Or left.HistoryIndex >= current.HistoryIndex Or current.HistoryIndex <> history.Count Then Exit Function
    Dim previous() As Double, accepted() As Double, field() As Double, oldControl As Double, control As Double, ratio As Double, i As Long
    previous = history(left.HistoryIndex): accepted = history(current.HistoryIndex)
    Dim previousLambda As Double: previousLambda = left.Lambda
    If Not mTrace Is Nothing Then
        Dim fraction As Double
        previous = history(current.HistoryIndex - 1)
        mTrace.ReadPoint current.HistoryIndex - 1, fraction, previousLambda
        If fraction <> 1# Then Exit Function
    End If
    oldControl = CoordinateOf(previous): control = CoordinateOf(accepted)
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
    mUsesTransverseControl = False
    left.Coordinate = previousCoordinate: current.Coordinate = currentCoordinate
    nominal = spacing: increment = 0.5 * spacing: minimum = increment * MIN_STEP_RATIO
    TryChangePostPeakControl = True
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
    coordinate = CoordinateOf(field)
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

' Передаёт подтверждённый максимум ветви без арматуры и оценку точности.
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

' Решает пробу при заданной мере. Проверки продвижения и касательной
' выполняются после выбора действующей проекции.
Private Function Probe(ByRef point As TBranchPoint, ByVal initialLambda As Double, _
        Optional ByVal recoverYieldControl As Boolean = False, Optional ByVal keepYieldControl As Boolean = False, _
        Optional ByVal arcFields As Collection = Nothing, Optional ByRef initialField As Variant) As Boolean
    If mProbeCount >= PROBE_LIMIT Then Exit Function
    mProbeCount = mProbeCount + 1
    NotifyProgress
    point.Uncertainty = 0#
    Dim solved As Boolean, recoverSameControl As Boolean
    recoverSameControl = keepYieldControl And mUsesYieldControl And mFixedYieldRecoveryProbes < FIXED_YIELD_RECOVERY_LIMIT
    solved = mSolver.TryDeformationStep(mOffset, mDirection, mProjection, point.Coordinate, initialLambda, _
        point.Lambda, mAbsTol, mRelTol, mMaxIterations, recoverYieldControl, mUsesYieldControl, recoverSameControl, initialField:=initialField)
    If mSolver.SameControlRecoveryTried Then
        mFixedYieldRecoveryProbes = mFixedYieldRecoveryProbes + 1
    ElseIf mSolver.YieldRecoveryTried Then
        mYieldRecoveryProbes = mYieldRecoveryProbes + 1
    End If
    If Not solved And Not arcFields Is Nothing And mProbeCount + 2 <= PROBE_LIMIT Then
        Dim extraProbes As Long
        solved = mSolver.TryRecoverDeformationByArc(arcFields, mTrace, initialLambda, mAbsTol, mRelTol, _
            mMaxIterations, point.Lambda, extraProbes)
        If extraProbes > 0 Then
            mArcRecoveryTried = True: mArcRecoveryProbes = extraProbes
            mProbeCount = mProbeCount + extraProbes
        End If
    End If
    If Not solved Then
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
        Exit Function
    End If
    Probe = True
End Function

' Передаёт наблюдателю только независимые копии принятого пути, без пробных состояний и ввода-вывода ядра.
' Полный снимок сохраняет смысл после замены хвоста при уточнении локального максимума.
Private Sub NotifyProgress()
    If mObserver Is Nothing Then Exit Sub
    Dim fields As Collection, trace As CCoupledLoadTrace, sourceTrace As CCoupledLoadTrace, meta As CResultMeta
    Set fields = AcceptedFields: Set sourceTrace = SavedTrace
    If fields Is Nothing Or sourceTrace Is Nothing Then Exit Sub
    Set trace = sourceTrace.CopyPrefix(sourceTrace.Count)
    Set meta = mMeta.Clone: meta.DiagnosticDetails = Details()
    mObserver.Report fields, trace, meta
End Sub

' Принимает новую координату на возрастающем участке. Последнее
' принятое и найденное поля образуют пару в одной мере. Проверки
' плеч и предела арматуры продолжают использовать эту же меру.
Private Function TryAdoptRecoveredControl(ByVal history As Collection, ByRef current As TBranchPoint, _
        ByRef trial As TBranchPoint, ByRef projection() As Double, ByRef nominal As Double, _
        ByRef increment As Double, ByRef minimum As Double) As Boolean
    If trial.Lambda <= current.Lambda + LoadResolution(current.Lambda) Or trial.Lambda > mMaximumLambda Then Exit Function
    If BeyondSteelLimit() Then Exit Function
    Dim previousField() As Double, acceptedField() As Double, field() As Double, resultants() As Double
    previousField = history(current.HistoryIndex - 1)
    acceptedField = history(current.HistoryIndex)
    mSolver.CopyTrial field, resultants
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
    mUsesTransverseControl = False
    current.Coordinate = oldCoordinate: trial.Coordinate = newCoordinate
    increment = 0.5 * nominal: minimum = increment * MIN_STEP_RATIO
    TryAdoptRecoveredControl = True
End Function

' Продвигает принятую материальную историю и фиксирует индекс её поля.
Private Sub AcceptPoint(ByVal history As Collection, ByRef point As TBranchPoint)
    Dim field() As Double
    mSolver.CopyTrial field, mCurrentResultants
    mSolver.AcceptStep
    history.Add field
    If Not mTrace Is Nothing Then mTrace.Add 1#, point.Lambda
    point.HistoryIndex = history.Count
    mBestBeforeCurrent = mBestLambda
    mBestLambda = GeomMax(mBestLambda, point.Lambda)
End Sub

' Повторяет правое поле после силовой пробы. Точка lambda = 1 должна
' принадлежать фактическому интервалу именно принятого повтора.
Private Function RepeatDesignTrial(ByRef current As TBranchPoint, ByRef trial As TBranchPoint, _
        ByVal targetCoordinate As Double, ByRef designVerified As Boolean) As Boolean
    Dim designCoordinate As Double
    designVerified = VerifyDesignLoad(current.Coordinate, trial.Coordinate, designCoordinate)
    trial.Coordinate = targetCoordinate
    If Not Probe(trial, current.Lambda, keepYieldControl:=True) Then
        StopSearch "Не удалось повторить пробу после проверки заданных усилий."
        Exit Function
    End If
    If Not UpdateTrialCoordinate(trial, current.Coordinate) Then
        StopSearch "Повторная проба после проверки заданных усилий не продвинулась по деформационной ветви."
        Exit Function
    End If
    If trial.Lambda > mMaximumLambda Then
        StopAtSearchBound
        Exit Function
    End If
    If trial.Lambda <= 0# Or (Not mFollowingLocalPeak And trial.Lambda <= current.Lambda + LoadResolution(current.Lambda)) Then
        mSolver.DiscardStep
        StopSearch "Повторная проба не подтвердила допустимое изменение множителя нагрузки."
        Exit Function
    End If
    If BeyondSteelLimit() Then
        mSolver.DiscardStep
        StopSearch "Повторная проба не подтвердила допустимость арматуры."
        Exit Function
    End If
    If designVerified Then designVerified = (designCoordinate >= current.Coordinate And designCoordinate <= trial.Coordinate)
    RepeatDesignTrial = True
End Function

' Проверяет точное lambda = 1 от левого плеча перед продвижением истории.
' Контрольная мера должна попасть внутрь проверяемого участка.
Private Function VerifyDesignLoad(ByVal leftCoordinate As Double, ByVal rightCoordinate As Double, _
        Optional ByRef coordinate As Double = 0#) As Boolean
    If mProbeCount >= PROBE_LIMIT Then Exit Function
    mProbeCount = mProbeCount + 1
    Dim load(1 To COMPONENT_COUNT) As Double, i As Long
    For i = 1 To COMPONENT_COUNT: load(i) = mOffset(i) + mDirection(i): Next i
    If Not mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), mAbsTol, mRelTol, mMaxIterations) Then Exit Function
    Dim field() As Double, resultants() As Double
    mSolver.CopyTrial field, resultants
    coordinate = CoordinateOf(field)
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
    If VerifyDesignLoad(CoordinateOf(leftField), CoordinateOf(rightField)) Then designIndex = leftIndex + 1
    VerifyRefinedDesign = mSolver.ReplayAcceptedFields(fields, fields.Count)
End Function


' Повторяет оба плеча от одного принятого левого состояния.
' Для предела без арматуры уточняет максимум до допуска. Для
' локального максимума железобетона сохраняет проверенное правое плечо.
Private Function VerifyPeak(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint, ByVal keepRight As Boolean) As Boolean
    If middle.Lambda <= left.Lambda + LoadResolution(middle.Lambda) Then Exit Function
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
    If keepRight Then
        If Not ReplayShoulders(history, checkedLeft, middle, checkedRight, 2, second, sampling, refinedFields, actual, True, retainedMaximum, refinedDesignLeft, _
                retainedPreviousMaximum) Then Exit Function
        accuracy = Abs(middle.Lambda - second.Lambda)
        If sampling > accuracy Then accuracy = sampling
        accuracy = accuracy + second.Uncertainty
        If mBestLambda - second.Lambda > accuracy Then accuracy = mBestLambda - second.Lambda
    Else
        If Not ReplayShoulders(history, left, middle, right, 2, first, sampling) Then Exit Function
        divisions = 4
        Do
            If Not ReplayShoulders(history, left, middle, right, divisions, second, sampling, refinedFields, actual, _
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
    mBestBeforeCurrent = retainedPreviousMaximum
    mBestLambda = retainedMaximum ' Равновесие удалённого крупного шага остаётся только в оценке погрешности.
    mMeta.SetResult rsSuccess, rcCheckPassed, rkCapacity, _
        "Подтверждён первый максимум прослеженной ветви.", Details() & " | Оценка погрешности lambda: " & CStr(accuracy) & _
        " | Координаты плеч: " & CStr(shoulderStart) & ", " & CStr(right.Coordinate)
    VerifyPeak = True
End Function

' Воспроизводит левые принятые поля, затем рассчитывает более мелкие
' шаги только вперёд. Сравнивает пик с обоими концами с учётом невязки.
' При сохранении спада передаёт последнюю пару точек уточнённой истории.
Private Function ReplayShoulders(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint, ByVal divisions As Long, _
        ByRef peak As TBranchPoint, ByRef sampling As Double, Optional ByRef limitFields As Collection = Nothing, _
        Optional ByRef limitResultants As Variant, Optional ByVal keepRight As Boolean = False, _
        Optional ByRef retainedMaximum As Double = 0#, Optional ByRef designLeftIndex As Long = 0, _
        Optional ByRef retainedPreviousMaximum As Double = 0#) As Boolean
    If Not mSolver.ReplayAcceptedFields(history, left.HistoryIndex) Then Exit Function
    Dim start As Double, finish As Double, previousLambda As Double, point As TBranchPoint
    Dim samples() As TBranchPoint, shoulder As Long, i As Long, index As Long
    Dim saveFields As Boolean, refined As Collection, integrated As Collection
    saveFields = Not IsMissing(limitResultants)
    If saveFields Then
        Set refined = New Collection: Set integrated = New Collection
    End If
    Dim field() As Double, resultants() As Double
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
            If Not Probe(point, previousLambda, keepYieldControl:=keepRight) Then Exit Function
            If Not UpdateTrialCoordinate(point, samples(index).Coordinate) Then Exit Function
            If BeyondSteelLimit() Or point.Lambda > mMaximumLambda Then
                mSolver.DiscardStep
                Exit Function
            End If
            If Not mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, point.Uncertainty) Then
                mSolver.DiscardStep
                mProbeDetails = "Не удалось оценить чувствительность множителя к невязке."
                Exit Function
            End If
            If keepRight And point.Lambda <= 0# Then
                mSolver.DiscardStep
                Exit Function
            End If
            If saveFields Then
                mSolver.CopyTrial field, resultants
                refined.Add field: integrated.Add resultants
            End If
            mSolver.AcceptStep
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
    If peakIndex = 0 Then Exit Function

    peak = samples(peakIndex)
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
        retainedMaximum = PrefixMaximumLambda(left.HistoryIndex): designLeftIndex = 0 ' Удалённые промежуточные поля не остаются подтверждённой границей нагрузки.
        For j = 1 To fieldCount
            limitFields.Add refined.Item(j)
            retainedPreviousMaximum = retainedMaximum
            retainedMaximum = GeomMax(retainedMaximum, samples(j).Lambda)
            If designLeftIndex = 0 And samples(j - 1).Lambda < 1# And samples(j).Lambda >= 1# Then _
                designLeftIndex = left.HistoryIndex + j - 1
        Next j
        If Not mTrace Is Nothing Then
            Set mLimitTrace = mTrace.CopyPrefix(left.HistoryIndex)
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

' Вычисляет подтверждённую нагрузку только по сохраняемому префиксу, включая максимум до локального спада.
' Без журнала левое поле непосредственно предшествует текущему и используется сохранённая общая граница.
Private Function PrefixMaximumLambda(ByVal lastIndex As Long) As Double
    If mTrace Is Nothing Then
        PrefixMaximumLambda = mBestBeforeCurrent
        Exit Function
    End If
    Dim i As Long, fraction As Double, factor As Double
    For i = 1 To lastIndex
        mTrace.ReadPoint i, fraction, factor
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
    If Len(reason) > 0 Then
        If Not mTrace Is Nothing Then
            For Each factor In multipliers: mTrace.Add 1#, factor: Next factor
        End If
        StopSearch reason
    Else
        mSteelLow = low: mSteelHigh = high: mSteelPoint = eventState.CriticalPoint
        Set mLimitFields = New Collection
        Dim i As Long
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
        If Not Probe(point, low.Lambda, keepYieldControl:=True) Then Exit Function
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
            If Not Probe(high, low.Lambda, keepYieldControl:=True) Then Exit Function
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
End Sub

' Записывает выбранную управляемую величину и область проверенной ветви.
Private Function Details() As String
    Details = "Мера деформации: " & mControlName & "; наибольший допустимый lambda: " & CStr(mBestLambda) & _
        "; равновесие при lambda = 1 подтверждено: " & CStr(mDesignHistoryIndex > 0) & "; проб: " & CStr(mProbeCount)
    If mUnverifiedDips > 0 Then Details = Details & "; неподтверждённых спадов: " & CStr(mUnverifiedDips) & _
        "; восстановлений после них: " & CStr(mRecoveredDips)
    If mUnverifiedDecline Then Details = Details & "; текущий спад не является подтверждённым физическим пределом"
    If mYieldRecoveryProbes > 0 Then Details = Details & "; проб смены меры: " & CStr(mYieldRecoveryProbes) & _
        "; принято после смены меры: " & CStr(mYieldRecoveredPoints)
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
