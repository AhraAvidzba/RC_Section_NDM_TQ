VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledNewtonSolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Нелинейное равновесие связанного сечения при заданном силовом винте
' ==========================================================================
' Использует те же геометрические операторы, что и упругий расчёт.
' Каждая проба демпфирования пересчитывает весь остаток. С боковыми
' историями градиент напряжений определяется по двум продольным точкам.
Implements ICoupledEquationSystem

Private Const COMPONENT_COUNT As Long = 6
Private Const GAUGE_COUNT As Long = 4
Private Const MIN_FRACTION As Double = 0.0078125
Private Const STEEL_MIN_FRACTION As Double = 0.000244140625 ' Дополнительные пробы силового шага после смены ветви стали.
Private Const MAX_STEEL_FIELD_TANGENTS As Long = 2 ' Ограничивает дорогие разностные сборки одним силовым шагом.
Private Const MIN_PROGRESS As Double = 0.0001
Private Const SMALL_STRAIN_MAX_NORM As Double = 2# ' Локальная поправка завершает состояние вблизи допуска.
Private Const SMALL_STRAIN_RADIUS As Double = 2.5 ' Радиус через силовую невязку в масштабированных столбцах.
Private Const SMALL_STRAIN_FORCE_FRACTION As Double = 0.92 ' Запас до допуска для нелинейной пробы.
Private Const SPATIAL_HALF_STEP As Double = 0.001 ' Полушаг секущей по безразмерной координате z/L.
Private Const GRADIENT_DIFFERENCE_STEP As Double = 0.0000006 ' Амплитуда возмущения максимальной компоненты деформации.
Private Const LOCAL_GRADIENT_DIFFERENCE_STEP As Double = 0.00000000025 ' Амплитуда пробы внутри перехода разностной материальной касательной.
Private Const GRADIENT_TANGENT_CHANGE_RATIO As Double = 0.1 ' Изменение касательной в долях опорного модуля для выбора линеаризации.
Private Const YIELD_TRACE_ITERATIONS As Long = 3
Private Const YIELD_DAMPING_STEPS As Long = 8
Private Const YIELD_ROOT_TOLERANCE As Double = 0.000000000002 ' МПа в пробном критерии бетона.
Private Const PRINCIPAL_EVENT_TOLERANCE As Double = 0.000000001 ' МПа: общая точка проверяется строже штатного равновесия.
Private Const PRINCIPAL_EVENT_RADIUS As Double = 0.25 ' Радиус локальной доводки в долях предыдущего механического приращения.
Private Const SUBSPACE_ITERATIONS As Long = 8 ' Прогноз расходует общий бюджет шага, оставляя итерацию полной системе.
Private Const HYBRID_DIFFERENCE_STEP As Double = 0.000001 ' Относительный шаг производной полной системы в подготовленных координатах.
Private Const HYBRID_STRAIN_SCALE As Double = 0.00001 ' Единичное перемещение численной координаты ограничено этой компонентой деформации.
Private Const HYBRID_AUXILIARY_SCALE As Double = 0.01 ' Масштаб множителей связей и lambda, без изменения их уравнений.

Private Type TYieldStation
    Point As Long
    Side As Long
    PrincipalIndex As Long ' 0 — поверхность CDP; 1..3 — главные напряжения прогноза ASD от старшего к младшему.
    FromValue As Double
    ToValue As Double
End Type

Private Type TSolveEffort
    Iterations As Long
    Evaluations As Long
    GradientBuilds As Long
    LocalGradientBuilds As Long
    SteelBranchBuilds As Long
    SteelFieldTangents As Long
    MaterialFailures As Long
    LastFailedPoint As Long
    LastFailedSide As Long
    HybridSteps As Long
    HybridMatrices As Long
    HybridEvaluations As Long
End Type

Private mSection As CCoupledElasticState
Private mResponse As ICoupledPointResponse
Private mPlusResponse As ICoupledPointResponse, mMinusResponse As ICoupledPointResponse
Private mUseSpatialHistory As Boolean
Private mLastFailedSide As Long ' 0 — центральная точка, +1 и -1 — продольные точки.
Private mSystem As CCoupledLinearSystem
Private mSteelCorrector As CCoupledSteelCorrection
Private mFieldSubspace As CCoupledFieldSubspace
Private mSubspaceChecked As Boolean, mSubspaceRecoveryActive As Boolean
Private mMeta As CResultMeta
Private mAccepted() As Double, mCandidate() As Double, mTrialResultants() As Double
Private mPointStrain(1 To COMPONENT_COUNT) As Double
Private mStress(1 To COMPONENT_COUNT) As Double
Private mTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mPlusTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mMinusTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mGradientTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mSpatialGradient(1 To COMPONENT_COUNT) As Double
Private mSlopeTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mPlusStrain(1 To COMPONENT_COUNT) As Double, mMinusStrain(1 To COMPONENT_COUNT) As Double
Private mPlusStress(1 To COMPONENT_COUNT) As Double, mMinusStress(1 To COMPONENT_COUNT) As Double
Private mMatrix() As Double, mExternal() As Double, mResidual() As Double
Private mDirectProofMatrix As Boolean ' Матрица собрана непосредственно в принятом силовом поле.
Private mCandidateReady As Boolean
Private mYieldRecoveryTried As Boolean, mSameControlRecoveryTried As Boolean
Private mRecoveredControlReady As Boolean
Private mSameControlInitialFailure As String
Private mRecoveredProjection() As Double
Private mConsistentControl As Boolean
Private mControlActive As Boolean
Private mControlProjection() As Double, mTargetControl As Double
Private mArcLoadProjection As Double, mCurrentControlLambda As Double
Private mWrongControlDirection As Boolean
Private mHasRejectedDirection As Boolean, mRejectedDirection() As Double
Private mAcceptedRevision As Long ' Ревизия всей принятой истории, включая восстановление и сброс.
Private mTransitionFrom() As Double, mTransitionTo() As Double, mTransitionRevision As Long
Private mHasTransitionSegment As Boolean, mPrincipalIncomingCandidate As Boolean
Private mPrincipalTransition As CCoupledPrincipalTransition
Private mPrincipalDesignVerified As Boolean
Private mPrincipalDesignField() As Double, mPrincipalDesignResultants() As Double
Private mPrincipalOutgoingCandidate As Boolean, mDepartureRevision As Long
Private mPrincipalDeparture As CCoupledPrincipalTransition
Private mPrincipalBaseLoad() As Double, mPrincipalLoadDirection() As Double
Private mBaseLoad() As Double, mLoadDirection() As Double
Private mBaseEquations() As Double, mDirectionEquations() As Double
Private mTrialLambda As Double
Private mDegreeCount As Long, mPointCount As Long, mGaugeStart As Long
Private mIterations As Long, mEvaluations As Long, mMaterialPoint As Long
Private mMaterialFailures As Long, mLastFailedPoint As Long
Private mGradientBuilds As Long
Private mLocalGradientBuilds As Long
Private mSteelBranchBuilds As Long, mSteelFieldTangents As Long
Private mReferenceModulus As Double, mLastNorm As Double, mMinimumPivot As Double
Private mFailureCode As ESolverFailureCode
Private mHybridActive As Boolean, mHybridScalesReady As Boolean
Private mHybridOriginRevision As Long, mHybridSeed() As Double, mHybridScales() As Double
Private mHybridSteps As Long, mHybridMatrices As Long, mHybridEvaluations As Long

' Подключает сечение и независимую центральную историю. Две отдельные
' боковые истории включают секущий продольный градиент напряжений.
' Модуль в МПа нужен для масштабирования уравнений перемещений.
Friend Sub Initialize(ByVal section As CCoupledElasticState, ByVal response As ICoupledPointResponse, _
        ByVal referenceModulus As Double, Optional ByVal plusSide As ICoupledPointResponse = Nothing, _
        Optional ByVal minusSide As ICoupledPointResponse = Nothing)
    If section Is Nothing Or response Is Nothing Or referenceModulus <= 0# Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Для равновесия требуется подготовленное сечение, материал и положительный модуль."
    If (plusSide Is Nothing) Xor (minusSide Is Nothing) Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Для секущего градиента требуются обе продольные истории."
    If Not plusSide Is Nothing Then
        If (plusSide Is minusSide) Or (plusSide Is response) Or (minusSide Is response) Then _
            Err.Raise 5, "CCoupledNewtonSolver", "Для секущего градиента требуются три независимые материальные истории."
    End If
    mDegreeCount = section.DegreeCount
    mPointCount = section.PointCount
    If mDegreeCount <= GAUGE_COUNT Or mPointCount < 1 Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Геометрические операторы связанного сечения не подготовлены."
    mGaugeStart = mDegreeCount - GAUGE_COUNT + 1
    Set mSection = section: Set mResponse = response
    Set mPlusResponse = plusSide: Set mMinusResponse = minusSide
    mUseSpatialHistory = Not (plusSide Is Nothing)
    Set mSystem = New CCoupledLinearSystem: Set mMeta = New CResultMeta
    Set mFieldSubspace = Nothing: mSubspaceChecked = False: mSubspaceRecoveryActive = False
    Set mSteelCorrector = Nothing
    If mUseSpatialHistory And section.ConcretePointCount < mPointCount Then
        If TypeOf response Is CCoupledReinforcedResponse And TypeOf plusSide Is CCoupledReinforcedResponse And _
                TypeOf minusSide Is CCoupledReinforcedResponse Then
            Dim center As CCoupledReinforcedResponse, positive As CCoupledReinforcedResponse, negative As CCoupledReinforcedResponse
            Set center = response: Set positive = plusSide: Set negative = minusSide
            Set mSteelCorrector = New CCoupledSteelCorrection
            mSteelCorrector.Initialize section, center, positive, negative, SPATIAL_HALF_STEP
        End If
    End If
    mReferenceModulus = referenceModulus
    mHybridScalesReady = False: Erase mHybridScales
    ResetHistory
End Sub

' Начинает независимую траекторию. Принятое поле и история других LC здесь не используются.
Friend Sub ResetHistory()
    mResponse.ClearHistory
    If mUseSpatialHistory Then
        mPlusResponse.ClearHistory
        mMinusResponse.ClearHistory
    End If
    ReDim mAccepted(1 To mDegreeCount)
    mAcceptedRevision = mAcceptedRevision + 1
    mCandidateReady = False
    mHasRejectedDirection = False: Erase mRejectedDirection
    mHasTransitionSegment = False: Set mPrincipalTransition = Nothing: mPrincipalIncomingCandidate = False
    Set mPrincipalDeparture = Nothing: mPrincipalOutgoingCandidate = False: mPrincipalDesignVerified = False
    mMeta.Clear
End Sub

' Ищет пробное равновесие от принятого поля или заданного прогноза. Моменты относятся
' к численному началу сечения; силы в Н, моменты в Н·мм.
' Само решение не публикует материальную историю.
Friend Function TryStep(ByVal force As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal torque As Double, ByVal qx As Double, ByVal qy As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        Optional ByRef initialField As Variant, Optional ByVal retainDirectProof As Boolean = False, _
        Optional ByRef fallbackField As Variant) As Boolean
    Dim load(1 To COMPONENT_COUNT) As Double
    load(1) = force: load(2) = mx: load(3) = my
    load(4) = torque: load(5) = qx: load(6) = qy
    mControlActive = False
    mConsistentControl = False
    mArcLoadProjection = 0#
    mSection.LoadEquations force, mx, my, torque, qx, qy, mExternal
    TryStep = SolveStep(load, 0#, absoluteTolerance, relativeTolerance, maxIterations, _
        initialField:=initialField, retainDirectProof:=retainDirectProof)
    If Not TryStep Then TryStep = TrySubspaceForceRecovery(load, absoluteTolerance, relativeTolerance, maxIterations, _
        initialField:=initialField, retainDirectProof:=retainDirectProof)
    If Not TryStep And Not IsMissing(fallbackField) Then TryStep = TryAlternativeForceField(load, fallbackField, _
        absoluteTolerance, relativeTolerance, maxIterations, retainDirectProof)
End Function

' Повторяет неудачную силовую пробу с независимым прогнозом только за оставшиеся итерации.
' Прогноз меняет начальное поле; нагрузка, принятая история и критерии равновесия сохраняются.
Private Function TryAlternativeForceField(ByRef load() As Double, ByRef prediction As Variant, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        ByVal retainDirectProof As Boolean) As Boolean
    If mCandidateReady Or mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
    If mMeta.ResultCode <> rcNumericalFailure And mMeta.ResultCode <> rcSingularTangent Then Exit Function
    Dim remaining As Long: remaining = maxIterations - mIterations
    If remaining < 1 Then Exit Function
    Dim beforeRetry As TSolveEffort, original As CResultMeta
    ReadSolveEffort beforeRetry: Set original = mMeta.Clone
    TryAlternativeForceField = SolveStep(load, 0#, absoluteTolerance, relativeTolerance, remaining, _
        initialField:=prediction, retainDirectProof:=retainDirectProof)
    If Not TryAlternativeForceField Then TryAlternativeForceField = TrySubspaceForceRecovery(load, _
        absoluteTolerance, relativeTolerance, remaining, initialField:=prediction, retainDirectProof:=retainDirectProof)
    AddSolveEffort beforeRetry
    Dim details As String
    If TryAlternativeForceField Then details = "подтверждён межузловой прогноз" Else details = "межузловой прогноз не принят"
    mMeta.DiagnosticDetails = DiagnosticSummary() & "; " & details & " | Исходная силовая проба: " & original.DiagnosticDetails
End Function

' Сообщает, доступна ли согласованная касательная для пространственной истории материалов.
Friend Property Get UsesSpatialHistory() As Boolean
    UsesSpatialHistory = mUseSpatialHistory
End Property

' Возвращает число механических координат поля перед множителями связей.
Friend Property Get PhysicalDegreeCount() As Long
    PhysicalDegreeCount = mGaugeStart - 1
End Property

' Решает равновесие при управлении одной линейной мерой деформации.
' Множитель нагрузки определяется вместе с полем и возвращается только
' для сошедшейся пробы. Материал остаётся в принятом ранее состоянии.
' После смены пространственной меры можно разрешить малую поправку CDP;
' полный разностный резерв включается явно после исчерпания деления шага.
' Оба резерва не применяются к псевдодуге и не меняют допуск принятия.
Friend Function TryDeformationStep(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef controlProjection() As Double, ByVal targetControl As Double, ByVal initialLambda As Double, _
        ByRef trialLambda As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maxIterations As Long, Optional ByVal recoverYieldControl As Boolean = False, _
        Optional ByVal consistentControl As Boolean = False, Optional ByVal recoverSameControl As Boolean = False, _
        Optional ByRef initialField As Variant, Optional ByVal loadProjection As Double = 0#, _
        Optional ByVal allowSmallCorrection As Boolean = False, Optional ByVal allowHybridRecovery As Boolean = False) As Boolean
    If LBound(baseLoad) <> 1 Or UBound(baseLoad) <> COMPONENT_COUNT Or _
            LBound(direction) <> 1 Or UBound(direction) <> COMPONENT_COUNT Or _
            LBound(controlProjection) <> 1 Or UBound(controlProjection) <> mDegreeCount Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Размер траектории или меры управляемой деформации не совпадает с решателем."
    Dim i As Long, hasDirection As Boolean, hasControl As Boolean
    For i = 1 To COMPONENT_COUNT
        If direction(i) <> 0# Then hasDirection = True
    Next i
    For i = 1 To mGaugeStart - 1
        If controlProjection(i) <> 0# Then hasControl = True
    Next i
    If Not hasDirection Or Not hasControl Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Для управления деформацией задайте ненулевую нагрузку и меру деформации."
    If loadProjection <> 0# And (recoverYieldControl Or recoverSameControl) Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Псевдодуга не использует восстановление со сменой ветви."
    ConfigureControlledLoad baseLoad, direction
    mArcLoadProjection = loadProjection
    mControlProjection = controlProjection: mTargetControl = targetControl
    mControlActive = True
    mConsistentControl = consistentControl And mUseSpatialHistory
    Dim load(1 To COMPONENT_COUNT) As Double
    TryDeformationStep = SolveStep(load, initialLambda, absoluteTolerance, relativeTolerance, maxIterations, _
        recoverYieldControl, initialField:=initialField, recoverSameControl:=recoverSameControl, _
        allowSmallCorrection:=(allowSmallCorrection And mUseSpatialHistory And Not HasPrincipalBoundaries And loadProjection = 0#), _
        allowHybridRecovery:=allowHybridRecovery)
    If TryDeformationStep Then trialLambda = mTrialLambda
End Function

' Готовит неизменную и масштабируемую части одной траектории для линейной либо материальной меры.
Private Sub ConfigureControlledLoad(ByRef baseLoad() As Double, ByRef direction() As Double)
    mBaseLoad = baseLoad: mLoadDirection = direction
    mSection.LoadEquations baseLoad(1), baseLoad(2), baseLoad(3), baseLoad(4), baseLoad(5), baseLoad(6), mBaseEquations
    mSection.LoadEquations direction(1), direction(2), direction(3), direction(4), direction(5), direction(6), mDirectionEquations
    ReDim mExternal(1 To mDegreeCount)
End Sub

' Проверяет короткую псевдодугу после измеренного расхождения касательной.
' Обе пробы используют принятую историю. Публикуется только проба при исходной мере.
Friend Function TryRecoverDeformationByArc(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, _
        ByVal acceptedLambda As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maxIterations As Long, ByRef trialLambda As Double, ByRef extraProbes As Long) As Boolean
    extraProbes = 0
    If mCandidateReady Or Not mControlActive Or Not mUseSpatialHistory Or mConsistentControl Then Exit Function
    If mYieldRecoveryTried Or mSameControlRecoveryTried Or Not mWrongControlDirection Then Exit Function
    If mMeta.InternalStatus <> rsNumericalFailure Or mMeta.ResultCode <> rcNumericalFailure Then Exit Function
    If maxIterations - mIterations <= 1 Then Exit Function

    Dim coefficients() As Double, projection() As Double, predicted() As Double
    Dim target As Double, loadProjection As Double, loadMetric As Double, radius As Double, factor As Double
    If Not PrepareArcMeasure(fields, trace, acceptedLambda, coefficients, projection, predicted, target, _
            loadProjection, loadMetric, radius, factor) Then Exit Function

    Dim originalMeta As CResultMeta, originalProjection() As Double, originalTarget As Double
    Dim base() As Double, direction() As Double, external() As Double, residual() As Double
    Dim originalResultants() As Double, resultants() As Double, matrix() As Double, field() As Double
    Dim originalFactor As Double, norm As Double, pivot As Double, failureCode As ESolverFailureCode
    Dim materialPoint As Long, lastPoint As Long, lastSide As Long
    Set originalMeta = mMeta.Clone
    originalProjection = mControlProjection: originalTarget = mTargetControl
    base = mBaseLoad: direction = mLoadDirection
    external = mExternal: residual = mResidual: originalResultants = mTrialResultants: matrix = mMatrix
    originalFactor = mCurrentControlLambda
    norm = mLastNorm: pivot = mMinimumPivot: failureCode = mFailureCode
    materialPoint = mMaterialPoint: lastPoint = mLastFailedPoint: lastSide = mLastFailedSide

    Dim initialIterations As Long, initialEvaluations As Long, materialFailures As Long
    Dim gradientBuilds As Long, localGradientBuilds As Long, guideIterations As Long, guideEvaluations As Long
    initialIterations = mIterations: initialEvaluations = mEvaluations: materialFailures = mMaterialFailures
    gradientBuilds = mGradientBuilds: localGradientBuilds = mLocalGradientBuilds
    Dim seed As Variant, guideFactor As Double, solved As Boolean, label As String, phaseDetails As String
    Dim errorNumber As Long, errorSource As String, errorText As String
    label = "псевдодуга"
    seed = predicted
    On Error GoTo InterruptedArc
    extraProbes = 1
    solved = TryDeformationStep(base, direction, projection, target, factor, guideFactor, _
        absoluteTolerance, relativeTolerance, maxIterations - initialIterations, initialField:=seed, loadProjection:=loadProjection)
    If Not solved Then GoTo RejectArc
    CopyTrial field, resultants
    If Not ArcFieldIsLocal(coefficients, projection, loadProjection, loadMetric, radius, field, acceptedLambda, guideFactor) Then
        label = label & " вне локальной окрестности"
        GoTo RejectArc
    End If
    If maxIterations - initialIterations - mIterations < 1 Then
        label = label & " исчерпала остаток итераций"
        GoTo RejectArc
    End If

    guideIterations = mIterations: guideEvaluations = mEvaluations
    materialFailures = materialFailures + mMaterialFailures
    gradientBuilds = gradientBuilds + mGradientBuilds: localGradientBuilds = localGradientBuilds + mLocalGradientBuilds
    DiscardStep ' Боковые и центральная истории остаются в одном исходном принятом поле.
    seed = field: label = "повтор при исходной мере"
    extraProbes = 2
    solved = TryDeformationStep(base, direction, originalProjection, originalTarget, guideFactor, factor, _
        absoluteTolerance, relativeTolerance, maxIterations - initialIterations - guideIterations, initialField:=seed)
    If Not solved Then GoTo RejectArc
    CopyTrial field, resultants
    If Not ArcFieldIsLocal(coefficients, projection, loadProjection, loadMetric, radius, field, acceptedLambda, factor) Then
        label = label & " вне локальной окрестности"
        GoTo RejectArc
    End If
    phaseDetails = mMeta.DiagnosticDetails
    trialLambda = factor
    TryRecoverDeformationByArc = True
    GoTo CountArc

RejectArc:
    phaseDetails = mMeta.DiagnosticDetails
    DiscardStep
    mControlProjection = originalProjection: mTargetControl = originalTarget: mArcLoadProjection = 0#
    mCurrentControlLambda = originalFactor
    mExternal = external: mResidual = residual: mTrialResultants = originalResultants: mMatrix = matrix
    If solved Or mMeta.InternalStatus = rsNumericalFailure Then
        Set mMeta = originalMeta
        mLastNorm = norm: mMinimumPivot = pivot: mFailureCode = failureCode
        mMaterialPoint = materialPoint: mLastFailedPoint = lastPoint: mLastFailedSide = lastSide
    End If
CountArc:
    mWrongControlDirection = False
    Dim arcIterations As Long, fixedIterations As Long
    If extraProbes = 1 Then
        arcIterations = mIterations
    Else
        arcIterations = guideIterations: fixedIterations = mIterations
    End If
    mIterations = initialIterations + guideIterations + mIterations
    mEvaluations = initialEvaluations + guideEvaluations + mEvaluations
    mMaterialFailures = materialFailures + mMaterialFailures
    mGradientBuilds = gradientBuilds + mGradientBuilds
    mLocalGradientBuilds = localGradientBuilds + mLocalGradientBuilds
    mMeta.DiagnosticDetails = DiagnosticSummary() & "; проверено вспомогательных проб: " & CStr(extraProbes) & _
        "; итерации исходной пробы / дуги / повтора: " & CStr(initialIterations) & " / " & _
        CStr(arcIterations) & " / " & CStr(fixedIterations) & "; общий лимит: " & CStr(maxIterations) & _
        "; последняя проба: " & label & " | Исходный отказ: " & originalMeta.DiagnosticDetails
    If Len(phaseDetails) > 0 Then mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & " | Проверка: " & phaseDetails
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, errorSource, errorText
    End If
    Exit Function
InterruptedArc:
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    Resume RejectArc
End Function

' Строит нормаль к принятой секущей. Масштаб нагрузки берётся по медиане трёх
' секущих на полной постоянной нагрузке. Неполную или расходящуюся историю не использует.
Private Function PrepareArcMeasure(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, _
        ByVal acceptedLambda As Double, ByRef coefficients() As Double, ByRef projection() As Double, _
        ByRef predicted() As Double, ByRef target As Double, ByRef loadProjection As Double, _
        ByRef loadMetric As Double, ByRef radius As Double, ByRef predictedLambda As Double) As Boolean
    If fields Is Nothing Or trace Is Nothing Then Exit Function
    If fields.Count < 4 Or trace.Count <> fields.Count Then Exit Function
    Dim factor(1 To 4) As Double, fraction As Double, i As Long, index As Long
    For i = 1 To 4
        trace.ReadPoint fields.Count - 4 + i, fraction, factor(i)
        If fraction <> 1# Or factor(i) <= 0# Then Exit Function
        If i > 1 Then
            If factor(i) <= factor(i - 1) Then Exit Function
        End If
    Next i
    If factor(4) <> acceptedLambda Then Exit Function

    Dim old() As Double, field() As Double, movement() As Double, physicalSquare As Double
    Dim slopes(1 To 3) As Double, lambdaMove As Double, fixedMove As Double, fixedDistance As Double
    ReDim movement(1 To mDegreeCount)
    field = fields.Item(fields.Count)
    If LBound(field) <> 1 Or UBound(field) <> mDegreeCount Then Exit Function
    For i = 1 To mDegreeCount
        If field(i) <> mAccepted(i) Then Exit Function
    Next i
    mSection.CopySpatialStrainCoefficients SPATIAL_HALF_STEP, coefficients
    For index = 1 To 3
        old = fields.Item(fields.Count - 4 + index): field = fields.Item(fields.Count - 3 + index)
        If LBound(old) <> 1 Or UBound(old) <> mDegreeCount Then Exit Function
        If LBound(field) <> 1 Or UBound(field) <> mDegreeCount Then Exit Function
        For i = 1 To mGaugeStart - 1: movement(i) = field(i) - old(i): Next i
        lambdaMove = factor(index + 1) - factor(index)
        physicalSquare = ArcPhysicalSquare(coefficients, movement, projection, index = 3)
        If physicalSquare <= 0# Then Exit Function
        slopes(index) = Sqr(physicalSquare) / lambdaMove
    Next index
    Dim median As Double, secant As Double, ratio As Double
    median = GeomMax(GeomMin(slopes(1), slopes(2)), GeomMin(GeomMax(slopes(1), slopes(2)), slopes(3)))
    loadMetric = median * median
    secant = Sqr(physicalSquare + loadMetric * lambdaMove ^ 2)
    If secant <= 0# Then Exit Function
    fixedDistance = mTargetControl
    For i = 1 To mGaugeStart - 1
        fixedMove = fixedMove + mControlProjection(i) * movement(i)
        fixedDistance = fixedDistance - mControlProjection(i) * mAccepted(i)
    Next i
    If fixedMove <= 0# Or fixedDistance <= 0# Then Exit Function
    ratio = 0.5 * GeomMin(1#, fixedDistance / fixedMove)
    radius = ratio * secant
    If radius <= 0# Then Exit Function
    loadProjection = loadMetric * lambdaMove / secant
    target = loadProjection * acceptedLambda + radius
    predictedLambda = acceptedLambda + ratio * lambdaMove
    predicted = mAccepted
    For i = 1 To mGaugeStart - 1
        projection(i) = projection(i) / secant
        target = target + projection(i) * mAccepted(i)
        predicted(i) = predicted(i) + ratio * movement(i)
    Next i
    PrepareArcMeasure = True
End Function

' Применяет физическую метрику к приращению поля. Девять строк оператора
' хранят три неизменные по z/L компоненты и три пары концов. Середина
' каждой пары равна среднему. Вес инженерного сдвига в тензорной норме равен 1/2.
Private Function ArcPhysicalSquare(ByRef coefficients() As Double, ByRef movement() As Double, _
        ByRef projection() As Double, Optional ByVal project As Boolean = False) As Double
    Dim strain(1 To 9) As Double, dual(1 To 9) As Double
    Dim point As Long, row As Long, first As Long, i As Long, x As Double, y As Double, area As Double
    Dim areaSum As Double, value As Double, middle As Double, weight As Double
    If project Then ReDim projection(1 To mDegreeCount)
    For point = 1 To mPointCount
        first = 9 * (point - 1)
        mSection.Point point, x, y, area
        areaSum = areaSum + 3# * area
        For row = 1 To 9
            value = 0#
            For i = 1 To mGaugeStart - 1
                value = value + coefficients(i, first + row) * movement(i)
            Next i
            strain(row) = value
        Next row
        dual(1) = 3# * strain(1): dual(2) = 3# * strain(2): dual(3) = 1.5 * strain(3)
        For row = 4 To 8 Step 2
            weight = 0.5
            If row = 4 Then weight = 1#
            middle = 0.25 * (strain(row) + strain(row + 1))
            dual(row) = weight * (strain(row) + middle)
            dual(row + 1) = weight * (strain(row + 1) + middle)
        Next row
        For row = 1 To 9
            value = area * dual(row)
            ArcPhysicalSquare = ArcPhysicalSquare + strain(row) * value
            If project Then
                For i = 1 To mGaugeStart - 1
                    projection(i) = projection(i) + coefficients(i, first + row) * value
                Next i
            End If
        Next row
    Next point
    ArcPhysicalSquare = ArcPhysicalSquare / areaSum
    If Not project Then Exit Function
    For i = 1 To mGaugeStart - 1: projection(i) = projection(i) / areaSum: Next i
End Function

' Проверяет положительный ход в малой окрестности принятого поля.
' Равновесие, исходная мера и связи уже проверены штатным решателем.
Private Function ArcFieldIsLocal(ByRef coefficients() As Double, ByRef projection() As Double, ByVal loadProjection As Double, _
        ByVal loadMetric As Double, ByVal radius As Double, ByRef field() As Double, ByVal acceptedLambda As Double, _
        ByVal factor As Double) As Boolean
    If factor <= 0# Then Exit Function
    Dim movement() As Double, unused() As Double, i As Long, progress As Double, distanceSquare As Double
    ReDim movement(1 To mDegreeCount)
    progress = loadProjection * (factor - acceptedLambda)
    For i = 1 To mGaugeStart - 1
        movement(i) = field(i) - mAccepted(i)
        progress = progress + projection(i) * movement(i)
    Next i
    If progress <= 0# Then Exit Function
    distanceSquare = ArcPhysicalSquare(coefficients, movement, unused) + loadMetric * (factor - acceptedLambda) ^ 2
    ArcFieldIsLocal = (distanceSquare <= 9# * radius ^ 2)
End Function

' Возвращает новую линейную меру только для подтверждённой пробы восстановления.
Friend Function CopyRecoveredControl(ByRef projection() As Double) As Boolean
    If Not mCandidateReady Or Not mRecoveredControlReady Then Exit Function
    projection = mRecoveredProjection
    CopyRecoveredControl = True
End Function

' Передаёт требование к касательной только для подтверждённой новой меры материального восстановления.
Friend Property Get RecoveredControlRequiresBranchTangent() As Boolean
    If mCandidateReady And mRecoveredControlReady Then RecoveredControlRequiresBranchTangent = mConsistentControl
End Property

' Сообщает о наличии знаковых границ у пространственного бетонного материала.
Friend Property Get HasPrincipalBoundaries() As Boolean
    If Not mUseSpatialHistory Or Not TypeOf mResponse Is CCoupledReinforcedResponse Then Exit Property
    Dim response As CCoupledReinforcedResponse: Set response = mResponse
    HasPrincipalBoundaries = response.HasPrincipalBoundaries
End Property

' Сообщает, что после действительной смены ветви собиралось аварийное направление.
Friend Property Get YieldRecoveryTried() As Boolean
    YieldRecoveryTried = mYieldRecoveryTried
End Property

' Сообщает о проверке перехода для повтора в той же мере. Эта проверка
' расходует отдельный бюджет, даже если пересечение не найдено.
Friend Property Get SameControlRecoveryTried() As Boolean
    SameControlRecoveryTried = mSameControlRecoveryTried
End Property

' Выбирает уверенно меняющуюся меру по двум последовательным равновесиям.
' Кривизны и закручивание уже приведены к габариту сечения. Средние
' инженерные сдвиги позволяют управлять продолжением при чистом Q.
Friend Function SelectDeformationControl(ByRef previous() As Double, ByRef current() As Double, _
        ByRef projection() As Double, ByRef currentControl As Double, ByRef controlIncrement As Double, _
        ByRef controlName As String) As Boolean
    Dim shearX() As Double, shearY() As Double, i As Long, candidate As Long, chosen As Long
    Dim changes(1 To COMPONENT_COUNT) As Double, signValue As Double, best As Double
    mSection.MeanShearProjections shearX, shearY
    For i = 1 To 4: changes(i) = current(i) - previous(i): Next i
    For i = 1 To mGaugeStart - 1
        changes(5) = changes(5) + shearX(i) * (current(i) - previous(i))
        changes(6) = changes(6) + shearY(i) * (current(i) - previous(i))
    Next i
    For candidate = 1 To COMPONENT_COUNT
        If Abs(changes(candidate)) > best Then
            chosen = candidate: best = Abs(changes(candidate))
        End If
    Next candidate
    If chosen = 0 Then Exit Function
    signValue = Sgn(changes(chosen))
    controlIncrement = best
    If chosen <= 4 Then
        ReDim projection(1 To mDegreeCount)
        projection(chosen) = signValue
    ElseIf chosen = 5 Then
        projection = shearX
    Else
        projection = shearY
    End If
    currentControl = 0#
    For i = 1 To mGaugeStart - 1
        If chosen > 4 Then projection(i) = signValue * projection(i)
        currentControl = currentControl + current(i) * projection(i)
    Next i
    Select Case chosen
        Case 1: controlName = "продольная деформация"
        Case 2: controlName = "кривизна изгиба вокруг X"
        Case 3: controlName = "кривизна изгиба вокруг Y"
        Case 4: controlName = "закручивание"
        Case 5: controlName = "средний инженерный сдвиг xz"
        Case 6: controlName = "средний инженерный сдвиг yz"
    End Select
    SelectDeformationControl = True
End Function

' Выбирает средний инженерный сдвиг, сопряжённый направлению изменяемой поперечной силы.
' Проекция не зависит от размера силы; отсутствие изменения сдвига сохраняет общий выбор меры.
Friend Function SelectTransverseDeformationControl(ByRef previous() As Double, ByRef current() As Double, _
        ByRef loadDirection() As Double, ByRef projection() As Double, ByRef currentControl As Double, _
        ByRef controlIncrement As Double, ByRef controlName As String) As Boolean
    If LBound(loadDirection) <> 1 Or UBound(loadDirection) <> COMPONENT_COUNT Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Для выбора сдвиговой меры передайте шесть компонент направления нагрузки."
    Dim forceScale As Double, alongX As Double, alongY As Double, length As Double
    forceScale = GeomMax(Abs(loadDirection(5)), Abs(loadDirection(6)))
    If forceScale = 0# Then Exit Function
    alongX = loadDirection(5) / forceScale: alongY = loadDirection(6) / forceScale
    length = Sqr(alongX * alongX + alongY * alongY)
    alongX = alongX / length: alongY = alongY / length
    Dim shearX() As Double, shearY() As Double, i As Long, change As Double, signValue As Double
    mSection.MeanShearProjections shearX, shearY
    ReDim projection(1 To mDegreeCount)
    For i = 1 To mGaugeStart - 1
        projection(i) = alongX * shearX(i) + alongY * shearY(i)
        change = change + projection(i) * (current(i) - previous(i))
    Next i
    If change = 0# Then Exit Function
    signValue = Sgn(change): currentControl = 0#: controlIncrement = Abs(change)
    For i = 1 To mGaugeStart - 1
        projection(i) = signValue * projection(i)
        currentControl = currentControl + projection(i) * current(i)
    Next i
    controlName = "средний инженерный сдвиг по направлению изменяемой поперечной силы"
    SelectTransverseDeformationControl = True
End Function

' Строит меру вдоль заданного направления в средней тензорной метрике трёх материальных станций.
' Масштаб направления не влияет на проекцию; множители связей исключаются из метрики и ориентации.
Friend Function SelectSpatialDirectionControl(ByRef previous() As Double, ByRef current() As Double, _
        ByRef movement() As Double, ByRef projection() As Double, ByRef currentControl As Double, _
        ByRef controlIncrement As Double) As Boolean
    Dim i As Long, fieldScale As Double
    For i = 1 To mGaugeStart - 1: fieldScale = GeomMax(fieldScale, Abs(movement(i))): Next i
    If fieldScale = 0# Then Exit Function
    Dim normalized() As Double, coefficients() As Double, square As Double, length As Double, change As Double
    ReDim normalized(1 To mDegreeCount)
    For i = 1 To mGaugeStart - 1: normalized(i) = movement(i) / fieldScale: Next i
    mSection.CopySpatialStrainCoefficients SPATIAL_HALF_STEP, coefficients
    square = ArcPhysicalSquare(coefficients, normalized, projection, True)
    If square <= 0# Then Exit Function
    length = Sqr(square)
    For i = 1 To mGaugeStart - 1
        projection(i) = projection(i) / length
        change = change + projection(i) * (current(i) - previous(i))
    Next i
    If change = 0# Then Exit Function
    currentControl = 0#: controlIncrement = Abs(change)
    For i = 1 To mGaugeStart - 1
        projection(i) = Sgn(change) * projection(i)
        currentControl = currentControl + projection(i) * current(i)
    Next i
    SelectSpatialDirectionControl = True
End Function

' Проверяет ход нового приращения вдоль предыдущей пространственной секущей в тензорной метрике трёх станций.
' Это численная ориентация продолжения, не критерий разрушения; служебные множители связей не участвуют.
Friend Function SpatialIncrementAdvances(ByRef previous() As Double, ByRef accepted() As Double, _
        ByRef trial() As Double, ByRef signedAdvance As Double, ByRef referenceIncrement As Double) As Boolean
    signedAdvance = 0#: referenceIncrement = 0#
    Dim movement() As Double, projection() As Double, coordinate As Double, i As Long
    ReDim movement(1 To mDegreeCount)
    For i = 1 To mGaugeStart - 1: movement(i) = accepted(i) - previous(i): Next i
    If Not SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, referenceIncrement) Then Exit Function
    For i = 1 To mGaugeStart - 1
        signedAdvance = signedAdvance + projection(i) * (trial(i) - accepted(i))
    Next i
    SpatialIncrementAdvances = (signedAdvance > 0#)
End Function

' Сохраняет только фактически отклонённую равновесную пробу для последующего выбора материального события.
Friend Sub PreserveRejectedSpatialTrial(ByRef previous() As Double)
    RequireCandidate
    Dim advance As Double, reference As Double
    If SpatialIncrementAdvances(previous, mAccepted, mCandidate, advance, reference) Then Exit Sub
    If HasPrincipalTransitionSegment Then Exit Sub ' Восстановленное равновесие может вернуться к прежнему знаку; фактическое промежуточное пересечение Ньютона сохраняется.
    mTransitionFrom = mAccepted: mTransitionTo = mCandidate
    mTransitionRevision = mAcceptedRevision: mHasTransitionSegment = True
End Sub

' Сообщает о действительном отклонённом отрезке той же принятой истории, пригодном для выбора спектральной границы.
Friend Property Get HasPrincipalTransitionSegment() As Boolean
    HasPrincipalTransitionSegment = mHasTransitionSegment And mTransitionRevision = mAcceptedRevision And HasPrincipalBoundaries
End Property

' Отбирает материальное событие до расходования новой пробы продолжения; сама граница ещё требует полного решения равновесия.
Friend Function CanPreparePrincipalTransition(ByRef previous() As Double, ByVal currentLambda As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double) As Boolean
    If Not HasPrincipalTransitionSegment Then Exit Function
    Dim station As TYieldStation
    CanPreparePrincipalTransition = SelectPrincipalTransitionStation(previous, station, _
        PrincipalCandidateTolerance(currentLambda, absoluteTolerance, relativeTolerance))
End Function

' Выбирает меру по последней заблокированной поправке только из того же принятого состояния.
' Направление служит параметризацией следующей пробы, а не принятым полем или признаком разрушения.
Friend Function SelectRejectedDeformationControl(ByRef previous() As Double, ByRef current() As Double, _
        ByRef projection() As Double, ByRef currentControl As Double, ByRef controlIncrement As Double, _
        ByRef controlName As String) As Boolean
    If Not mUseSpatialHistory Or Not mControlActive Or mCandidateReady Or Not mHasRejectedDirection Then Exit Function
    If mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
    Dim i As Long
    For i = 1 To mDegreeCount: If current(i) <> mAccepted(i) Then Exit Function
    Next i
    If Not SelectSpatialDirectionControl(previous, current, mRejectedDirection, projection, currentControl, controlIncrement) Then Exit Function
    controlName = "пространственная мера по отвергнутой поправке"
    SelectRejectedDeformationControl = True
End Function

' После численного отказа уточнённого силового шага ищет ограниченный прогноз.
' Успех разрешает только полное повторное решение с исходной нагрузкой и бюджетом.
Private Function TrySubspaceForceRecovery(ByRef load() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        Optional ByRef initialField As Variant, Optional ByVal retainDirectProof As Boolean = False) As Boolean
    If mControlActive Or Not mUseSpatialHistory Or mSubspaceRecoveryActive Or mCandidateReady Then Exit Function
    If Not mSection.HasRefinedField Then Exit Function ' Прогноз ограничивает дополнительные пространственные моды уточнённого поля.
    If mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
    If mMeta.ResultCode <> rcNumericalFailure And mMeta.ResultCode <> rcSingularTangent Then Exit Function
    Dim remaining As Long: remaining = maxIterations - mIterations
    If remaining < 2 Then Exit Function
    If Not mSubspaceChecked Then
        Dim prepared As New CCoupledFieldSubspace
        If prepared.Prepare(mSection) Then Set mFieldSubspace = prepared
        mSubspaceChecked = True
    End If
    If mFieldSubspace Is Nothing Then Exit Function

    Dim original As CResultMeta, beforeRetry As TSolveEffort
    Set original = mMeta.Clone
    Dim external() As Double, residual() As Double, matrix() As Double, resultants() As Double
    Dim originalNorm As Double, originalPivot As Double
    Dim originalCode As ESolverFailureCode, originalPoint As Long, originalLastPoint As Long, originalSide As Long
    external = mExternal: residual = mResidual: matrix = mMatrix: resultants = mTrialResultants
    originalNorm = mLastNorm: originalPivot = mMinimumPivot
    originalCode = mFailureCode: originalPoint = mMaterialPoint
    originalLastPoint = mLastFailedPoint: originalSide = mLastFailedSide
    Dim seed() As Double, predicted() As Double, initial As Variant
    If IsMissing(initialField) Then seed = mAccepted Else seed = initialField
    Dim auxiliaryIterations As Long, details As String, predictedReady As Boolean, predictionCharged As Boolean
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    On Error GoTo RecoveryFailed
    mSubspaceRecoveryActive = True
    predictedReady = BuildSubspaceForcePrediction(seed, load, absoluteTolerance, relativeTolerance, _
        GeomMin(SUBSPACE_ITERATIONS, remaining - 1), predicted, auxiliaryIterations, details)
    mIterations = mIterations + auxiliaryIterations: predictionCharged = True
    If predictedReady Then
        ReadSolveEffort beforeRetry
        initial = predicted
        TrySubspaceForceRecovery = SolveStep(load, 0#, absoluteTolerance, relativeTolerance, _
            maxIterations - mIterations, initialField:=initial, retainDirectProof:=retainDirectProof)
        AddSolveEffort beforeRetry
        If TrySubspaceForceRecovery Then
            mMeta.DiagnosticDetails = DiagnosticSummary() & "; подтверждён квадратичный прогноз; " & details & _
                " | Исходная численная проба: " & original.DiagnosticDetails
            GoTo Finished
        End If
        If mMeta.InternalStatus <> rsNumericalFailure Then
            mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & " | После квадратичного прогноза; " & details
            GoTo Finished
        End If
        details = details & "; полная система не подтвердила прогноз"
    End If

RestoreFailure:
    mCandidateReady = False: Set mMeta = original
    mExternal = external: mResidual = residual: mMatrix = matrix: mTrialResultants = resultants
    mLastNorm = originalNorm: mMinimumPivot = originalPivot
    mFailureCode = originalCode: mMaterialPoint = originalPoint
    If mLastFailedPoint = 0 Then mLastFailedPoint = originalLastPoint: mLastFailedSide = originalSide
    mMeta.DiagnosticDetails = DiagnosticSummary() & "; квадратичный прогноз не принят; " & details & _
        " | Исходная численная проба: " & original.DiagnosticDetails
Finished:
    mSubspaceRecoveryActive = False
    Exit Function

RecoveryFailed:
    failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
    If failureNumber = 6 Or failureNumber = 11 Then
        If Not predictionCharged Then mIterations = mIterations + auxiliaryIterations
        details = "не удалось вычислить конечный квадратичный прогноз: " & failureDescription
        GoTo RestoreFailure
    End If
    mSubspaceRecoveryActive = False
    Err.Raise failureNumber, failureSource, failureDescription
End Function

' Решает вспомогательные уравнения P-transpose r(Pv) при заданных шести усилиях.
' История материальных станций только читается; прогноз не публикует равновесие.
Private Function BuildSubspaceForcePrediction(ByRef seed() As Double, ByRef load() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        ByRef predicted() As Double, ByRef iterations As Long, ByRef details As String) As Boolean
    Erase predicted: iterations = 0
    Dim coordinates() As Double, current() As Double, reduced() As Double, smallMatrix() As Double
    Dim correction() As Double
    If Not mFieldSubspace.TryCoordinates(seed, coordinates) Then
        details = "не подготовлены координаты прогноза": Exit Function
    End If
    Dim count As Long, i As Long
    count = mFieldSubspace.Count
    Dim norm As Double, trialNorm As Double, tolerance As Double
    Dim fraction As Double, accepted As Boolean, trialCoordinates() As Double, candidate() As Double
    Dim system As New CCoupledLinearSystem
    Do
        mFieldSubspace.Expand coordinates, current
        If Not Evaluate(current, True, True, recordTrial:=False) Then
            details = "не получен материальный отклик прогноза": Exit Function
        End If
        norm = SubspaceForceResidualNorm(load, reduced)
        tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
        details = "итераций прогноза: " & CStr(iterations) & "; ограниченная невязка, МПа: " & CStr(norm)
        If norm <= tolerance Then
            predicted = current
            BuildSubspaceForcePrediction = True: Exit Function
        End If
        If iterations >= maxIterations Then
            details = details & "; достигнут бюджет прогноза": Exit Function
        End If
        mFieldSubspace.PullMatrix mMatrix, smallMatrix
        correction = reduced
        For i = 1 To count
            correction(i) = -reduced(i)
        Next i
        If Not system.Solve(smallMatrix, correction, count) Then
            details = details & "; вырождена касательная прогноза": Exit Function
        End If
        ReDim trialCoordinates(1 To count)
        fraction = 1#: accepted = False
        Do While fraction >= STEEL_MIN_FRACTION
            For i = 1 To count: trialCoordinates(i) = coordinates(i) + fraction * correction(i): Next i
            mFieldSubspace.Expand trialCoordinates, candidate
            If Evaluate(candidate, False, recordTrial:=False) Then
                trialNorm = SubspaceForceResidualNorm(load, reduced)
                If trialNorm < norm Then
                    coordinates = trialCoordinates
                    accepted = True: Exit Do
                End If
            End If
            fraction = fraction * 0.5
        Loop
        iterations = iterations + 1 ' Неудачное направление тоже расходует общий бюджет шага.
        If Not accepted Then
            details = details & "; демпфирование не уменьшило ограниченную невязку": Exit Function
        End If
    Loop
End Function

' Объединяет ограниченные слабые уравнения, связи и независимые интегралы шести усилий.
Private Function SubspaceForceResidualNorm(ByRef load() As Double, ByRef reduced() As Double) As Double
    mFieldSubspace.PullVector mResidual, reduced
    SubspaceForceResidualNorm = GeomMax(mFieldSubspace.ResidualNorm(reduced, mReferenceModulus), _
        mSection.ResultantResidualNorm(mTrialResultants, load))
End Function

' Сохраняет затраты исходной пробы и прогноза перед полным повторным решением.
Private Sub ReadSolveEffort(ByRef effort As TSolveEffort)
    effort.Iterations = mIterations: effort.Evaluations = mEvaluations
    effort.GradientBuilds = mGradientBuilds: effort.LocalGradientBuilds = mLocalGradientBuilds
    effort.SteelBranchBuilds = mSteelBranchBuilds: effort.SteelFieldTangents = mSteelFieldTangents
    effort.MaterialFailures = mMaterialFailures
    effort.LastFailedPoint = mLastFailedPoint: effort.LastFailedSide = mLastFailedSide
    effort.HybridSteps = mHybridSteps: effort.HybridMatrices = mHybridMatrices: effort.HybridEvaluations = mHybridEvaluations
End Sub

' Добавляет предыдущие затраты после повторного решения; счётчики не скрывают вспомогательную работу.
Private Sub AddSolveEffort(ByRef effort As TSolveEffort)
    mIterations = mIterations + effort.Iterations: mEvaluations = mEvaluations + effort.Evaluations
    mGradientBuilds = mGradientBuilds + effort.GradientBuilds
    mLocalGradientBuilds = mLocalGradientBuilds + effort.LocalGradientBuilds
    mSteelBranchBuilds = mSteelBranchBuilds + effort.SteelBranchBuilds
    mSteelFieldTangents = mSteelFieldTangents + effort.SteelFieldTangents
    mHybridSteps = mHybridSteps + effort.HybridSteps: mHybridMatrices = mHybridMatrices + effort.HybridMatrices
    mHybridEvaluations = mHybridEvaluations + effort.HybridEvaluations
    mMaterialFailures = mMaterialFailures + effort.MaterialFailures
    If mLastFailedPoint = 0 And effort.MaterialFailures > 0 Then
        mLastFailedPoint = effort.LastFailedPoint: mLastFailedSide = effort.LastFailedSide
    End If
End Sub

' Очищает только пробные данные и затраты одного решения, сохраняя принятую материальную историю.
Private Sub ResetSolveState()
    mDirectProofMatrix = False: mMeta.Clear: mCandidateReady = False
    mYieldRecoveryTried = False: mSameControlRecoveryTried = False: mWrongControlDirection = False
    mSameControlInitialFailure = vbNullString
    mRecoveredControlReady = False: Erase mRecoveredProjection
    mHasRejectedDirection = False: Erase mRejectedDirection
    mHasTransitionSegment = False: mPrincipalIncomingCandidate = False: mPrincipalDesignVerified = False
    mPrincipalOutgoingCandidate = False: Set mPrincipalDeparture = Nothing
    mIterations = 0: mEvaluations = 0: mMaterialPoint = 0: mGradientBuilds = 0
    mLocalGradientBuilds = 0: mSteelBranchBuilds = 0: mSteelFieldTangents = 0
    mMaterialFailures = 0: mLastFailedPoint = 0: mLastFailedSide = 0
    mLastNorm = 0#: mMinimumPivot = 0#
    mHybridSteps = 0: mHybridMatrices = 0: mHybridEvaluations = 0
End Sub

' Решает одну силовую или деформационно управляемую пробу с общей
' невязкой, ограничителями поправок и фиксацией материальных состояний.
Private Function SolveStep(ByRef load() As Double, ByVal initialLambda As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        Optional ByVal recoverYieldControl As Boolean = False, Optional ByRef initialField As Variant, _
        Optional ByVal retainDirectProof As Boolean = False, Optional ByVal recoverSameControl As Boolean = False, _
        Optional ByVal allowSmallCorrection As Boolean = False, Optional ByVal allowHybridRecovery As Boolean = False) As Boolean
    ResetSolveState
    If absoluteTolerance <= 0# Or relativeTolerance <= 0# Or maxIterations < 1 Then
        RecordFailure sfcInvalidConfiguration, "Для связанного равновесия задайте положительные допуски и лимит итераций."
        Exit Function
    End If
    On Error GoTo EvaluationFailed
    Dim current() As Double, correction() As Double, norm As Double, tolerance As Double
    Dim candidateNorm As Double, candidateTolerance As Double, fraction As Double, i As Long, accepted As Boolean
    Dim bestTrialNorm As Double, hasValidTrial As Boolean
    Dim governingRow As Long, baseEquation As Double, trialEquation As Double, trialFraction As Double
    Dim useGradientTangent As Boolean, built As Boolean
    Dim useSteelCorrection As Boolean, useConcreteFieldTangent As Boolean, minimumFraction As Double
    useSteelCorrection = Not mControlActive And Not (mSteelCorrector Is Nothing)
    Dim currentLambda As Double, candidateLambda As Double
    currentLambda = initialLambda: candidateLambda = initialLambda
    If mControlActive Then SetControlledLoad currentLambda, load
    tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    If IsMissing(initialField) Then
        current = mAccepted
    Else
        current = initialField ' Прогноз меняет только первое поле итераций, материалы читают принятую историю.
        If LBound(current) <> 1 Or UBound(current) <> mDegreeCount Then
            RecordFailure sfcInvalidInput, "Размер начального поля не совпадает с решателем."
            Exit Function
        End If
    End If
    Do
RecomputeTangent:
        If mControlActive Then SetControlledLoad currentLambda, load
        built = Evaluate(current, True, useGradientTangent, Not mConsistentControl, mConsistentControl)
        If Not built And mConsistentControl And mFailureCode = sfcNumericalFailure Then
            built = Evaluate(current, True, useGradientTangent, recordTrial:=False) ' Например, при вырождении односторонних разностей в другой точке.
        End If
        If Not built Then
            Dim assemblyFailure As ESolverFailureCode, tangentPoint As Long
            assemblyFailure = mFailureCode: tangentPoint = mMaterialPoint
            If assemblyFailure = sfcNumericalFailure Or assemblyFailure = sfcSingularTangent Then
                If Evaluate(current, False) Then ' В точном равновесии поправка и её касательная не нужны.
                    If EquilibriumNorm(load, current) <= tolerance And GaugesVerified(tolerance) Then
                        mCandidate = current
                        PublishCandidate currentLambda
                        SolveStep = True
                        Exit Function
                    End If
                    mMaterialPoint = tangentPoint ' Сохраняем точку отказа касательной после проверки невязки.
                Else
                    assemblyFailure = mFailureCode
                End If
            End If
            RecordFailure assemblyFailure, "Не удалось получить материальный отклик для равновесия связанного сечения."
            Exit Function
        End If
        norm = EquilibriumNorm(load, current)
        If mConsistentControl And norm > tolerance And norm <= SMALL_STRAIN_MAX_NORM * tolerance Then
            If TryWarpGaugeEquilibrium(current, load, tolerance) Then norm = mLastNorm
        End If
        If norm <= tolerance Then
            If mConsistentControl Then
                If Not Evaluate(current, False) Then
                    RecordFailure mFailureCode, "Не удалось записать пробу подтверждённого равновесия."
                    Exit Function
                End If
                If EquilibriumNorm(load, current) > tolerance Then
                    RecordFailure sfcNumericalFailure, "Невязка изменилась при повторе принятой пробы."
                    Exit Function
                End If
            End If
            If Not GaugesVerified(tolerance) Then
                RecordFailure sfcNumericalFailure, "Уравновешенное поле нарушает условия однозначности перемещений."
                Exit Function
            End If
            mCandidate = current
            mDirectProofMatrix = retainDirectProof And Not mControlActive
            PublishCandidate currentLambda
            SolveStep = True
            Exit Function
        End If
        If mIterations >= maxIterations Then Exit Do
        If Not SolveCorrection(current, correction) Then
            mMinimumPivot = mSystem.MinimumPivot
            If Not mUseSpatialHistory And Not useGradientTangent And mSection.HasAxialGradient(current) Then
                useGradientTangent = True
                GoTo RecomputeTangent
            End If
            RecordFailure mSystem.FailureCode, "Касательная связанного сечения вырождена или не решается с заданной точностью."
            Exit Function
        End If
        mMinimumPivot = mSystem.MinimumPivot
        If useSteelCorrection Then RefineSteelCorrection current, correction, useConcreteFieldTangent
        If mControlActive And (useGradientTangent Or mUseSpatialHistory) Then
            governingRow = LargestEquilibriumResidualRow()
            baseEquation = mResidual(governingRow)
        End If
        ReDim mCandidate(1 To mDegreeCount)
        fraction = 1#: accepted = False: hasValidTrial = False
        candidateTolerance = tolerance
        minimumFraction = MIN_FRACTION
        If mSteelBranchBuilds > 0 Then minimumFraction = STEEL_MIN_FRACTION
        Do While fraction >= minimumFraction
            For i = 1 To mDegreeCount
                mCandidate(i) = current(i) + fraction * correction(i)
            Next i
            If mControlActive Then
                candidateLambda = currentLambda + fraction * correction(mDegreeCount + 1)
                SetControlledLoad candidateLambda, load
            End If
            If Evaluate(mCandidate, False) Then
                If mControlActive Then candidateTolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
                candidateNorm = EquilibriumNorm(load, mCandidate)
                If Not hasValidTrial Or candidateNorm < bestTrialNorm Then bestTrialNorm = candidateNorm
                hasValidTrial = True
                If governingRow > 0 Then
                    trialEquation = mResidual(governingRow)
                    trialFraction = fraction
                End If
                accepted = (candidateNorm <= candidateTolerance Or candidateNorm < norm * (1# - MIN_PROGRESS * fraction))
            ElseIf mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
                RecordFailure mFailureCode, "Материальная проба не может быть повторена с меньшей поправкой."
                Exit Function
            End If
            If accepted Then Exit Do
            fraction = fraction * 0.5
        Loop
        If Not accepted Then
            If useSteelCorrection And mSteelBranchBuilds > 0 And Not useConcreteFieldTangent And _
                    mSteelFieldTangents < MAX_STEEL_FIELD_TANGENTS Then
                useConcreteFieldTangent = True ' Та же точка и та же принятая история, другой бетонный градиент.
                GoTo RecomputeTangent
            End If
            If Not mUseSpatialHistory And Not useGradientTangent And mSection.HasAxialGradient(current) Then
                useGradientTangent = True
                GoTo RecomputeTangent
            End If
            If mControlActive And (mConsistentControl Or allowSmallCorrection) And norm <= SMALL_STRAIN_MAX_NORM * tolerance Then
                Dim rejectedNorm As Double, rejectedCode As ESolverFailureCode
                Dim rejectedPoint As Long, rejectedLastPoint As Long, rejectedSide As Long, rejectedFailures As Long
                Dim rejectedLoad(1 To COMPONENT_COUNT) As Double, rejectedExternal() As Double
                Dim rejectedResidual() As Double, rejectedResultants() As Double, rejectedMatrix() As Double
                rejectedNorm = mLastNorm: rejectedCode = mFailureCode
                rejectedPoint = mMaterialPoint: rejectedLastPoint = mLastFailedPoint: rejectedSide = mLastFailedSide
                rejectedFailures = mMaterialFailures
                rejectedExternal = mExternal: rejectedResidual = mResidual: rejectedResultants = mTrialResultants
                rejectedMatrix = mMatrix
                For i = 1 To COMPONENT_COUNT: rejectedLoad(i) = load(i): Next i
                If TrySmallStrainEquilibrium(current, currentLambda, load, norm, tolerance, _
                        absoluteTolerance, relativeTolerance, candidateLambda) Then
                    mIterations = mIterations + 1
                    PublishCandidate candidateLambda
                    mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & "; поправка с ограниченным приращением деформации"
                    SolveStep = True
                    Exit Function
                End If
                mLastNorm = rejectedNorm: mFailureCode = rejectedCode
                mMaterialPoint = rejectedPoint: mLastFailedPoint = rejectedLastPoint: mLastFailedSide = rejectedSide
                mMaterialFailures = rejectedFailures
                mExternal = rejectedExternal: mResidual = rejectedResidual: mTrialResultants = rejectedResultants
                mMatrix = rejectedMatrix
                For i = 1 To COMPONENT_COUNT: load(i) = rejectedLoad(i): Next i
            End If
            RecordFailure sfcNumericalFailure, "Уменьшение поправки не снизило невязку связанного сечения."
            Dim details As String, mismatch As String
            details = mMeta.DiagnosticDetails
            If hasValidTrial Then
                details = details & "; невязка перед демпфированием, МПа: " & CStr(norm) & _
                    "; лучшая невязка проб, МПа: " & CStr(bestTrialNorm)
                If governingRow > 0 Then
                    mismatch = DirectionMismatchDetails(governingRow, baseEquation, _
                        (trialEquation - baseEquation) / trialFraction, correction)
                    details = details & mismatch
                End If
            Else
                details = details & "; ни одна уменьшенная поправка не дала материального отклика"
            End If
            mMeta.DiagnosticDetails = details
            If hasValidTrial And mControlActive And mUseSpatialHistory Then ' Максимум невязки может перейти в другую строку при верном знаке исходной.
                ReDim mRejectedDirection(1 To mDegreeCount)
                For i = 1 To mGaugeStart - 1: mRejectedDirection(i) = correction(i): Next i
                mHasRejectedDirection = True
                mTransitionFrom = current: mTransitionTo = current
                For i = 1 To mGaugeStart - 1: mTransitionTo(i) = current(i) + correction(i): Next i
                mTransitionRevision = mAcceptedRevision: mHasTransitionSegment = True
            End If
            If Len(mismatch) > 0 And mUseSpatialHistory Then
                mWrongControlDirection = True
                If recoverYieldControl Then
                    Dim recoveryIterationLimit As Long
                    If HasPrincipalBoundaries Then recoveryIterationLimit = maxIterations
                    SolveStep = RecoverYieldStep(current, currentLambda, correction, trialFraction, tolerance, _
                        absoluteTolerance, relativeTolerance, recoveryIterationLimit)
                ElseIf recoverSameControl And mControlActive And mConsistentControl And Not mSameControlRecoveryTried Then
                    If mIterations + 1 < maxIterations Then
                        mSameControlRecoveryTried = True
                        mSameControlInitialFailure = details
                        If TrySameControlSeed(current, currentLambda, correction, trialFraction, tolerance, _
                                absoluteTolerance, relativeTolerance, maxIterations - 1) Then
                            mIterations = mIterations + 1 ' Заблокированное направление расходует одну итерацию этого шага.
                            SetControlledLoad currentLambda, load
                            tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
                            GoTo RecomputeTangent
                        End If
                        RecordFailure sfcNumericalFailure, "Уменьшение поправки не снизило невязку связанного сечения."
                    End If
                End If
            End If
            If Not SolveStep And allowHybridRecovery Then SolveStep = TryHybridControlledField(current, currentLambda, load, _
                absoluteTolerance, relativeTolerance, maxIterations)
            Exit Function
        End If
        mIterations = mIterations + 1
        If candidateNorm <= candidateTolerance And GaugesVerified(candidateTolerance) Then
            PublishCandidate candidateLambda
            SolveStep = True
            Exit Function
        End If
        useConcreteFieldTangent = False
        current = mCandidate
        currentLambda = candidateLambda
        tolerance = candidateTolerance
    Loop
    RecordFailure sfcNumericalFailure, "Достигнут лимит итераций связанного равновесия."
    Exit Function

EvaluationFailed:
    Dim number As Long, source As String, description As String
    number = Err.Number: source = Err.Source: description = Err.Description
    If number = 18 Then Err.Raise number, source, description ' Стандартная отмена Excel передаётся внешнему владельцу.
    If number = 6 Or number = 11 Then
        RecordFailure sfcNumericalFailure, "Не удалось вычислить конечную поправку поля. " & description
    Else
        RecordFailure sfcInternalError, "Ошибка сборки связанного равновесия. " & description
    End If
End Function

' Проверяет явно запрошенный гибридный резерв после численного отказа демпфирования пространственного CDP.
' Каждая dogleg-поправка расходует остаток maxIterations; все разностные чтения входят в mEvaluations.
' Для спектральных границ ASD и псевдодуги сохраняются их собственные правила продолжения.
Private Function TryHybridControlledField(ByRef current() As Double, ByVal initialLambda As Double, ByRef load() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long) As Boolean
    If Not mControlActive Or mArcLoadProjection <> 0# Or Not SupportsHybridRecovery Then Exit Function
    If mCandidateReady Or mHybridActive Or mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
    If mMeta.ResultCode <> rcNumericalFailure And mMeta.ResultCode <> rcSingularTangent Then Exit Function
    Dim remaining As Long: remaining = maxIterations - mIterations
    If remaining < 2 Then Exit Function ' Отклонённая поправка Ньютона и хотя бы одна новая поправка делят общий лимит.
    Dim original As CResultMeta: Set original = mMeta.Clone
    Dim external() As Double, residual() As Double, matrix() As Double, resultants() As Double
    Dim originalNorm As Double, originalLambda As Double, originalCode As ESolverFailureCode, originalPoint As Long
    external = mExternal: residual = mResidual: matrix = mMatrix: resultants = mTrialResultants
    originalNorm = mLastNorm: originalLambda = mCurrentControlLambda: originalCode = mFailureCode: originalPoint = mMaterialPoint
    Dim originalLoad(1 To COMPONENT_COUNT) As Double, i As Long
    For i = 1 To COMPONENT_COUNT: originalLoad(i) = load(i): Next i
    On Error GoTo Failed
    If Not PrepareHybridScales() Then GoTo RestoreFailure
    mIterations = mIterations + 1: remaining = remaining - 1 ' Учитывается уже вычисленная, но отвергнутая поправка Ньютона.
    Dim size As Long: size = mDegreeCount + 1
    ReDim mHybridSeed(1 To size)
    For i = 1 To mDegreeCount: mHybridSeed(i) = current(i): Next i
    mHybridSeed(size) = initialLambda: mHybridOriginRevision = mAcceptedRevision
    Dim coordinates() As Double, solution() As Double, field() As Double
    ReDim coordinates(1 To size): ReDim field(1 To mDegreeCount)
    Dim hybrid As New CCoupledHybridSolver, equations As ICoupledEquationSystem
    Set equations = Me
    Dim functionLimit As Long: functionLimit = CLng(CDbl(remaining + 1) * (2# * size + 1#) + remaining)
    mHybridActive = True
    Dim converged As Boolean
    converged = hybrid.TrySolve(equations, coordinates, functionLimit, remaining, GeomMin(absoluteTolerance, 0.0000000001), _
        0.00000000001, 0.2, solution, relativeDifferenceStep:=HYBRID_DIFFERENCE_STEP)
    mHybridActive = False
    mIterations = mIterations + hybrid.Steps
    mHybridSteps = mHybridSteps + hybrid.Steps: mHybridMatrices = mHybridMatrices + hybrid.JacobianCalls
    mHybridEvaluations = mHybridEvaluations + hybrid.FunctionCalls
    If mFailureCode <> sfcNone And mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
        RecordFailure mFailureCode, "Материальная проба глобальной касательной не может быть исправлена уменьшением поправки."
        Exit Function
    End If
    If Not hybrid.HasBestPoint Then GoTo RestoreFailure
    For i = 1 To mDegreeCount: field(i) = mHybridSeed(i) + mHybridScales(i) * solution(i): Next i
    Dim lambda As Double: lambda = mHybridSeed(size) + mHybridScales(size) * solution(size)
    If mAcceptedRevision <> mHybridOriginRevision Then Err.Raise 5, "CCoupledNewtonSolver", "Принятая история изменилась во время гибридного решения."
    SetControlledLoad lambda, load
    If Not Evaluate(field, False) Then
        If mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
            RecordFailure mFailureCode, "Не удалось записать материальную пробу итогового гибридного поля."
            Exit Function
        End If
        GoTo RestoreFailure
    End If
    Dim tolerance As Double: tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    If EquilibriumNorm(load, field) > tolerance Or Not GaugesVerified(tolerance) Then GoTo RestoreFailure
    mCandidate = field: PublishCandidate lambda
    mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & "; равновесие подтверждено после гибридного резерва" & _
        " | Исходная численная проба: " & original.DiagnosticDetails
    TryHybridControlledField = True
    Exit Function
RestoreFailure:
    mHybridActive = False: mCandidateReady = False: Set mMeta = original
    mExternal = external: mResidual = residual: mMatrix = matrix: mTrialResultants = resultants
    mLastNorm = originalNorm: mCurrentControlLambda = originalLambda: mFailureCode = originalCode: mMaterialPoint = originalPoint
    For i = 1 To COMPONENT_COUNT: load(i) = originalLoad(i): Next i
    If mHybridEvaluations > 0 Then mMeta.DiagnosticDetails = DiagnosticSummary() & _
        "; гибридный резерв не подтвердил равновесие | Исходная численная проба: " & original.DiagnosticDetails
    Exit Function
Failed:
    Dim errorNumber As Long, errorSource As String, errorDescription As String
    errorNumber = Err.Number: errorSource = Err.Source: errorDescription = Err.Description
    mHybridActive = False
    If errorNumber = 6 Or errorNumber = 11 Then Resume RestoreFailure
    Err.Raise errorNumber, errorSource, errorDescription
End Function

' Разрешает контроллеру запросить полный разностный резерв только для пространственной истории без спектральных границ.
' Псевдодуга и действующая мера дополнительно проверяются в самой управляемой пробе.
Friend Property Get SupportsHybridRecovery() As Boolean
    If Not mUseSpatialHistory Then Exit Property
    SupportsHybridRecovery = Not (HybridResponseHasPrincipalBoundaries(mResponse) Or _
        HybridResponseHasPrincipalBoundaries(mPlusResponse) Or HybridResponseHasPrincipalBoundaries(mMinusResponse))
End Property

' Проверяет ограничение резерва для отдельной центральной или боковой истории, с арматурой и без неё.
Private Function HybridResponseHasPrincipalBoundaries(ByVal response As ICoupledPointResponse) As Boolean
    If response Is Nothing Then Exit Function
    If TypeOf response Is CCoupledReinforcedResponse Then
        Dim reinforced As CCoupledReinforcedResponse: Set reinforced = response
        HybridResponseHasPrincipalBoundaries = reinforced.HasPrincipalBoundaries
    ElseIf TypeOf response Is CCoupledConcreteResponse Then
        Dim concrete As CCoupledConcreteResponse: Set concrete = response
        HybridResponseHasPrincipalBoundaries = concrete.HasPrincipalBoundaries
    End If
End Function

' Готовит неизменные масштабы геометрических координат только при первом запросе резерва.
' Повторная Initialize сбрасывает подготовку; нагрузка и материальная история в масштабы не входят.
Private Function PrepareHybridScales() As Boolean
    If mHybridScalesReady Then PrepareHybridScales = True: Exit Function
    ReDim mHybridScales(1 To mDegreeCount + 1)
    Dim unit() As Double: ReDim unit(1 To mDegreeCount)
    Dim strain(1 To COMPONENT_COUNT) As Double, slope(1 To COMPONENT_COUNT) As Double
    Dim degree As Long, point As Long, component As Long, largest As Double
    For degree = 1 To mGaugeStart - 1
        unit(degree) = 1#: largest = 0#
        For point = 1 To mPointCount
            mSection.PointStrain point, unit, strain: mSection.AxialGradientAtPoint point, unit, slope
            For component = 1 To COMPONENT_COUNT
                largest = GeomMax(largest, Abs(strain(component)) + SPATIAL_HALF_STEP * Abs(slope(component)))
            Next component
        Next point
        If largest = 0# Then Exit Function
        mHybridScales(degree) = HYBRID_STRAIN_SCALE / largest: unit(degree) = 0#
    Next degree
    For degree = mGaugeStart To mDegreeCount + 1: mHybridScales(degree) = HYBRID_AUXILIARY_SCALE: Next degree
    mHybridScalesReady = True: PrepareHybridScales = True
End Function

' Читающий интерфейс полной управляемой системы в подготовленных координатах.
' Пробы не записывают материал; итоговый кандидат проходит отдельную штатную проверку перед публикацией.
Private Function ICoupledEquationSystem_TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
    If Not mHybridActive Or mAcceptedRevision <> mHybridOriginRevision Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Нет действующего неизменённого контекста гибридной пробы."
    Dim size As Long: size = mDegreeCount + 1
    If LBound(coordinates) <> 1 Or UBound(coordinates) <> size Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Размер гибридной координаты не совпадает с системой равновесия."
    Dim field() As Double, load(1 To COMPONENT_COUNT) As Double, lambda As Double, i As Long, j As Long
    ReDim field(1 To mDegreeCount): ReDim residual(1 To size)
    For i = 1 To mDegreeCount: field(i) = mHybridSeed(i) + mHybridScales(i) * coordinates(i): Next i
    lambda = mHybridSeed(size) + mHybridScales(size) * coordinates(size)
    SetControlledLoad lambda, load
    If Not Evaluate(field, withJacobian, recordTrial:=False) Then Exit Function
    For i = 1 To mDegreeCount
        residual(i) = mResidual(i)
        If i >= mGaugeStart Then residual(i) = residual(i) * mReferenceModulus
    Next i
    residual(size) = ControlResidual(field)
    If withJacobian Then
        BuildControlledMatrix jacobian, size
        For j = 1 To size
            For i = 1 To size
                jacobian(j, i) = jacobian(j, i) * mHybridScales(j)
                If i >= mGaugeStart And i <= mDegreeCount Then jacobian(j, i) = jacobian(j, i) * mReferenceModulus
            Next i
        Next j
    End If
    ICoupledEquationSystem_TryEvaluate = True
End Function

' Проверяет поправку служебного множителя без записи материальной пробы.
' Поле передаётся штатной ветви принятия только после полной проверки равновесия.
Private Function TryWarpGaugeEquilibrium(ByRef field() As Double, ByRef load() As Double, _
        ByVal tolerance As Double) As Boolean
    If Not GaugesVerified(tolerance) Or Abs(ControlResidual(field)) > tolerance Then Exit Function
    If mSection.ResultantResidualNorm(mTrialResultants, load) > tolerance Then Exit Function
    Dim gauge As Long, change As Double
    If Not mSection.TryWarpGaugeCorrection(mResidual, tolerance, gauge, change) Then Exit Function

    Dim savedGauge As Double, savedNorm As Double, savedResidual() As Double, savedResultants() As Double
    Dim savedCode As ESolverFailureCode, savedPoint As Long, savedLastPoint As Long
    Dim savedSide As Long, savedFailures As Long
    Dim failureNumber As Long, failureSource As String, failureDescription As String
    savedGauge = field(gauge): savedNorm = mLastNorm
    savedResidual = mResidual: savedResultants = mTrialResultants
    savedCode = mFailureCode: savedPoint = mMaterialPoint: savedLastPoint = mLastFailedPoint
    savedSide = mLastFailedSide: savedFailures = mMaterialFailures
    On Error GoTo ProbeFailed
    field(gauge) = savedGauge + change
    If field(gauge) = savedGauge Then Exit Function ' Поправка меньше точности представления координаты.
    If Evaluate(field, False, recordTrial:=False) Then
        If EquilibriumNorm(load, field) <= tolerance And GaugesVerified(tolerance) Then
            TryWarpGaugeEquilibrium = True
            Exit Function
        End If
    End If

RestoreProbe:
    On Error GoTo 0
    field(gauge) = savedGauge
    mResidual = savedResidual: mTrialResultants = savedResultants: mLastNorm = savedNorm
    mFailureCode = savedCode: mMaterialPoint = savedPoint: mLastFailedPoint = savedLastPoint
    mLastFailedSide = savedSide: mMaterialFailures = savedFailures
    If failureNumber <> 0 Then Err.Raise failureNumber, failureSource, failureDescription
    Exit Function

ProbeFailed:
    If Err.Number <> 6 And Err.Number <> 11 Then
        failureNumber = Err.Number: failureSource = Err.Source: failureDescription = Err.Description
    End If
    Resume RestoreProbe
End Function

' Находит наибольшую невязку уравнений равновесия в МПа без уравнений закрепления.
Private Function LargestEquilibriumResidualRow() As Long
    Dim i As Long, row As Long: row = 1
    For i = 2 To mGaugeStart - 1
        If Abs(mResidual(i)) > Abs(mResidual(row)) Then row = i
    Next i
    LargestEquilibriumResidualRow = row
End Function

' Сравнивает действие касательной расширенной системы с последней вычисленной
' пробой при той же принятой истории. Дополнительного материального расчёта нет.
Private Function DirectionMismatchDetails(ByVal row As Long, ByVal baseEquation As Double, _
        ByVal measured As Double, ByRef correction() As Double) As String
    If measured = 0# Or Sgn(measured) <> Sgn(baseEquation) Then Exit Function
    Dim i As Long, calculated As Double
    calculated = -mDirectionEquations(row) * correction(mDegreeCount + 1) ' Вклад изменения множителя нагрузки в невязку.
    For i = 1 To mDegreeCount
        calculated = calculated + mMatrix(i, row) * correction(i)
    Next i
    If Sgn(calculated) <> -Sgn(baseEquation) Then Exit Function
    DirectionMismatchDetails = "; уравнение " & CStr(row) & " (" & mSection.EquilibriumEquationDescription(row) & _
        "); действие касательной, МПа: " & CStr(calculated) & _
        "; фактическое изменение на единицу доли поправки, МПа: " & CStr(measured)
End Function


' Использует равновесие у перехода только как первое приближение той же
' пробы. Контроль и история исходные, промежуточное поле не принимается.
Private Function TrySameControlSeed(ByRef field() As Double, ByRef lambda As Double, ByRef correction() As Double, _
        ByVal fraction As Double, ByVal tolerance As Double, ByVal absoluteTolerance As Double, _
        ByVal relativeTolerance As Double, ByVal iterationLimit As Long) As Boolean
    Dim originalProjection() As Double, originalTarget As Double
    originalProjection = mControlProjection: originalTarget = mTargetControl
    If Not RecoverYieldStep(field, lambda, correction, fraction, tolerance, absoluteTolerance, relativeTolerance, iterationLimit) Then Exit Function
    field = mCandidate: lambda = mTrialLambda
    mControlProjection = originalProjection: mTargetControl = originalTarget
    DiscardStep
    mRecoveredControlReady = False: Erase mRecoveredProjection
    TrySameControlSeed = True
End Function

' Ищет соседнюю траекторию только после наблюдаемой смены ветви на
' отклонённой поправке. Принятые буферы материала остаются исходными.
Private Function RecoverYieldStep(ByRef field() As Double, ByVal lambda As Double, ByRef correction() As Double, _
        ByVal fraction As Double, ByVal tolerance As Double, ByVal absoluteTolerance As Double, _
        ByVal relativeTolerance As Double, Optional ByVal iterationLimit As Long = 0) As Boolean
    If Not (TypeOf mResponse Is CCoupledReinforcedResponse And TypeOf mPlusResponse Is CCoupledReinforcedResponse And _
            TypeOf mMinusResponse Is CCoupledReinforcedResponse) Then Exit Function
    Dim ends() As Double, i As Long
    ReDim ends(1 To mDegreeCount)
    For i = 1 To mDegreeCount: ends(i) = field(i) + fraction * correction(i): Next i
    Dim station As TYieldStation
    If Not FirstYieldCrossing(field, ends, station) Then Exit Function
    mYieldRecoveryTried = True

    Dim previousProjection() As Double, previousTarget As Double
    previousProjection = mControlProjection: previousTarget = mTargetControl
    Dim result() As Double, trialLambda As Double, newProjection() As Double
    Dim errorNumber As Long, errorSource As String, errorText As String
    On Error GoTo Failed
    mControlActive = False
    If TraceYieldCrossing(station, field, ends, lambda, lambda + fraction * correction(mDegreeCount + 1), _
            tolerance, absoluteTolerance, relativeTolerance, True, result, trialLambda, newProjection, iterationLimit) Then
        Dim load(1 To COMPONENT_COUNT) As Double, coordinate As Double, trialTolerance As Double
        For i = 1 To mGaugeStart - 1: coordinate = coordinate + newProjection(i) * result(i): Next i
        mControlProjection = newProjection: mTargetControl = coordinate
        mControlActive = True
        SetControlledLoad trialLambda, load
        trialTolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
        If Evaluate(result, False) Then
            If EquilibriumNorm(load, result) <= trialTolerance And GaugesVerified(trialTolerance) Then
                mRecoveredProjection = newProjection: mRecoveredControlReady = True
                mCandidate = result
                mConsistentControl = (station.PrincipalIndex = 0)
                PublishCandidate trialLambda
                If station.PrincipalIndex = 0 Then
                    mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & "; равновесие восстановлено после смены ветви бетона"
                Else
                    mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & "; равновесие восстановлено у знаковой границы ASD; точка: " & _
                        CStr(station.Point) & "; станция: " & CStr(station.Side) & "; главное напряжение: " & CStr(station.PrincipalIndex)
                End If
                RecoverYieldStep = True
                Exit Function
            End If
        End If
    End If
Restore:
    mControlActive = True
    mControlProjection = previousProjection: mTargetControl = previousTarget
    On Error GoTo 0
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorText
    Exit Function
Failed:
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    Resume Restore
End Function

' Выбирает ближайшее пересечение CDP или знаковой границы ASD на действительной отклонённой поправке.
' Для CDP дополнительно подтверждает смену упругой и пластической ветвей локальным возвратом.
Private Function FirstYieldCrossing(ByRef field() As Double, ByRef ends() As Double, ByRef first As TYieldStation) As Boolean
    Dim probe As TYieldStation, fraction As Double, nearest As Double
    Dim point As Long, side As Long, principalIndex As Long
    nearest = 1#
    For point = 1 To mPointCount
        For side = -1 To 1
            For principalIndex = 0 To 3
                If ReadYieldCrossing(point, side, field, ends, probe, principalIndex) Then
                    If LocateYieldCoordinate(probe, field, ends, 0#, fraction) Then
                        If Not FirstYieldCrossing Or fraction < nearest Then
                            nearest = fraction: first = probe: FirstYieldCrossing = True
                        End If
                    End If
                End If
            Next principalIndex
        Next side
    Next point
End Function

' Проверяет поверхность CDP или выбранную знаковую границу ASD и сохраняет исходные значения в МПа.
Private Function ReadYieldCrossing(ByVal point As Long, ByVal side As Long, ByRef field() As Double, _
        ByRef ends() As Double, ByRef station As TYieldStation, Optional ByVal principalIndex As Long = 0) As Boolean
    Dim response As CCoupledReinforcedResponse: Set response = YieldResponse(side)
    Dim leftStrain(1 To COMPONENT_COUNT) As Double, rightStrain(1 To COMPONENT_COUNT) As Double
    Dim unused(1 To COMPONENT_COUNT) As Double, leftF As Double, rightF As Double
    YieldStrain point, side, field, leftStrain
    YieldStrain point, side, ends, rightStrain
    If principalIndex > 0 Then
        If Not response.TryReadPrincipalBoundary(point, leftStrain, principalIndex, leftF, unused, False) Then Exit Function
        If Not response.TryReadPrincipalBoundary(point, rightStrain, principalIndex, rightF, unused, False) Then Exit Function
    Else
        If Not response.TryReadTrialYield(point, leftStrain, leftF, unused, False) Then Exit Function
        If Not response.TryReadTrialYield(point, rightStrain, rightF, unused, False) Then Exit Function
    End If
    If (leftF < 0#) = (rightF < 0#) Or leftF = 0# Or rightF = 0# Then Exit Function
    If principalIndex = 0 Then
        Dim leftPlastic As Boolean, rightPlastic As Boolean, principal As Long, leftSteps As Long, rightSteps As Long
        If Not response.TryReadConcreteBranch(point, leftStrain, leftPlastic, principal, leftSteps) Then Exit Function
        If Not response.TryReadConcreteBranch(point, rightStrain, rightPlastic, principal, rightSteps) Then Exit Function
        If leftSteps > 1 Or rightSteps > 1 Or leftPlastic = rightPlastic Then Exit Function
        If leftPlastic <> (leftF > 0#) Or rightPlastic <> (rightF > 0#) Then Exit Function
    End If
    station.Point = point: station.Side = side: station.PrincipalIndex = principalIndex
    station.FromValue = leftF: station.ToValue = rightF
    ReadYieldCrossing = True
End Function

' Читает координату и производную именно выбранной материальной границы без изменения принятой истории.
Private Function ReadTransitionValue(ByRef station As TYieldStation, ByRef strain() As Double, ByRef value As Double, _
        ByRef gradient() As Double, Optional ByVal needGradient As Boolean = True) As Boolean
    Dim response As CCoupledReinforcedResponse: Set response = YieldResponse(station.Side)
    If station.PrincipalIndex > 0 Then
        ReadTransitionValue = response.TryReadPrincipalBoundary(station.Point, strain, station.PrincipalIndex, value, gradient, needGradient)
    Else
        ReadTransitionValue = response.TryReadTrialYield(station.Point, strain, value, gradient, needGradient)
    End If
End Function

' Передаёт бетон за центральной или боковой материальной историей.
Private Function YieldResponse(ByVal side As Long) As CCoupledReinforcedResponse
    Select Case side
        Case -1: Set YieldResponse = mMinusResponse
        Case 0: Set YieldResponse = mResponse
        Case 1: Set YieldResponse = mPlusResponse
    End Select
End Function

' Вычисляет деформацию или её направленное приращение на станции z/L.
Private Sub YieldStrain(ByVal point As Long, ByVal side As Long, ByRef field() As Double, ByRef strain() As Double)
    mSection.PointStrain point, field, strain
    If side = 0 Then Exit Sub
    Dim slope(1 To COMPONENT_COUNT) As Double, i As Long
    mSection.AxialGradientAtPoint point, field, slope
    For i = 1 To COMPONENT_COUNT: strain(i) = strain(i) + side * SPATIAL_HALF_STEP * slope(i): Next i
End Sub

' Делит фактический отрезок до заданной материальной координаты, лежащей между концами.
Private Function LocateYieldCoordinate(ByRef station As TYieldStation, ByRef field() As Double, _
        ByRef ends() As Double, ByVal target As Double, ByRef fraction As Double) As Boolean
    Dim response As CCoupledReinforcedResponse: Set response = YieldResponse(station.Side)
    Dim leftStrain(1 To COMPONENT_COUNT) As Double, rightStrain(1 To COMPONENT_COUNT) As Double
    Dim strain(1 To COMPONENT_COUNT) As Double, unused(1 To COMPONENT_COUNT) As Double
    YieldStrain station.Point, station.Side, field, leftStrain
    YieldStrain station.Point, station.Side, ends, rightStrain
    Dim low As Double, high As Double, value As Double, j As Long, i As Long
    high = 1#
    For j = 1 To 40
        fraction = 0.5 * (low + high)
        For i = 1 To COMPONENT_COUNT: strain(i) = leftStrain(i) + fraction * (rightStrain(i) - leftStrain(i)): Next i
        If Not ReadTransitionValue(station, strain, value, unused, False) Then Exit Function
        If Abs(value - target) <= YIELD_ROOT_TOLERANCE Then
            LocateYieldCoordinate = True
            Exit Function
        End If
        If (value < target) = (station.FromValue < target) Then low = fraction Else high = fraction
    Next j
End Function

' Для CDP начинает со стороны исходного поля, для ASD сначала проверяет сторону обнаруженного изменения знака.
' Вторая сторона остаётся запасной пробой; малый отступ отделяет обе цели от границы округления.
Private Function TraceYieldCrossing(ByRef station As TYieldStation, ByRef field() As Double, ByRef ends() As Double, _
        ByVal fromLambda As Double, ByVal toLambda As Double, ByVal tolerance As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal allowSuccessor As Boolean, _
        ByRef result() As Double, ByRef lambda As Double, ByRef projection() As Double, _
        Optional ByVal iterationLimit As Long = 0) As Boolean
    Dim target As Double, fraction As Double, trial As Long, i As Long, anchor() As Double
    target = Sgn(station.FromValue) * GeomMin(0.125 * tolerance, _
        0.25 * GeomMin(Abs(station.FromValue), Abs(station.ToValue)))
    If station.PrincipalIndex > 0 Then target = -target ' Повтор принятого состояния перед нулём не заменяет переход на соседнюю спектральную ветвь.
    ReDim anchor(1 To mDegreeCount)
    For trial = 1 To 2
        If iterationLimit > 0 And mIterations >= iterationLimit Then Exit Function
        If LocateYieldCoordinate(station, field, ends, target, fraction) Then
            For i = 1 To mDegreeCount: anchor(i) = field(i) + fraction * (ends(i) - field(i)): Next i
            If FollowYieldCoordinate(station, anchor, fromLambda + fraction * (toLambda - fromLambda), _
                    target, absoluteTolerance, relativeTolerance, allowSuccessor, result, lambda, projection, iterationLimit) Then
                TraceYieldCrossing = True
                Exit Function
            End If
        End If
        target = -target
    Next trial
End Function

' Ведёт короткую серию читающих проб при заданной материальной координате. Две правые
' части одной факторизации задают поправку и направление её точного удержания.
' Следующая станция допускается по фактической ошибке текущего шага.
Private Function FollowYieldCoordinate(ByRef station As TYieldStation, ByRef startField() As Double, _
        ByVal lambda As Double, ByVal target As Double, ByVal absoluteTolerance As Double, _
        ByVal relativeTolerance As Double, ByVal allowSuccessor As Boolean, _
        ByRef result() As Double, ByRef resultLambda As Double, ByRef projection() As Double, _
        Optional ByVal iterationLimit As Long = 0) As Boolean
    Dim field() As Double, trial() As Double, load(1 To COMPONENT_COUNT) As Double
    Dim correction() As Double, nullDirection() As Double, eventRow() As Double, nextStation As TYieldStation
    Dim strainRate(1 To COMPONENT_COUNT) As Double, i As Long, iteration As Long, damping As Long
    Dim baseNorm As Double, norm As Double, value As Double, beta As Double, fraction As Double, trialLambda As Double
    Dim trialTolerance As Double
    Dim row As Long, baseRow As Double, expected As Double, previousTrySuccessor As Boolean, accepted As Boolean
    field = startField
    ReDim trial(1 To mDegreeCount)
    For iteration = 1 To YIELD_TRACE_ITERATIONS
        If iterationLimit > 0 Then
            If mIterations >= iterationLimit Then Exit Function
            mIterations = mIterations + 1 ' Направляющая итерация расходует общий остаток, даже при отказе.
        End If
        SetControlledLoad lambda, load
        If Not Evaluate(field, True, recordTrial:=False, consistentBranch:=(station.PrincipalIndex = 0)) Then Exit Function
        baseNorm = EquilibriumNorm(load, field)
        row = LargestEquilibriumResidualRow(): baseRow = mResidual(row)
        If Not YieldDirections(station, field, target, value, eventRow, correction, nullDirection) Then Exit Function
        baseNorm = GeomMax(baseNorm, Abs(value - target))
        YieldStrain station.Point, station.Side, nullDirection, strainRate
        fraction = 1#: accepted = False
        For damping = 0 To YIELD_DAMPING_STEPS
            If ProjectYieldTrial(station, target, field, correction, nullDirection, strainRate, fraction, trial, beta) Then
                trialLambda = lambda + fraction * correction(mDegreeCount + 1) + beta * nullDirection(mDegreeCount + 1)
                SetControlledLoad trialLambda, load
                If Evaluate(trial, False, recordTrial:=False) Then
                    norm = EquilibriumNorm(load, trial)
                    trialTolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
                    If GaugesVerified(trialTolerance) Then
                        If norm <= trialTolerance And trialLambda > 0# Then
                            ReDim projection(1 To mDegreeCount)
                            For i = 1 To mGaugeStart - 1: projection(i) = eventRow(i) / mReferenceModulus: Next i
                            result = trial: resultLambda = trialLambda
                            FollowYieldCoordinate = True
                            Exit Function
                        End If
                        accepted = (norm < baseNorm * (1# - MIN_PROGRESS * fraction))
                    End If
                    If accepted Then Exit For
                    If allowSuccessor And Not previousTrySuccessor And damping = 0 Then
                        expected = -mDirectionEquations(row) * (trialLambda - lambda)
                        For i = 1 To mDegreeCount: expected = expected + mMatrix(i, row) * (trial(i) - field(i)): Next i
                        If FindYieldDeparture(station, field, trial, row, mResidual(row) - baseRow - expected, nextStation) Then
                            previousTrySuccessor = True
                            If TraceYieldCrossing(nextStation, field, trial, lambda, trialLambda, trialTolerance, _
                                    absoluteTolerance, relativeTolerance, False, result, resultLambda, projection, iterationLimit) Then
                                FollowYieldCoordinate = True
                                Exit Function
                            End If
                            ' Пробы другой ветви заменили рабочую касательную. Данное направление уже рассчитано.
                        End If
                    End If
                End If
            End If
            fraction = fraction * 0.5
        Next damping
        If Not accepted Then Exit Function
        field = trial: lambda = trialLambda
        If iterationLimit = 0 Then mIterations = mIterations + 1
    Next iteration
End Function

' Заменяет одно уравнение управления градиентом выбранной материальной координаты по полю.
' Четыре уравнения однозначности и столбец нагрузки сохраняются.
Private Function YieldDirections(ByRef station As TYieldStation, ByRef field() As Double, ByVal target As Double, _
        ByRef value As Double, ByRef eventRow() As Double, ByRef correction() As Double, ByRef nullDirection() As Double) As Boolean
    Dim response As CCoupledReinforcedResponse: Set response = YieldResponse(station.Side)
    Dim strain(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double
    YieldStrain station.Point, station.Side, field, strain
    If Not ReadTransitionValue(station, strain, value, gradient) Then Exit Function
    Dim size As Long, i As Long, j As Long, basis() As Double, augmented() As Double
    size = mDegreeCount + 1
    BuildControlledMatrix augmented, size
    ReDim basis(1 To mDegreeCount), correction(1 To size), nullDirection(1 To size), eventRow(1 To size)
    For i = 1 To mGaugeStart - 1
        basis(i) = 1#
        YieldStrain station.Point, station.Side, basis, strain
        For j = 1 To COMPONENT_COUNT: eventRow(i) = eventRow(i) + gradient(j) * strain(j): Next j
        basis(i) = 0#
        augmented(i, size) = eventRow(i)
    Next i
    For i = 1 To mDegreeCount: correction(i) = -mResidual(i): Next i
    correction(size) = target - value
    nullDirection(size) = 1#
    If Not mSystem.Solve(augmented, correction, size) Then Exit Function
    YieldDirections = mSystem.SolveFactored(nullDirection)
End Function

' Удерживает фактическую материальную координату на каждой доле поправки. Скалярная
' доводка вдоль нулевого направления сохраняет линейные силовые строки.
Private Function ProjectYieldTrial(ByRef station As TYieldStation, ByVal target As Double, ByRef field() As Double, _
        ByRef correction() As Double, ByRef nullDirection() As Double, ByRef strainRate() As Double, _
        ByVal fraction As Double, ByRef trial() As Double, ByRef beta As Double) As Boolean
    Const MAX_NULL_SHIFT As Double = 0.002 ' МПа по координате Ftrial, а не масштаб нагрузки.
    Dim response As CCoupledReinforcedResponse: Set response = YieldResponse(station.Side)
    Dim strain(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double
    Dim i As Long, iteration As Long, slope As Double, value As Double
    beta = 0#
    For iteration = 1 To 7
        For i = 1 To mDegreeCount: trial(i) = field(i) + fraction * correction(i) + beta * nullDirection(i): Next i
        YieldStrain station.Point, station.Side, trial, strain
        If Not ReadTransitionValue(station, strain, value, gradient) Then Exit Function
        If Abs(value - target) <= YIELD_ROOT_TOLERANCE Then
            ProjectYieldTrial = True
            Exit Function
        End If
        slope = 0#
        For i = 1 To COMPONENT_COUNT: slope = slope + gradient(i) * strainRate(i): Next i
        If slope < 0.5 Or slope > 1.5 Then Exit Function
        beta = beta + (target - value) / slope
        If Abs(beta) > MAX_NULL_SHIFT Then Exit Function
    Next iteration
End Function

' Выбирает станцию, на которой фактический переход объясняет рост
' определяющего уравнения. Станции без пересечения не меняют координату.
Private Function FindYieldDeparture(ByRef held As TYieldStation, ByRef field() As Double, ByRef trial() As Double, _
        ByVal row As Long, ByVal mismatch As Double, ByRef selected As TYieldStation) As Boolean
    If mismatch = 0# Then Exit Function
    Dim point As Long, side As Long, principalIndex As Long, crossing As TYieldStation
    Dim departure As Double, largest As Double
    For point = 1 To mPointCount
        For side = -1 To 1
            For principalIndex = 0 To 3
                If point <> held.Point Or side <> held.Side Or principalIndex <> held.PrincipalIndex Then
                    If ReadYieldCrossing(point, side, field, trial, crossing, principalIndex) Then
                        If YieldPointDeviation(crossing, field, trial, row, departure) Then
                            If Sgn(departure) = Sgn(mismatch) And Abs(departure) > largest Then
                                largest = Abs(departure): selected = crossing
                                FindYieldDeparture = True
                            End If
                        End If
                    End If
                End If
            Next principalIndex
        Next side
    Next point
End Function

' Проецирует поправку к действию касательной тем же оператором, что и
' сборка. Для стержня используется уже собранный отклик стали и бетона.
Private Function YieldPointDeviation(ByRef station As TYieldStation, ByRef field() As Double, _
        ByRef trial() As Double, ByVal row As Long, ByRef departure As Double) As Boolean
    Dim response As ICoupledPointResponse: Set response = YieldResponse(station.Side)
    Dim beforeStrain(1 To COMPONENT_COUNT) As Double, afterStrain(1 To COMPONENT_COUNT) As Double
    Dim beforeStress(1 To COMPONENT_COUNT) As Double, afterStress(1 To COMPONENT_COUNT) As Double
    Dim tangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
    YieldStrain station.Point, station.Side, field, beforeStrain
    YieldStrain station.Point, station.Side, trial, afterStrain
    If response.EvaluatePoint(station.Point, beforeStrain, beforeStress, tangent, ctmFullBranchConsistent, False) <> sfcNone Then Exit Function
    If response.EvaluatePoint(station.Point, afterStrain, afterStress, tangent, ctmStressOnly, False) <> sfcNone Then Exit Function
    Dim center(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double
    Dim resultants(1 To COMPONENT_COUNT) As Double, residual() As Double, unused() As Double
    Dim j As Long, i As Long, value As Double
    For i = 1 To COMPONENT_COUNT
        value = afterStress(i) - beforeStress(i)
        For j = 1 To COMPONENT_COUNT: value = value - tangent(i, j) * (afterStrain(j) - beforeStrain(j)): Next j
        If station.Side = 0 Then center(i) = value Else gradient(i) = station.Side * value / (2# * SPATIAL_HALF_STEP)
    Next i
    ReDim residual(1 To mDegreeCount)
    mSection.AddSpatialPointEquilibrium station.Point, center, mTangent, gradient, mGradientTangent, _
        mSlopeTangent, residual, unused, False, resultants
    departure = residual(row)
    YieldPointDeviation = True
End Function

' Публикует только пробное поле, чья невязка проверена по всей системе.
Private Sub PublishCandidate(ByVal lambda As Double)
    mTrialLambda = lambda
    mCandidateReady = True
    mMeta.SetResult rsSuccess, rcNone, rkDirectState
    mMeta.DiagnosticDetails = DiagnosticSummary()
End Sub

' Обновляет шесть внешних усилий и их заранее нормированные уравнения.
Private Sub SetControlledLoad(ByVal lambda As Double, ByRef load() As Double)
    mCurrentControlLambda = lambda
    Dim i As Long
    For i = 1 To COMPONENT_COUNT
        load(i) = mBaseLoad(i) + lambda * mLoadDirection(i)
    Next i
    For i = 1 To mDegreeCount
        mExternal(i) = mBaseEquations(i) + lambda * mDirectionEquations(i)
    Next i
End Sub

' Решает поправку Ньютона. При управлении деформацией добавляет столбец
' производной нагрузки по lambda и строку выбранной меры деформации.
Private Function SolveCorrection(ByRef field() As Double, ByRef correction() As Double) As Boolean
    Dim i As Long
    If Not mControlActive Then
        correction = mResidual
        For i = 1 To mDegreeCount: correction(i) = -correction(i): Next i
        SolveCorrection = mSystem.Solve(mMatrix, correction, mDegreeCount)
        Exit Function
    End If

    Dim augmented() As Double, size As Long
    size = mDegreeCount + 1
    ReDim correction(1 To size)
    BuildControlledMatrix augmented, size
    For i = 1 To mDegreeCount: correction(i) = -mResidual(i): Next i
    correction(size) = -ControlResidual(field)
    SolveCorrection = mSystem.Solve(augmented, correction, size)
End Function

' Пересматривает ветви стали без изменения принятой истории и базовой матрицы.
' После отказа локального направления использует касательную бетонного остатка.
Private Sub RefineSteelCorrection(ByRef field() As Double, ByRef correction() As Double, ByVal useFieldTangent As Boolean)
    If Not useFieldTangent Then
        mSteelBranchBuilds = mSteelBranchBuilds + mSteelCorrector.Refine(field, mMatrix, mResidual, correction)
        Exit Sub
    End If
    Dim concrete() As Double
    mSteelFieldTangents = mSteelFieldTangents + 1
    If Not mSteelCorrector.TryConcreteTangent(field, mExternal, concrete, mEvaluations) Then Exit Sub
    mSteelBranchBuilds = mSteelBranchBuilds + mSteelCorrector.Refine(field, concrete, mResidual, correction, concreteMatrix:=True)
End Sub

' Завершает близкую к равновесию пространственную пробу при том же
' управлении. Малая поправка нужна там, где демпфирование Ньютона
' не уменьшает действительный остаток после возврата материалов.
Private Function TrySmallStrainEquilibrium(ByRef field() As Double, ByVal lambda As Double, _
        ByRef load() As Double, ByVal norm As Double, ByVal tolerance As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByRef trialLambda As Double) As Boolean
    Dim force() As Double, forceResidual() As Double, equality() As Double, equalityResidual() As Double
    Dim scales() As Double, strain() As Double, correction() As Double, trial() As Double
    Dim optimizer As New CCoupledStepOptimizer, i As Long, row As Long, fraction As Double, radius As Double
    Dim trialTolerance As Double
    On Error GoTo UnavailableCorrection
    SetControlledLoad lambda, load
    If Not Evaluate(field, True, recordTrial:=False, refinedBranch:=True) Then Exit Function
    If Not BuildSmallStrainSystem(field, load, tolerance, force, forceResidual, equality, equalityResidual, scales) Then Exit Function
    mSection.CopySpatialStrainCoefficients SPATIAL_HALF_STEP, strain
    For row = 1 To UBound(strain, 2)
        For i = 1 To UBound(scales)
            strain(i, row) = strain(i, row) / scales(i)
        Next i
    Next row
    radius = SMALL_STRAIN_RADIUS * norm ' Столбцы касательной приведены к единичному максимуму.
    If Not optimizer.TryMinimizeStrain(force, forceResidual, equality, equalityResidual, strain, _
            radius, SMALL_STRAIN_FORCE_FRACTION * tolerance, correction) Then Exit Function

    For i = 1 To UBound(scales)
        correction(i) = correction(i) / scales(i)
    Next i
    ReDim trial(1 To mDegreeCount)
    fraction = 1#
    Do While fraction >= 0.125
        For i = 1 To mDegreeCount
            trial(i) = field(i) + fraction * correction(i)
        Next i
        trialLambda = lambda + fraction * correction(mDegreeCount + 1)
        SetControlledLoad trialLambda, load
        trialTolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
        If Evaluate(trial, False, recordTrial:=False) Then
            If EquilibriumNorm(load, trial) <= trialTolerance And GaugesVerified(trialTolerance) Then
                If Evaluate(trial, False) Then
                    If EquilibriumNorm(load, trial) <= trialTolerance And GaugesVerified(trialTolerance) Then
                        mCandidate = trial
                        TrySmallStrainEquilibrium = True
                        Exit Function
                    End If
                ElseIf mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
                    Exit Function
                End If
            End If
        ElseIf mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
            Exit Function
        End If
        fraction = fraction * 0.5
    Loop
    Exit Function

UnavailableCorrection:
    If Err.Number <> 6 And Err.Number <> 7 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Приводит касательную и правую часть к общей шкале координат.
' Шесть усилий проверяются как отдельные силовые строки, а связи
' и уравнение управления задаются точными равенствами.
Private Function BuildSmallStrainSystem(ByRef field() As Double, ByRef load() As Double, ByVal tolerance As Double, _
        ByRef force() As Double, ByRef forceResidual() As Double, ByRef equality() As Double, _
        ByRef equalityResidual() As Double, ByRef scales() As Double) As Boolean
    Dim matrix() As Double, residual() As Double, projection() As Double
    Dim size As Long, weakRows As Long, i As Long, j As Long, k As Long, row As Long
    Dim maximum As Double, factor As Double, normalized As Double, direct As Double, area As Double, lengthScale As Double
    size = mDegreeCount + 1: weakRows = mGaugeStart - 1
    BuildControlledMatrix matrix, size
    ReDim residual(1 To size), scales(1 To size)
    For i = 1 To mDegreeCount
        residual(i) = mResidual(i)
    Next i
    residual(size) = ControlResidual(field)
    For i = mGaugeStart To mDegreeCount
        residual(i) = residual(i) * mReferenceModulus
        For j = 1 To size
            matrix(j, i) = matrix(j, i) * mReferenceModulus
        Next j
    Next i
    For j = 1 To size
        maximum = 0#
        For i = 1 To size
            maximum = GeomMax(maximum, Abs(matrix(j, i)))
        Next i
        If maximum = 0# Then Exit Function
        scales(j) = maximum
        For i = 1 To size
            matrix(j, i) = matrix(j, i) / maximum
        Next i
    Next j
    ReDim force(1 To size, 1 To weakRows + COMPONENT_COUNT), forceResidual(1 To weakRows + COMPONENT_COUNT)
    ReDim equality(1 To size, 1 To GAUGE_COUNT + 1), equalityResidual(1 To GAUGE_COUNT + 1)
    For i = 1 To weakRows
        forceResidual(i) = residual(i)
        For j = 1 To size
            force(j, i) = matrix(j, i)
        Next j
    Next i
    For i = mGaugeStart To size
        row = i - weakRows: maximum = 0#
        For j = 1 To size
            maximum = GeomMax(maximum, Abs(matrix(j, i)))
        Next j
        If maximum = 0# Then Exit Function
        equalityResidual(row) = residual(i) / maximum
        For j = 1 To size
            equality(j, row) = matrix(j, i) / maximum
        Next j
    Next i

    mSection.CopyResultantResidualProjection projection
    mSection.CopyResultantResidualScales area, lengthScale
    For k = 1 To COMPONENT_COUNT
        row = weakRows + k: normalized = 0#
        For i = 1 To weakRows
            factor = projection(i, k)
            If factor <> 0# Then
                normalized = normalized + factor * residual(i)
                For j = 1 To size
                    force(j, row) = force(j, row) + factor * matrix(j, i)
                Next j
            End If
        Next i
        direct = (mTrialResultants(k) - load(k)) / area
        If k >= 2 And k <= 4 Then direct = direct / lengthScale
        If Abs(normalized - direct) > 0.001 * tolerance Then Exit Function ' Проверяем замыкание проекции на исходном поле.
        forceResidual(row) = direct
    Next k
    BuildSmallStrainSystem = True
End Function

' Добавляет неизвестный множитель и уравнение управления деформацией.
Private Sub BuildControlledMatrix(ByRef augmented() As Double, ByVal size As Long)
    Dim i As Long, j As Long
    ReDim augmented(1 To size, 1 To size)
    For i = 1 To mDegreeCount
        For j = 1 To mDegreeCount: augmented(j, i) = mMatrix(j, i): Next j
        augmented(size, i) = -mDirectionEquations(i)
        If i < mGaugeStart Then augmented(i, size) = mControlProjection(i) * mReferenceModulus
    Next i
    If mArcLoadProjection <> 0# Then augmented(size, size) = mArcLoadProjection * mReferenceModulus
End Sub

' Оценивает чувствительность множителя к невязке при том же принятом
' состоянии. Сопряжённая система даёт строку обратного оператора по lambda.
' Ошибка повторной сборки исключает принятие пробы.
Friend Function TrialLambdaUncertainty(ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByRef uncertainty As Double) As Boolean
    RequireCandidate
    If Not mControlActive Then Exit Function
    If Not Evaluate(mCandidate, True, True, recordTrial:=Not mUseSpatialHistory, consistentBranch:=mConsistentControl) Then
        mCandidateReady = False
        Exit Function
    End If
    Dim matrix() As Double, weights() As Double, i As Long, j As Long, size As Long
    size = mDegreeCount + 1
    BuildControlledMatrix matrix, size
    For i = mGaugeStart To mDegreeCount
        For j = 1 To size: matrix(j, i) = matrix(j, i) * mReferenceModulus: Next j
    Next i
    Dim temp As Double
    For i = 1 To size
        For j = i + 1 To size
            temp = matrix(j, i): matrix(j, i) = matrix(i, j): matrix(i, j) = temp
        Next j
    Next i
    ReDim weights(1 To size)
    weights(size) = 1#
    Dim adjoint As New CCoupledLinearSystem
    If Not adjoint.Solve(matrix, weights, size) Then Exit Function
    Dim norm As Double, residual As Double
    norm = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    For i = 1 To mDegreeCount
        residual = Abs(mResidual(i))
        If i >= mGaugeStart Then residual = residual * mReferenceModulus
        If residual > norm Then norm = residual
    Next i
    residual = Abs(ControlResidual(mCandidate))
    If residual > norm Then norm = residual
    uncertainty = 0#
    For i = 1 To size: uncertainty = uncertainty + Abs(weights(i)): Next i
    uncertainty = uncertainty * norm + 2E-12 * (1# + Abs(mTrialLambda))
    TrialLambdaUncertainty = True
End Function

' Приводит невязку контролируемой деформации к масштабу силовых уравнений.
Private Function ControlResidual(ByRef field() As Double) As Double
    Dim i As Long, value As Double
    value = -mTargetControl
    If mArcLoadProjection <> 0# Then value = value + mArcLoadProjection * mCurrentControlLambda
    For i = 1 To mGaugeStart - 1
        value = value + field(i) * mControlProjection(i)
    Next i
    ControlResidual = value * mReferenceModulus
End Function

' Публикует пробное поле и материальную историю после инженерной приёмки шага.
' Повторный вызов требует нового успешного решения.
Friend Sub AcceptStep()
    RequireCandidate
    Dim keepTransition As Boolean
    If mPrincipalIncomingCandidate Then
        If Not mPrincipalTransition Is Nothing Then keepTransition = mPrincipalTransition.MatchesOrigin(mAcceptedRevision, mAccepted)
    End If
    Dim completed As CCoupledPrincipalTransition
    If mPrincipalOutgoingCandidate Then
        If Not mPrincipalTransition Is Nothing Then
            If mPrincipalTransition.MatchesIncoming(mAcceptedRevision, mAccepted) Then Set completed = mPrincipalTransition
        End If
    End If
    Dim acceptedField() As Double: acceptedField = mCandidate
    AcceptMaterialTrials
    mAccepted = acceptedField
    mAcceptedRevision = mAcceptedRevision + 1
    mCandidateReady = False
    mHasRejectedDirection = False: Erase mRejectedDirection
    mHasTransitionSegment = False: mPrincipalIncomingCandidate = False
    If keepTransition Then
        If Not mPrincipalTransition.MatchesIncoming(mAcceptedRevision, mAccepted) Then Set mPrincipalTransition = Nothing
    Else
        Set mPrincipalTransition = Nothing
    End If
    Set mPrincipalDeparture = completed: mDepartureRevision = mAcceptedRevision: mPrincipalOutgoingCandidate = False
End Sub

' Отбрасывает вычисленную пробу после поиска или пользовательской отмены.
Friend Sub DiscardStep()
    mCandidateReady = False
    mPrincipalIncomingCandidate = False
    mPrincipalOutgoingCandidate = False
    If Not mPrincipalTransition Is Nothing Then
        If Not mPrincipalTransition.MatchesIncoming(mAcceptedRevision, mAccepted) Then Set mPrincipalTransition = Nothing
    End If
End Sub

' Передаёт независимые копии равновесного пробного поля и шести интегралов напряжений.
Friend Sub CopyTrial(ByRef field() As Double, ByRef resultants() As Double)
    RequireCandidate
    field = mCandidate
    resultants = mTrialResultants
End Sub

' Возвращает ранее принятое поле без пробных поправок.
Friend Sub CopyAccepted(ByRef field() As Double)
    field = mAccepted
End Sub

' Возвращает ревизию принятого состояния; пробные вычисления её не изменяют.
' Восстановление истории создаёт новые ревизии, даже если итоговые координаты совпали.
Friend Property Get AcceptedRevision() As Long
    AcceptedRevision = mAcceptedRevision
End Property

' Восстанавливает состояние после известных принятых полей. Напряжения
' и история вычисляются по той же последовательности без поиска НДС.
' Вызывающий передаёт поля этой секции с начала той же траектории.
Friend Function ReplayAcceptedFields(ByVal history As Collection, ByVal lastField As Long) As Boolean
    If history Is Nothing Then Err.Raise 5, "CCoupledNewtonSolver", "Для восстановления не сохранена история полей."
    If lastField < 0 Or lastField > history.Count Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Точка восстановления вне сохранённой истории полей."
    ResetHistory
    Dim index As Long, point As Long, field() As Double, failedSide As Long
    For index = 1 To lastField
        field = history(index)
        For point = 1 To mPointCount
            mSection.PointStrain point, field, mPointStrain
            If EvaluatePointResponses(point, field, ctmStressOnly, True, failedSide) <> sfcNone Then
                ResetHistory
                Exit Function
            End If
        Next point
        AcceptMaterialTrials
        mAccepted = field
        mAcceptedRevision = mAcceptedRevision + 1
    Next index
    ReplayAcceptedFields = True
End Function

' Продвигает материальные буферы после успешной пробы всех точек.
' Цикл проб гарантирует заполнение каждой истории до смены буфера.
Private Sub AcceptMaterialTrials()
    mResponse.AcceptTrial
    If mUseSpatialHistory Then
        mPlusResponse.AcceptTrial
        mMinusResponse.AcceptTrial
    End If
End Sub

' Возвращает независимый статус последнего шага.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Считает отклик всех точек от фиксированной принятой истории.
' При демпфировании собирает только невязку и интегралы. Читающая проба
' касательной сохраняет материальные буферы ранее подтверждённого поля.
Private Function Evaluate(ByRef field() As Double, ByVal withMatrix As Boolean, _
        Optional ByVal useGradientTangent As Boolean = False, Optional ByVal recordTrial As Boolean = True, _
        Optional ByVal consistentBranch As Boolean = False, Optional ByVal refinedBranch As Boolean = False) As Boolean
    mEvaluations = mEvaluations + 1
    mSection.StartEquations field, mExternal, mResidual, mMatrix, withMatrix
    Dim point As Long, useAxialGradient As Boolean, tangentMode As ECoupledTangentMode
    If Not mUseSpatialHistory Then useAxialGradient = mSection.HasAxialGradient(field)
    tangentMode = ctmFull
    If withMatrix And consistentBranch Then tangentMode = ctmFullBranchConsistent
    If withMatrix And refinedBranch Then tangentMode = ctmRefinedBranchConsistent
    If Not withMatrix Then
        If useAxialGradient Then tangentMode = ctmAxialGradient Else tangentMode = ctmStressOnly
    End If
    useGradientTangent = useGradientTangent And withMatrix And useAxialGradient
    If useGradientTangent Then mGradientBuilds = mGradientBuilds + 1
    Dim differenceScale As Double, plusOk As Boolean, minusOk As Boolean, failedSide As Long
    ReDim mTrialResultants(1 To COMPONENT_COUNT)
    For point = 1 To mPointCount
        mMaterialPoint = point
        mSection.PointStrain point, field, mPointStrain
        If useGradientTangent Then TakeGradientProbes point, field, differenceScale, plusOk, minusOk
        mFailureCode = EvaluatePointResponses(point, field, tangentMode, recordTrial, failedSide)
        If mFailureCode <> sfcNone Then
            mMaterialFailures = mMaterialFailures + 1: mLastFailedPoint = point: mLastFailedSide = failedSide
            Exit Function
        End If
        If mUseSpatialHistory Then
            BuildSpatialGradient withMatrix
            mSection.AddSpatialPointEquilibrium point, mStress, mTangent, mSpatialGradient, mGradientTangent, _
                mSlopeTangent, mResidual, mMatrix, withMatrix, mTrialResultants
        Else
            mSection.AddPointEquilibrium point, field, mStress, mTangent, mResidual, mMatrix, withMatrix, mTrialResultants, useAxialGradient
        End If
        If differenceScale > 0# Then
            If Not plusOk And Not minusOk Then
                mFailureCode = sfcNumericalFailure
                mMaterialFailures = mMaterialFailures + 1: mLastFailedPoint = point
                Exit Function
            End If
            If TryLocalGradientDerivative(point, field, plusOk, minusOk) Then
                mLocalGradientBuilds = mLocalGradientBuilds + 1
                mSection.AddPointGradientTangent point, mGradientTangent, mMatrix
            Else
                AddGradientDerivative point, differenceScale, plusOk, minusOk
            End If
        End If
    Next point
    mMaterialPoint = 0
    Evaluate = True
End Function

' Получает центральный отклик и отклики продольных точек при необходимости.
' Каждая история читает свой принятый буфер. Стороны вычисляются и при g=0,
' поскольку ранее различавшиеся истории могут давать разные напряжения.
Private Function EvaluatePointResponses(ByVal point As Long, ByRef field() As Double, _
        ByVal tangentMode As ECoupledTangentMode, ByVal recordTrial As Boolean, _
        ByRef failedSide As Long) As ESolverFailureCode
    failedSide = 0
    EvaluatePointResponses = mResponse.EvaluatePoint(point, mPointStrain, mStress, mTangent, tangentMode, recordTrial)
    If EvaluatePointResponses <> sfcNone Or Not mUseSpatialHistory Then Exit Function

    Dim gradient(1 To COMPONENT_COUNT) As Double, i As Long, offset As Double
    mSection.AxialGradientAtPoint point, field, gradient
    For i = 1 To COMPONENT_COUNT
        offset = SPATIAL_HALF_STEP * gradient(i) ' Приращение деформации от центра до продольной точки.
        mPlusStrain(i) = mPointStrain(i) + offset
        mMinusStrain(i) = mPointStrain(i) - offset
    Next i
    failedSide = 1
    EvaluatePointResponses = mPlusResponse.EvaluatePoint(point, mPlusStrain, mPlusStress, mPlusTangent, tangentMode, recordTrial)
    If EvaluatePointResponses <> sfcNone Then Exit Function
    failedSide = -1
    EvaluatePointResponses = mMinusResponse.EvaluatePoint(point, mMinusStrain, mMinusStress, mMinusTangent, tangentMode, recordTrial)
End Function

' Собирает секущую по z/L в МПа и её производные. Производная по
' центральной деформации берётся из разности материальных касательных,
' производная по уклону g — из их среднего.
Private Sub BuildSpatialGradient(ByVal withMatrix As Boolean)
    Dim i As Long, j As Long
    For i = 1 To COMPONENT_COUNT
        mSpatialGradient(i) = (mPlusStress(i) - mMinusStress(i)) / (2# * SPATIAL_HALF_STEP)
    Next i
    If Not withMatrix Then Exit Sub
    For j = 1 To COMPONENT_COUNT
        For i = 1 To COMPONENT_COUNT
            mGradientTangent(i, j) = (mPlusTangent(i, j) - mMinusTangent(i, j)) / (2# * SPATIAL_HALF_STEP)
            mSlopeTangent(i, j) = 0.5 * (mPlusTangent(i, j) + mMinusTangent(i, j))
        Next i
    Next j
End Sub

' При резком изменении касательной бетона на деформационном шаге уточняет
' производную фактического Dg по деформации. В g входят zz, xz и yz.
' Каждая проба читает ту же принятую историю и сохраняет базовый отклик.
Private Function TryLocalGradientDerivative(ByVal point As Long, ByRef field() As Double, _
        ByVal plusOk As Boolean, ByVal minusOk As Boolean) As Boolean
    If Not mControlActive Or point > mSection.ConcretePointCount Then Exit Function
    If Not plusOk Or Not minusOk Then Exit Function
    If Not HasLargeGradientTangentChange() Then Exit Function

    Dim gradient(1 To COMPONENT_COUNT) As Double, strain(1 To COMPONENT_COUNT) As Double
    Dim plusTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
    Dim minusTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
    Dim stress(1 To COMPONENT_COUNT) As Double, perturbation As Double
    Dim row As Long, column As Long
    mSection.AxialGradientAtPoint point, field, gradient
    For column = 1 To COMPONENT_COUNT: strain(column) = mPointStrain(column): Next column
    For column = 1 To COMPONENT_COUNT
        perturbation = LOCAL_GRADIENT_DIFFERENCE_STEP * (1# + Abs(strain(column)))
        strain(column) = mPointStrain(column) + perturbation
        If mResponse.EvaluatePoint(point, strain, stress, plusTangent, ctmAxialGradient, recordTrial:=False) <> sfcNone Then Exit Function
        strain(column) = mPointStrain(column) - perturbation
        If mResponse.EvaluatePoint(point, strain, stress, minusTangent, ctmAxialGradient, recordTrial:=False) <> sfcNone Then Exit Function
        strain(column) = mPointStrain(column)
        For row = 1 To COMPONENT_COUNT
            mGradientTangent(row, column) = ((plusTangent(row, 3) - minusTangent(row, 3)) * gradient(3) + _
                (plusTangent(row, 5) - minusTangent(row, 5)) * gradient(5) + _
                (plusTangent(row, 6) - minusTangent(row, 6)) * gradient(6)) / (2# * perturbation)
        Next row
    Next column
    TryLocalGradientDerivative = True
End Function

' Сравнивает успешные касательные широкой пробы вдоль Bp·d. Порог
' применяется при сборке поправки Ньютона.
Private Function HasLargeGradientTangentChange() As Boolean
    Dim row As Long, column As Long, threshold As Double
    threshold = GRADIENT_TANGENT_CHANGE_RATIO * mReferenceModulus
    For column = 1 To COMPONENT_COUNT
        For row = 1 To COMPONENT_COUNT
            If Abs(mPlusTangent(row, column) - mMinusTangent(row, column)) > threshold Then
                HasLargeGradientTangentChange = True
                Exit Function
            End If
        Next row
    Next column
End Function

' При неизменной принятой истории берёт две пробные касательные вдоль Bp·d.
' Возмущения сохраняют отклик и готовые столбцы базовой деформации.
Private Sub TakeGradientProbes(ByVal point As Long, ByRef field() As Double, _
        ByRef differenceScale As Double, ByRef plusOk As Boolean, ByRef minusOk As Boolean)
    Dim gradient(1 To COMPONENT_COUNT) As Double, perturbed(1 To COMPONENT_COUNT) As Double
    Dim perturbedStress(1 To COMPONENT_COUNT) As Double, magnitude As Double, i As Long
    mSection.AxialGradientAtPoint point, field, gradient
    For i = 1 To COMPONENT_COUNT
        If Abs(gradient(i)) > magnitude Then magnitude = Abs(gradient(i))
    Next i
    differenceScale = magnitude / GRADIENT_DIFFERENCE_STEP
    If differenceScale = 0# Then Exit Sub
    For i = 1 To COMPONENT_COUNT
        perturbed(i) = mPointStrain(i) + (gradient(i) / magnitude) * GRADIENT_DIFFERENCE_STEP
    Next i
    plusOk = mResponse.EvaluatePoint(point, perturbed, perturbedStress, mPlusTangent, recordTrial:=False) = sfcNone
    For i = 1 To COMPONENT_COUNT
        perturbed(i) = mPointStrain(i) - (gradient(i) / magnitude) * GRADIENT_DIFFERENCE_STEP
    Next i
    minusOk = mResponse.EvaluatePoint(point, perturbed, perturbedStress, mMinusTangent, recordTrial:=False) = sfcNone
End Sub

' На гладком материальном участке смешанные производные напряжений
' перестановочны. Поэтому производную D·(Bp·d) по деформации можно
' вычислить направленной производной D вдоль Bp·d.
Private Sub AddGradientDerivative(ByVal point As Long, ByVal differenceScale As Double, _
        ByVal plusOk As Boolean, ByVal minusOk As Boolean)
    Dim i As Long, j As Long, changed As Boolean
    For j = 1 To COMPONENT_COUNT
        For i = 1 To COMPONENT_COUNT
            If plusOk And minusOk Then
                mGradientTangent(i, j) = 0.5 * differenceScale * (mPlusTangent(i, j) - mMinusTangent(i, j))
            ElseIf plusOk Then
                mGradientTangent(i, j) = differenceScale * (mPlusTangent(i, j) - mTangent(i, j))
            Else
                mGradientTangent(i, j) = differenceScale * (mTangent(i, j) - mMinusTangent(i, j))
            End If
            If mGradientTangent(i, j) <> 0# Then changed = True
        Next i
    Next j
    If changed Then mSection.AddPointGradientTangent point, mGradientTangent, mMatrix
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует невязку и линеаризацию для независимой проверки разностями.
' Проба использует принятую историю, которую этот вызов не публикует.
Friend Function TestEquilibriumLinearization(ByRef field() As Double, ByVal withMatrix As Boolean, _
        ByVal useGradientTangent As Boolean, ByRef residual() As Double, ByRef tangent() As Double) As Boolean
    Dim savedExternal() As Double, savedResultants() As Double
    If mCandidateReady Then
        savedExternal = mExternal: savedResultants = mTrialResultants
    End If
    ReDim mExternal(1 To mDegreeCount)
    If Evaluate(field, withMatrix, useGradientTangent, recordTrial:=False) Then
        residual = mResidual
        If withMatrix Then tangent = mMatrix
        TestEquilibriumLinearization = True
    End If
    If mCandidateReady Then
        mExternal = savedExternal: mTrialResultants = savedResultants
    End If
End Function

' Уточняет сторону предела на пробной истории. Если знак не разделён,
' сохраняет только то силовое поле, которое проверила та же система уравнений.
Friend Function VerifyDirectSteelTrial(ByRef load() As Double, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        ByRef details As String, ByRef unresolved As ECoupledSteelUncertainty) As Boolean
    Const CHECK_TOLERANCE As Double = 1E-12 ' МПа или относительный масштаб повторной проверки равновесия.
    RequireCandidate
    unresolved = csuNone: details = vbNullString
    If mControlActive Then Err.Raise 5, , "Проверка стали требует силовой пробы."
    If Not reinforcement.HasSteel Then VerifyDirectSteelTrial = True: Exit Function

    Dim original() As Double, correction() As Double, point As Long
    original = mCandidate
    If Not ReadDirectScreen(original, correction) Then
        unresolved = reinforcement.PresentSteelCriteria
        details = "Равновесие сохранено. Чувствительность предела стали не подтверждена."
        VerifyDirectSteelTrial = True
        Exit Function
    End If
    If Not reinforcement.SteelBoundaryWithinCorrection(mSection, original, correction, point, unresolved) Then
        VerifyDirectSteelTrial = True
        Exit Function
    End If

    ' Уточнение начинается из принятого поля. Остаток и знак меняются с кривизной ветви.
    Dim seed As Variant, strictAbsolute As Double, strictRelative As Double
    seed = original
    strictAbsolute = GeomMin(absoluteTolerance, CHECK_TOLERANCE)
    strictRelative = GeomMin(relativeTolerance, CHECK_TOLERANCE)
    If Not TryStep(load(1), load(2), load(3), load(4), load(5), load(6), strictAbsolute, strictRelative, maxIterations, seed) Then
        If mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
        details = "Уточнение стороны предела: " & mMeta.ResultComment
        VerifyDirectSteelTrial = RestoreDirectForceField(original, load, absoluteTolerance, relativeTolerance)
        Exit Function
    End If
    If Not WithinDirectCorrection(original, correction, mCandidate) Then
        details = "Уточнение стороны предела вышло из окрестности силовой пробы."
        VerifyDirectSteelTrial = RestoreDirectForceField(original, load, absoluteTolerance, relativeTolerance)
        Exit Function
    End If

    Dim polished() As Double, refinedCorrection() As Double, refinedUncertainty As ECoupledSteelUncertainty
    polished = mCandidate
    If ReadDirectCorrection(polished, refinedCorrection) Then
        If Not reinforcement.SteelBoundaryWithinCorrection(mSection, polished, refinedCorrection, point, refinedUncertainty) Then
            unresolved = csuNone
            details = "Сторона предела стали подтверждена уточнённым силовым равновесием."
            VerifyDirectSteelTrial = True
            Exit Function
        End If
        unresolved = refinedUncertainty
    Else
        unresolved = reinforcement.PresentSteelCriteria
        details = "Силовое равновесие уточнено. Чувствительность предела стали не подтверждена."
        VerifyDirectSteelTrial = True
        Exit Function
    End If

    ' Предельное поле даёт диагностическую невязку, даже при равенстве в пределах арифметики.
    ' Напряжения для публикации всегда остаются на подтверждённом силовом поле.
    Dim boundary() As Double: boundary = polished
    If point > 0 Then
        If reinforcement.ProjectSteelBoundary(mSection, point, boundary) Then
            If WithinDirectCorrection(original, correction, boundary) Then
                If CheckDirectBoundaryField(boundary, load, strictAbsolute, strictRelative) Then _
                    details = "Невязка поля на границе стали, МПа: " & CStr(mLastNorm) & "."
                VerifyDirectSteelTrial = RestoreDirectForceField(polished, load, strictAbsolute, strictRelative)
                Exit Function
            End If
        End If
    End If
    VerifyDirectSteelTrial = True
End Function

' Пересчитывает фиксированное силовое поле из прежнего принятого буфера.
' Проба удовлетворяет всем усилиям, слабым условиям и связям однозначности.
Private Function RestoreDirectForceField(ByRef field() As Double, ByRef load() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double) As Boolean
    mCandidateReady = False
    If Evaluate(field, False) Then
        Dim tolerance As Double: tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
        If EquilibriumNorm(load, field) <= tolerance And GaugesVerified(tolerance) Then
            mCandidate = field
            PublishCandidate 0#
            RestoreDirectForceField = True
            Exit Function
        End If
        RecordFailure sfcNumericalFailure, "Повторная проверка не подтвердила равновесие сохранённого силового поля."
    Else
        RecordFailure mFailureCode, "Не восстановлен материальный отклик сохранённого силового поля."
    End If
End Function

' Оценивает чувствительность по матрице принятого поля. После принятой
' поправки нужна новая касательная, поскольку наклон диаграммы мог измениться.
Private Function ReadDirectScreen(ByRef field() As Double, ByRef correction() As Double) As Boolean
    If mDirectProofMatrix Then
        correction = mResidual
        Dim i As Long
        For i = 1 To mDegreeCount: correction(i) = -correction(i): Next i
        ReadDirectScreen = mSystem.Solve(mMatrix, correction, mDegreeCount)
        If ReadDirectScreen Then Exit Function
    End If
    ReadDirectScreen = ReadDirectCorrection(field, correction)
End Function

' Оценивает поправку на пробном поле без замены принятой истории, пробных
' напряжений и исходного результата силового шага.
Private Function ReadDirectCorrection(ByRef field() As Double, ByRef correction() As Double) As Boolean
    Dim savedResidual() As Double, savedForces() As Double, savedNorm As Double
    Dim savedFailure As ESolverFailureCode, savedPoint As Long, savedLastPoint As Long, savedSide As Long
    Dim savedEvaluations As Long, savedGradient As Long, savedLocalGradient As Long, savedFailures As Long
    savedResidual = mResidual: savedForces = mTrialResultants: savedNorm = mLastNorm
    savedFailure = mFailureCode: savedPoint = mMaterialPoint: savedLastPoint = mLastFailedPoint: savedSide = mLastFailedSide
    savedEvaluations = mEvaluations: savedGradient = mGradientBuilds: savedLocalGradient = mLocalGradientBuilds
    savedFailures = mMaterialFailures
    Dim useGradient As Boolean: useGradient = mSection.HasAxialGradient(field)
    If Evaluate(field, True, useGradient, recordTrial:=False) Then
        ReadDirectCorrection = SolveCorrection(field, correction)
    End If
    mResidual = savedResidual: mTrialResultants = savedForces: mLastNorm = savedNorm
    mFailureCode = savedFailure: mMaterialPoint = savedPoint: mLastFailedPoint = savedLastPoint: mLastFailedSide = savedSide
    mEvaluations = savedEvaluations: mGradientBuilds = savedGradient: mLocalGradientBuilds = savedLocalGradient
    mMaterialFailures = savedFailures
    If Not ReadDirectCorrection Then ReDim correction(1 To mDegreeCount)
End Function

' Ограничивает уточнение окрестностью исходной поправки Ньютона.
' Множитель два резервирует следующее приращение в сжимающейся итерации.
Private Function WithinDirectCorrection(ByRef reference() As Double, ByRef correction() As Double, ByRef trial() As Double) As Boolean
    Const ROUNDING_STEP As Double = 2.22044604925031E-16
    Dim i As Long, change As Double, radius As Double, fieldMagnitude As Double
    For i = 1 To mGaugeStart - 1
        change = GeomMax(change, Abs(trial(i) - reference(i)))
        radius = GeomMax(radius, Abs(correction(i)))
        fieldMagnitude = GeomMax(fieldMagnitude, Abs(reference(i)))
    Next i
    WithinDirectCorrection = (change <= 2# * radius + 8# * ROUNDING_STEP * fieldMagnitude)
End Function

' Проверяет неподвижное поле на пределах арифметической точности
' интегралов и более строгом допуске силовой задачи.
Private Function CheckDirectBoundaryField(ByRef field() As Double, ByRef load() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double) As Boolean
    Const ROUNDING_STEP As Double = 2.22044604925031E-16
    Dim forceMagnitude As Double: forceMagnitude = 1# + StressScale(mExternal)
    Dim tolerance As Double: tolerance = GeomMin(EquilibriumTolerance(absoluteTolerance, relativeTolerance), _
        (mPointCount + mDegreeCount) * ROUNDING_STEP * forceMagnitude)
    If Not Evaluate(field, False) Then Exit Function
    If EquilibriumNorm(load, field) > tolerance Then Exit Function
    CheckDirectBoundaryField = GaugesVerified(tolerance)
End Function

' Возвращает максимум по всем силовым уравнениям и интегралам в масштабе МПа.
' Связи однозначности заданы в масштабе деформации и проверяются отдельно.
Private Function EquilibriumNorm(ByRef load() As Double, ByRef field() As Double) As Double
    Dim i As Long, value As Double
    mLastNorm = 0#
    For i = 1 To mGaugeStart - 1
        value = Abs(mResidual(i))
        If value > mLastNorm Then mLastNorm = value
    Next i
    Dim resultantNorm As Double: resultantNorm = mSection.ResultantResidualNorm(mTrialResultants, load)
    If resultantNorm > mLastNorm Then mLastNorm = resultantNorm
    If mControlActive Then
        value = Abs(ControlResidual(field))
        If value > mLastNorm Then mLastNorm = value
    End If
    EquilibriumNorm = mLastNorm
End Function

' Проверяет кинематические ограничения через эквивалентный допуск по деформации.
Private Function GaugesVerified(ByVal tolerance As Double) As Boolean
    Dim i As Long
    For i = mGaugeStart To mDegreeCount
        If Abs(mResidual(i)) > tolerance / mReferenceModulus Then Exit Function
    Next i
    GaugesVerified = True
End Function

' Нормирует относительный допуск на максимальную приведённую нагрузку в МПа.
Private Function StressScale(ByRef values() As Double) As Double
    Dim i As Long, value As Double
    For i = 1 To mGaugeStart - 1
        value = Abs(values(i))
        If value > StressScale Then StressScale = value
    Next i
End Function

' Вычисляет допуск в МПа по внешним уравнениям текущей пробы. При управлении
' деформацией они меняются вместе с множителем нагрузки.
Private Function EquilibriumTolerance(ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double) As Double
    EquilibriumTolerance = absoluteTolerance + relativeTolerance * StressScale(mExternal)
End Function

' Проверяет, что кандидат относится к последней успешной пробе.
Private Sub RequireCandidate()
    If Not mCandidateReady Then Err.Raise 5, "CCoupledNewtonSolver", "Нет подтверждённого равновесия пробного шага."
End Sub

' Формирует общий статус численного отказа, сохраняя диагностический контекст.
Private Sub RecordFailure(ByVal code As ESolverFailureCode, ByVal reason As String)
    mMeta.SetSolverFailure code, rkDirectState, reason, (mEvaluations > 0)
    mMeta.DiagnosticDetails = DiagnosticSummary()
    If Len(mSameControlInitialFailure) > 0 Then _
        mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & " | До восстанавливающей пробы: " & mSameControlInitialFailure
End Sub

' Возвращает параметры последней итерации, необходимые для проверки отказа.
Private Function DiagnosticSummary() As String
    DiagnosticSummary = "Итерации: " & CStr(mIterations) & "; пробы: " & CStr(mEvaluations) & _
        "; невязка, МПа: " & CStr(mLastNorm) & "; мин. ведущий элемент: " & CStr(mMinimumPivot)
    If mMaterialPoint > 0 Then DiagnosticSummary = DiagnosticSummary & "; материальная точка: " & CStr(mMaterialPoint)
    If mMaterialFailures > 0 Then DiagnosticSummary = DiagnosticSummary & "; отказов материала: " & CStr(mMaterialFailures) & _
        "; последняя точка отказа: " & CStr(mLastFailedPoint)
    If mLastFailedSide <> 0 Then DiagnosticSummary = DiagnosticSummary & _
        "; продольное положение последней точки отказа: " & Format$(mLastFailedSide * SPATIAL_HALF_STEP, "+0.000;-0.000") & " z/L"
    If mGradientBuilds > 0 Then DiagnosticSummary = DiagnosticSummary & "; уточнений касательной: " & CStr(mGradientBuilds)
    If mSameControlRecoveryTried Then DiagnosticSummary = DiagnosticSummary & "; проверен переход для исходной меры"
    If mLocalGradientBuilds > 0 Then DiagnosticSummary = DiagnosticSummary & _
        "; локальных производных градиента напряжений: " & CStr(mLocalGradientBuilds)
    If mSteelBranchBuilds > 0 Then DiagnosticSummary = DiagnosticSummary & _
        "; пересборок ветвей стали: " & CStr(mSteelBranchBuilds)
    If mSteelFieldTangents > 0 Then DiagnosticSummary = DiagnosticSummary & _
        "; разностных бетонных касательных по полю: " & CStr(mSteelFieldTangents)
    If mHybridEvaluations > 0 Then DiagnosticSummary = DiagnosticSummary & "; гибридных поправок: " & CStr(mHybridSteps) & _
        "; глобальных касательных: " & CStr(mHybridMatrices) & "; разностных и прочих гибридных проб: " & CStr(mHybridEvaluations)
End Function

' Выбирает границу по действительному отклонённому отрезку и решает общую точку и входящую сторону от одной истории.
' Радиус определяется предыдущим механическим шагом; выбранный номер точки не зависит от формы сечения.
Friend Function TryPreparePrincipalTransition(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef previous() As Double, ByVal currentLambda As Double, ByVal absoluteTolerance As Double, _
        ByVal relativeTolerance As Double, ByVal maxIterations As Long, ByRef transition As CCoupledPrincipalTransition, _
        ByRef trialLambda As Double, ByRef projection() As Double) As Boolean
    Set transition = Nothing
    If Not mPrincipalTransition Is Nothing Then Exit Function
    If Not HasPrincipalBoundaries Or Not mHasTransitionSegment Or mTransitionRevision <> mAcceptedRevision Then Exit Function
    If absoluteTolerance <= 0# Or relativeTolerance <= 0# Or maxIterations < 3 Or currentLambda <= 0# Then Exit Function
    If Not SamePrincipalLoadPath(baseLoad, direction, mBaseLoad, mLoadDirection) Then Exit Function
    Dim station As TYieldStation
    If Not SelectPrincipalTransitionStation(previous, station, _
            PrincipalCandidateTolerance(currentLambda, absoluteTolerance, relativeTolerance)) Then Exit Function
    If station.PrincipalIndex = 0 Then Exit Function
    Dim strain(1 To COMPONENT_COUNT) As Double, unused(1 To COMPONENT_COUNT) As Double, value As Double
    YieldStrain station.Point, station.Side, mAccepted, strain
    If Not ReadTransitionValue(station, strain, value, unused, False) Then Exit Function
    If value = 0# Or Sgn(value) <> Sgn(station.FromValue) Then Exit Function
    Dim reference As Double: reference = SpatialFieldDistance(previous, mAccepted)
    If reference <= 0# Then Exit Function
    Dim originalMeta As CResultMeta, originalProjection() As Double, originalTarget As Double, originalLambda As Double
    Dim originalDirection() As Double, hadDirection As Boolean, segmentFrom() As Double, segmentTo() As Double
    Set originalMeta = mMeta.Clone
    originalProjection = mControlProjection: originalTarget = mTargetControl: originalLambda = mCurrentControlLambda
    hadDirection = mHasRejectedDirection
    If hadDirection Then originalDirection = mRejectedDirection
    segmentFrom = mTransitionFrom: segmentTo = mTransitionTo
    ConfigureControlledLoad baseLoad, direction
    ResetSolveState
    Set mPrincipalTransition = Nothing
    mControlActive = False: mArcLoadProjection = 0#: mConsistentControl = False
    Dim load(1 To COMPONENT_COUNT) As Double, gap As Double, tolerance As Double, radius As Double
    SetControlledLoad currentLambda, load
    gap = GeomMin(0.25 * Abs(value), 0.125 * EquilibriumTolerance(absoluteTolerance, relativeTolerance))
    If gap <= 16# * YIELD_ROOT_TOLERANCE Then GoTo Unsuccessful
    tolerance = GeomMin(PRINCIPAL_EVENT_TOLERANCE, 0.125 * gap)
    radius = PRINCIPAL_EVENT_RADIUS * reference
    Dim face() As Double, incoming() As Double, faceLambda As Double
    On Error GoTo Failed
    If Not FollowPrincipalCoordinate(station, mAccepted, currentLambda, 0#, tolerance, maxIterations, _
            mAccepted, radius, face, faceLambda, projection) Then GoTo Unsuccessful
    If Not PrincipalSteelWithinLimits(face) Then GoTo Unsuccessful
    Dim advance As Double, oldIncrement As Double
    If Not SpatialIncrementAdvances(previous, mAccepted, face, advance, oldIncrement) Then GoTo Unsuccessful
    If Not FollowPrincipalCoordinate(station, face, faceLambda, Sgn(value) * gap, tolerance, maxIterations, _
            mAccepted, radius, incoming, trialLambda, projection) Then GoTo Unsuccessful
    If Not SpatialIncrementAdvances(previous, mAccepted, incoming, advance, oldIncrement) Then GoTo Unsuccessful
    Dim designVerified As Boolean
    If Not VerifyPrincipalDesignLoad(station, mAccepted, currentLambda, incoming, trialLambda, radius, _
            absoluteTolerance, relativeTolerance, maxIterations, designVerified) Then GoTo Unsuccessful
    Dim prepared As New CCoupledPrincipalTransition
    prepared.Initialize mAcceptedRevision, station.Point, station.Side, station.PrincipalIndex, Sgn(value), gap, _
        reference, faceLambda, trialLambda, mAccepted, face, incoming
    If Not PublishPrincipalCandidate(incoming, trialLambda, projection, tolerance) Then GoTo Unsuccessful
    Set mPrincipalTransition = prepared: Set transition = prepared
    mPrincipalBaseLoad = baseLoad: mPrincipalLoadDirection = direction
    mPrincipalIncomingCandidate = True: mPrincipalDesignVerified = designVerified
    mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & "; подтверждена входящая сторона спектрального угла; точка: " & _
        CStr(station.Point) & "; станция: " & CStr(station.Side) & "; главное напряжение: " & CStr(station.PrincipalIndex)
    TryPreparePrincipalTransition = True
    Exit Function
Unsuccessful:
    mControlActive = True: mCandidateReady = False
    Dim failureDetails As String: failureDetails = DiagnosticSummary()
    mControlProjection = originalProjection: mTargetControl = originalTarget: mCurrentControlLambda = originalLambda
    mHasRejectedDirection = hadDirection
    If hadDirection Then mRejectedDirection = originalDirection Else Erase mRejectedDirection
    mTransitionFrom = segmentFrom: mTransitionTo = segmentTo: mTransitionRevision = mAcceptedRevision
    mHasTransitionSegment = True: Set mPrincipalTransition = Nothing: mPrincipalIncomingCandidate = False
    Set mMeta = originalMeta
    mMeta.DiagnosticDetails = "Не подтверждено входящее равновесие спектрального угла; " & failureDetails & _
        " | Исходная проба: " & originalMeta.DiagnosticDetails
    Exit Function
Failed:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Unsuccessful
End Function

' Решает исходящую сторону только после принятия сохранённого входящего поля.
' Общая точка повторно проверяется уже с новой историей; чужой или устаревший контекст не используется.
Friend Function TryCompletePrincipalTransition(ByVal transition As CCoupledPrincipalTransition, _
        ByRef baseLoad() As Double, ByRef direction() As Double, ByVal currentLambda As Double, _
        ByVal maxIterations As Long, ByRef trialLambda As Double, ByRef projection() As Double, _
        Optional ByVal absoluteTolerance As Double = 0.0000005, Optional ByVal relativeTolerance As Double = 0.0000005) As Boolean
    If transition Is Nothing Or mPrincipalTransition Is Nothing Then Exit Function
    If Not (transition Is mPrincipalTransition) Then Exit Function
    If Not transition.MatchesIncoming(mAcceptedRevision, mAccepted) Or maxIterations < 2 Then Exit Function
    If Not SamePrincipalLoadPath(baseLoad, direction, mPrincipalBaseLoad, mPrincipalLoadDirection) Then Exit Function
    Dim incoming() As Double, incomingLambda As Double
    transition.CopyIncoming incoming, incomingLambda
    If currentLambda <> incomingLambda Then Exit Function
    Dim station As TYieldStation
    station.Point = transition.Point: station.Side = transition.Side: station.PrincipalIndex = transition.PrincipalIndex
    ConfigureControlledLoad baseLoad, direction
    ResetSolveState
    mControlActive = False: mArcLoadProjection = 0#: mConsistentControl = False
    Dim tolerance As Double, radius As Double, face() As Double, faceLambda As Double, outgoing() As Double
    tolerance = GeomMin(PRINCIPAL_EVENT_TOLERANCE, 0.125 * transition.PrincipalGap)
    radius = PRINCIPAL_EVENT_RADIUS * transition.ReferenceIncrement
    On Error GoTo Failed
    If Not FollowPrincipalCoordinate(station, mAccepted, currentLambda, 0#, tolerance, maxIterations, _
            mAccepted, radius, face, faceLambda, projection) Then GoTo Unsuccessful
    If Not PrincipalSteelWithinLimits(face) Then GoTo Unsuccessful
    Dim originalFace() As Double, originalLambda As Double
    transition.CopyFace originalFace, originalLambda
    If SpatialFieldDistance(originalFace, face) > radius Then GoTo Unsuccessful
    If Not FollowPrincipalCoordinate(station, face, faceLambda, -transition.IncomingSign * transition.PrincipalGap, _
            tolerance, maxIterations, mAccepted, radius, outgoing, trialLambda, projection) Then GoTo Unsuccessful
    Dim designVerified As Boolean
    If Not VerifyPrincipalDesignLoad(station, mAccepted, currentLambda, outgoing, trialLambda, radius, _
            absoluteTolerance, relativeTolerance, maxIterations, designVerified) Then GoTo Unsuccessful
    If Not PublishPrincipalCandidate(outgoing, trialLambda, projection, tolerance) Then GoTo Unsuccessful
    mPrincipalDesignVerified = designVerified: mPrincipalOutgoingCandidate = True
    mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & "; подтверждена исходящая сторона спектрального угла; точка: " & _
        CStr(station.Point) & "; станция: " & CStr(station.Side) & "; главное напряжение: " & CStr(station.PrincipalIndex)
    TryCompletePrincipalTransition = True
    Exit Function
Unsuccessful:
    mControlActive = True: mCandidateReady = False
    RecordFailure sfcNumericalFailure, "Не подтверждено локальное равновесие исходящей стороны материального перехода."
    Exit Function
Failed:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
    Resume Unsuccessful
End Function

' Сообщает о проверенном точном lambda=1 только для ещё не принятого кандидата материального перехода.
Friend Property Get PrincipalDesignLoadVerified() As Boolean
    PrincipalDesignLoadVerified = mCandidateReady And mPrincipalDesignVerified
End Property

' Передаёт точное проверенное заданное сочетание до принятия следующего состояния, не восстанавливая материал скрыто.
Friend Function CopyPrincipalDesignState(ByRef field() As Double, ByRef resultants() As Double) As Boolean
    If Not PrincipalDesignLoadVerified Then Exit Function
    field = mPrincipalDesignField: resultants = mPrincipalDesignResultants
    CopyPrincipalDesignState = True
End Function

' При пересечении единицы отдельно решает заданные усилия от той же принятой истории в остатке бюджета.
' Читающая силовая проба не принимает материал и должна остаться между спектральными координатами концов.
Private Function VerifyPrincipalDesignLoad(ByRef station As TYieldStation, ByRef first() As Double, _
        ByVal firstLambda As Double, ByRef last() As Double, ByVal lastLambda As Double, ByVal radius As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        ByRef verified As Boolean) As Boolean
    verified = False
    If (firstLambda - 1#) * (lastLambda - 1#) > 0# Then VerifyPrincipalDesignLoad = True: Exit Function
    Dim remaining As Long: remaining = maxIterations - mIterations
    If remaining < 1 Or absoluteTolerance <= 0# Or relativeTolerance <= 0# Then Exit Function
    Dim effort As TSolveEffort: ReadSolveEffort effort
    Dim strain(1 To COMPONENT_COUNT) As Double, unused(1 To COMPONENT_COUNT) As Double
    Dim firstValue As Double, lastValue As Double, value As Double, field() As Double, actual() As Double
    YieldStrain station.Point, station.Side, first, strain
    If Not ReadTransitionValue(station, strain, firstValue, unused, False) Then Exit Function
    YieldStrain station.Point, station.Side, last, strain
    If Not ReadTransitionValue(station, strain, lastValue, unused, False) Then Exit Function
    Dim load(1 To COMPONENT_COUNT) As Double, i As Long, predicted() As Double, initial As Variant, ratio As Double
    predicted = first
    If lastLambda <> firstLambda Then ratio = (1# - firstLambda) / (lastLambda - firstLambda)
    For i = 1 To mGaugeStart - 1: predicted(i) = first(i) + ratio * (last(i) - first(i)): Next i
    For i = 1 To COMPONENT_COUNT: load(i) = mBaseLoad(i) + mLoadDirection(i): Next i
    initial = predicted
    If TryStep(load(1), load(2), load(3), load(4), load(5), load(6), absoluteTolerance, relativeTolerance, remaining, initialField:=initial) Then
        CopyTrial field, actual
        YieldStrain station.Point, station.Side, field, strain
        If ReadTransitionValue(station, strain, value, unused, False) Then
            verified = (value >= GeomMin(firstValue, lastValue) - YIELD_ROOT_TOLERANCE And _
                value <= GeomMax(firstValue, lastValue) + YIELD_ROOT_TOLERANCE And SpatialFieldDistance(first, field) <= radius)
            Dim response As CCoupledReinforcedResponse: Set response = mResponse
            If response.TrialSteelUtilization > 1# Then verified = False
            If verified Then mPrincipalDesignField = field: mPrincipalDesignResultants = actual
        End If
    End If
    DiscardStep
    AddSolveEffort effort
    mControlActive = False
    VerifyPrincipalDesignLoad = verified
End Function

' Решает G=0 и одну спектральную координату расширенным Ньютоном с общим бюджетом итераций.
' Читающие пробы и поиск доли не принимают историю; радиус исключает скачок к удалённому равновесию.
Private Function FollowPrincipalCoordinate(ByRef station As TYieldStation, ByRef initial() As Double, _
        ByVal initialLambda As Double, ByVal target As Double, ByVal tolerance As Double, ByVal maxIterations As Long, _
        ByRef anchor() As Double, ByVal radius As Double, ByRef result() As Double, _
        ByRef resultLambda As Double, ByRef projection() As Double) As Boolean
    Dim field() As Double, trial() As Double, correction() As Double, nullDirection() As Double, eventRow() As Double
    Dim load(1 To COMPONENT_COUNT) As Double, strain(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double
    Dim value As Double, baseNorm As Double, trialNorm As Double, fraction As Double, trialLambda As Double
    Dim damping As Long, i As Long, accepted As Boolean
    field = initial: resultLambda = initialLambda
    Do While mIterations < maxIterations
        mIterations = mIterations + 1
        SetControlledLoad resultLambda, load
        If Not Evaluate(field, True, recordTrial:=False) Then Exit Function
        If Not YieldDirections(station, field, target, value, eventRow, correction, nullDirection) Then Exit Function
        mMinimumPivot = mSystem.MinimumPivot
        baseNorm = GeomMax(EquilibriumNorm(load, field), Abs(value - target))
        If baseNorm <= tolerance And GaugesVerified(tolerance) Then
            If SpatialFieldDistance(anchor, field) > radius Or resultLambda <= 0# Then Exit Function
            ReDim projection(1 To mDegreeCount)
            For i = 1 To mGaugeStart - 1: projection(i) = eventRow(i) / mReferenceModulus: Next i
            result = field: FollowPrincipalCoordinate = True
            Exit Function
        End If
        fraction = 1#: accepted = False
        For damping = 0 To 16
            trial = field
            For i = 1 To mDegreeCount: trial(i) = field(i) + fraction * correction(i): Next i
            trialLambda = resultLambda + fraction * correction(mDegreeCount + 1)
            If trialLambda > 0# And SpatialFieldDistance(anchor, trial) <= radius Then
                SetControlledLoad trialLambda, load
                If Evaluate(trial, False, recordTrial:=False) Then
                    YieldStrain station.Point, station.Side, trial, strain
                    If ReadTransitionValue(station, strain, value, gradient, False) Then
                        trialNorm = GeomMax(EquilibriumNorm(load, trial), Abs(value - target))
                        If GaugesVerified(tolerance) Then
                            accepted = (trialNorm <= tolerance Or trialNorm < baseNorm * (1# - MIN_PROGRESS * fraction))
                            If accepted Then Exit For
                        End If
                    End If
                End If
            End If
            fraction = fraction * 0.5
        Next damping
        If Not accepted Then Exit Function
        field = trial: resultLambda = trialLambda
    Loop
End Function

' Проверяет арматуру в общей непринятой точке до перехода на другую сторону; пробные буферы следующего поля заполняются заново.
Private Function PrincipalSteelWithinLimits(ByRef field() As Double) As Boolean
    If Not Evaluate(field, False) Then Exit Function
    Dim response As CCoupledReinforcedResponse: Set response = mResponse
    PrincipalSteelWithinLimits = (response.TrialSteelUtilization <= 1#)
End Function

' Заполняет все пробные материальные буферы подтверждённого поля и возвращает обычный контракт кандидата.
Private Function PublishPrincipalCandidate(ByRef field() As Double, ByVal lambda As Double, _
        ByRef projection() As Double, ByVal tolerance As Double) As Boolean
    mControlProjection = projection: mTargetControl = 0#: mControlActive = True
    Dim i As Long, load(1 To COMPONENT_COUNT) As Double
    For i = 1 To mGaugeStart - 1: mTargetControl = mTargetControl + projection(i) * field(i): Next i
    SetControlledLoad lambda, load
    If Not Evaluate(field, False) Then Exit Function
    If EquilibriumNorm(load, field) > tolerance Or Not GaugesVerified(tolerance) Then Exit Function
    mCandidate = field: PublishCandidate lambda
    PublishPrincipalCandidate = True
End Function

' Вычисляет механическое расстояние в общей тензорной метрике трёх станций без множителей связей.
Private Function SpatialFieldDistance(ByRef first() As Double, ByRef second() As Double) As Double
    Dim movement() As Double, coefficients() As Double, unused() As Double, i As Long, square As Double
    ReDim movement(1 To mDegreeCount)
    For i = 1 To mGaugeStart - 1: movement(i) = second(i) - first(i): Next i
    mSection.CopySpatialStrainCoefficients SPATIAL_HALF_STEP, coefficients
    square = ArcPhysicalSquare(coefficients, movement, unused)
    If square > 0# Then SpatialFieldDistance = Sqr(square)
End Function

' Не переносит разрешение материального перехода на другую постоянную или масштабируемую часть нагрузки.
Private Function SamePrincipalLoadPath(ByRef firstBase() As Double, ByRef firstDirection() As Double, _
        ByRef secondBase() As Double, ByRef secondDirection() As Double) As Boolean
    If LBound(firstBase) <> 1 Or UBound(firstBase) <> COMPONENT_COUNT Or _
            LBound(firstDirection) <> 1 Or UBound(firstDirection) <> COMPONENT_COUNT Then Exit Function
    Dim i As Long
    For i = 1 To COMPONENT_COUNT
        If firstBase(i) <> secondBase(i) Or firstDirection(i) <> secondDirection(i) Then Exit Function
    Next i
    SamePrincipalLoadPath = True
End Function

' Предпочитает наблюдавшееся пересечение; при возврате Ньютона к прежнему знаку ищет ближайшую границу впереди принятой секущей.
' Линейная оценка ограничена радиусом локальной доводки и сама не подтверждает ни переход, ни разрушение.
Private Function SelectPrincipalTransitionStation(ByRef previous() As Double, ByRef station As TYieldStation, _
        ByVal tolerance As Double) As Boolean
    If FirstYieldCrossing(mAccepted, mTransitionFrom, station) Then SelectPrincipalTransitionStation = True: Exit Function
    If FirstYieldCrossing(mAccepted, mTransitionTo, station) Then SelectPrincipalTransitionStation = True: Exit Function
    Dim movement() As Double, point As Long, side As Long, principal As Long, i As Long
    ReDim movement(1 To mDegreeCount)
    For i = 1 To mGaugeStart - 1: movement(i) = mAccepted(i) - previous(i): Next i
    Dim strain(1 To COMPONENT_COUNT) As Double, rateStrain(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double
    Dim value As Double, rate As Double, fraction As Double, nearest As Double, candidate As TYieldStation
    nearest = PRINCIPAL_EVENT_RADIUS
    For point = 1 To mPointCount
        For side = -1 To 1
            YieldStrain point, side, mAccepted, strain
            YieldStrain point, side, movement, rateStrain
            For principal = 1 To 3
                candidate.Point = point: candidate.Side = side: candidate.PrincipalIndex = principal
                If ReadTransitionValue(candidate, strain, value, gradient) Then
                    rate = 0#
                    For i = 1 To COMPONENT_COUNT: rate = rate + gradient(i) * rateStrain(i): Next i
                    If value * rate < 0# Then
                        fraction = -value / rate
                        If fraction <= nearest And Abs(value) > 16# * YIELD_ROOT_TOLERANCE Then
                            candidate.FromValue = value: candidate.ToValue = -value
                            station = candidate: nearest = fraction: SelectPrincipalTransitionStation = True
                        End If
                    End If
                End If
            Next principal
        Next side
    Next point
    If Not SelectPrincipalTransitionStation Then SelectPrincipalTransitionStation = NearestPrincipalStation(previous, station, tolerance)
End Function

' Определяет масштаб выбора границы от принятого множителя и исходной траектории без изменения внешних уравнений пробы.
Private Function PrincipalCandidateTolerance(ByVal lambda As Double, ByVal absoluteTolerance As Double, _
        ByVal relativeTolerance As Double) As Double
    Dim loadScale As Double, value As Double, i As Long
    For i = 1 To mGaugeStart - 1
        value = Abs(mBaseEquations(i) + lambda * mDirectionEquations(i))
        If value > loadScale Then loadScale = value
    Next i
    PrincipalCandidateTolerance = absoluteTolerance + relativeTolerance * loadScale
End Function

' При сильном повороте ищет простую границу только в пределах численно различимого напряжения текущего равновесия.
' Допуск ограничивает выбор кандидата; общая точка затем решается со своим более строгим допуском и проверкой направления.
Private Function NearestPrincipalStation(ByRef previous() As Double, ByRef station As TYieldStation, _
        ByVal tolerance As Double) As Boolean
    Dim nearest As Double: nearest = GeomMin(tolerance, _
        mReferenceModulus * PRINCIPAL_EVENT_RADIUS * SpatialFieldDistance(previous, mAccepted))
    If nearest <= 0# Then Exit Function
    Dim point As Long, side As Long, principal As Long, candidate As TYieldStation
    Dim strain(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double, value As Double
    For point = 1 To mPointCount
        For side = -1 To 1
            YieldStrain point, side, mAccepted, strain
            For principal = 1 To 3
                candidate.Point = point: candidate.Side = side: candidate.PrincipalIndex = principal
                If ReadTransitionValue(candidate, strain, value, gradient) Then
                    If Abs(value) > 16# * YIELD_ROOT_TOLERANCE And Abs(value) < nearest Then
                        candidate.FromValue = value: candidate.ToValue = -value
                        station = candidate: nearest = Abs(value): NearestPrincipalStation = True
                    End If
                End If
            Next principal
        Next side
    Next point
End Function

' Находит касательную исходящей ветви после принятия её истории; знак задаёт дальнейшее удаление от выбранной границы.
' Разрешение одноразового перехода нельзя переносить на другую ревизию или на другой экземпляр контекста.
Friend Function TryPrincipalDepartureControl(ByVal transition As CCoupledPrincipalTransition, ByVal lambda As Double, _
        ByRef movement() As Double, ByRef lambdaDirection As Double, ByRef projection() As Double, _
        ByRef coordinate As Double, Optional ByRef linearizationField As Variant, _
        Optional ByRef linearizationLambda As Double = 0#, Optional ByVal maxIterations As Long = 32) As Boolean
    If transition Is Nothing Or mPrincipalDeparture Is Nothing Then Exit Function
    If Not (transition Is mPrincipalDeparture) Or mDepartureRevision <> mAcceptedRevision Or lambda <> mTrialLambda Then Exit Function
    Dim station As TYieldStation, strain(1 To COMPONENT_COUNT) As Double, gradient(1 To COMPONENT_COUNT) As Double
    station.Point = transition.Point: station.Side = transition.Side: station.PrincipalIndex = transition.PrincipalIndex
    YieldStrain station.Point, station.Side, mAccepted, strain
    Dim value As Double
    If Not ReadTransitionValue(station, strain, value, gradient, False) Then Exit Function
    Dim load(1 To COMPONENT_COUNT) As Double, wasControlled As Boolean
    wasControlled = mControlActive: mControlActive = False
    On Error GoTo Failed
    Dim tolerance As Double: tolerance = GeomMin(PRINCIPAL_EVENT_TOLERANCE, 0.125 * transition.PrincipalGap)
    Dim probe() As Double, probeLambda As Double, target As Double
    Dim departureGap As Double
    departureGap = GeomMax(transition.PrincipalGap, mReferenceModulus * (2# ^ (-26))) ' sqrt машинного epsilon отделяет производную от округления почти совпавших историй.
    target = value - transition.IncomingSign * departureGap
    If Not FollowPrincipalCoordinate(station, mAccepted, lambda, target, tolerance, maxIterations, mAccepted, _
            PRINCIPAL_EVENT_RADIUS * transition.ReferenceIncrement, probe, probeLambda, projection) Then GoTo Finished
    SetControlledLoad probeLambda, load
    If Not Evaluate(probe, True, recordTrial:=False) Then GoTo Finished
    If EquilibriumNorm(load, probe) > tolerance Or Not GaugesVerified(tolerance) Then GoTo Finished
    Dim correction() As Double, nullDirection() As Double, row() As Double, prior() As Double, length As Double, i As Long
    If Not YieldDirections(station, probe, target, value, row, correction, nullDirection) Then GoTo Finished
    ReDim movement(1 To mDegreeCount)
    prior = mAccepted
    For i = 1 To mDegreeCount
        movement(i) = -transition.IncomingSign * nullDirection(i)
        prior(i) = mAccepted(i) - movement(i)
    Next i
    If Not SelectSpatialDirectionControl(prior, mAccepted, movement, projection, coordinate, length) Then GoTo Finished
    For i = 1 To mDegreeCount: movement(i) = movement(i) / length: Next i
    lambdaDirection = -transition.IncomingSign * nullDirection(mDegreeCount + 1) / length
    If Not IsMissing(linearizationField) Then linearizationField = probe
    linearizationLambda = probeLambda
    TryPrincipalDepartureControl = True
Finished:
    mControlActive = wasControlled
    Exit Function
Failed:
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    mControlActive = wasControlled
    If number <> 6 And number <> 11 Then Err.Raise number, source, reason
End Function
