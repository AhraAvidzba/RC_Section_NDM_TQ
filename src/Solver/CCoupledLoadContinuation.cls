VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledLoadContinuation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Прослеживание равновесной ветви связанного сечения по усилиям
' ==========================================================================
' Сначала нагружает постоянной частью, затем изменяет lambda. Численный
' отказ уменьшает шаг, не меняя последнее принятое состояние материалов.

Private Const COMPONENT_COUNT As Long = 6
Private Const MIN_STEP_RATIO As Double = 0.015625 ' Не менее 1/64 начального приращения.
Private Const STEP_GROWTH As Double = 1.5 ' Возвращает начальный шаг за несколько удачных подшагов.
Private Const DOUBLE_RELATIVE_ROUNDING As Double = 2.22044604925031E-16 ' Относительный шаг представления Double около единицы.
Private Const TWIST_DEGREE As Long = 4 ' Закручивание, приведённое к габариту сечения.

Private mSolver As CCoupledNewtonSolver
Private mOffset() As Double, mDirection() As Double
Private mField() As Double, mResultants() As Double, mPreviousField() As Double
Private mHasOffset As Boolean, mHasDirection As Boolean, mHasAccepted As Boolean
Private mOffsetFraction As Double, mLambda As Double
Private mPreviousLambda As Double
Private mPreferDeformationOnStall As Boolean
Private mPreferDeformationAfterInitial As Boolean
Private mExactTargetRecovery As String ' Способ подтверждения точной конечной нагрузки после численного отказа.
Private mSecantFailure As String ' Причина отказа прогноза для диагностики всей запрошенной ветви.
Private mTwistGuideDetails As String ' Исход вспомогательной пробы, которая не входит в принятую историю.
Private mInitialSteps As Long, mMaxAttempts As Long, mMaxIterations As Long
Private mAbsoluteTolerance As Double, mRelativeTolerance As Double
Private mRejected As Long, mAccepted As Long, mAttempts As Long
Private mMeta As CResultMeta
Private mAcceptedFields As Collection ' Копии полей принятых шагов для восстановления истории в точках вывода.
Private mAcceptedTrace As CCoupledLoadTrace
Private mDirectSteel As CCoupledReinforcedResponse ' Чтение границы стали на точной силовой цели.
Private mDirectTarget As Double
Private mDirectSteelDetails As String
Private mDirectSteelUncertainty As ECoupledSteelUncertainty
Private mReinforcement As CCoupledReinforcedResponse ' Контроль предельных деформаций включается при поиске.
Private mLimitTolerance As Double
Private mHasSteelLimit As Boolean, mSteelLimitOnOffset As Boolean
Private mSteelLimitLow As Double, mSteelLimitHigh As Double
Private mSteelLimitPoint As Long ' Материальная точка, определившая численно уточнённую верхнюю границу.
Private mLimitFields As Collection
Private mLimitResultants() As Double
Private mLimitTrace As CCoupledLoadTrace

' Начинает новое независимое сочетание. Копирует траекторию и обнуляет
' материалы, чтобы результат не зависел от предыдущей строки таблицы.
Friend Sub Initialize(ByVal solver As CCoupledNewtonSolver, ByVal path As CLoadPathVector, _
        ByVal initialSteps As Long, ByVal maxAttempts As Long, ByVal maxIterations As Long, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, Optional ByVal saveAcceptedFields As Boolean = False)
    If solver Is Nothing Or path Is Nothing Then _
        Err.Raise 5, "CCoupledLoadContinuation", "Для прослеживания ветви передайте решатель и траекторию."
    If initialSteps < 1 Or maxAttempts < 1 Or maxIterations < 1 Or absoluteTolerance <= 0# Or relativeTolerance <= 0# Then _
        Err.Raise 5, "CCoupledLoadContinuation", "Количество шагов, итераций и допуски должны быть положительными."
    Set mSolver = solver
    path.Components mOffset, mDirection
    Dim i As Long
    mHasOffset = False: mHasDirection = False
    For i = 1 To COMPONENT_COUNT
        If mOffset(i) <> 0# Then mHasOffset = True
        If mDirection(i) <> 0# Then mHasDirection = True
    Next i
    mInitialSteps = initialSteps: mMaxAttempts = maxAttempts: mMaxIterations = maxIterations
    mAbsoluteTolerance = absoluteTolerance: mRelativeTolerance = relativeTolerance
    mOffsetFraction = 0#: mLambda = 0#: mPreviousLambda = -1#: mHasAccepted = False
    mPreferDeformationOnStall = False: mPreferDeformationAfterInitial = False
    Set mReinforcement = Nothing: Set mDirectSteel = Nothing
    mHasSteelLimit = False: mSteelLimitOnOffset = False
    mSteelLimitPoint = 0
    Set mLimitFields = Nothing: Erase mLimitResultants
    Erase mField: Erase mResultants: Erase mPreviousField
    Set mAcceptedFields = Nothing: Set mAcceptedTrace = Nothing: Set mLimitTrace = Nothing
    If saveAcceptedFields Then
        Set mAcceptedFields = New Collection
        Set mAcceptedTrace = New CCoupledLoadTrace
    End If
    Set mMeta = New CResultMeta
    mSolver.ResetHistory
End Sub

' Ищет первое пересечение принятого ограничения арматуры на свежей
' траектории. Всегда проверяет расчётное lambda = 1, если оно в диапазоне.
' Достигнутая численная или заданная граница отдельно хранится в ResultMeta.
Friend Sub TraceSteelLimit(ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal initialLambda As Double, ByVal maximumLambda As Double, ByVal lambdaTolerance As Double, _
        Optional ByVal preferDeformationOnStall As Boolean = False, Optional ByVal preferDeformationAfterInitial As Boolean = False, _
        Optional ByVal normalizeInitialStep As Boolean = False)
    If reinforcement Is Nothing Or mHasAccepted Or initialLambda <= 0# Or maximumLambda <= 0# Or lambdaTolerance <= 0# Then _
        Err.Raise 5, "CCoupledLoadContinuation", "Для поиска предела передайте свежую ветвь и положительные настройки Capacity."
    Set mReinforcement = reinforcement
    mLimitTolerance = lambdaTolerance
    mPreferDeformationOnStall = preferDeformationOnStall: mPreferDeformationAfterInitial = preferDeformationAfterInitial
    If Not mHasDirection Then
        If Advance(0#) Then mMeta.SetResult rsNotApplicable, rcNotApplicable, rkCapacity, _
            "Выбранный путь не изменяет усилия. Задайте путь с ненулевой переменной частью."
        Exit Sub
    End If

    Dim target As Double, increment As Double, boundary As Double
    increment = LowerOf(initialLambda, 1#)
    If preferDeformationAfterInitial And normalizeInitialStep Then increment = mSolver.InitialVariableIncrement(mDirection, increment)
    Do
        boundary = maximumLambda
        If mLambda < 1# Then boundary = LowerOf(boundary, 1#)
        target = boundary
        If increment < boundary - mLambda Then target = mLambda + increment
        If target <= mLambda Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkCapacity, _
                "Приращение меньше точности представления множителя.", True
            Exit Sub
        End If
        If Not Advance(target) Then Exit Sub
        If preferDeformationAfterInitial And target < maximumLambda Then
            If SavedFieldCount >= 2 And PreviousAcceptedLambda < mLambda Then
                mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, _
                    "Начальная силовая точка подтверждена; поиск продолжается при управлении деформацией.", ProgressDetails()
                Exit Sub
            End If
        End If
        If target = maximumLambda Then Exit Do
        If increment < maximumLambda / STEP_GROWTH Then increment = increment * STEP_GROWTH
    Loop
    mMeta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, _
        "Достигнута граница Capacity.MaxLambda. Предельная деформация арматуры на пройденном участке не достигнута."
    mMeta.DiagnosticDetails = ProgressDetails()
End Sub

' Начинает деформационный поиск от конца подтверждённой силовой ветви
' после численного отказа. Техническая конечная граница уже является исходом.
Friend Function ContinueByDeformation(ByVal maximumLambda As Double, _
        Optional ByVal progress As ICoupledContinuationProgress = Nothing) As CCoupledDeformationContinuation
    If mMeta.InternalStatus <> rsNumericalFailure Or mMeta.ResultCode = rcSearchBoundReached Then Exit Function
    If mHasSteelLimit Or Not HasAcceptedVariableState Or SavedFieldCount < 2 Or PreviousAcceptedLambda >= mLambda Then Exit Function
    Dim controlled As New CCoupledDeformationContinuation
    controlled.Trace mSolver, mReinforcement, mOffset, mDirection, mAcceptedFields, _
        PreviousAcceptedLambda, mLambda, mResultants, maximumLambda, mLimitTolerance, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance, mAcceptedTrace, progress
    controlled.ReportAcceptedProgress
    Set ContinueByDeformation = controlled
End Function

' Возвращает интервал арматурного события по равновесным пробам.
' Если он охватывает заданную нагрузку, ResultMeta сообщает NumFail.
' Для преднагружения границы относятся к его доле.
Friend Function CopySteelLimit(ByRef low As Double, ByRef high As Double, ByRef onOffset As Boolean, _
        Optional ByRef criticalPoint As Long = 0) As Boolean
    If Not mHasSteelLimit Then Exit Function
    low = mSteelLimitLow: high = mSteelLimitHigh: onOffset = mSteelLimitOnOffset
    criticalPoint = mSteelLimitPoint
    CopySteelLimit = True
End Function

' Нижняя граница относится к завершённому преднагружению и принятым пробам.
Friend Property Get HasAcceptedVariableState() As Boolean
    HasAcceptedVariableState = mHasAccepted And (mOffsetFraction = 1#)
End Property

' Число принятых полей сохранённой траектории, включая постоянную часть.
Friend Property Get SavedFieldCount() As Long
    If Not mAcceptedFields Is Nothing Then SavedFieldCount = mAcceptedFields.Count
End Property

' Снимает отдельную копию степеней свободы принятого шага для вывода.
Friend Sub CopySavedField(ByVal index As Long, ByRef field() As Double)
    field = mAcceptedFields.Item(index)
End Sub

' Передаёт принятые поля для восстановления напряжений в точках вывода.
' Получатель читает готовую последовательность без изменения.
Friend Property Get AcceptedFields() As Collection
    Set AcceptedFields = mAcceptedFields
End Property

' Возвращает параметры ветви до предельной точки либо последнего принятого шага.
Friend Property Get SavedTrace() As CCoupledLoadTrace
    If mLimitTrace Is Nothing Then Set SavedTrace = mAcceptedTrace Else Set SavedTrace = mLimitTrace
End Property

' Передаёт историю и усилия допустимого плеча стального предела.
' Пробы бисации используют одну историю, поэтому добавляется только
' окончательная допустимая точка, если она отличается от левого поля.
Friend Function CopyLimitState(ByRef fields As Collection, ByRef resultants() As Double) As Boolean
    If mLimitFields Is Nothing Then Exit Function
    Set fields = mLimitFields
    resultants = mLimitResultants
    CopyLimitState = True
End Function

' Продвигается до заданного lambda по текущей истории. После отказа
' доступны достигнутая точка и повтор от неё. Уменьшение lambda требует
' восстановления снимка истории или повторной подготовки траектории.
Friend Function Advance(ByVal targetLambda As Double, Optional ByVal verifyDirectSteel As CCoupledReinforcedResponse = Nothing) As Boolean
    mMeta.Clear
    Set mDirectSteel = verifyDirectSteel: mDirectTarget = targetLambda: mDirectSteelDetails = vbNullString
    mDirectSteelUncertainty = csuNone
    mRejected = 0: mAccepted = 0: mAttempts = 0
    mExactTargetRecovery = vbNullString
    mSecantFailure = vbNullString
    mTwistGuideDetails = vbNullString
    If targetLambda < mLambda Then
        mMeta.SetSolverFailure sfcInvalidInput, rkDirectState, _
            "Целевой множитель ниже достигнутого. Начните траекторию заново.", False
        Exit Function
    End If

    If mOffsetFraction < 1# Then
        If mHasOffset Then
            If Not FollowSegment(mOffsetFraction, 1#, True) Then Exit Function
        Else
            mOffsetFraction = 1#
        End If
    End If
    If targetLambda > mLambda Then
        If Not FollowSegment(mLambda, targetLambda, False) Then Exit Function
    End If
    If Not mHasAccepted Then
        If Not SolveAndAccept(0#, False) Then Exit Function ' В том числе нулевое сочетание с подтверждением состояния.
    End If
    mMeta.SetResult rsSuccess, rcNone, rkDirectState
    mMeta.DiagnosticDetails = ProgressDetails()
    Advance = True
End Function

' Сообщает, какая проверка стали осталась открытой при подтверждённом НДС.
Friend Property Get DirectSteelUncertainty() As ECoupledSteelUncertainty
    DirectSteelUncertainty = mDirectSteelUncertainty
End Property

' Возвращает копию результата последнего принятого шага текущей ветви.
' При отказе это подтверждённая промежуточная точка с собственным lambda.
Friend Sub CopyAccepted(ByRef field() As Double, ByRef resultants() As Double)
    If Not mHasAccepted Then Err.Raise 5, "CCoupledLoadContinuation", "На текущей траектории нет принятого состояния."
    field = mField: resultants = mResultants
End Sub

' Возвращает достигнутый множитель переменной части.
Friend Property Get AcceptedLambda() As Double
    AcceptedLambda = mLambda
End Property

' Возвращает множитель предыдущей принятой точки переменной части.
' До двух принятых точек возвращает -1.
Friend Property Get PreviousAcceptedLambda() As Double
    PreviousAcceptedLambda = mPreviousLambda
End Property

' Возвращает долю подтверждённой постоянной части при первом нагружении.
Friend Property Get AcceptedOffsetFraction() As Double
    AcceptedOffsetFraction = mOffsetFraction
End Property

' Число отказов, после которых нагрузка делилась на меньшие приращения.
Friend Property Get RejectedAttempts() As Long
    RejectedAttempts = mRejected
End Property

' Возвращает исход последнего запрошенного участка, включая причину отказа.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Проходит отрезок, уменьшая приращение после численного отказа.
' Дополнительные пробы точной цели ограничены общим бюджетом и последним принятым интервалом.
Private Function FollowSegment(ByRef reached As Double, ByVal target As Double, ByVal offsetStage As Boolean) As Boolean
    Dim nominal As Double, minimum As Double, increment As Double, candidate As Double
    Dim attemptedIncrement As Double
    Dim segmentStart As Double, triedExactTarget As Boolean, triedSecant As Boolean
    segmentStart = reached
    nominal = (target - segmentStart) / mInitialSteps
    minimum = nominal * MIN_STEP_RATIO
    increment = nominal
    Do While reached < target
        If mAttempts >= mMaxAttempts Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkDirectState, "Достигнут лимит проб нагружения связанного сечения.", True
            mMeta.DiagnosticDetails = ProgressDetails()
            Exit Function
        End If
        attemptedIncrement = increment
        If increment >= target - reached Then
            candidate = target ' Последняя проба должна получить нагрузку в запрошенной точке точно.
        Else
            candidate = reached + increment
            If candidate > reached And target - candidate <= target * DOUBLE_RELATIVE_ROUNDING Then _
                candidate = target ' Точная конечная нагрузка проверяется вместо отдельной пробы на остатке округления.
        End If
        If candidate <= reached Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkDirectState, "Шаг нагружения меньше точности представления множителя.", True
            mMeta.DiagnosticDetails = ProgressDetails()
            Exit Function
        End If
        If candidate = target Then
            triedExactTarget = True
            attemptedIncrement = target - reached ' Учитываем укороченную длину точной конечной пробы.
        End If
        If SolveAndAccept(candidate, offsetStage) Then
            reached = candidate
            If mPreferDeformationAfterInitial And Not offsetStage Then
                If SavedFieldCount >= 2 And mPreviousLambda < mLambda Then
                    mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, _
                        "Начальная силовая пара подтверждена; поиск продолжается при управлении деформацией.", ProgressDetails()
                    Exit Function
                End If
            End If
            triedExactTarget = False
            If increment < nominal / STEP_GROWTH Then
                increment = increment * STEP_GROWTH
            Else
                increment = nominal
            End If
        Else
            If mHasSteelLimit Or mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
            mRejected = mRejected + 1
            If candidate = target And Not offsetStage Then
                If TryRecoverTargetFromPrediction(target, triedSecant) Then
                    reached = target
                    FollowSegment = True
                    Exit Function
                End If
            End If
            If mPreferDeformationOnStall And Not offsetStage And mRejected >= 2 Then
                If SavedFieldCount >= 2 And mPreviousLambda < mLambda Then
                    mMeta.DiagnosticDetails = ProgressDetails() & " | " & mMeta.DiagnosticDetails
                    Exit Function
                End If
            End If
            increment = attemptedIncrement * 0.5 ' Вычитание близких множителей может потерять граничный шаг 1/64.
            If increment < minimum Then
                If Not offsetStage And Not triedExactTarget And mReinforcement Is Nothing And _
                        reached > segmentStart And target - reached <= nominal And mAttempts < mMaxAttempts Then
                    If TryRecoverExactTarget(target, triedSecant) Then
                        reached = target
                        FollowSegment = True
                        Exit Function
                    End If
                End If
                mMeta.ResultComment = "Не сошёлся минимальный подшаг. " & mMeta.ResultComment
                mMeta.DiagnosticDetails = ProgressDetails() & " | " & mMeta.DiagnosticDetails
                Exit Function
            End If
        End If
    Loop
    FollowSegment = True
End Function

' Проверяет конечную нагрузку от последнего принятого поля. При отказе
' сохраняет причину остановки на минимальном подшаге и историю ветви.
Private Function TryRecoverExactTarget(ByVal target As Double, ByRef triedSecant As Boolean) As Boolean
    Dim stall As CResultMeta: Set stall = mMeta
    If SolveAndAccept(target, False) Then
        mExactTargetRecovery = "точной конечной пробой."
        TryRecoverExactTarget = True
    Else
        mRejected = mRejected + 1
        Dim canPredict As Boolean: canPredict = (mMeta.InternalStatus = rsNumericalFailure)
        stall.DiagnosticDetails = stall.DiagnosticDetails & " | Точная конечная проба: " & mMeta.ResultComment
        Set mMeta = stall
        If canPredict Then TryRecoverExactTarget = TryRecoverTargetFromPrediction(target, triedSecant)
    End If
End Function

' Повторяет точную нагрузку от секущего прогноза двух принятых полей. После
' численного отказа допускает вспомогательную пробу с управлением закручиванием.
' Восстановление доступно только в прямом расчёте с завершённой постоянной частью.
Private Function TryRecoverTargetFromPrediction(ByVal target As Double, ByRef triedSecant As Boolean) As Boolean
    If triedSecant Or mAttempts >= mMaxAttempts Or Not mReinforcement Is Nothing Then Exit Function
    If mOffsetFraction <> 1# Or mPreviousLambda <= 0# Then Exit Function
    Dim previousSpan As Double: previousSpan = mLambda - mPreviousLambda
    Dim remaining As Double: remaining = target - mLambda
    If previousSpan <= mLambda * DOUBLE_RELATIVE_ROUNDING Or remaining <= 0# Or remaining > previousSpan Then Exit Function
    Dim initialField As Variant
    Dim ratio As Double: ratio = remaining / previousSpan
    If Not BuildSecantField(ratio, initialField) Then Exit Function

    Dim failure As CResultMeta: Set failure = mMeta
    triedSecant = True
    If TryEquilibrium(target, False, initialField) Then
        SaveCandidate target, False
        mExactTargetRecovery = "точной конечной пробой с секущим прогнозом."
        TryRecoverTargetFromPrediction = True
    Else
        mRejected = mRejected + 1
        mSecantFailure = mMeta.ResultComment
        Dim canGuide As Boolean: canGuide = (mMeta.InternalStatus = rsNumericalFailure)
        Set mMeta = failure
        If canGuide Then TryRecoverTargetFromPrediction = TryRecoverTargetWithTwist(target, ratio, previousSpan, remaining)
    End If
End Function

' Использует равновесие при заданном закручивании только как начальное поле.
' Принятие возможно лишь в отдельной силовой пробе на точной конечной нагрузке.
Private Function TryRecoverTargetWithTwist(ByVal target As Double, ByVal ratio As Double, _
        ByVal previousSpan As Double, ByVal remaining As Double) As Boolean
    If mMaxAttempts - mAttempts < 2 Or Not mHasDirection Or Not mSolver.UsesSpatialHistory Then Exit Function
    Dim currentTwist As Double: currentTwist = mField(TWIST_DEGREE)
    Dim previousTwist As Double: previousTwist = mPreviousField(TWIST_DEGREE)
    Dim delta As Double: delta = currentTwist - previousTwist ' Конечность обоих полей проверена в секущем прогнозе.
    If Abs(delta) <= (Abs(currentTwist) + Abs(previousTwist)) * DOUBLE_RELATIVE_ROUNDING Then Exit Function
    Dim projection() As Double: ReDim projection(1 To UBound(mField))
    projection(TWIST_DEGREE) = Sgn(delta)
    Dim currentControl As Double: currentControl = projection(TWIST_DEGREE) * currentTwist
    Dim targetControl As Double: targetControl = currentControl + ratio * Abs(delta)
    If targetControl = currentControl Then Exit Function ' Прогноз приращения меньше точности представления.

    Dim guideLambda As Double, failure As CResultMeta: Set failure = mMeta
    mAttempts = mAttempts + 1
    If Not mSolver.TryDeformationStep(mOffset, mDirection, projection, targetControl, target, guideLambda, _
            mAbsoluteTolerance, mRelativeTolerance, mMaxIterations, consistentControl:=True) Then
        mRejected = mRejected + 1
        mTwistGuideDetails = mSolver.ResultMeta.ResultComment
        mSolver.DiscardStep
        Exit Function
    End If
    If guideLambda <> guideLambda Or guideLambda <= mLambda Or _
            guideLambda - mLambda > previousSpan Or guideLambda - target > remaining Then
        mRejected = mRejected + 1
        mTwistGuideDetails = "Вспомогательный множитель вне проверяемого интервала: " & CStr(guideLambda) & "."
        mSolver.DiscardStep
        Exit Function
    End If

    Dim guideField() As Double, guideResultants() As Double, initialField As Variant
    mSolver.CopyTrial guideField, guideResultants
    mSolver.DiscardStep
    initialField = guideField
    mTwistGuideDetails = "Вспомогательный множитель: " & CStr(guideLambda) & "."
    If TryEquilibrium(target, False, initialField) Then
        SaveCandidate target, False
        mExactTargetRecovery = "точной конечной пробой после управления закручиванием."
        TryRecoverTargetWithTwist = True
    Else
        mRejected = mRejected + 1
        mTwistGuideDetails = mTwistGuideDetails & " Точная силовая проба: " & mMeta.ResultComment
        Set mMeta = failure
    End If
End Function

' Возвращает отдельный прогноз только при конечных коэффициентах и приращении поля.
Private Function BuildSecantField(ByVal ratio As Double, ByRef initialField As Variant) As Boolean
    Dim field() As Double: field = mField
    Dim i As Long, changed As Boolean, previousValue As Double
    For i = LBound(field) To UBound(field)
        previousValue = mPreviousField(i)
        If field(i) <> field(i) Or previousValue <> previousValue Then Exit Function
        If Abs(field(i)) > 1E+100 Or Abs(previousValue) > 1E+100 Then Exit Function ' Защита арифметики экстраполяции.
        field(i) = field(i) + ratio * (field(i) - previousValue)
        If field(i) <> mField(i) Then changed = True
    Next i
    If Not changed Then Exit Function
    initialField = field
    BuildSecantField = True
End Function

' Подставляет усилия этапа и принимает только сошедшееся состояние полностью.
Private Function SolveAndAccept(ByVal parameter As Double, ByVal offsetStage As Boolean) As Boolean
    If Not TryEquilibrium(parameter, offsetStage) Then Exit Function
    If Not mReinforcement Is Nothing Then
        If mReinforcement.TrialSteelUtilization > 1# Then
            mSolver.DiscardStep
            RefineSteelLimit parameter, offsetStage
            Exit Function
        End If
    End If
    SaveCandidate parameter, offsetStage
    SolveAndAccept = True
End Function

' Ищет равновесие одной пробы без принятия истории материалов.
Private Function TryEquilibrium(ByVal parameter As Double, ByVal offsetStage As Boolean, _
        Optional ByRef initialField As Variant) As Boolean
    Dim load(1 To COMPONENT_COUNT) As Double, i As Long
    For i = 1 To COMPONENT_COUNT
        If offsetStage Then
            load(i) = parameter * mOffset(i)
        Else
            load(i) = mOffset(i) + parameter * mDirection(i)
        End If
    Next i
    Dim checkSteel As Boolean
    If Not mDirectSteel Is Nothing Then
        checkSteel = (Not offsetStage And parameter = mDirectTarget) Or _
            (offsetStage And parameter = 1# And mDirectTarget = 0#)
        checkSteel = checkSteel And mDirectSteel.HasSteel
    End If
    mAttempts = mAttempts + 1
    If Not mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), _
            mAbsoluteTolerance, mRelativeTolerance, mMaxIterations, initialField, retainDirectProof:=checkSteel) Then
        Set mMeta = mSolver.ResultMeta
        Exit Function
    End If
    If checkSteel Then
        If Not mSolver.VerifyDirectSteelTrial(load, mDirectSteel, mAbsoluteTolerance, mRelativeTolerance, _
                mMaxIterations, mDirectSteelDetails, mDirectSteelUncertainty) Then
            Set mMeta = mSolver.ResultMeta
            Exit Function
        End If
    End If
    TryEquilibrium = True
End Function

' Принимает равновесную пробу целиком и при необходимости копирует поле.
Private Sub SaveCandidate(ByVal parameter As Double, ByVal offsetStage As Boolean)
    If mHasAccepted Then mPreviousLambda = mLambda
    If Not offsetStage And mLambda > 0# And mReinforcement Is Nothing Then mPreviousField = mField
    mSolver.CopyTrial mField, mResultants
    mSolver.AcceptStep
    mHasAccepted = True: mAccepted = mAccepted + 1
    If Not mAcceptedFields Is Nothing Then
        mAcceptedFields.Add mField
        If offsetStage Then mAcceptedTrace.Add parameter, 0# Else mAcceptedTrace.Add mOffsetFraction, parameter
    End If
End Sub

' Уточняет пересечение с откатом каждой пробы к одному принятому состоянию.
' При численной неудаче внешний шаг делится и подтверждается заново.
Private Sub RefineSteelLimit(ByVal high As Double, ByVal offsetStage As Boolean)
    Dim low As Double, candidate As Double, tolerance As Double
    Dim lowRatio As Double, highRatio As Double, ratio As Double, limitPoint As Long
    Dim finalField() As Double, finalResultants() As Double, keepFields As Boolean
    keepFields = (Not offsetStage And Not mAcceptedFields Is Nothing)
    low = mLambda
    If offsetStage Then low = mOffsetFraction
    lowRatio = mReinforcement.AcceptedSteelUtilization
    highRatio = mReinforcement.TrialSteelUtilization
    limitPoint = mReinforcement.TrialSteelCriticalPoint
    tolerance = mLimitTolerance
    If Not offsetStage And high >= 0.95 And low <= 1.05 Then tolerance = LowerOf(tolerance, 0.0001) ' Уточнение рядом с расчётным множителем.
    Do While high - low > tolerance Or low = 0#
        If mAttempts >= mMaxAttempts Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkCapacity, _
                "Исчерпан лимит проб уточнения предельной деформации арматуры.", True
            Exit Sub
        End If
        candidate = low + 0.5 * (high - low)
        If candidate = low Or candidate = high Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkCapacity, _
                "Машинная точность множителя ограничила уточнение предельной деформации арматуры.", True
            Exit Sub
        End If
        If Not TryEquilibrium(candidate, offsetStage) Then Exit Sub
        ratio = mReinforcement.TrialSteelUtilization
        If ratio > 1# Then
            high = candidate: highRatio = ratio
            limitPoint = mReinforcement.TrialSteelCriticalPoint
        Else
            low = candidate: lowRatio = ratio
            If keepFields Then mSolver.CopyTrial finalField, finalResultants
        End If
        mSolver.DiscardStep
    Loop
    mSteelLimitLow = low: mSteelLimitHigh = high: mSteelLimitOnOffset = offsetStage
    mSteelLimitPoint = limitPoint
    mHasSteelLimit = True
    If keepFields And (mHasAccepted Or low > mLambda) Then
        Set mLimitFields = New Collection
        Set mLimitTrace = mAcceptedTrace.CopyPrefix(mAcceptedTrace.Count)
        Dim i As Long
        For i = 1 To mAcceptedFields.Count: mLimitFields.Add mAcceptedFields.Item(i): Next i
        If low > mLambda Then
            mLimitFields.Add finalField
            mLimitTrace.Add 1#, low
            mLimitResultants = finalResultants
        Else
            mLimitResultants = mResultants
        End If
    End If
    Dim limitDetails As String
    limitDetails = ProgressDetails() & " | Интервал: [" & CStr(low) & ", " & CStr(high) & _
        "], отношения деформаций к пределам: " & CStr(lowRatio) & " и " & CStr(highRatio) & "."
    If Not offsetStage And low < 1# And high >= 1# Then
        mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, _
            "Интервал достижения предельной деформации арматуры охватывает расчётную нагрузку (lambda = 1).", limitDetails
    Else
        mMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
            "На ветви достигнута предельная деформация арматуры.", limitDetails
    End If
End Sub

' Ограничивает шаг верхней границей без Excel API.
Private Function LowerOf(ByVal value As Double, ByVal maximum As Double) As Double
    LowerOf = maximum
    If value < maximum Then LowerOf = value
End Function

' Собирает показатели ветви для отчёта, сохраняя точку последнего принятия.
Private Function ProgressDetails() As String
    ProgressDetails = "Доля постоянной части: " & CStr(mOffsetFraction) & "; lambda: " & CStr(mLambda) & _
        "; принято: " & CStr(mAccepted) & "; отклонено: " & CStr(mRejected) & "; проб: " & CStr(mAttempts)
    If Len(mDirectSteelDetails) > 0 Then ProgressDetails = ProgressDetails & " | " & mDirectSteelDetails
    If Len(mExactTargetRecovery) > 0 Then ProgressDetails = ProgressDetails & " | Равновесие подтверждено " & mExactTargetRecovery
    If Len(mSecantFailure) > 0 Then ProgressDetails = ProgressDetails & " | Секущий прогноз: " & mSecantFailure
    If Len(mTwistGuideDetails) > 0 Then ProgressDetails = ProgressDetails & " | Управление закручиванием: " & mTwistGuideDetails
End Function

' Восстанавливает точное заданное сочетание по допустимому деформационному пути.
' Интерполяция задаёт только начальное поле; результат принимается после отдельного решения всех уравнений при lambda = 1.
' Вызывается для ASD после численного отказа силовых шагов и при полностью приложенной постоянной части.
Friend Function TryRecoverDirectByDeformation(Optional ByVal progress As ICoupledContinuationProgress = Nothing) As Boolean
    If mMeta.InternalStatus <> rsNumericalFailure Or mMeta.ResultCode = rcSearchBoundReached Then Exit Function
    If mDirectSteel Is Nothing Or Not HasAcceptedVariableState Or SavedFieldCount < 2 Then Exit Function
    If PreviousAcceptedLambda >= mLambda Or mDirectTarget <> 1# Then Exit Function
    Dim failure As CResultMeta: Set failure = mMeta.Clone
    Dim history As New Collection, seedTrace As CCoupledLoadTrace, field() As Double, index As Long
    Dim prefixCount As Long, variablePoints As Long, fraction As Double, factor As Double
    For index = 1 To mAcceptedTrace.Count
        mAcceptedTrace.ReadPoint index, fraction, factor
        If fraction = 1# And factor > 0# Then variablePoints = variablePoints + 1
        If variablePoints = 2 Then prefixCount = index: Exit For
    Next index
    If prefixCount < 2 Then Exit Function
    For index = 1 To prefixCount
        field = mAcceptedFields.Item(index): history.Add field
    Next index
    Set seedTrace = mAcceptedTrace.CopyPrefix(history.Count) ' Постоянная часть сохраняется полностью, начало переменной ветви не содержит измельчённых отказавших шагов.
    Dim seedPrevious As Double, seedLambda As Double, seedResultants() As Double, seedInitial As Variant
    seedTrace.ReadPoint history.Count - 1, fraction, seedPrevious
    seedTrace.ReadPoint history.Count, fraction, seedLambda
    seedInitial = history.Item(history.Count)
    mSolver.DiscardStep
    If Not mSolver.ReplayAcceptedFields(history, history.Count - 1) Then Set mMeta = failure: Exit Function
    If Not TryEquilibrium(seedLambda, False, seedInitial) Then
        mSolver.DiscardStep: Set mMeta = failure: Exit Function
    End If
    mSolver.CopyTrial field, seedResultants: mSolver.AcceptStep
    history.Remove history.Count: history.Add field ' В начальный узел передаются фактические поле и усилия повторно решённой пробы.
    Dim controlled As New CCoupledDeformationContinuation
    controlled.Trace mSolver, mDirectSteel, mOffset, mDirection, history, seedPrevious, seedLambda, _
        seedResultants, 2#, 0.001, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance, seedTrace, progress
    controlled.ReportAcceptedProgress
    Dim branchFields As Collection, branchTrace As CCoupledLoadTrace, controlMeta As CResultMeta
    Set branchFields = controlled.AcceptedFields: Set branchTrace = controlled.SavedTrace
    Set controlMeta = controlled.ResultMeta
    If Not branchFields Is Nothing And Not branchTrace Is Nothing Then
        If branchFields.Count <> branchTrace.Count Then Err.Raise 5, "CCoupledLoadContinuation", "Поля восстановленной ветви не совпадают с её нагрузочной историей."
        Dim leftFraction As Double, rightFraction As Double, leftLambda As Double, rightLambda As Double
        Dim left() As Double, right() As Double, predicted() As Double, ratio As Double, degree As Long, initial As Variant
        For index = 2 To branchFields.Count
            branchTrace.ReadPoint index - 1, leftFraction, leftLambda
            branchTrace.ReadPoint index, rightFraction, rightLambda
            If leftFraction = 1# And rightFraction = 1# And leftLambda <= 1# And rightLambda >= 1# And rightLambda > leftLambda Then
                left = branchFields.Item(index - 1): right = branchFields.Item(index): predicted = left
                ratio = (1# - leftLambda) / (rightLambda - leftLambda)
                For degree = LBound(predicted) To UBound(predicted)
                    predicted(degree) = left(degree) + ratio * (right(degree) - left(degree))
                Next degree
                If mSolver.ReplayAcceptedFields(branchFields, index - 1) Then
                    initial = predicted
                    Dim exactSolved As Boolean
                    exactSolved = TryEquilibrium(1#, False)
                    If Not exactSolved Then
                        mSolver.DiscardStep
                        exactSolved = TryEquilibrium(1#, False, initial)
                    End If
                    If exactSolved Then
                        Set mAcceptedFields = New Collection
                        For degree = 1 To index - 1
                            field = branchFields.Item(degree): mAcceptedFields.Add field
                        Next degree
                        Set mAcceptedTrace = branchTrace.CopyPrefix(index - 1)
                        mField = left: mLambda = leftLambda: mAccepted = mAcceptedFields.Count
                        SaveCandidate 1#, False: mLambda = 1#
                        mMeta.SetResult rsSuccess, rcNone, rkDirectState
                        mMeta.DiagnosticDetails = "Силовой участок: " & failure.DiagnosticDetails & _
                            " | Деформационное продолжение: " & controlMeta.DiagnosticDetails & _
                            " | Точная нагрузка lambda = 1 решена от левого принятого состояния с исходными допусками."
                        TryRecoverDirectByDeformation = True
                        Exit Function
                    End If
                    mSolver.DiscardStep
                End If
            End If
        Next index
    End If
    Dim peak As Double, accuracy As Double, designVerified As Boolean, low As Double, high As Double, criticalPoint As Long
    Dim mechanism As String
    If controlled.CopyPeak(peak, accuracy, designVerified) Then
        If peak + accuracy < 1# Then mechanism = "Заданная нагрузка превышает подтверждённый максимум допустимой ветви. Lambda предела: " & CStr(peak) & "; оценка погрешности: " & CStr(accuracy) & "."
    ElseIf controlled.CopySteelLimit(low, high, criticalPoint) Then
        If high < 1# And controlled.BestAdmissibleLambda < 1# Then _
            mechanism = "На допустимой ветви предел деформации арматуры достигнут раньше заданной нагрузки. Интервал lambda: [" & CStr(low) & "; " & CStr(high) & "]."
    End If
    If Len(mechanism) > 0 Then
        Dim actual() As Double
        If controlled.CopyLimitState(history, actual) Then
            Set mAcceptedFields = history: Set mAcceptedTrace = branchTrace.CopyPrefix(history.Count)
            mField = history.Item(history.Count): mResultants = actual
            mAcceptedTrace.ReadPoint history.Count, mOffsetFraction, mLambda
        End If
        mMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkDirectState, mechanism, controlMeta.DiagnosticDetails
    Else
        Set mMeta = failure
        mMeta.DiagnosticDetails = failure.DiagnosticDetails & " | Деформационное восстановление: " & controlMeta.ResultComment & " | " & controlMeta.DiagnosticDetails
    End If
End Function
