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

Private mSolver As CCoupledNewtonSolver
Private mOffset() As Double, mDirection() As Double
Private mField() As Double, mResultants() As Double
Private mHasOffset As Boolean, mHasDirection As Boolean, mHasAccepted As Boolean
Private mOffsetFraction As Double, mLambda As Double
Private mPreviousLambda As Double
Private mPreferDeformationOnStall As Boolean
Private mInitialSteps As Long, mMaxAttempts As Long, mMaxIterations As Long
Private mAbsoluteTolerance As Double, mRelativeTolerance As Double
Private mRejected As Long, mAccepted As Long, mAttempts As Long
Private mMeta As CResultMeta
Private mAcceptedFields As Collection ' Копии полей принятых шагов для восстановления истории в точках вывода.
Private mReinforcement As CCoupledReinforcedResponse ' Контроль предельных деформаций включается при поиске.
Private mLimitTolerance As Double
Private mHasSteelLimit As Boolean, mSteelLimitOnOffset As Boolean
Private mSteelLimitLow As Double, mSteelLimitHigh As Double
Private mSteelLimitPoint As Long ' Материальная точка, определившая численно уточнённую верхнюю границу.
Private mLimitFields As Collection
Private mLimitResultants() As Double

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
    mPreferDeformationOnStall = False
    Set mReinforcement = Nothing
    mHasSteelLimit = False: mSteelLimitOnOffset = False
    mSteelLimitPoint = 0
    Set mLimitFields = Nothing: Erase mLimitResultants
    Erase mField: Erase mResultants
    Set mAcceptedFields = Nothing
    If saveAcceptedFields Then Set mAcceptedFields = New Collection
    Set mMeta = New CResultMeta
    mSolver.ResetHistory
End Sub

' Ищет первое пересечение принятого ограничения арматуры на свежей
' траектории. Всегда проверяет расчётное lambda = 1, если оно в диапазоне.
' Достигнутая численная или заданная граница отдельно хранится в ResultMeta.
Friend Sub TraceSteelLimit(ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal initialLambda As Double, ByVal maximumLambda As Double, ByVal lambdaTolerance As Double, _
        Optional ByVal preferDeformationOnStall As Boolean = False)
    If reinforcement Is Nothing Or mHasAccepted Or initialLambda <= 0# Or maximumLambda <= 0# Or lambdaTolerance <= 0# Then _
        Err.Raise 5, "CCoupledLoadContinuation", "Для поиска предела передайте свежую ветвь и положительные настройки Capacity."
    Set mReinforcement = reinforcement
    mLimitTolerance = lambdaTolerance
    mPreferDeformationOnStall = preferDeformationOnStall
    If Not mHasDirection Then
        If Advance(0#) Then mMeta.SetResult rsNotApplicable, rcNotApplicable, rkCapacity, _
            "Выбранный путь не изменяет усилия. Задайте путь с ненулевой переменной частью."
        Exit Sub
    End If

    Dim target As Double, increment As Double, boundary As Double
    increment = LowerOf(initialLambda, 1#)
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
        If target = maximumLambda Then Exit Do
        If increment < maximumLambda / STEP_GROWTH Then increment = increment * STEP_GROWTH
    Loop
    mMeta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, _
        "Достигнута граница Capacity.MaxLambda. Предельная деформация арматуры на пройденном участке не достигнута."
    mMeta.DiagnosticDetails = ProgressDetails()
End Sub

' Начинает деформационный поиск от конца подтверждённой силовой ветви
' после численного отказа. Техническая конечная граница уже является исходом.
Friend Function ContinueByDeformation(ByVal maximumLambda As Double) As CCoupledDeformationContinuation
    If mMeta.InternalStatus <> rsNumericalFailure Or mMeta.ResultCode = rcSearchBoundReached Then Exit Function
    If mHasSteelLimit Or Not HasAcceptedVariableState Or SavedFieldCount < 2 Or mPreviousLambda >= mLambda Then Exit Function
    Dim controlled As New CCoupledDeformationContinuation
    controlled.Trace mSolver, mReinforcement, mOffset, mDirection, mAcceptedFields, _
        mPreviousLambda, mLambda, mResultants, maximumLambda, mLimitTolerance, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance
    Set ContinueByDeformation = controlled
End Function

' Возвращает интервал пересечения арматурного критерия, подтверждённый
' равновесными пробами. Для преднагружения границы относятся к его доле.
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
Friend Function Advance(ByVal targetLambda As Double) As Boolean
    mMeta.Clear
    mRejected = 0: mAccepted = 0: mAttempts = 0
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

' Проходит один отрезок параметра со стабильным ограничением числа проб.
' Повтор численно трудного шага допускается только с меньшим приращением.
Private Function FollowSegment(ByRef reached As Double, ByVal target As Double, ByVal offsetStage As Boolean) As Boolean
    Dim nominal As Double, minimum As Double, increment As Double, candidate As Double
    nominal = (target - reached) / mInitialSteps
    minimum = nominal * MIN_STEP_RATIO
    increment = nominal
    Do While reached < target
        If mAttempts >= mMaxAttempts Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkDirectState, "Достигнут лимит проб нагружения связанного сечения.", True
            mMeta.DiagnosticDetails = ProgressDetails()
            Exit Function
        End If
        If increment >= target - reached Then
            candidate = target ' Последняя проба должна получить нагрузку в запрошенной точке точно.
        Else
            candidate = reached + increment
        End If
        If candidate <= reached Then
            mMeta.SetSolverFailure sfcNumericalFailure, rkDirectState, "Шаг нагружения меньше точности представления множителя.", True
            mMeta.DiagnosticDetails = ProgressDetails()
            Exit Function
        End If
        If SolveAndAccept(candidate, offsetStage) Then
            reached = candidate
            If increment < nominal / STEP_GROWTH Then
                increment = increment * STEP_GROWTH
            Else
                increment = nominal
            End If
        Else
            If mHasSteelLimit Or mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
            mRejected = mRejected + 1
            If mPreferDeformationOnStall And Not offsetStage And mRejected >= 2 Then
                If SavedFieldCount >= 2 And mPreviousLambda < mLambda Then
                    mMeta.DiagnosticDetails = ProgressDetails() & " | " & mMeta.DiagnosticDetails
                    Exit Function
                End If
            End If
            increment = (candidate - reached) * 0.5
            If increment < minimum Then
                mMeta.ResultComment = "Не сошёлся минимальный подшаг. " & mMeta.ResultComment
                mMeta.DiagnosticDetails = ProgressDetails() & " | " & mMeta.DiagnosticDetails
                Exit Function
            End If
        End If
    Loop
    FollowSegment = True
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
    SaveCandidate
    SolveAndAccept = True
End Function

' Ищет равновесие одной пробы без принятия истории материалов.
Private Function TryEquilibrium(ByVal parameter As Double, ByVal offsetStage As Boolean) As Boolean
    Dim load(1 To COMPONENT_COUNT) As Double, i As Long
    For i = 1 To COMPONENT_COUNT
        If offsetStage Then
            load(i) = parameter * mOffset(i)
        Else
            load(i) = mOffset(i) + parameter * mDirection(i)
        End If
    Next i
    mAttempts = mAttempts + 1
    If Not mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), _
            mAbsoluteTolerance, mRelativeTolerance, mMaxIterations) Then
        Set mMeta = mSolver.ResultMeta
        Exit Function
    End If
    TryEquilibrium = True
End Function

' Принимает равновесную пробу целиком и при необходимости копирует поле.
Private Sub SaveCandidate()
    If mHasAccepted Then mPreviousLambda = mLambda
    mSolver.CopyTrial mField, mResultants
    mSolver.AcceptStep
    mHasAccepted = True: mAccepted = mAccepted + 1
    If Not mAcceptedFields Is Nothing Then mAcceptedFields.Add mField
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
        Dim i As Long
        For i = 1 To mAcceptedFields.Count: mLimitFields.Add mAcceptedFields.Item(i): Next i
        If low > mLambda Then
            mLimitFields.Add finalField
            mLimitResultants = finalResultants
        Else
            mLimitResultants = mResultants
        End If
    End If
    mMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
        "На ветви достигнута предельная деформация арматуры.", _
        ProgressDetails() & " | Интервал: [" & CStr(low) & ", " & CStr(high) & _
        "], отношения деформаций к пределам: " & CStr(lowRatio) & " и " & CStr(highRatio) & "."
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
End Function
