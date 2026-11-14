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
Private mHasOffset As Boolean, mHasAccepted As Boolean
Private mOffsetFraction As Double, mLambda As Double
Private mInitialSteps As Long, mMaxAttempts As Long, mMaxIterations As Long
Private mAbsoluteTolerance As Double, mRelativeTolerance As Double
Private mRejected As Long, mAccepted As Long, mAttempts As Long
Private mMeta As CResultMeta
Private mAcceptedFields As Collection ' Копии полей принятых шагов для восстановления истории в точках вывода.

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
    mHasOffset = False
    For i = 1 To COMPONENT_COUNT
        If mOffset(i) <> 0# Then mHasOffset = True: Exit For
    Next i
    mInitialSteps = initialSteps: mMaxAttempts = maxAttempts: mMaxIterations = maxIterations
    mAbsoluteTolerance = absoluteTolerance: mRelativeTolerance = relativeTolerance
    mOffsetFraction = 0#: mLambda = 0#: mHasAccepted = False
    Erase mField: Erase mResultants
    Set mAcceptedFields = Nothing
    If saveAcceptedFields Then Set mAcceptedFields = New Collection
    Set mMeta = New CResultMeta
    mSolver.ResetHistory
End Sub

' Число принятых полей сохранённой траектории, включая постоянную часть.
Friend Property Get SavedFieldCount() As Long
    If Not mAcceptedFields Is Nothing Then SavedFieldCount = mAcceptedFields.Count
End Property

' Снимает отдельную копию степеней свободы принятого шага для вывода.
Friend Sub CopySavedField(ByVal index As Long, ByRef field() As Double)
    field = mAcceptedFields.Item(index)
End Sub

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
            If mMeta.InternalStatus <> rsNumericalFailure Then Exit Function
            mRejected = mRejected + 1
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
    Dim field() As Double, resultants() As Double
    mSolver.CopyTrial field, resultants
    mSolver.AcceptStep
    mField = field: mResultants = resultants
    mHasAccepted = True: mAccepted = mAccepted + 1
    If Not mAcceptedFields Is Nothing Then mAcceptedFields.Add field
    SolveAndAccept = True
End Function

' Собирает показатели ветви для отчёта, сохраняя точку последнего принятия.
Private Function ProgressDetails() As String
    ProgressDetails = "Доля постоянной части: " & CStr(mOffsetFraction) & "; lambda: " & CStr(mLambda) & _
        "; принято: " & CStr(mAccepted) & "; отклонено: " & CStr(mRejected) & "; проб: " & CStr(mAttempts)
End Function
