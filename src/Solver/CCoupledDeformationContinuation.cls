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

Private Const COMPONENT_COUNT As Long = 6
Private Const STEP_LIMIT As Long = 32
Private Const PROBE_LIMIT As Long = 96 ' Общий предел проб продолжения, деления и повторного прохождения плеч.
Private Const MIN_STEP_RATIO As Double = 0.015625
Private Const NEAR_ONE_TOLERANCE As Double = 0.0001

Private mSolver As CCoupledNewtonSolver, mReinforcement As CCoupledReinforcedResponse
Private mOffset() As Double, mDirection() As Double, mProjection() As Double
Private mAbsTol As Double, mRelTol As Double, mLambdaTolerance As Double
Private mMaxIterations As Long, mMaximumLambda As Double, mBestLambda As Double
Private mControlName As String, mProbeDetails As String, mMeta As CResultMeta
Private mPeak As TBranchPoint, mPeakAccuracy As Double
Private mHasPeak As Boolean, mDesignVerified As Boolean
Private mLocalPeak As Double, mLocalPeakAccuracy As Double
Private mFollowingLocalPeak As Boolean, mRecoveredPeaks As Long
Private mReusedRightPoints As Long
Private mShoulderProbes As Long
Private mSteelLow As Double, mSteelHigh As Double, mSteelPoint As Long
Private mProbeCount As Long
Private mLimitFields As Collection
Private mLimitResultants() As Double
Private mCurrentResultants() As Double
Private mTrace As CCoupledLoadTrace, mLimitTrace As CCoupledLoadTrace

' Начинает деформационный участок после двух принятых силовых точек.
' Выбор меры производится по их фактической разности. История первых
' точек принадлежит той же траектории, что и текущий решатель.
Friend Sub Trace(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef offset() As Double, ByRef direction() As Double, ByVal history As Collection, _
        ByVal previousLambda As Double, ByVal currentLambda As Double, ByRef initialResultants() As Double, _
        ByVal maximumLambda As Double, ByVal lambdaTolerance As Double, ByVal maximumIterations As Long, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, Optional ByVal loadTrace As CCoupledLoadTrace = Nothing)
    Set mSolver = solver: Set mReinforcement = reinforcement
    mOffset = offset: mDirection = direction
    mMaxIterations = maximumIterations: mMaximumLambda = maximumLambda
    mLambdaTolerance = lambdaTolerance: mAbsTol = absoluteTolerance: mRelTol = relativeTolerance
    mBestLambda = currentLambda: mDesignVerified = (currentLambda >= 1#)
    mHasPeak = False: mSteelPoint = 0: mProbeCount = 0
    mLocalPeak = 0#: mLocalPeakAccuracy = 0#
    mFollowingLocalPeak = False: mRecoveredPeaks = 0
    mReusedRightPoints = 0
    mShoulderProbes = 0
    Set mLimitFields = Nothing: Erase mLimitResultants
    Set mTrace = Nothing: Set mLimitTrace = Nothing
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

    Dim left As TBranchPoint, current As TBranchPoint, trial As TBranchPoint
    Dim previousField() As Double, field() As Double, increment As Double, minimum As Double, nominal As Double
    left.HistoryIndex = history.Count - 1: left.Lambda = previousLambda
    current.HistoryIndex = history.Count: current.Lambda = currentLambda
    previousField = history(left.HistoryIndex): field = history(current.HistoryIndex)
    If Not mSolver.SelectDeformationControl(previousField, field, mProjection, current.Coordinate, nominal, mControlName) Then
        StopSearch "Не удалось выбрать меняющуюся меру деформации."
        Exit Sub
    End If
    left.Coordinate = current.Coordinate - nominal
    increment = nominal * 0.5
    minimum = increment * MIN_STEP_RATIO

    Dim attempt As Long, reinforced As Boolean
    If Not reinforcement Is Nothing Then reinforced = reinforcement.HasSteel
    For attempt = 1 To PROBE_LIMIT
        If mProbeCount >= PROBE_LIMIT Then Exit For
        If attempt > STEP_LIMIT And mLocalPeak = 0# Then Exit For
        trial.Coordinate = current.Coordinate + increment
        If increment < minimum Or trial.Coordinate <= current.Coordinate Then Exit For
        If Not Probe(trial, current.Lambda, False) Then
            increment = increment * 0.5
        ElseIf trial.Lambda > maximumLambda Then
            mSolver.DiscardStep
            mMeta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, _
                "Достигнута граница Capacity.MaxLambda до подтверждения предельного состояния.", Details()
            Exit Sub
        ElseIf BeyondSteelLimit() Then
            If ResolveSteelCrossing(history, current, trial) Then Exit Sub
            increment = increment * 0.5
        ElseIf mFollowingLocalPeak Or trial.Lambda > current.Lambda + LoadResolution(current.Lambda) Then
            If trial.Lambda <= 0# Then
                mSolver.DiscardStep
                StopSearch "После локального максимума множитель вышел из области положительной нагрузки."
                Exit Sub
            End If
            If Not mDesignVerified And current.Lambda < 1# And trial.Lambda >= 1# Then
                VerifyDesignLoad current, trial
                If Not Probe(trial, current.Lambda, False) Then
                    StopSearch "Не удалось повторить пробу после проверки заданных усилий."
                    Exit Sub
                End If
                If BeyondSteelLimit() Then
                    mSolver.DiscardStep
                    StopSearch "Повторная проба не подтвердила допустимость арматуры."
                    Exit Sub
                End If
            End If
            AcceptPoint history, trial
            left = current: current = trial
            If current.Lambda > mBestLambda Then mBestLambda = current.Lambda
            If mFollowingLocalPeak Then CheckRecovery current.Lambda
            If increment < nominal / 1.25 Then increment = increment * 1.25
        ElseIf trial.Lambda < current.Lambda - LoadResolution(current.Lambda) Then
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
                increment = increment * 0.5
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

' Продолжает от уже проверенного правого плеча. Последняя уточняющая
' проходка завершилась именно в этом материальном состоянии.
Private Sub ResumeAfterLocalPeak(ByRef history As Collection, ByRef current As TBranchPoint, ByRef right As TBranchPoint)
    mLocalPeak = mPeak.Lambda: mLocalPeakAccuracy = mPeakAccuracy
    mHasPeak = False: mFollowingLocalPeak = True
    Set history = mLimitFields
    Set mTrace = mLimitTrace: Set mLimitTrace = Nothing
    mCurrentResultants = mLimitResultants
    Set mLimitFields = Nothing: Erase mLimitResultants
    current = right
    mReusedRightPoints = mReusedRightPoints + current.HistoryIndex - mPeak.HistoryIndex
End Sub

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
    lambda = mPeak.Lambda: accuracy = mPeakAccuracy: designVerified = mDesignVerified
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

' Передаёт арматурный предел на восходящем деформационном участке.
Friend Function CopySteelLimit(ByRef low As Double, ByRef high As Double, ByRef criticalPoint As Long) As Boolean
    If mSteelPoint = 0 Then Exit Function
    low = mSteelLow: high = mSteelHigh: criticalPoint = mSteelPoint
    CopySteelLimit = True
End Function

' Возвращает независимый статус контролируемого участка.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Решает пробу при заданной мере и по запросу оценивает влияние невязки.
Private Function Probe(ByRef point As TBranchPoint, ByVal initialLambda As Double, _
        ByVal estimateUncertainty As Boolean) As Boolean
    If mProbeCount >= PROBE_LIMIT Then Exit Function
    mProbeCount = mProbeCount + 1
    point.Uncertainty = 0#
    If Not mSolver.TryDeformationStep(mOffset, mDirection, mProjection, point.Coordinate, initialLambda, _
            point.Lambda, mAbsTol, mRelTol, mMaxIterations) Then
        mProbeDetails = mSolver.ResultMeta.DiagnosticDetails
        Exit Function
    End If
    If estimateUncertainty Then
        If Not mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, point.Uncertainty) Then
            mSolver.DiscardStep
            mProbeDetails = "Не удалось оценить чувствительность множителя к невязке."
            Exit Function
        End If
    End If
    Probe = True
End Function

' Продвигает принятую материальную историю и фиксирует индекс её поля.
Private Sub AcceptPoint(ByVal history As Collection, ByRef point As TBranchPoint)
    Dim field() As Double
    mSolver.CopyTrial field, mCurrentResultants
    mSolver.AcceptStep
    history.Add field
    If Not mTrace Is Nothing Then mTrace.Add 1#, point.Lambda
    point.HistoryIndex = history.Count
End Sub

' Проверяет точное lambda = 1 от левого плеча перед продвижением истории.
' Контрольная мера должна попасть внутрь проверяемого участка.
Private Sub VerifyDesignLoad(ByRef left As TBranchPoint, ByRef right As TBranchPoint)
    If mProbeCount >= PROBE_LIMIT Then Exit Sub
    mProbeCount = mProbeCount + 1
    Dim load(1 To COMPONENT_COUNT) As Double, i As Long
    For i = 1 To COMPONENT_COUNT: load(i) = mOffset(i) + mDirection(i): Next i
    If Not mSolver.TryStep(load(1), load(2), load(3), load(4), load(5), load(6), mAbsTol, mRelTol, mMaxIterations) Then Exit Sub
    Dim field() As Double, resultants() As Double, coordinate As Double
    mSolver.CopyTrial field, resultants
    coordinate = CoordinateOf(field)
    If Not BeyondSteelLimit() Then mDesignVerified = (coordinate >= left.Coordinate And coordinate <= right.Coordinate)
    mSolver.DiscardStep
End Sub

' Повторяет оба плеча от одного принятого левого состояния.
' Для предела без арматуры уточняет максимум до допуска. Для
' локального максимума железобетона сохраняет проверенное правое плечо.
Private Function VerifyPeak(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint, ByVal keepRight As Boolean) As Boolean
    If middle.Lambda <= left.Lambda + LoadResolution(middle.Lambda) Then Exit Function
    Dim first As TBranchPoint, second As TBranchPoint, tolerance As Double
    tolerance = mLambdaTolerance
    If Abs(middle.Lambda - 1#) <= 0.05 And tolerance > NEAR_ONE_TOLERANCE Then tolerance = NEAR_ONE_TOLERANCE
    Dim sampling As Double, refinedFields As Collection, actual As Variant
    Dim accuracy As Double, divisions As Long
    If keepRight Then
        If Not ReplayShoulders(history, left, middle, right, 2, second, sampling, refinedFields, actual, True) Then Exit Function
        accuracy = Abs(middle.Lambda - second.Lambda)
        If sampling > accuracy Then accuracy = sampling
        accuracy = accuracy + second.Uncertainty
        If mBestLambda - second.Lambda > accuracy Then accuracy = mBestLambda - second.Lambda
    Else
        If Not ReplayShoulders(history, left, middle, right, 2, first, sampling) Then Exit Function
        divisions = 4
        Do
            If Not ReplayShoulders(history, left, middle, right, divisions, second, sampling, refinedFields, actual) Then Exit Function
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
    mPeak = second: mPeakAccuracy = accuracy: mHasPeak = True
    Set mLimitFields = refinedFields: mLimitResultants = actual
    If second.Lambda > mBestLambda Then mBestLambda = second.Lambda
    mMeta.SetResult rsSuccess, rcCheckPassed, rkCapacity, _
        "Подтверждён первый максимум прослеженной ветви.", Details() & " | Оценка погрешности lambda: " & CStr(accuracy) & _
        " | Координаты плеч: " & CStr(left.Coordinate) & ", " & CStr(right.Coordinate)
    VerifyPeak = True
End Function

' Воспроизводит левые принятые поля, затем рассчитывает более мелкие
' шаги только вперёд. Сравнивает пик с обоими концами с учётом невязки.
Private Function ReplayShoulders(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint, ByVal divisions As Long, _
        ByRef peak As TBranchPoint, ByRef sampling As Double, Optional ByRef limitFields As Collection = Nothing, _
        Optional ByRef limitResultants As Variant, Optional ByVal keepRight As Boolean = False) As Boolean
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
            If mProbeCount < PROBE_LIMIT Then mShoulderProbes = mShoulderProbes + 1
            If Not Probe(point, previousLambda, True) Then Exit Function
            If BeyondSteelLimit() Or point.Lambda > mMaximumLambda Then
                mSolver.DiscardStep
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
    Dim dropLeft As Double, dropRight As Double
    For i = 1 To index - 1
        point = samples(i)
        If point.Lambda - point.Uncertainty > samples(i - 1).Lambda + samples(i - 1).Uncertainty And _
                point.Lambda - point.Uncertainty > samples(i + 1).Lambda + samples(i + 1).Uncertainty Then
            If point.Lambda - point.Uncertainty <= left.Lambda + LoadResolution(point.Lambda) Then Exit Function
            If point.Lambda - point.Uncertainty <= samples(index).Lambda + samples(index).Uncertainty Then Exit Function
            peak = point
            dropLeft = peak.Lambda - samples(i - 1).Lambda + samples(i - 1).Uncertainty
            dropRight = peak.Lambda - samples(i + 1).Lambda + samples(i + 1).Uncertainty
            sampling = dropLeft
            If dropRight > sampling Then sampling = dropRight
            If saveFields Then
                Set limitFields = New Collection
                Dim j As Long, fieldCount As Long
                fieldCount = i
                If keepRight Then
                    fieldCount = index
                    right = samples(index)
                End If
                For j = 1 To left.HistoryIndex: limitFields.Add history.Item(j): Next j
                For j = 1 To fieldCount: limitFields.Add refined.Item(j): Next j
                If Not mTrace Is Nothing Then
                    Set mLimitTrace = mTrace.CopyPrefix(left.HistoryIndex)
                    For j = 1 To fieldCount: mLimitTrace.Add 1#, samples(j).Lambda: Next j
                End If
                limitResultants = integrated.Item(fieldCount)
            End If
            ReplayShoulders = True
            Exit Function
        End If
    Next i
End Function

' Уточняет пересечение стального критерия от одного принятого поля.
' Падение множителя внутри интервала оставляет порядок механизмов
' неопределённым и требует сокращения внешнего шага.
Private Function ResolveSteelCrossing(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef right As TBranchPoint) As Boolean
    Dim low As TBranchPoint, high As TBranchPoint, probePoint As TBranchPoint
    Dim point As Long, attempt As Long, tolerance As Double
    Dim field() As Double, resultants() As Double
    low = left: high = right
    point = mReinforcement.TrialSteelCriticalPoint
    mSolver.DiscardStep
    If high.Lambda <= low.Lambda Then Exit Function
    tolerance = mLambdaTolerance
    If high.Lambda >= 0.95 And low.Lambda <= 1.05 And tolerance > NEAR_ONE_TOLERANCE Then tolerance = NEAR_ONE_TOLERANCE
    For attempt = 1 To STEP_LIMIT
        If high.Lambda - low.Lambda <= tolerance Then
            If low.Lambda <= 0# Then Exit Function
            If mLocalPeak > 0# And high.Lambda < mBestLambda - mLocalPeakAccuracy Then
                StopSearch "Арматурный предел достигнут при нагрузке ниже ранее подтверждённого равновесия. Нужна оценка порядка предельных состояний."
                ResolveSteelCrossing = True
                Exit Function
            End If
            mSteelLow = low.Lambda: mSteelHigh = high.Lambda: mSteelPoint = point
            Set mLimitFields = New Collection
            Dim i As Long
            For i = 1 To left.HistoryIndex: mLimitFields.Add history.Item(i): Next i
            If Not mTrace Is Nothing Then Set mLimitTrace = mTrace.CopyPrefix(left.HistoryIndex)
            If low.Lambda > left.Lambda Then
                mLimitFields.Add field
                If Not mLimitTrace Is Nothing Then mLimitTrace.Add 1#, low.Lambda
                mLimitResultants = resultants
            Else
                mLimitResultants = mCurrentResultants
            End If
            mMeta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, _
                "Достигнута предельная деформация арматуры на возрастающем участке.", _
                Details() & " | Интервал: [" & CStr(mSteelLow) & ", " & CStr(mSteelHigh) & "]"
            ResolveSteelCrossing = True
            Exit Function
        End If
        probePoint.Coordinate = (low.Coordinate + high.Coordinate) * 0.5
        If probePoint.Coordinate <= low.Coordinate Or probePoint.Coordinate >= high.Coordinate Then Exit Function
        If Not Probe(probePoint, low.Lambda, False) Then Exit Function
        If probePoint.Lambda <= low.Lambda Or probePoint.Lambda >= high.Lambda Then
            mSolver.DiscardStep
            Exit Function
        End If
        If BeyondSteelLimit() Then
            point = mReinforcement.TrialSteelCriticalPoint: high = probePoint
        Else
            low = probePoint
            mSolver.CopyTrial field, resultants
        End If
        mSolver.DiscardStep
    Next attempt
End Function

' Определяет превышение по пробной точке до фиксации истории.
Private Function BeyondSteelLimit() As Boolean
    If Not mReinforcement Is Nothing Then BeyondSteelLimit = (mReinforcement.TrialSteelUtilization > 1#)
End Function

' Считывает выбранную физическую меру из поля.
Private Function CoordinateOf(ByRef field() As Double) As Double
    Dim i As Long
    For i = 1 To UBound(field): CoordinateOf = CoordinateOf + mProjection(i) * field(i): Next i
End Function

' Выделяет изменение множителя для выбора пробного интервала. В итоговой
' проверке дополнительно используется чувствительность к невязке.
Private Function LoadResolution(ByVal lambda As Double) As Double
    LoadResolution = 8# * mRelTol * (1# + Abs(lambda))
End Function

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
        "; равновесие при lambda = 1 подтверждено: " & CStr(mDesignVerified) & "; проб: " & CStr(mProbeCount)
    If mLocalPeak > 0# Then
        Details = Details & "; локальный максимум lambda: " & CStr(mLocalPeak) & _
            "; погрешность локального максимума: " & CStr(mLocalPeakAccuracy) & _
            "; максимумов превзойдено после спада: " & CStr(mRecoveredPeaks)
        Details = Details & "; повторно использовано точек правого плеча: " & CStr(mReusedRightPoints)
        Details = Details & "; проб при проверке плеч: " & CStr(mShoulderProbes)
    End If
    If Len(mProbeDetails) > 0 Then Details = Details & " | Последняя неудачная проба: " & mProbeDetails
End Function
