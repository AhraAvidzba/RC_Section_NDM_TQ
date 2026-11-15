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
' Поиск первого максимума связанной ветви при управлении деформацией
' ==========================================================================
' Приращения меры деформации положительны. Плечи предполагаемого максимума
' проверяются повторно с меньшим шагом от восстановленной истории левого
' плеча. Численный отказ, плато и предел диапазона не считаются максимумом.

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
Private mSteelLow As Double, mSteelHigh As Double, mSteelPoint As Long
Private mProbeCount As Long

' Начинает деформационный участок после двух принятых силовых точек.
' Выбор меры производится по их фактической разности. История первых
' точек принадлежит той же траектории, что и текущий решатель.
Friend Sub Trace(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef offset() As Double, ByRef direction() As Double, ByVal history As Collection, _
        ByVal previousLambda As Double, ByVal currentLambda As Double, _
        ByVal maximumLambda As Double, ByVal lambdaTolerance As Double, ByVal maximumIterations As Long, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double)
    Set mSolver = solver: Set mReinforcement = reinforcement
    mOffset = offset: mDirection = direction
    mMaxIterations = maximumIterations: mMaximumLambda = maximumLambda
    mLambdaTolerance = lambdaTolerance: mAbsTol = absoluteTolerance: mRelTol = relativeTolerance
    mBestLambda = currentLambda: mDesignVerified = (currentLambda >= 1#)
    mHasPeak = False: mSteelPoint = 0: mProbeCount = 0
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

    Dim attempt As Long
    For attempt = 1 To STEP_LIMIT
        If mProbeCount >= PROBE_LIMIT Then Exit For
        trial.Coordinate = current.Coordinate + increment
        If increment < minimum Or trial.Coordinate <= current.Coordinate Then Exit For
        If Not Probe(trial, current.Lambda, False) Then
            increment = increment * 0.5
        ElseIf trial.Lambda > maximumLambda Then
            mSolver.DiscardStep
            mMeta.SetResult rsNumericalFailure, rcSearchBoundReached, rkCapacity, _
                "Достигнута граница Capacity.MaxLambda до подтверждения максимума.", Details()
            Exit Sub
        ElseIf BeyondSteelLimit() Then
            If ResolveSteelCrossing(current, trial) Then Exit Sub
            increment = increment * 0.5
        ElseIf trial.Lambda > current.Lambda + LoadResolution(current.Lambda) Then
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
            If increment < nominal / 1.25 Then increment = increment * 1.25
        ElseIf trial.Lambda < current.Lambda - LoadResolution(current.Lambda) Then
            mSolver.DiscardStep
            If VerifyPeak(history, left, current, trial) Then Exit Sub
            If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
                StopSearch "Не удалось восстановить историю возрастающей ветви."
                Exit Sub
            End If
            increment = increment * 0.5
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

' Возвращает максимальный подтверждённый множитель на пройденной
' допустимой ветви, даже если общий предел остался неопределённым.
Friend Property Get BestAdmissibleLambda() As Double
    BestAdmissibleLambda = mBestLambda
End Property

' Передаёт подтверждённый максимум и оценку его численной точности.
Friend Function CopyPeak(ByRef lambda As Double, ByRef accuracy As Double, ByRef designVerified As Boolean) As Boolean
    If Not mHasPeak Then Exit Function
    lambda = mPeak.Lambda: accuracy = mPeakAccuracy: designVerified = mDesignVerified
    CopyPeak = True
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
    Dim field() As Double, resultants() As Double
    mSolver.CopyTrial field, resultants
    mSolver.AcceptStep
    history.Add field
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

' Повторяет оба плеча от одного принятого левого состояния. Делит
' интервалы до тех пор, пока не стабилизируется оценка максимума.
Private Function VerifyPeak(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef middle As TBranchPoint, ByRef right As TBranchPoint) As Boolean
    If middle.Lambda <= left.Lambda + LoadResolution(middle.Lambda) Then Exit Function
    Dim first As TBranchPoint, second As TBranchPoint, tolerance As Double
    tolerance = mLambdaTolerance
    If Abs(middle.Lambda - 1#) <= 0.05 And tolerance > NEAR_ONE_TOLERANCE Then tolerance = NEAR_ONE_TOLERANCE
    Dim sampling As Double
    If Not ReplayShoulders(history, left, middle, right, 2, first, sampling) Then Exit Function
    Dim accuracy As Double, divisions As Long
    divisions = 4
    Do
        If Not ReplayShoulders(history, left, middle, right, divisions, second, sampling) Then Exit Function
        accuracy = Abs(first.Lambda - second.Lambda)
        If sampling > accuracy Then accuracy = sampling
        accuracy = accuracy + first.Uncertainty + second.Uncertainty
        If mBestLambda - second.Lambda > accuracy Then accuracy = mBestLambda - second.Lambda
        If accuracy <= tolerance Then Exit Do
        first = second
        divisions = divisions * 2
    Loop While divisions <= 16
    If accuracy > tolerance Then Exit Function
    mPeak = second: mPeakAccuracy = accuracy: mHasPeak = True
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
        ByRef peak As TBranchPoint, ByRef sampling As Double) As Boolean
    If Not mSolver.ReplayAcceptedFields(history, left.HistoryIndex) Then Exit Function
    Dim start As Double, finish As Double, previousLambda As Double, point As TBranchPoint
    Dim samples() As TBranchPoint, shoulder As Long, i As Long, index As Long
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
            If Not Probe(point, previousLambda, True) Then Exit Function
            If BeyondSteelLimit() Or point.Lambda > mMaximumLambda Then
                mSolver.DiscardStep
                Exit Function
            End If
            mSolver.AcceptStep
            index = index + 1
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
            ReplayShoulders = True
            Exit Function
        End If
    Next i
End Function

' Уточняет пересечение стального критерия от одного принятого поля.
' Падение множителя внутри интервала оставляет порядок механизмов
' неопределённым и требует сокращения внешнего шага.
Private Function ResolveSteelCrossing(ByRef left As TBranchPoint, ByRef right As TBranchPoint) As Boolean
    Dim low As TBranchPoint, high As TBranchPoint, probePoint As TBranchPoint
    Dim point As Long, attempt As Long, tolerance As Double
    low = left: high = right
    point = mReinforcement.TrialSteelCriticalPoint
    mSolver.DiscardStep
    If high.Lambda <= low.Lambda Then Exit Function
    tolerance = mLambdaTolerance
    If high.Lambda >= 0.95 And low.Lambda <= 1.05 And tolerance > NEAR_ONE_TOLERANCE Then tolerance = NEAR_ONE_TOLERANCE
    For attempt = 1 To STEP_LIMIT
        If high.Lambda - low.Lambda <= tolerance Then
            If low.Lambda <= 0# Then Exit Function
            mSteelLow = low.Lambda: mSteelHigh = high.Lambda: mSteelPoint = point
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
    If mBestLambda > 0# Then reason = reason & " Допустимое равновесие подтверждено до lambda = " & Format$(mBestLambda, "0.000000") & "."
    mMeta.SetResult rsNumericalFailure, rcNumericalFailure, rkCapacity, reason, Details()
End Sub

' Записывает выбранную управляемую величину и область проверенной ветви.
Private Function Details() As String
    Details = "Мера деформации: " & mControlName & "; наибольший допустимый lambda: " & CStr(mBestLambda) & _
        "; равновесие при lambda = 1 подтверждено: " & CStr(mDesignVerified) & "; проб: " & CStr(mProbeCount)
    If Len(mProbeDetails) > 0 Then Details = Details & " | Последняя неудачная проба: " & mProbeDetails
End Function
