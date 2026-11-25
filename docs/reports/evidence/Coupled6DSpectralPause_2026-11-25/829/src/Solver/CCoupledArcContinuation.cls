VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledArcContinuation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Локально проверенное продолжение пространственной ветви по знаковой дуге
' ==========================================================================
' Полная и половинная пробы относятся к одной принятой истории. При изломе
' допускается отдельное решение первого начала пластического нагружения.
' Класс возвращает только непринятый кандидат; разрушение не устанавливает.

Private Const ACCURACY_SCALE As Double = 0.0001
Private Const FIELD_ERROR_LIMIT As Double = 0.005
Private mSolver As CCoupledNewtonSolver
Private mBase() As Double, mDirection() As Double
Private mAbsolute As Double, mRelative As Double, mModulus As Double
Private mMaximumIterations As Long, mCalls As Long, mPrimary As Long
Private mCertificate As Variant, mDetails As String, mError As Double
Private mApproaches As Collection

' Сохраняет подготовленное ядро и исходную траекторию; внутренние пробы используют более строгую точность без изменения материала.
Friend Sub Initialize(ByVal solver As CCoupledNewtonSolver, ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maximumIterations As Long)
    If solver Is Nothing Or absoluteTolerance <= 0# Or relativeTolerance <= 0# Or maximumIterations < 2 Then _
        Err.Raise 5, "CCoupledArcContinuation", "Не заданы решатель, точность или бюджет дугового продолжения."
    Set mSolver = solver: mBase = baseLoad: mDirection = direction
    mAbsolute = absoluteTolerance * ACCURACY_SCALE: mRelative = relativeTolerance * ACCURACY_SCALE
    mMaximumIterations = maximumIterations: mModulus = solver.ReferenceModulus
End Sub

' Решает одну пару и при необходимости одно материальное событие; любой выход сохраняет принятую ревизию и общий лимит вызовов.
' Успех оставляет штатный CopyTrial с чисто пространственной мерой, пригодной для существующих проверок стали и принятия истории.
Friend Function TryStep(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, ByVal ratio As Double, _
        ByVal callAllowance As Long, ByRef resultLambda As Double, ByRef resultProjection() As Double) As Boolean
    mCalls = 0: mPrimary = 0: mError = -1#: mDetails = vbNullString: mCertificate = Empty
    Set mApproaches = New Collection
    If mSolver Is Nothing Then Err.Raise 5, "CCoupledArcContinuation", "Дуговой контекст не подготовлен."
    If fields Is Nothing Or trace Is Nothing Then Exit Function
    If fields.Count < 4 Or trace.Count <> fields.Count Or callAllowance < 3 Or ratio <= 0# Or ratio > 2# Then Exit Function
    Dim accepted() As Double, previous() As Double, movement() As Double, projection() As Double
    Dim count As Long: count = fields.Count
    accepted = fields(count): previous = fields(count - 1): movement = accepted
    Dim actualAccepted() As Double, degree As Long: mSolver.CopyAccepted actualAccepted
    If UBound(actualAccepted) <> UBound(accepted) Then Err.Raise 5, "CCoupledArcContinuation", "Размер принятого поля не совпадает с историей продолжения."
    For degree = 1 To UBound(accepted)
        If actualAccepted(degree) <> accepted(degree) Then Err.Raise 5, "CCoupledArcContinuation", "Принятая материальная история не соответствует последнему полю продолжения."
    Next degree
    Dim i As Long, spacing As Double, coordinate As Double
    For i = 1 To UBound(movement): movement(i) = accepted(i) - previous(i): Next i
    If Not mSolver.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then Exit Function
    Dim factor(1 To 4) As Double, slope(1 To 3) As Double, fraction As Double, a() As Double, b() As Double
    For i = 1 To 4
        trace.ReadPoint count - 4 + i, fraction, factor(i)
        If fraction <> 1# Or factor(i) <= 0# Then Exit Function
    Next i
    For i = 1 To 3
        If factor(i + 1) = factor(i) Then Exit Function
        a = fields(count - 4 + i): b = fields(count - 3 + i)
        slope(i) = mSolver.SpatialFieldDistance(a, b) / Abs(factor(i + 1) - factor(i))
    Next i
    Dim alpha As Double, dl As Double, secant As Double, radius As Double, q As Double, target As Double
    alpha = GeomMax(GeomMin(slope(1), slope(2)), GeomMin(GeomMax(slope(1), slope(2)), slope(3)))
    dl = factor(4) - factor(3): secant = Sqr(spacing * spacing + alpha * alpha * dl * dl)
    If secant <= 0# Then Exit Function
    radius = ratio * secant: q = alpha * alpha * dl / secant: target = q * factor(4)
    For i = 1 To UBound(projection)
        projection(i) = projection(i) * spacing / secant: target = target + projection(i) * accepted(i)
    Next i
    Dim revision As Long: revision = mSolver.AcceptedRevision
    Dim halfSeed() As Double, fullSeed() As Double, half() As Double, full() As Double, ah() As Double, af() As Double
    halfSeed = accepted: fullSeed = accepted
    For i = 1 To mSolver.PhysicalDegreeCount
        halfSeed(i) = accepted(i) + 0.5 * ratio * movement(i): fullSeed(i) = accepted(i) + ratio * movement(i)
    Next i
    Dim loadingPrediction() As Double: loadingPrediction = fullSeed
    Dim sh As Boolean, sf As Boolean, lh As Double, lf As Double, halfSteps As Long, pairSteps As Long, prediction As Variant
    Dim recent As Variant, loadingDeparture As Boolean
    loadingDeparture = mSolver.TryReadRecentLoadingDeparture(loadingPrediction, recent)
    Dim fullPrediction As Double: fullPrediction = factor(4) + ratio * dl
    If Not loadingDeparture Then
        prediction = halfSeed: mCalls = mCalls + 1
        sh = mSolver.TryDeformationStep(mBase, mDirection, projection, target + 0.5 * radius, factor(4) + 0.5 * ratio * dl, _
            lh, mAbsolute, mRelative, mMaximumIterations \ 2, consistentControl:=True, initialField:=prediction, _
            loadProjection:=q, allowHybridRecovery:=True, conditionHybridRecovery:=True)
        halfSteps = mSolver.IterationCount: mPrimary = halfSteps
        If sh Then mSolver.CopyTrial half, ah Else ReDim half(1 To 1): ReDim ah(1 To 1)
        mSolver.DiscardStep
        If sh Then
            For i = 1 To mSolver.PhysicalDegreeCount: fullSeed(i) = 2# * half(i) - accepted(i): Next i
            fullPrediction = 2# * lh - factor(4)
        End If
        prediction = fullSeed: mCalls = mCalls + 1
        sf = mSolver.TryDeformationStep(mBase, mDirection, projection, target + radius, fullPrediction, lf, _
            mAbsolute, mRelative, mMaximumIterations - halfSteps, consistentControl:=True, initialField:=prediction, _
            loadProjection:=q, allowHybridRecovery:=True, conditionHybridRecovery:=True)
        pairSteps = halfSteps + mSolver.IterationCount: mPrimary = pairSteps
        If sf Then mSolver.CopyTrial full, af Else ReDim full(1 To 1): ReDim af(1 To 1)
        mSolver.DiscardStep
    Else
        ReDim half(1 To 1): ReDim full(1 To 1): ReDim ah(1 To 1): ReDim af(1 To 1)
    End If
    Dim verified As Boolean, lambdaError As Double, lambdaBound As Double, doubled() As Double, length As Double
    lambdaError = -1#: lambdaBound = 0#
    If sf And sh Then
        length = mSolver.SpatialFieldDistance(accepted, full)
        If length > 0# Then
            doubled = half
            For i = 1 To mSolver.PhysicalDegreeCount: doubled(i) = 2# * half(i) - accepted(i): Next i
            mError = mSolver.SpatialFieldDistance(full, doubled) / length
            lambdaError = Abs(2# * lh - factor(4) - lf)
            lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(lf - factor(4)), 8# * mRelative * (1# + Abs(factor(4))))
            verified = mError <= FIELD_ERROR_LIMIT And lambdaError <= lambdaBound
            verified = verified And 0.5 * radius * mModulus > 8# * GeomMax(Bound(factor(4)), GeomMax(Bound(lf), Bound(lh)))
            verified = verified And LocalRoot(accepted, full, lf, factor(4), projection, q, alpha, radius)
            verified = verified And LocalRoot(accepted, half, lh, factor(4), projection, q, alpha, 0.5 * radius)
        End If
    End If
    Dim candidate() As Double, candidateLambda As Double, phaseSteps As Long, eventInfo As Variant, face() As Double, faceHalf() As Double
    Dim faceActual() As Double, faceHalfActual() As Double, eventHalfLambda As Double, eventError As Double, eventSolved As Boolean
    Dim eventLambda As Double, selectedKind As Long
    eventInfo = Array(0#): ReDim face(1 To 1): ReDim faceHalf(1 To 1): ReDim faceActual(1 To 1): ReDim faceHalfActual(1 To 1)
    If verified Then
        candidate = full: candidateLambda = lf: phaseSteps = pairSteps: selectedKind = 1
    ElseIf mCalls + 3 <= callAllowance Then
        If loadingDeparture Then
            eventInfo = recent: ReDim candidate(1 To 1)
            verified = TryLoadingDeparture(accepted, previous, loadingPrediction, factor(4), dl * ratio, recent, candidate, _
                candidateLambda, faceHalf, eventHalfLambda, phaseSteps)
            face = candidate: eventError = mError: eventSolved = verified: eventLambda = candidateLambda
            If verified Then selectedKind = 4
        Else
            Dim bracket() As Double, bracketLambda As Double
            If sf Then
                bracket = full: bracketLambda = lf
            ElseIf sh Then
                bracket = half: bracketLambda = lh
            Else
                bracket = fullSeed: bracketLambda = fullPrediction ' Прогноз только задаёт отрезок поиска поверхности; равновесие и пригодность найденного события проверяются заново.
            End If
            verified = TryEvent(accepted, bracket, factor(4), bracketLambda, projection, q, candidate, candidateLambda, phaseSteps, _
                eventInfo, face, faceHalf, faceActual, faceHalfActual, eventHalfLambda, eventError, eventSolved, callAllowance, eventLambda, selectedKind)
        End If
    End If
    Dim restored As Boolean, restoredLambda As Double, restoredField() As Double, actual() As Double
    If verified And mMaximumIterations - phaseSteps >= 1 And mCalls < callAllowance Then
        movement = candidate
        For i = 1 To UBound(movement): movement(i) = candidate(i) - accepted(i): Next i
        If mSolver.SelectSpatialDirectionControl(accepted, candidate, movement, resultProjection, coordinate, length) Then
            prediction = candidate: mCalls = mCalls + 1
            restored = mSolver.TryDeformationStep(mBase, mDirection, resultProjection, coordinate, candidateLambda, restoredLambda, _
                mAbsolute, mRelative, mMaximumIterations - phaseSteps, consistentControl:=True, initialField:=prediction)
            phaseSteps = phaseSteps + mSolver.IterationCount: mPrimary = mPrimary + mSolver.IterationCount
            If restored Then
                mSolver.CopyTrial restoredField, actual
                restored = restoredLambda = candidateLambda And mSolver.IterationCount = 0
                If restored Then
                    For i = 1 To UBound(candidate)
                        If candidate(i) <> restoredField(i) Then restored = False: Exit For
                    Next i
                End If
            End If
        End If
    End If
    If mSolver.AcceptedRevision <> revision Or pairSteps > mMaximumIterations Or phaseSteps > mMaximumIterations Or mCalls > callAllowance Then _
        Err.Raise 5, "CCoupledArcContinuation", "Дуговая проверка изменила принятую историю или превысила общий бюджет."
    resultLambda = candidateLambda
    mCertificate = Array(Array(Abs(CLng(sf)), Abs(CLng(sh)), Abs(CLng(verified)), Abs(CLng(restored)), lf, lh, factor(4), factor(3), _
        ratio, alpha, secant, radius, q, target, revision, mSolver.AcceptedRevision, pairSteps, phaseSteps, mCalls, mPrimary, _
        mError, lambdaError, lambdaBound, candidateLambda, eventHalfLambda, eventError, Abs(CLng(eventSolved)), eventLambda, selectedKind, mApproaches.Count), _
        projection, accepted, halfSeed, fullSeed, half, full, ah, af, eventInfo, face, faceHalf, faceActual, faceHalfActual)
    mDetails = "Знаковая дуга: парных итераций " & CStr(pairSteps) & "; всех основных итераций " & CStr(mPrimary) & _
        "; вызовов " & CStr(mCalls) & "; ошибка поля " & CStr(mError) & "; событие " & CStr(eventSolved) & "; восстановление " & CStr(restored)
    If Not restored Then mSolver.DiscardStep
    TryStep = restored
End Function

' Проверяет радиус и ориентацию в совместной метрике поля и множителя; знак изменения нагрузки заранее не ограничивается.
Private Function LocalRoot(ByRef accepted() As Double, ByRef field() As Double, ByVal factor As Double, ByVal reference As Double, _
        ByRef projection() As Double, ByVal q As Double, ByVal alpha As Double, ByVal radius As Double) As Boolean
    Dim distance As Double, progress As Double, i As Long
    distance = mSolver.SpatialFieldDistance(accepted, field)
    distance = Sqr(distance * distance + alpha * alpha * (factor - reference) ^ 2)
    progress = q * (factor - reference)
    For i = 1 To UBound(projection): progress = progress + projection(i) * (field(i) - accepted(i)): Next i
    LocalRoot = factor > 0# And distance <= 3# * radius And progress > 0#
End Function

' Решает найденную границу нагружения и половинный подход от той же истории; обе пробы делят самостоятельный бюджет фазы.
Private Function TryEvent(ByRef accepted() As Double, ByRef bracket() As Double, ByVal reference As Double, ByVal bracketLambda As Double, _
        ByRef arcProjection() As Double, ByVal q As Double, ByRef candidate() As Double, ByRef factor As Double, ByRef phaseSteps As Long, _
        ByRef eventInfo As Variant, ByRef face() As Double, ByRef half() As Double, ByRef actualFace() As Double, _
        ByRef actualHalf() As Double, ByRef halfLambda As Double, ByRef error As Double, ByRef eventSolved As Boolean, _
        ByVal callAllowance As Long, ByRef eventLambda As Double, ByRef selectedKind As Long) As Boolean
    mCalls = mCalls + 1
    eventSolved = mSolver.TryLoadingYieldBoundary(mBase, mDirection, bracket, reference, bracketLambda, mAbsolute, mRelative, _
        mMaximumIterations \ 2, face, factor, eventInfo)
    phaseSteps = mSolver.IterationCount: mPrimary = mPrimary + phaseSteps
    If Not eventSolved Then Exit Function
    eventLambda = factor
    Dim projection() As Double, movement() As Double, coordinate As Double, spacing As Double, i As Long
    movement = face
    For i = 1 To UBound(movement): movement(i) = face(i) - accepted(i): Next i
    If Not mSolver.SelectSpatialDirectionControl(accepted, face, movement, projection, coordinate, spacing) Then Exit Function
    Dim seed() As Double, prediction As Variant, solved As Boolean
    seed = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = 0.5 * (face(i) + accepted(i)): Next i
    prediction = seed: mCalls = mCalls + 1
    solved = mSolver.TryDeformationStep(mBase, mDirection, projection, coordinate - 0.5 * spacing, 0.5 * (reference + factor), _
        halfLambda, mAbsolute, mRelative, mMaximumIterations - phaseSteps, consistentControl:=True, initialField:=prediction, _
        allowHybridRecovery:=True, conditionHybridRecovery:=True)
    phaseSteps = phaseSteps + mSolver.IterationCount: mPrimary = mPrimary + mSolver.IterationCount
    If Not solved Then mSolver.DiscardStep: Exit Function
    mSolver.CopyTrial half, actualHalf: mSolver.DiscardStep
    Dim doubled() As Double, lambdaBound As Double, progress As Double
    doubled = half
    For i = 1 To mSolver.PhysicalDegreeCount: doubled(i) = 2# * half(i) - accepted(i): Next i
    error = mSolver.SpatialFieldDistance(face, doubled) / spacing
    lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))
    progress = q * (factor - reference)
    For i = 1 To UBound(arcProjection): progress = progress + arcProjection(i) * (face(i) - accepted(i)): Next i
    Dim halfYield As Double, beforeSurface As Boolean
    beforeSurface = mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), half, halfYield)
    If error <= FIELD_ERROR_LIMIT And Abs(2# * halfLambda - reference - factor) <= lambdaBound And progress > 0# And _
            factor > 0# And halfLambda > 0# And 0.5 * spacing * mModulus > 8# * GeomMax(Bound(factor), Bound(halfLambda)) And beforeSurface And halfYield < 0# Then
        candidate = face: mError = error: selectedKind = 2: TryEvent = True
        Exit Function
    End If
    If TryPreEventApproach(accepted, half, actualHalf, halfLambda, reference, projection, coordinate, spacing, _
            arcProjection, q, callAllowance, eventInfo, phaseSteps, candidate, factor) Then
        selectedKind = 3: TryEvent = True
    End If
End Function

' Читает исходное равновесное разрешение в принятой норме уравнений с фактическим внутренним допуском.
Private Function Bound(ByVal factor As Double) As Double
    Bound = mSolver.ControlledLoadTolerance(mBase, mDirection, factor, mAbsolute, mRelative)
End Function

' Возвращает число нелинейных вызовов для общего счётчика владельца, включая неудачные события и восстановление.
Friend Property Get Calls() As Long
    Calls = mCalls
End Property

' Возвращает оценку полного/половинного расхождения для адаптации следующей доли дуги.
Friend Property Get FieldError() As Double
    FieldError = mError
End Property

' Возвращает описание фактической численной работы; оно не заменяет статус физической проверки.
Friend Property Get Details() As String
    Details = mDetails
End Property

' Возвращает независимые копии исходных полей и результатов проверок одной операции без продвижения материальной истории.
Friend Function Verification() As Variant
    Verification = mCertificate
End Function

' Проверяет уменьшаемый кандидат перед событием от неизменной принятой истории; каждая новая половинная проба расходует остаток той же фазы.
' Уже найденное поле переиспользуется только как непринятый кандидат; выбор следующей величины не является принятием материала.
Private Function TryPreEventApproach(ByRef accepted() As Double, ByRef initial() As Double, ByRef initialActual() As Double, _
        ByVal initialLambda As Double, ByVal reference As Double, ByRef projection() As Double, ByVal faceCoordinate As Double, _
        ByVal faceSpacing As Double, ByRef arcProjection() As Double, ByVal q As Double, ByVal callAllowance As Long, _
        ByRef eventInfo As Variant, ByRef phaseSteps As Long, ByRef selected() As Double, ByRef selectedLambda As Double) As Boolean
    Dim candidate() As Double, actualCandidate() As Double, factor As Double, fraction As Double
    candidate = initial: actualCandidate = initialActual: factor = initialLambda: fraction = 0.5
    Dim level As Long, revision As Long: revision = mSolver.AcceptedRevision
    For level = 1 To 4
        If mCalls + 2 > callAllowance Or mMaximumIterations - phaseSteps < 2 Then Exit For ' Один вызов и одна допустимая поправка остаются для точного восстановления кандидата.
        Dim seed() As Double, prediction As Variant, i As Long, probe() As Double, actualProbe() As Double
        seed = accepted
        For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = 0.5 * (accepted(i) + candidate(i)): Next i
        prediction = seed
        Dim solved As Boolean, probeLambda As Double, steps As Long
        mCalls = mCalls + 1
        solved = mSolver.TryDeformationStep(mBase, mDirection, projection, faceCoordinate - (1# - 0.5 * fraction) * faceSpacing, _
            0.5 * (reference + factor), probeLambda, mAbsolute, mRelative, mMaximumIterations - phaseSteps - 1, _
            consistentControl:=True, initialField:=prediction, allowHybridRecovery:=True, conditionHybridRecovery:=True)
        steps = mSolver.IterationCount: phaseSteps = phaseSteps + steps: mPrimary = mPrimary + steps
        If solved Then mSolver.CopyTrial probe, actualProbe Else ReDim probe(1 To 1): ReDim actualProbe(1 To 1)
        mSolver.DiscardStep
        Dim error As Double, lambdaError As Double, lambdaBound As Double, length As Double, resolved As Boolean, verified As Boolean
        error = -1#: lambdaError = -1#: lambdaBound = 0#: length = mSolver.SpatialFieldDistance(accepted, candidate)
        resolved = False: verified = False
        If solved And length > 0# Then
            Dim doubled() As Double: doubled = probe
            For i = 1 To mSolver.PhysicalDegreeCount: doubled(i) = 2# * probe(i) - accepted(i): Next i
            error = mSolver.SpatialFieldDistance(candidate, doubled) / length
            lambdaError = Abs(2# * probeLambda - reference - factor)
            lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))
            resolved = 0.5 * length * mModulus > 8# * GeomMax(Bound(factor), Bound(probeLambda))
            verified = error <= FIELD_ERROR_LIMIT And lambdaError <= lambdaBound And resolved And factor > 0# And probeLambda > 0#
            verified = verified And PositiveArcProgress(accepted, candidate, factor, reference, arcProjection, q)
            verified = verified And PositiveArcProgress(accepted, probe, probeLambda, reference, arcProjection, q)
            verified = verified And length <= 1.5 * fraction * faceSpacing
            verified = verified And mSolver.SpatialFieldDistance(accepted, probe) <= 0.75 * fraction * faceSpacing
            Dim candidateYield As Double, probeYield As Double
            If Not mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), candidate, candidateYield) Then verified = False
            If Not mSolver.TryReadLoadingSurface(CLng(eventInfo(1)), CLng(eventInfo(2)), probe, probeYield) Then verified = False
            verified = verified And candidateYield < 0# And probeYield < 0#
        End If
        mApproaches.Add Array(Array(level, fraction, factor, probeLambda, reference, length, error, lambdaError, lambdaBound, _
            Abs(CLng(resolved)), Abs(CLng(verified)), revision, mSolver.AcceptedRevision, phaseSteps, steps, mCalls), _
            candidate, probe, projection, actualCandidate, actualProbe)
        If mSolver.AcceptedRevision <> revision Or phaseSteps > mMaximumIterations Or mCalls >= callAllowance Then _
            Err.Raise 5, "CCoupledArcContinuation", "Уточнение подхода нарушило историю или остаток общего бюджета."
        If verified Then
            selected = candidate: selectedLambda = factor: mError = error
            TryPreEventApproach = True: Exit Function
        End If
        If Not solved Or Not resolved Then Exit For
        candidate = probe: actualCandidate = actualProbe: factor = probeLambda: fraction = 0.5 * fraction
    Next level
End Function

' Проверяет ориентированное продвижение по исходной знаковой дуге без требования роста множителя нагрузки.
Private Function PositiveArcProgress(ByRef accepted() As Double, ByRef field() As Double, ByVal factor As Double, _
        ByVal reference As Double, ByRef projection() As Double, ByVal q As Double) As Boolean
    Dim progress As Double, i As Long: progress = q * (factor - reference)
    For i = 1 To UBound(projection): progress = progress + projection(i) * (field(i) - accepted(i)): Next i
    PositiveArcProgress = progress > 0#
End Function

' Возвращает отдельный набор свидетельств уточнения; изменение коллекции у потребителя не меняет рабочий объект.
Friend Property Get ApproachVerification() As Collection
    Dim result As New Collection, record As Variant
    If Not mApproaches Is Nothing Then
        For Each record In mApproaches: result.Add record: Next record
    End If
    Set ApproachVerification = result
End Property


' Проверяет полный и половинный выход по единственной недавно принятой границе CDP в общей заранее заданной окрестности.
' Направление проверяется по физическому полю и положительной материальной координате; входящая дуговая плоскость не навязывает рост нагрузки.
Private Function TryLoadingDeparture(ByRef accepted() As Double, ByRef previous() As Double, ByRef prediction() As Double, _
        ByVal reference As Double, ByVal predictedLambdaChange As Double, ByRef recent As Variant, ByRef full() As Double, _
        ByRef factor As Double, ByRef half() As Double, ByRef halfLambda As Double, ByRef phaseSteps As Long) As Boolean
    Dim spacing As Double: spacing = mSolver.SpatialFieldDistance(accepted, prediction)
    If spacing <= 0# Then Exit Function
    Dim halfTarget As Double, fullTarget As Double
    halfTarget = recent(3) + 0.25 * (recent(4) - recent(3))
    fullTarget = recent(3) + 0.5 * (recent(4) - recent(3))
    If halfTarget <= 0# Or fullTarget <= halfTarget Then Exit Function
    Dim seed() As Double, i As Long, solved As Boolean
    seed = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = accepted(i) + 0.25 * (prediction(i) - accepted(i)): Next i
    mCalls = mCalls + 1
    solved = mSolver.TryRecentLoadingCoordinate(mBase, mDirection, seed, reference + 0.25 * predictedLambdaChange, _
        halfTarget, mAbsolute, mRelative, mMaximumIterations \ 2, 1.5 * spacing, half, halfLambda)
    phaseSteps = mSolver.IterationCount: mPrimary = mPrimary + phaseSteps: mSolver.DiscardStep
    If Not solved Or mMaximumIterations - phaseSteps < 1 Then Exit Function
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = 2# * half(i) - accepted(i): Next i
    mCalls = mCalls + 1
    solved = mSolver.TryRecentLoadingCoordinate(mBase, mDirection, seed, reference + 0.5 * predictedLambdaChange, _
        fullTarget, mAbsolute, mRelative, mMaximumIterations - phaseSteps, 3# * spacing, full, factor)
    phaseSteps = phaseSteps + mSolver.IterationCount: mPrimary = mPrimary + mSolver.IterationCount: mSolver.DiscardStep
    If Not solved Then Exit Function
    Dim length As Double: length = mSolver.SpatialFieldDistance(accepted, full)
    If length <= 0# Then Exit Function
    Dim doubled() As Double: doubled = half
    For i = 1 To mSolver.PhysicalDegreeCount: doubled(i) = 2# * half(i) - accepted(i): Next i
    mError = mSolver.SpatialFieldDistance(full, doubled) / length
    Dim lambdaBound As Double, advance As Double, oldLength As Double
    lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))
    If mError > FIELD_ERROR_LIMIT Or Abs(2# * halfLambda - reference - factor) > lambdaBound Then Exit Function
    If 0.5 * length * mModulus <= 8# * GeomMax(Bound(reference), GeomMax(Bound(factor), Bound(halfLambda))) Then Exit Function
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, full, advance, oldLength) Then Exit Function
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, half, advance, oldLength) Then Exit Function
    TryLoadingDeparture = True
End Function
