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
' Каждая проба демпфирования пересчитывает весь остаток. При недостаточной
' касательной добавляются производные D вдоль продольного градиента.

Private Const COMPONENT_COUNT As Long = 6
Private Const GAUGE_COUNT As Long = 4
Private Const MIN_FRACTION As Double = 0.0078125
Private Const MIN_PROGRESS As Double = 0.0001
Private Const GRADIENT_DIFFERENCE_STEP As Double = 0.0000006 ' Амплитуда возмущения максимальной компоненты деформации.

Private mSection As CCoupledElasticState
Private mResponse As ICoupledPointResponse
Private mSystem As CCoupledLinearSystem
Private mMeta As CResultMeta
Private mAccepted() As Double, mCandidate() As Double, mTrialResultants() As Double
Private mPointStrain(1 To COMPONENT_COUNT) As Double
Private mStress(1 To COMPONENT_COUNT) As Double
Private mTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mPlusTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mMinusTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mGradientTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mMatrix() As Double, mExternal() As Double, mResidual() As Double
Private mCandidateReady As Boolean
Private mControlActive As Boolean
Private mControlProjection() As Double, mTargetControl As Double
Private mBaseLoad() As Double, mLoadDirection() As Double
Private mBaseEquations() As Double, mDirectionEquations() As Double
Private mTrialLambda As Double
Private mDegreeCount As Long, mPointCount As Long, mGaugeStart As Long
Private mIterations As Long, mEvaluations As Long, mMaterialPoint As Long
Private mMaterialFailures As Long, mLastFailedPoint As Long
Private mGradientBuilds As Long
Private mReferenceModulus As Double, mLastNorm As Double, mMinimumPivot As Double
Private mFailureCode As ESolverFailureCode

' Подключает готовое сечение и независимый набор материальной истории.
' Модуль в МПа нужен для масштабирования уравнений перемещений.
Friend Sub Initialize(ByVal section As CCoupledElasticState, ByVal response As ICoupledPointResponse, _
        ByVal referenceModulus As Double)
    If section Is Nothing Or response Is Nothing Or referenceModulus <= 0# Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Для равновесия требуется подготовленное сечение, материал и положительный модуль."
    mDegreeCount = section.DegreeCount
    mPointCount = section.PointCount
    If mDegreeCount <= GAUGE_COUNT Or mPointCount < 1 Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Геометрические операторы связанного сечения не подготовлены."
    mGaugeStart = mDegreeCount - GAUGE_COUNT + 1
    Set mSection = section: Set mResponse = response
    Set mSystem = New CCoupledLinearSystem: Set mMeta = New CResultMeta
    mReferenceModulus = referenceModulus
    ResetHistory
End Sub

' Начинает независимую траекторию. Принятое поле и история других LC здесь не используются.
Friend Sub ResetHistory()
    mResponse.ClearHistory
    ReDim mAccepted(1 To mDegreeCount)
    mCandidateReady = False
    mMeta.Clear
End Sub

' Ищет пробное равновесие от последнего принятого поля. Моменты относятся
' к численному началу сечения; силы в Н, моменты в Н·мм.
' Само решение не публикует материальную историю.
Friend Function TryStep(ByVal force As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal torque As Double, ByVal qx As Double, ByVal qy As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long) As Boolean
    Dim load(1 To COMPONENT_COUNT) As Double
    load(1) = force: load(2) = mx: load(3) = my
    load(4) = torque: load(5) = qx: load(6) = qy
    mControlActive = False
    mSection.LoadEquations force, mx, my, torque, qx, qy, mExternal
    TryStep = SolveStep(load, 0#, absoluteTolerance, relativeTolerance, maxIterations)
End Function

' Решает равновесие при управлении одной линейной мерой деформации.
' Множитель нагрузки определяется вместе с полем и возвращается только
' для сошедшейся пробы. Материал остаётся в принятом ранее состоянии.
Friend Function TryDeformationStep(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef controlProjection() As Double, ByVal targetControl As Double, ByVal initialLambda As Double, _
        ByRef trialLambda As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maxIterations As Long) As Boolean
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
    mBaseLoad = baseLoad: mLoadDirection = direction
    mControlProjection = controlProjection: mTargetControl = targetControl
    mSection.LoadEquations baseLoad(1), baseLoad(2), baseLoad(3), baseLoad(4), baseLoad(5), baseLoad(6), mBaseEquations
    mSection.LoadEquations direction(1), direction(2), direction(3), direction(4), direction(5), direction(6), mDirectionEquations
    ReDim mExternal(1 To mDegreeCount)
    mControlActive = True
    Dim load(1 To COMPONENT_COUNT) As Double
    TryDeformationStep = SolveStep(load, initialLambda, absoluteTolerance, relativeTolerance, maxIterations)
    If TryDeformationStep Then trialLambda = mTrialLambda
End Function

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

' Решает одну силовую или деформационно управляемую пробу с общей
' невязкой, ограничителями поправок и фиксацией материальных состояний.
Private Function SolveStep(ByRef load() As Double, ByVal initialLambda As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long) As Boolean
    mMeta.Clear
    mCandidateReady = False
    mIterations = 0: mEvaluations = 0: mMaterialPoint = 0: mGradientBuilds = 0
    mMaterialFailures = 0: mLastFailedPoint = 0
    mLastNorm = 0#: mMinimumPivot = 0#
    If absoluteTolerance <= 0# Or relativeTolerance <= 0# Or maxIterations < 1 Then
        RecordFailure sfcInvalidConfiguration, "Для связанного равновесия задайте положительные допуски и лимит итераций."
        Exit Function
    End If
    On Error GoTo EvaluationFailed
    Dim current() As Double, correction() As Double, norm As Double, tolerance As Double
    Dim candidateNorm As Double, fraction As Double, i As Long, accepted As Boolean
    Dim bestTrialNorm As Double, hasValidTrial As Boolean
    Dim useGradientTangent As Boolean
    Dim currentLambda As Double, candidateLambda As Double
    currentLambda = initialLambda: candidateLambda = initialLambda
    If mControlActive Then SetControlledLoad currentLambda, load
    tolerance = absoluteTolerance + relativeTolerance * StressScale(mExternal)
    current = mAccepted
    Do
RecomputeTangent:
        If mControlActive Then SetControlledLoad currentLambda, load
        If Not Evaluate(current, True, useGradientTangent) Then
            RecordFailure mFailureCode, "Не удалось получить материальный отклик для равновесия связанного сечения."
            Exit Function
        End If
        norm = EquilibriumNorm(load, current)
        If norm <= tolerance Then
            If Not GaugesVerified(tolerance) Then
                RecordFailure sfcNumericalFailure, "Уравновешенное поле нарушает условия однозначности перемещений."
                Exit Function
            End If
            mCandidate = current
            PublishCandidate currentLambda
            SolveStep = True
            Exit Function
        End If
        If mIterations >= maxIterations Then Exit Do
        If Not SolveCorrection(current, correction) Then
            mMinimumPivot = mSystem.MinimumPivot
            If Not useGradientTangent And mSection.HasAxialGradient(current) Then
                useGradientTangent = True
                GoTo RecomputeTangent
            End If
            RecordFailure mSystem.FailureCode, "Касательная связанного сечения вырождена или не решается с заданной точностью."
            Exit Function
        End If
        mMinimumPivot = mSystem.MinimumPivot
        ReDim mCandidate(1 To mDegreeCount)
        fraction = 1#: accepted = False: hasValidTrial = False
        Do While fraction >= MIN_FRACTION
            For i = 1 To mDegreeCount
                mCandidate(i) = current(i) + fraction * correction(i)
            Next i
            If mControlActive Then
                candidateLambda = currentLambda + fraction * correction(mDegreeCount + 1)
                SetControlledLoad candidateLambda, load
            End If
            If Evaluate(mCandidate, False) Then
                candidateNorm = EquilibriumNorm(load, mCandidate)
                If Not hasValidTrial Or candidateNorm < bestTrialNorm Then bestTrialNorm = candidateNorm
                hasValidTrial = True
                accepted = (candidateNorm <= tolerance Or candidateNorm < norm * (1# - MIN_PROGRESS * fraction))
            ElseIf mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
                RecordFailure mFailureCode, "Материальная проба не может быть повторена с меньшей поправкой."
                Exit Function
            End If
            If accepted Then Exit Do
            fraction = fraction * 0.5
        Loop
        If Not accepted Then
            If Not useGradientTangent And mSection.HasAxialGradient(current) Then
                useGradientTangent = True
                GoTo RecomputeTangent
            End If
            RecordFailure sfcNumericalFailure, "Уменьшение поправки не снизило невязку связанного сечения."
            Dim details As String: details = mMeta.DiagnosticDetails
            If hasValidTrial Then
                details = details & "; невязка перед демпфированием, МПа: " & CStr(norm) & _
                    "; лучшая невязка проб, МПа: " & CStr(bestTrialNorm)
            Else
                details = details & "; ни одна уменьшенная поправка не дала материального отклика"
            End If
            mMeta.DiagnosticDetails = details
            Exit Function
        End If
        mIterations = mIterations + 1
        If candidateNorm <= tolerance And GaugesVerified(tolerance) Then
            PublishCandidate candidateLambda
            SolveStep = True
            Exit Function
        End If
        current = mCandidate
        currentLambda = candidateLambda
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


' Публикует только пробное поле, чья невязка проверена по всей системе.
Private Sub PublishCandidate(ByVal lambda As Double)
    mTrialLambda = lambda
    mCandidateReady = True
    mMeta.SetResult rsSuccess, rcNone, rkDirectState
    mMeta.DiagnosticDetails = DiagnosticSummary()
End Sub

' Обновляет шесть внешних усилий и их заранее нормированные уравнения.
Private Sub SetControlledLoad(ByVal lambda As Double, ByRef load() As Double)
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

' Добавляет неизвестный множитель и уравнение управления деформацией.
Private Sub BuildControlledMatrix(ByRef augmented() As Double, ByVal size As Long)
    Dim i As Long, j As Long
    ReDim augmented(1 To size, 1 To size)
    For i = 1 To mDegreeCount
        For j = 1 To mDegreeCount: augmented(j, i) = mMatrix(j, i): Next j
        augmented(size, i) = -mDirectionEquations(i)
        If i < mGaugeStart Then augmented(i, size) = mControlProjection(i) * mReferenceModulus
    Next i
End Sub

' Оценивает чувствительность множителя к невязке при том же принятом
' состоянии. Сопряжённая система даёт строку обратного оператора по lambda.
' Ошибка повторной сборки исключает принятие пробы.
Friend Function TrialLambdaUncertainty(ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByRef uncertainty As Double) As Boolean
    RequireCandidate
    If Not mControlActive Then Exit Function
    If Not Evaluate(mCandidate, True, True) Then
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
    norm = absoluteTolerance + relativeTolerance * StressScale(mExternal)
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
    For i = 1 To mGaugeStart - 1
        value = value + field(i) * mControlProjection(i)
    Next i
    ControlResidual = value * mReferenceModulus
End Function

' Публикует пробное поле и материальную историю после инженерной приёмки шага.
' Повторный вызов требует нового успешного решения.
Friend Sub AcceptStep()
    RequireCandidate
    Dim acceptedField() As Double: acceptedField = mCandidate
    mResponse.AcceptTrial
    mAccepted = acceptedField
    mCandidateReady = False
End Sub

' Отбрасывает вычисленную пробу после поиска или пользовательской отмены.
Friend Sub DiscardStep()
    mCandidateReady = False
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

' Восстанавливает состояние после известных принятых полей. Напряжения
' и история вычисляются по той же последовательности без поиска НДС.
' Вызывающий передаёт поля этой секции с начала той же траектории.
Friend Function ReplayAcceptedFields(ByVal history As Collection, ByVal lastField As Long) As Boolean
    If history Is Nothing Then Err.Raise 5, "CCoupledNewtonSolver", "Для восстановления не сохранена история полей."
    If lastField < 0 Or lastField > history.Count Then _
        Err.Raise 5, "CCoupledNewtonSolver", "Точка восстановления вне сохранённой истории полей."
    ResetHistory
    Dim index As Long, point As Long, field() As Double
    For index = 1 To lastField
        field = history(index)
        For point = 1 To mPointCount
            mSection.PointStrain point, field, mPointStrain
            If mResponse.EvaluatePoint(point, mPointStrain, mStress, mTangent, ctmStressOnly) <> sfcNone Then
                ResetHistory
                Exit Function
            End If
        Next point
        mResponse.AcceptTrial
        mAccepted = field
    Next index
    ReplayAcceptedFields = True
End Function

' Возвращает независимый статус последнего шага.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Считает отклик всех точек от фиксированной принятой истории.
' Line search оставляет только невязку и интегралы, без заполнения матрицы.
Private Function Evaluate(ByRef field() As Double, ByVal withMatrix As Boolean, _
        Optional ByVal useGradientTangent As Boolean = False) As Boolean
    mEvaluations = mEvaluations + 1
    mSection.StartEquations field, mExternal, mResidual, mMatrix, withMatrix
    Dim point As Long, useAxialGradient As Boolean, tangentMode As ECoupledTangentMode
    useAxialGradient = mSection.HasAxialGradient(field)
    If Not withMatrix Then
        If useAxialGradient Then tangentMode = ctmAxialGradient Else tangentMode = ctmStressOnly
    End If
    useGradientTangent = useGradientTangent And withMatrix And useAxialGradient
    If useGradientTangent Then mGradientBuilds = mGradientBuilds + 1
    Dim differenceScale As Double, plusOk As Boolean, minusOk As Boolean
    ReDim mTrialResultants(1 To COMPONENT_COUNT)
    For point = 1 To mPointCount
        mMaterialPoint = point
        mSection.PointStrain point, field, mPointStrain
        If useGradientTangent Then TakeGradientProbes point, field, differenceScale, plusOk, minusOk
        mFailureCode = mResponse.EvaluatePoint(point, mPointStrain, mStress, mTangent, tangentMode)
        If mFailureCode <> sfcNone Then
            mMaterialFailures = mMaterialFailures + 1: mLastFailedPoint = point
            Exit Function
        End If
        mSection.AddPointEquilibrium point, field, mStress, mTangent, mResidual, mMatrix, withMatrix, mTrialResultants, useAxialGradient
        If differenceScale > 0# Then
            If Not plusOk And Not minusOk Then
                mFailureCode = sfcNumericalFailure
                mMaterialFailures = mMaterialFailures + 1: mLastFailedPoint = point
                Exit Function
            End If
            AddGradientDerivative point, differenceScale, plusOk, minusOk
        End If
    Next point
    mMaterialPoint = 0
    Evaluate = True
End Function

' При неизменной принятой истории берёт две пробные касательные вдоль Bp·d.
' Базовая проба пересчитывается после возмущений и сохраняет полную историю.
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
    plusOk = mResponse.EvaluatePoint(point, perturbed, perturbedStress, mPlusTangent) = sfcNone
    For i = 1 To COMPONENT_COUNT
        perturbed(i) = mPointStrain(i) - (gradient(i) / magnitude) * GRADIENT_DIFFERENCE_STEP
    Next i
    minusOk = mResponse.EvaluatePoint(point, perturbed, perturbedStress, mMinusTangent) = sfcNone
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
    ReDim mExternal(1 To mDegreeCount)
    If Not Evaluate(field, withMatrix, useGradientTangent) Then Exit Function
    residual = mResidual
    If withMatrix Then tangent = mMatrix
    TestEquilibriumLinearization = True
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

' Проверяет, что кандидат относится к последней успешной пробе.
Private Sub RequireCandidate()
    If Not mCandidateReady Then Err.Raise 5, "CCoupledNewtonSolver", "Нет подтверждённого равновесия пробного шага."
End Sub

' Формирует общий статус численного отказа, сохраняя диагностический контекст.
Private Sub RecordFailure(ByVal code As ESolverFailureCode, ByVal reason As String)
    mMeta.SetSolverFailure code, rkDirectState, reason, (mEvaluations > 0)
    mMeta.DiagnosticDetails = DiagnosticSummary()
End Sub

' Возвращает параметры последней итерации, необходимые для проверки отказа.
Private Function DiagnosticSummary() As String
    DiagnosticSummary = "Итерации: " & CStr(mIterations) & "; пробы: " & CStr(mEvaluations) & _
        "; невязка, МПа: " & CStr(mLastNorm) & "; мин. ведущий элемент: " & CStr(mMinimumPivot)
    If mMaterialPoint > 0 Then DiagnosticSummary = DiagnosticSummary & "; материальная точка: " & CStr(mMaterialPoint)
    If mMaterialFailures > 0 Then DiagnosticSummary = DiagnosticSummary & "; отказов материала: " & CStr(mMaterialFailures) & _
        "; последняя точка отказа: " & CStr(mLastFailedPoint)
    If mGradientBuilds > 0 Then DiagnosticSummary = DiagnosticSummary & "; уточнений касательной: " & CStr(mGradientBuilds)
End Function
