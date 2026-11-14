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
' Касательная замораживает производные D вдоль продольного градиента.
' Каждая проба демпфирования пересчитывает весь остаток и материальный отклик.

Private Const COMPONENT_COUNT As Long = 6
Private Const GAUGE_COUNT As Long = 4
Private Const MIN_FRACTION As Double = 0.0078125
Private Const MIN_PROGRESS As Double = 0.0001

Private mSection As CCoupledElasticState
Private mResponse As ICoupledPointResponse
Private mSystem As CCoupledLinearSystem
Private mMeta As CResultMeta
Private mAccepted() As Double, mCandidate() As Double, mTrialResultants() As Double
Private mPointStrain(1 To COMPONENT_COUNT) As Double
Private mStress(1 To COMPONENT_COUNT) As Double
Private mTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
Private mMatrix() As Double, mExternal() As Double, mResidual() As Double
Private mCandidateReady As Boolean
Private mDegreeCount As Long, mPointCount As Long, mGaugeStart As Long
Private mIterations As Long, mEvaluations As Long, mMaterialPoint As Long
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
    mMeta.Clear
    mCandidateReady = False
    mIterations = 0: mEvaluations = 0: mMaterialPoint = 0
    mLastNorm = 0#: mMinimumPivot = 0#
    If absoluteTolerance <= 0# Or relativeTolerance <= 0# Or maxIterations < 1 Then
        RecordFailure sfcInvalidConfiguration, "Для связанного равновесия задайте положительные допуски и лимит итераций."
        Exit Function
    End If
    On Error GoTo EvaluationFailed
    mSection.LoadEquations force, mx, my, torque, qx, qy, mExternal
    Dim load(1 To COMPONENT_COUNT) As Double
    load(1) = force: load(2) = mx: load(3) = my
    load(4) = torque: load(5) = qx: load(6) = qy
    Dim current() As Double, correction() As Double, norm As Double, tolerance As Double
    Dim candidateNorm As Double, fraction As Double, i As Long, accepted As Boolean
    tolerance = absoluteTolerance + relativeTolerance * StressScale(mExternal)
    current = mAccepted
    Do
        If Not Evaluate(current, True) Then
            RecordFailure mFailureCode, "Не удалось получить материальный отклик для равновесия связанного сечения."
            Exit Function
        End If
        norm = EquilibriumNorm(load)
        If norm <= tolerance Then
            If Not GaugesVerified(tolerance) Then
                RecordFailure sfcNumericalFailure, "Уравновешенное поле нарушает условия однозначности перемещений."
                Exit Function
            End If
            mCandidate = current
            mCandidateReady = True
            mMeta.SetResult rsSuccess, rcNone, rkDirectState
            mMeta.DiagnosticDetails = DiagnosticSummary()
            TryStep = True
            Exit Function
        End If
        If mIterations >= maxIterations Then Exit Do
        correction = mResidual
        For i = 1 To mDegreeCount
            correction(i) = -correction(i)
        Next i
        If Not mSystem.Solve(mMatrix, correction, mDegreeCount) Then
            mMinimumPivot = mSystem.MinimumPivot
            RecordFailure mSystem.FailureCode, "Касательная связанного сечения вырождена или не решается с заданной точностью."
            Exit Function
        End If
        mMinimumPivot = mSystem.MinimumPivot
        ReDim mCandidate(1 To mDegreeCount)
        fraction = 1#: accepted = False
        Do While fraction >= MIN_FRACTION
            For i = 1 To mDegreeCount
                mCandidate(i) = current(i) + fraction * correction(i)
            Next i
            If Evaluate(mCandidate, False) Then
                candidateNorm = EquilibriumNorm(load)
                accepted = (candidateNorm <= tolerance Or candidateNorm < norm * (1# - MIN_PROGRESS * fraction))
            ElseIf mFailureCode <> sfcNumericalFailure And mFailureCode <> sfcSingularTangent Then
                RecordFailure mFailureCode, "Материальная проба не может быть повторена с меньшей поправкой."
                Exit Function
            End If
            If accepted Then Exit Do
            fraction = fraction * 0.5
        Loop
        If Not accepted Then
            RecordFailure sfcNumericalFailure, "Уменьшение поправки не снизило невязку связанного сечения."
            Exit Function
        End If
        current = mCandidate
        mIterations = mIterations + 1
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

' Возвращает независимый статус последнего шага.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Считает отклик всех точек от фиксированной принятой истории.
' Line search оставляет только невязку и интегралы, без заполнения матрицы.
Private Function Evaluate(ByRef field() As Double, ByVal withMatrix As Boolean) As Boolean
    mEvaluations = mEvaluations + 1
    mSection.StartEquations field, mExternal, mResidual, mMatrix, withMatrix
    Dim point As Long
    ReDim mTrialResultants(1 To COMPONENT_COUNT)
    For point = 1 To mPointCount
        mMaterialPoint = point
        mSection.PointStrain point, field, mPointStrain
        mFailureCode = mResponse.EvaluatePoint(point, mPointStrain, mStress, mTangent)
        If mFailureCode <> sfcNone Then Exit Function
        mSection.AddPointEquilibrium point, field, mStress, mTangent, mResidual, mMatrix, withMatrix, mTrialResultants
    Next point
    mMaterialPoint = 0
    Evaluate = True
End Function

' Возвращает максимум по всем силовым уравнениям и интегралам в масштабе МПа.
' Связи однозначности заданы в масштабе деформации и проверяются отдельно.
Private Function EquilibriumNorm(ByRef load() As Double) As Double
    Dim i As Long, value As Double
    mLastNorm = 0#
    For i = 1 To mGaugeStart - 1
        value = Abs(mResidual(i))
        If value > mLastNorm Then mLastNorm = value
    Next i
    Dim resultantNorm As Double: resultantNorm = mSection.ResultantResidualNorm(mTrialResultants, load)
    If resultantNorm > mLastNorm Then mLastNorm = resultantNorm
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
End Function
