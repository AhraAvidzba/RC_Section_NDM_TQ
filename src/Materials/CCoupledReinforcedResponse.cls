VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledReinforcedResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Материальный отклик бетона с направленной ненапрягаемой арматурой
' ==========================================================================
' Бетонный закон вычисляется также в точках замещённого объёма.
' В этих точках суммируется отклик стали и вычитается совпадающий
' отклик бетона. История закладывается для каждой точки интегрирования.

Implements ICoupledPointResponse

Private Const COMPONENT_COUNT As Long = 6
Private mConcrete As ICoupledPointResponse
Private mSteel() As CCoupledSteelModel
Private mDirection() As Double
Private mPlastic() As Double, mBackstress() As Double
Private mAxialStress() As Double ' Фактическое продольное напряжение стали до вычитания замещённого бетона.
Private mFirstSteel As Long, mLastSteel As Long, mFirstTransverse As Long
Private mAcceptedBuffer As Long, mTrialBuffer As Long
Private mTrialSteelUtilization As Double
Private mAcceptedSteelUtilization As Double
Private mTrialSteelPoint As Long, mAcceptedSteelPoint As Long
Private mTrialTransversePoint As Long, mAcceptedTransversePoint As Long
Private mTrialTransverseUtilization As Double, mAcceptedTransverseUtilization As Double
Private mTrialLongitudinalUtilization As Double, mAcceptedLongitudinalUtilization As Double
Private mTrialLongitudinalPoint As Long, mAcceptedLongitudinalPoint As Long

' Один раз сопоставляет точкам их диаграммы и проекции.
' Models индексируется классом стали, указанным в сформированной геометрии.
Friend Sub Initialize(ByVal operators As CCoupledElasticState, ByVal concrete As ICoupledPointResponse, _
        ByVal models As Object)
    If operators Is Nothing Or concrete Is Nothing Then _
        Err.Raise 5, "CCoupledReinforcedResponse", "Для связанного отклика требуется подготовленное сечение и бетон."
    Set mConcrete = Nothing
    Erase mSteel: Erase mDirection
    mFirstSteel = operators.ConcretePointCount + 1
    mLastSteel = operators.PointCount
    mFirstTransverse = mLastSteel + 1
    If mLastSteel >= mFirstSteel Then
        If models Is Nothing Then Err.Raise 5, "CCoupledReinforcedResponse", "Не подготовлены диаграммы стали."
        ReDim mSteel(mFirstSteel To mLastSteel)
        ReDim mDirection(1 To COMPONENT_COUNT, mFirstSteel To mLastSteel)
        Dim point As Long, j As Long, projection(1 To COMPONENT_COUNT) As Double, material As String
        For point = mFirstSteel To mLastSteel
            If mFirstTransverse > mLastSteel And operators.TransverseGroupAt(point) > 0 Then mFirstTransverse = point
            material = operators.SteelClassAt(point)
            If Not models.Exists(material) Then Err.Raise vbObjectError + 1781, "CCoupledReinforcedResponse", _
                "Для связанного расчёта нет параметров стали класса " & material & "."
            Set mSteel(point) = models.Item(material)
            operators.SteelProjection point, projection
            For j = 1 To COMPONENT_COUNT
                mDirection(j, point) = projection(j)
            Next j
        Next point
    End If
    Set mConcrete = concrete
    ICoupledPointResponse_ClearHistory
End Sub

' Показывает наличие продольной или поперечной арматуры в балансе сечения.
Friend Property Get HasSteel() As Boolean
    HasSteel = (mFirstSteel > 0 And mLastSteel >= mFirstSteel)
End Property

' Показывает максимальное отношение деформации к пределу на последнем
' полностью вычисленном пробном поле. Первый вызов точки сбрасывает максимум.
Friend Property Get TrialSteelUtilization() As Double
    TrialSteelUtilization = mTrialSteelUtilization
End Property

' Показывает максимум принятого состояния, не изменяемый отклонёнными пробами.
Friend Property Get AcceptedSteelUtilization() As Double
    AcceptedSteelUtilization = mAcceptedSteelUtilization
End Property

' Возвращает точку максимального использования на последней полной пробе.
Friend Property Get TrialSteelCriticalPoint() As Long
    TrialSteelCriticalPoint = mTrialSteelPoint
End Property

' Возвращает контрольную точку принятого состояния по всей арматуре.
Friend Property Get AcceptedSteelCriticalPoint() As Long
    AcceptedSteelCriticalPoint = mAcceptedSteelPoint
End Property

' Возвращает точку максимального использования поперечной арматуры.
' При её отсутствии сохраняется нулевой индекс.
Friend Property Get AcceptedTransverseCriticalPoint() As Long
    AcceptedTransverseCriticalPoint = mAcceptedTransversePoint
End Property

' Возвращает использование по деформации в принятой точке хомута или ветви.
Friend Property Get AcceptedTransverseSteelUtilization() As Double
    AcceptedTransverseSteelUtilization = mAcceptedTransverseUtilization
End Property

' Возвращает максимальное использование отдельно по продольным стержням.
Friend Property Get AcceptedLongitudinalUtilization() As Double
    AcceptedLongitudinalUtilization = mAcceptedLongitudinalUtilization
End Property

' Возвращает продольную контрольную точку отдельно от хомутов.
Friend Property Get AcceptedLongitudinalPoint() As Long
    AcceptedLongitudinalPoint = mAcceptedLongitudinalPoint
End Property

' Считывает принятое напряжение стержня материальной точки в МПа.
Friend Property Get AcceptedSteelStress(ByVal point As Long) As Double
    AcceptedSteelStress = mAxialStress(point, mAcceptedBuffer)
End Property

' Обновляет независимую материальную точку. В поправке замещения площадь
' берётся у геометрических операторов и здесь не учитывается повторно.
Private Function ICoupledPointResponse_EvaluatePoint(ByVal point As Long, ByRef strain() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double, _
        Optional ByVal tangentMode As ECoupledTangentMode = ctmFull, _
        Optional ByVal recordTrial As Boolean = True) As ESolverFailureCode
    If (tangentMode = ctmFullBranchConsistent Or tangentMode = ctmRefinedBranchConsistent) And recordTrial Then
        ICoupledPointResponse_EvaluatePoint = sfcInvalidConfiguration
        Exit Function
    End If
    If point = 1 And recordTrial Then
        mTrialSteelUtilization = 0#: mTrialSteelPoint = 0
        mTrialTransverseUtilization = 0#: mTrialTransversePoint = 0
        mTrialLongitudinalUtilization = 0#: mTrialLongitudinalPoint = 0
    End If
    ICoupledPointResponse_EvaluatePoint = mConcrete.EvaluatePoint(point, strain, stress, tangent, tangentMode, recordTrial)
    If ICoupledPointResponse_EvaluatePoint <> sfcNone Or point < mFirstSteel Then Exit Function
    Dim j As Long, k As Long, axialStrain As Double, axialStress As Double, axialTangent As Double
    Dim utilization As Double, plastic As Double, backstress As Double
    For j = 1 To COMPONENT_COUNT
        axialStrain = axialStrain + mDirection(j, point) * strain(j)
    Next j
    mSteel(point).EvaluateTrial axialStrain, mPlastic(point, mAcceptedBuffer), mBackstress(point, mAcceptedBuffer), _
        axialStress, axialTangent, plastic, backstress
    If recordTrial Then
        mPlastic(point, mTrialBuffer) = plastic: mBackstress(point, mTrialBuffer) = backstress
        mAxialStress(point, mTrialBuffer) = axialStress
        utilization = mSteel(point).StrainUtilization(axialStrain)
        If utilization > mTrialSteelUtilization Then
            mTrialSteelUtilization = utilization: mTrialSteelPoint = point
        End If
        If point >= mFirstTransverse Then
            If utilization > mTrialTransverseUtilization Or mTrialTransversePoint = 0 Then
                mTrialTransverseUtilization = utilization: mTrialTransversePoint = point
            End If
        ElseIf utilization > mTrialLongitudinalUtilization Or mTrialLongitudinalPoint = 0 Then
            mTrialLongitudinalUtilization = utilization: mTrialLongitudinalPoint = point
        End If
    End If
    For j = 1 To COMPONENT_COUNT
        stress(j) = axialStress * mDirection(j, point) - stress(j)
        If tangentMode <> ctmStressOnly Then
            For k = 1 To COMPONENT_COUNT
                If tangentMode <> ctmAxialGradient Or k = 3 Or k >= 5 Then _
                    tangent(j, k) = axialTangent * mDirection(j, point) * mDirection(k, point) - tangent(j, k)
            Next k
        End If
    Next j
End Function

' Читает пробный критерий вытесненного или обычного бетона.
' Коэффициенты знака замещения учитываются в уравнениях, а не здесь.
Friend Function TryReadTrialYield(ByVal point As Long, ByRef strain() As Double, ByRef value As Double, _
        ByRef gradient() As Double, Optional ByVal needGradient As Boolean = True) As Boolean
    If Not TypeOf mConcrete Is CCoupledConcreteResponse Then Exit Function
    Dim concrete As CCoupledConcreteResponse: Set concrete = mConcrete
    TryReadTrialYield = concrete.TryReadTrialYield(point, strain, value, gradient, needGradient)
End Function

' Читает фактическую ветвь бетонной составляющей от принятой истории
' без записи отклика бетона или стали в пробные буферы.
Friend Function TryReadConcreteBranch(ByVal point As Long, ByRef strain() As Double, _
        ByRef plasticIncrement As Boolean, ByRef principalSign As Long, ByRef substeps As Long) As Boolean
    If Not TypeOf mConcrete Is CCoupledConcreteResponse Then Exit Function
    Dim concrete As CCoupledConcreteResponse: Set concrete = mConcrete
    TryReadConcreteBranch = concrete.TryReadConcreteBranch(point, strain, plasticIncrement, principalSign, substeps)
End Function

' Публикует новый буфер после принятия бетонных проб, без копирования массивов.
Private Sub ICoupledPointResponse_AcceptTrial()
    mConcrete.AcceptTrial
    mAcceptedBuffer = mTrialBuffer
    mTrialBuffer = 3 - mAcceptedBuffer
    mAcceptedSteelUtilization = mTrialSteelUtilization
    mAcceptedLongitudinalUtilization = mTrialLongitudinalUtilization: mAcceptedLongitudinalPoint = mTrialLongitudinalPoint
    mAcceptedSteelPoint = mTrialSteelPoint
    mAcceptedTransversePoint = mTrialTransversePoint
    mAcceptedTransverseUtilization = mTrialTransverseUtilization
End Sub

' Удаляет историю независимого сочетания в обоих материалах.
Private Sub ICoupledPointResponse_ClearHistory()
    mConcrete.ClearHistory
    mAcceptedBuffer = 1: mTrialBuffer = 2
    mTrialSteelUtilization = 0#: mAcceptedSteelUtilization = 0#
    mTrialLongitudinalUtilization = 0#: mAcceptedLongitudinalUtilization = 0#
    mTrialLongitudinalPoint = 0: mAcceptedLongitudinalPoint = 0
    mTrialSteelPoint = 0: mAcceptedSteelPoint = 0
    mTrialTransversePoint = 0: mAcceptedTransversePoint = 0
    mTrialTransverseUtilization = 0#: mAcceptedTransverseUtilization = 0#
    Erase mPlastic: Erase mBackstress: Erase mAxialStress
    If mLastSteel >= mFirstSteel Then
        ReDim mPlastic(mFirstSteel To mLastSteel, 1 To 2)
        ReDim mBackstress(mFirstSteel To mLastSteel, 1 To 2)
        ReDim mAxialStress(mFirstSteel To mLastSteel, 1 To 2)
    End If
End Sub

' Ищет знаковый предел внутри линейной оценки поправки поля. Оценка
' применяется только для запуска уточнения на той же материальной истории.
Friend Function SteelBoundaryWithinCorrection(ByVal operators As CCoupledElasticState, ByRef field() As Double, _
        ByRef correction() As Double, ByRef governingPoint As Long, ByRef unresolved As ECoupledSteelUncertainty) As Boolean
    Const ROUNDING_STEP As Double = 2.22044604925031E-16
    Dim point As Long, degree As Long, strain As Double, limit As Double, otherLimit As Double, margin As Double
    Dim projection() As Double, movement As Double, arithmeticScale As Double
    Dim family As Long, best(0 To 1) As Double, candidates(0 To 1) As Long, exceeded(0 To 1) As Boolean
    governingPoint = 0
    For point = mFirstSteel To mLastSteel
        family = 0
        If point >= mFirstTransverse Then family = 1
        ReadSteelStrain operators, point, field, strain, limit
        operators.CopySteelStrainProjection point, projection
        movement = 0#: arithmeticScale = Abs(limit)
        For degree = 1 To UBound(projection)
            movement = movement + Abs(projection(degree) * correction(degree))
            arithmeticScale = arithmeticScale + Abs(projection(degree) * field(degree))
        Next degree
        otherLimit = mSteel(point).SignedStrainLimit(-limit)
        margin = Abs(strain - limit)
        If Abs(strain - otherLimit) < margin Then
            limit = otherLimit: margin = Abs(strain - limit)
            arithmeticScale = arithmeticScale + Abs(limit)
        End If
        ' Две геометрические компоненты на координату и шесть проекций оси.
        If margin <= movement + (4# * UBound(projection) + 12#) * ROUNDING_STEP * arithmeticScale Then
            margin = margin / Abs(limit)
            If candidates(family) = 0 Or margin < best(family) Then
                candidates(family) = point: best(family) = margin
            End If
        ElseIf mSteel(point).StrainUtilization(strain) > 1# Then
            exceeded(family) = True ' Дополнительная точка не меняет уже подтверждённый статус этой группы.
        End If
    Next point
    unresolved = csuNone
    If Not exceeded(0) And candidates(0) <> 0 Then
        unresolved = csuLongitudinal: governingPoint = candidates(0)
    End If
    If Not exceeded(1) And candidates(1) <> 0 Then
        unresolved = unresolved Or csuTransverse
        If governingPoint = 0 Then governingPoint = candidates(1)
    End If
    SteelBoundaryWithinCorrection = (governingPoint <> 0)
End Function

' Меняет одну координату поля по наиболее чувствительному коэффициенту
' проекции. Полученное поле требует отдельной проверки равновесия и стали.
Friend Function ProjectSteelBoundary(ByVal operators As CCoupledElasticState, ByVal point As Long, _
        ByRef field() As Double) As Boolean
    Dim strain As Double, limit As Double, projection() As Double
    ReadSteelStrain operators, point, field, strain, limit
    operators.CopySteelStrainProjection point, projection
    Dim i As Long, chosen As Long, maximum As Double
    For i = 1 To UBound(projection)
        If Abs(projection(i)) > maximum Then
            maximum = Abs(projection(i)): chosen = i
        End If
    Next i
    If chosen = 0 Then Exit Function
    field(chosen) = field(chosen) + (limit - strain) / projection(chosen)
    ReadSteelStrain operators, point, field, strain, limit
    ProjectSteelBoundary = (strain = limit)
End Function

' Считывает деформацию и предел её знака тем же порядком вычислений,
' который используется для материальной точки стали.
Private Sub ReadSteelStrain(ByVal operators As CCoupledElasticState, ByVal point As Long, _
        ByRef field() As Double, ByRef strain As Double, ByRef limit As Double)
    Dim tensor(1 To COMPONENT_COUNT) As Double, i As Long
    operators.PointStrain point, field, tensor
    strain = 0#
    For i = 1 To COMPONENT_COUNT
        strain = strain + mDirection(i, point) * tensor(i)
    Next i
    limit = mSteel(point).SignedStrainLimit(strain)
End Sub

' Возвращает группы заданной арматуры. При отсутствии подтверждённой
' чувствительности каждая фактически заданная группа требует уточнения.
Friend Property Get PresentSteelCriteria() As ECoupledSteelUncertainty
    If Not HasSteel Then Exit Property
    Dim groups As ECoupledSteelUncertainty
    If mFirstSteel < mFirstTransverse Then groups = csuLongitudinal
    If mFirstTransverse <= mLastSteel Then groups = groups Or csuTransverse
    PresentSteelCriteria = groups
End Property
