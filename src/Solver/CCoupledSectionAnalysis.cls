VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledSectionAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Связанное НДС для пакета сочетаний одного сечения
' ==========================================================================
' Подготавливает геометрические операторы один раз на запуск. Материальные
' параметры берутся из выбранного набора характеристик и явной калибровки.
' Каждое сочетание начинает нагружение с нулевой материальной историей.

Private Const INITIAL_STEPS As Long = 4
Private Const STEP_ATTEMPTS As Long = 64
Private Const STRESS_TOLERANCE As Double = 0.0000005 ' Нормированная невязка равновесия, МПа.
Private Const RELATIVE_TOLERANCE As Double = 0.0000005

Private mSection As CSectionModel, mTransverse As CTransverseRebarLayout
Private mProvider As CMaterialModelProvider, mOperators As CCoupledElasticState
Private mRevision As Long, mProviderRevision As Long, mMaxIterations As Long
Private mPoisson As Double, mGt As Double, mGc As Double, mLength As Double
Private mYieldRatio As Double, mPeakStrain As Double, mSteelClass As String
Private mEb As Double, mCenterX As Double, mCenterY As Double
Private mSupplementarySteel As Object, mSteelReferences As Object ' Модели по набору характеристик, ссылки применённых классов.
Private mSteelCache As Object, mOperatorCache As Object ' Неизменяемые модели и операторы по точному набору характеристик в пределах запуска.
Private mDiagramCache As Object ' Точки фактических диаграмм для публикации одного набора характеристик.

' Принимает калибровку в мм, МПа и Н/мм, а размеры деформаций в долях единицы.
' Геометрия и диаграммы остаются неизменными между строками текущего пакета.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal transverse As CTransverseRebarLayout, _
        ByVal provider As CMaterialModelProvider, ByVal poisson As Double, ByVal gt As Double, ByVal gc As Double, _
        ByVal localizationLength As Double, ByVal yieldRatio As Double, ByVal peakStrain As Double, _
        ByVal steelClass As String, ByVal maxIterations As Long, Optional ByVal supplementarySteel As Object = Nothing)
    If section Is Nothing Or transverse Is Nothing Or provider Is Nothing Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Передайте сечение, поперечное армирование и параметры материалов."
    If Not transverse.IsSpecified Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "В блоке Transverse задайте поперечную арматуру или выберите «Нет»."
    If poisson < 0# Or poisson >= 0.5 Or gt <= 0# Or gc <= 0# Or localizationLength <= 0# Or _
            yieldRatio <= 0# Or yieldRatio >= 1# Or peakStrain <= 0# Or maxIterations < 1 Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Проверьте коэффициент Пуассона, энергии, характерную длину и параметры пика бетона."
    Set mSection = section: Set mTransverse = transverse: Set mProvider = provider
    mRevision = section.Revision: mProviderRevision = provider.Revision
    mPoisson = poisson: mGt = gt: mGc = gc: mLength = localizationLength
    mYieldRatio = yieldRatio: mPeakStrain = peakStrain
    mSteelClass = Trim$(steelClass): mMaxIterations = maxIterations
    Set mSupplementarySteel = supplementarySteel
    Set mSteelCache = CreateObject("Scripting.Dictionary")
    Set mOperatorCache = CreateObject("Scripting.Dictionary")
    Set mDiagramCache = CreateObject("Scripting.Dictionary")
    Set mSteelReferences = CreateObject("Scripting.Dictionary"): mSteelReferences.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To section.RebarCount
        RegisterSteelClass section.RebarSteelClassForMaterial(i, mSteelClass), "rngNDMSectionGeometry, стержень " & section.RebarID(i)
    Next i
    For i = 1 To transverse.Count
        RegisterSteelClass transverse.SteelClass(i), transverse.SourceReference(i)
    Next i
    Set mOperators = Nothing
    Dim area As Double
    section.ConcreteMeshCentroid mCenterX, mCenterY, area
End Sub

' Число различных применённых классов для предварительной проверки таблиц.
Friend Property Get SteelClassCount() As Long
    SteelClassCount = mSteelReferences.Count
End Property

' Передаёт неизменяемые пути подготовленных групп для вывода принятого НДС.
Friend Property Get TransverseLayout() As CTransverseRebarLayout
    Set TransverseLayout = mTransverse
End Property

' Решает в выбранной точке траектории и возвращает только подтверждённое
' состояние заданных усилий. Численный отказ не публикует итоги промежуточной ступени.
Friend Function SolveTarget(ByVal loads As CSectionLoadState, ByVal loadPath As String, _
        ByVal spec As CMaterialModelSpec, ByRef field() As Double, ByRef resultants() As Double, _
        Optional ByVal captureFields As Boolean = False, Optional ByRef acceptedState As CSectionStateResult = Nothing, _
        Optional ByRef acceptedTrace As CCoupledLoadTrace = Nothing) As CResultMeta
    Erase field: Erase resultants
    Set acceptedState = Nothing
    Set acceptedTrace = Nothing
    Dim concrete As New CCoupledConcreteResponse, fc As Double, ft As Double
    Dim meta As CResultMeta, response As New CCoupledReinforcedResponse, models As Object
    Set meta = PrepareResponse(spec, concrete, response, fc, ft, models)
    If meta.InternalStatus <> rsUnset Then Set SolveTarget = meta: Exit Function

    Dim solver As CCoupledNewtonSolver, continuation As New CCoupledLoadContinuation, path As CLoadPathVector
    Set path = BuildPath(loads, loadPath)
    Set solver = CreateNewtonSolver(response, fc, ft, models)
    continuation.Initialize solver, path, INITIAL_STEPS, STEP_ATTEMPTS, mMaxIterations, _
        STRESS_TOLERANCE, RELATIVE_TOLERANCE, captureFields
    If continuation.Advance(1#) Then
        continuation.CopyAccepted field, resultants
        Set meta = continuation.ResultMeta
        If response.AcceptedTransverseSteelUtilization > 1# Then
            meta.SetResult rsCheckFailed, rcTransverseStrainExceeded, rkDirectState, _
                SteelLimitComment(response.AcceptedTransverseCriticalPoint) & ". Использование по деформации " & _
                    Format$(response.AcceptedTransverseSteelUtilization, "0.000") & ".", meta.DiagnosticDetails
        ElseIf response.AcceptedSteelUtilization > 1# Then
            meta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkDirectState, _
                SteelLimitComment(response.AcceptedSteelCriticalPoint) & ". Использование по деформации " & _
                    Format$(response.AcceptedSteelUtilization, "0.000") & ".", meta.DiagnosticDetails
        Else
            meta.ResultComment = "Связанное равновесие подтверждено по шести усилиям."
        End If
        If captureFields Then
            Dim target() As Double, warning As String
            path.TargetSix 1#, target
            Set acceptedState = CaptureAcceptedState(sstStrengthState, spec, continuation.AcceptedFields, _
                concrete, response, models, target, resultants, fc, ft, meta, warning)
            If Len(warning) > 0 Then AppendStateWarning meta, warning
        End If
    Else
        Set meta = continuation.ResultMeta
    End If
    Set acceptedTrace = continuation.SavedTrace
    Set SolveTarget = meta
End Function

' Ищет физический предел на независимой ветви, при необходимости управляя
' деформацией. Вектор относится к пользовательской точке приложения.
Friend Function SolveCapacity(ByVal loads As CSectionLoadState, ByVal loadPath As String, _
        ByVal spec As CMaterialModelSpec, ByVal initialLambda As Double, ByVal maximumLambda As Double, _
        ByVal lambdaTolerance As Double, ByVal maxIterations As Long, _
        Optional ByRef acceptedTrace As CCoupledLoadTrace = Nothing) As CCapacityResult
    Dim capacity As New CCapacityResult, meta As CResultMeta, reason As String
    Set SolveCapacity = capacity
    Set acceptedTrace = Nothing
    If initialLambda <= 0# Then reason = "Capacity.InitialLambda"
    If maximumLambda <= 0# Then reason = "Capacity.MaxLambda"
    If lambdaTolerance <= 0# Then reason = "Capacity.ToleranceLambda"
    If maxIterations < 1 Then reason = "Capacity.SolverMaxIterations"
    If Len(reason) > 0 Then
        capacity.InitializeUnsupportedConfiguration "На Config укажите положительное значение " & reason & "."
        Exit Function
    End If

    Dim path As CLoadPathVector
    Set path = BuildPath(loads, loadPath, reason)
    If Len(reason) > 0 Then
        capacity.InitializeSkipped reason, loadPath
        Exit Function
    End If

    Dim concrete As New CCoupledConcreteResponse, response As New CCoupledReinforcedResponse
    Dim models As Object, fc As Double, ft As Double
    Set meta = PrepareResponse(spec, concrete, response, fc, ft, models)
    If meta.InternalStatus <> rsUnset Then
        capacity.InitializeCoupledUnresolved meta, loadPath
        Exit Function
    End If

    Dim solver As CCoupledNewtonSolver, branch As New CCoupledLoadContinuation
    Set solver = CreateNewtonSolver(response, fc, ft, models)
    branch.Initialize solver, path, INITIAL_STEPS, STEP_ATTEMPTS, maxIterations, STRESS_TOLERANCE, RELATIVE_TOLERANCE, True
    branch.TraceSteelLimit response, initialLambda, maximumLambda, lambdaTolerance, preferDeformationOnStall:=True
    Set acceptedTrace = branch.SavedTrace
    Set meta = branch.ResultMeta
    Dim low As Double, high As Double, onOffset As Boolean, ultimate() As Double, criticalPoint As Long
    Dim hasSteel As Boolean, controlled As CCoupledDeformationContinuation
    Dim limitHistory As Collection, actual() As Double, limitTarget() As Double
    hasSteel = branch.CopySteelLimit(low, high, onOffset, criticalPoint)
    If Not hasSteel Then
        Set controlled = branch.ContinueByDeformation(maximumLambda)
        If Not controlled Is Nothing Then
            Set acceptedTrace = controlled.SavedTrace
            Dim forceDetails As String: forceDetails = meta.DiagnosticDetails
            Set meta = controlled.ResultMeta
            If Len(forceDetails) > 0 Then meta.DiagnosticDetails = forceDetails & " | " & meta.DiagnosticDetails
            hasSteel = controlled.CopySteelLimit(low, high, criticalPoint)
            Dim peak As Double, accuracy As Double, designVerified As Boolean
            If controlled.CopyPeak(peak, accuracy, designVerified) Then
                path.TargetSix peak, limitTarget
                ultimate = limitTarget
                loads.MoveResultantsToLoadPoint mCenterX, mCenterY, ultimate
                capacity.InitializeCoupledPeak peak, accuracy, designVerified, ultimate, loadPath, meta.DiagnosticDetails
                controlled.CopyLimitState limitHistory, actual
                SaveCapacityFields capacity, solver, spec, limitHistory, response, concrete, models, limitTarget, actual, fc, ft
                Exit Function
            End If
        End If
    End If
    If hasSteel Then
        reason = SteelLimitComment(criticalPoint)
        Dim transverseLimit As Boolean: transverseLimit = (mOperators.TransverseGroupAt(criticalPoint) > 0)
        If onOffset Then
            capacity.InitializeCoupledBaseFail reason, loadPath, meta.DiagnosticDetails, transverseLimit
        Else
            path.TargetSix low, limitTarget
            ultimate = limitTarget
            loads.MoveResultantsToLoadPoint mCenterX, mCenterY, ultimate
            capacity.InitializeCoupledLimit low, ultimate, loadPath, reason, Nothing, meta.DiagnosticDetails, transverseLimit
            If controlled Is Nothing Then
                branch.CopyLimitState limitHistory, actual
            Else
                controlled.CopyLimitState limitHistory, actual
            End If
            If Not transverseLimit Then criticalPoint = 0
            SaveCapacityFields capacity, solver, spec, limitHistory, response, concrete, models, limitTarget, actual, fc, ft, criticalPoint
        End If
    Else
        low = -1#
        If branch.HasAcceptedVariableState Then low = branch.AcceptedLambda
        If Not controlled Is Nothing Then low = controlled.BestAdmissibleLambda
        capacity.InitializeCoupledUnresolved meta, loadPath, low
    End If
End Function

' Восстанавливает материалы только до проверенной предельной точки.
' При неполных полях сохраняет интегральный результат и предупреждение.
Private Sub SaveCapacityFields(ByVal capacity As CCapacityResult, ByVal solver As CCoupledNewtonSolver, _
        ByVal spec As CMaterialModelSpec, ByVal history As Collection, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal concrete As CCoupledConcreteResponse, ByVal models As Object, ByRef target() As Double, _
        ByRef actual() As Double, ByVal fc As Double, ByVal ft As Double, Optional ByVal governingTransversePoint As Long = 0)
    Dim warning As String, state As CSectionStateResult
    If history Is Nothing Then
        warning = "История предельной точки недоступна для вывода полей."
    ElseIf Not solver.ReplayAcceptedFields(history, history.Count) Then
        warning = "Не удалось восстановить историю материалов в предельной точке."
    Else
        Dim stateMeta As New CResultMeta, capacityMeta As CResultMeta
        Set capacityMeta = capacity.ResultMeta
        stateMeta.SetResult rsSuccess, rcNone, rkCapacity, "Предельное равновесие подтверждено.", capacityMeta.DiagnosticDetails
        Set state = CaptureAcceptedState(sstCapacityState, spec, history, concrete, reinforcement, models, _
            target, actual, fc, ft, stateMeta, warning, governingTransversePoint)
    End If
    capacity.AttachCoupledState state, warning
End Sub

' Формулирует проверяемый механизм. Для поперечной арматуры группа и класс
' стоят в начале строки, чтобы их было видно в однострочной таблице.
Private Function SteelLimitComment(ByVal point As Long) As String
    If mOperators.TransverseGroupAt(point) > 0 Then
        SteelLimitComment = SteelControlDescription(point, True)
    Else
        SteelLimitComment = "Достигнута предельная деформация арматуры. " & SteelControlDescription(point)
    End If
End Function

' Показывает группу или стержень, класс всей ветви и координаты расчётной
' точки в осях сечения относительно центра бетонной сетки, мм. Для
' предельной точки поперечной арматуры добавляет причину после класса.
Private Function SteelControlDescription(ByVal point As Long, Optional ByVal atTransverseLimit As Boolean = False) As String
    Dim group As Long, x As Double, y As Double, weight As Double, rebarIndex As Long
    group = mOperators.TransverseGroupAt(point)
    If group = 0 Then
        rebarIndex = point - mOperators.ConcretePointCount
        SteelControlDescription = "Продольный стержень " & mSection.RebarID(rebarIndex) & ", класс " & mOperators.SteelClassAt(point)
    Else
        Dim path As CGeometryPath: Set path = mTransverse.Path(group)
        SteelControlDescription = path.ID & " (" & mOperators.SteelClassAt(point) & ")"
        If atTransverseLimit Then SteelControlDescription = SteelControlDescription & " — предел деформации"
    End If
    mOperators.Point point, x, y, weight
    SteelControlDescription = SteelControlDescription & ", x = " & Format$(x - mCenterX, "0.00") & _
        " мм, y = " & Format$(y - mCenterY, "0.00") & " мм"
End Function

' Готовит независимую историю расчёта и переиспользует неизменяемые
' геометрические операторы и модели стали того же набора характеристик.
Private Function PrepareResponse(ByVal spec As CMaterialModelSpec, ByVal concrete As CCoupledConcreteResponse, _
        ByVal response As CCoupledReinforcedResponse, ByRef fc As Double, ByRef ft As Double, ByRef models As Object) As CResultMeta
    If mSection.Revision <> mRevision Or mProvider.Revision <> mProviderRevision Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Повторно подготовьте связанный расчёт после изменения сечения или материалов."
    If spec.ConcreteTension <> ctkUseDiagram Then
        Set PrepareResponse = ConfigurationError("Для связанного НДС выберите UseDiagram для растянутого бетона в модели прочности профиля.")
        Exit Function
    End If
    If spec.SteelDiagram <> sdkTwoLine And (mSection.RebarCount > 0 Or mTransverse.Count > 0) Then
        Set PrepareResponse = ConfigurationError("Для связанного НДС выберите TwoLine для стали в модели прочности профиля.")
        Exit Function
    End If

    On Error GoTo InvalidMaterialOrGeometry
    Dim meta As CResultMeta
    Set models = SteelModels(spec)
    PrepareOperators spec, models
    Set meta = mOperators.ResultMeta
    Set PrepareResponse = meta
    If meta.InternalStatus <> rsUnset Then Exit Function
    fc = mProvider.ConcreteCompressionResistanceFromSpec(spec)
    ft = mProvider.ConcreteTensionResistanceFromSpec(spec)
    InitializePointResponse concrete, response, fc, ft, models
    Exit Function

InvalidMaterialOrGeometry:
    Dim number As Long, reason As String, source As String
    number = Err.Number: reason = Err.Description: source = Err.Source
    If number = 5 Or (number >= vbObjectError + 1776 And number <= vbObjectError + 1783) Then
        Set PrepareResponse = ConfigurationError(reason)
    Else
        Err.Raise number, source, reason
    End If
End Function

' Инициализирует одну независимую историю бетона и стали. История использует
' общие операторы геометрии и неизменяемые диаграммы всех классов стали.
Private Sub InitializePointResponse(ByVal concrete As CCoupledConcreteResponse, ByVal response As CCoupledReinforcedResponse, _
        ByVal fc As Double, ByVal ft As Double, ByVal models As Object)
    concrete.Initialize mOperators.PointCount, mEb, mPoisson, mYieldRatio * fc, fc, ft, mPeakStrain, mGt, mGc, mLength
    response.Initialize mOperators, concrete, models
End Sub

' Подключает центральную и две боковые истории. Боковые точки участвуют
' в градиенте напряжений, пределы и выходные данные определяются по центру.
Private Function CreateNewtonSolver(ByVal response As CCoupledReinforcedResponse, ByVal fc As Double, ByVal ft As Double, _
        ByVal models As Object) As CCoupledNewtonSolver
    Dim plusConcrete As New CCoupledConcreteResponse, minusConcrete As New CCoupledConcreteResponse
    Dim plusResponse As New CCoupledReinforcedResponse, minusResponse As New CCoupledReinforcedResponse
    InitializePointResponse plusConcrete, plusResponse, fc, ft, models
    InitializePointResponse minusConcrete, minusResponse, fc, ft, models
    Dim solver As New CCoupledNewtonSolver
    solver.Initialize mOperators, response, mEb, plusResponse, minusResponse
    Set CreateNewtonSolver = solver
End Function

' Копирует точки фактических моделей один раз для всех сочетаний этого набора.
Private Function CalibrationCurves(ByVal spec As CMaterialModelSpec, ByVal concrete As CCoupledConcreteResponse, _
        ByVal steel As Object) As CCoupledMaterialSnapshot
    Dim key As String: key = spec.ValueSetText
    If mDiagramCache.Exists(key) Then
        Set CalibrationCurves = mDiagramCache(key)
    Else
        Dim snapshot As New CCoupledMaterialSnapshot
        snapshot.Initialize concrete, mEb, steel
        mDiagramCache.Add key, snapshot
        Set CalibrationCurves = snapshot
    End If
End Function

' Сохраняет осевое и пространственное поля принятой траектории. Отказ
' отдельной точки вывода сохраняет подтверждённое интегральное равновесие.
Private Function CaptureAcceptedState(ByVal stateType As ESectionStateType, ByVal spec As CMaterialModelSpec, _
        ByVal history As Collection, ByVal concrete As CCoupledConcreteResponse, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal models As Object, ByRef target() As Double, ByRef resultants() As Double, ByVal fc As Double, ByVal ft As Double, _
        ByVal meta As CResultMeta, ByRef warning As String, Optional ByVal governingTransversePoint As Long = 0) As CSectionStateResult
    Dim observer As New CCoupledConcreteResponse, fields As New CCoupledFieldSnapshot
    observer.Initialize mSection.ConcreteCount, mEb, mPoisson, mYieldRatio * fc, fc, ft, mPeakStrain, mGt, mGc, mLength
    fields.Capture mSection, mOperators, history, observer, concrete, reinforcement, warning
    Dim fieldMeta As CResultMeta: Set fieldMeta = meta.Clone
    If Len(warning) > 0 Then AppendStateWarning fieldMeta, warning
    Dim eps0 As Double, kx As Double, ky As Double, twist As Double, state As New CSectionStateResult, field() As Double
    field = history.Item(history.Count)
    mOperators.GeneralizedDeformation field, eps0, kx, ky, twist
    Dim residualArea As Double, residualLength As Double
    mOperators.CopyResultantResidualScales residualArea, residualLength
    state.InitializeFromCoupled stateType, mSection, spec, target, resultants, eps0, kx, ky, twist, fields, fieldMeta, _
        (reinforcement.AcceptedSteelUtilization <= 1#), residualArea, residualLength
    state.SetCoupledMaterialCurves CalibrationCurves(spec, concrete, models)
    SetTransverseControl state, reinforcement, models, field, governingTransversePoint
    Set CaptureAcceptedState = state
End Function

' Считывает напряжение и вычисляет проекцию деформации для одной точки
' принятого поля. При пределе использует точку, определившую арматурный критерий.
Private Sub SetTransverseControl(ByVal state As CSectionStateResult, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByVal models As Object, ByRef field() As Double, ByVal point As Long)
    If point = 0 Then point = reinforcement.AcceptedTransverseCriticalPoint
    If point = 0 Then Exit Sub
    Dim tensor(1 To 6) As Double, direction(1 To 6) As Double, strain As Double, strainLimit As Double, i As Long
    mOperators.PointStrain point, field, tensor
    mOperators.SteelProjection point, direction
    For i = 1 To 6
        strain = strain + tensor(i) * direction(i) ' Тензорная проекция на ось стержня.
    Next i
    Dim steel As CCoupledSteelModel: Set steel = models.Item(mOperators.SteelClassAt(point))
    strainLimit = steel.SignedStrainLimit(strain)
    state.SetCoupledTransverseControl SteelControlDescription(point), _
        steel.StrainUtilization(strain), reinforcement.AcceptedSteelStress(point), strain, strainLimit
End Sub

' Добавляет предупреждение о поле, сохраняя статус физической проверки.
Private Sub AppendStateWarning(ByVal meta As CResultMeta, ByVal warning As String)
    Dim comment As String: comment = Trim$(meta.ResultComment & " " & warning)
    If meta.InternalStatus = rsSuccess Then
        meta.SetResult rsSuccessWithWarning, meta.ResultCode, meta.ResultKind, comment, meta.DiagnosticDetails
    Else
        meta.ResultComment = comment
    End If
End Sub

' Готовит операторы с модулями фактически назначенных моделей. Профили
' с одним набором материалов переиспользуют готовые операторы запуска.
Private Sub PrepareOperators(ByVal spec As CMaterialModelSpec, ByVal models As Object)
    Dim key As String: key = spec.ValueSetText
    mEb = mProvider.ConcreteElasticModulusFromSpec(spec)
    If mOperatorCache.Exists(key) Then
        Set mOperators = mOperatorCache(key)
    Else
        Dim prepared As New CCoupledElasticState
        prepared.PrepareModel mSection, mEb, mPoisson, 1, 1, 5, mTransverse, steelModels:=models, importedGrade:=mSteelClass
        mOperatorCache.Add key, prepared
        Set mOperators = prepared
    End If
End Sub

' Запоминает место задания класса для сообщения при отсутствии его модели.
Private Sub RegisterSteelClass(ByVal grade As String, ByVal inputLocation As String)
    If Len(grade) = 0 Then Err.Raise vbObjectError + 1780, "CCoupledSectionAnalysis", _
        "Укажите класс арматуры в геометрии сечения (" & inputLocation & ")."
    If Not mSteelReferences.Exists(grade) Then mSteelReferences.Add grade, inputLocation
End Sub

' Возвращает только применённые классы для набора характеристик профиля.
' Ошибка строки одного набора не мешает использованию другого набора.
Private Function SteelModels(ByVal spec As CMaterialModelSpec) As Object
    Dim key As String: key = spec.ValueSetText
    If mSteelCache.Exists(key) Then Set SteelModels = mSteelCache(key): Exit Function

    Dim available As Object
    If Not mSupplementarySteel Is Nothing Then
        If mSupplementarySteel.Exists(key) Then Set available = mSupplementarySteel(key)
    End If
    Dim models As Object: Set models = CreateObject("Scripting.Dictionary")
    models.CompareMode = vbTextCompare
    Dim grade As Variant, model As CCoupledSteelModel
    For Each grade In mSteelReferences.Keys
        If StrComp(CStr(grade), mSteelClass, vbTextCompare) = 0 Then
            Set model = PrimarySteelModel(spec)
        Else
            Set model = SupplementarySteelModel(available, CStr(grade), key)
        End If
        models.Add grade, model
    Next grade
    mSteelCache.Add key, models
    Set SteelModels = models
End Function

' Использует штатную двухлинейную диаграмму для явно сопоставленного класса.
Private Function PrimarySteelModel(ByVal spec As CMaterialModelSpec) As CCoupledSteelModel
    Dim es As Double, slope As Double, model As New CCoupledSteelModel
    es = mProvider.SteelElasticModulusFromSpec(spec)
    If mProvider.SteelParameters.Esc <> es Then Err.Raise vbObjectError + 1780, "CCoupledSectionAnalysis", _
        "Для связанного расчёта задайте одинаковый модуль Es при растяжении и сжатии (rngSteelMaterialParameters)."
    slope = mProvider.PlateauSlopeFactor * es
    model.Initialize es, mProvider.SteelTensionResistanceFromSpec(spec), _
        mProvider.SteelCompressionResistanceFromSpec(spec), slope, slope, _
        mProvider.SteelTensionLimitFromSpec(spec), -mProvider.SteelCompressionLimitFromSpec(spec) ' Поставщик возвращает знаковый предел сжатия.
    Set PrimarySteelModel = model
End Function

' Объясняет отсутствие диаграммы с указанием класса, профиля и геометрии.
Private Function SupplementarySteelModel(ByVal available As Object, ByVal grade As String, ByVal key As String) As CCoupledSteelModel
    If Not available Is Nothing Then
        If available.Exists(grade) Then
            If Not IsObject(available(grade)) Then Err.Raise vbObjectError + 1780, "CCoupledSectionAnalysis", CStr(available(grade))
            Set SupplementarySteelModel = available(grade)
            Exit Function
        End If
    End If
    Err.Raise vbObjectError + 1780, "CCoupledSectionAnalysis", _
        "Для класса «" & grade & "» (" & mSteelReferences(grade) & ") заполните строку " & key & " в rngCoupledSteelClasses на Config."
End Function

' Auto и All6 пропорционально наращивают шесть усилий. Явные пути T/Q
' сохраняют остальные усилия в точке приложения при любом множителе.
' Моменты переносятся к тому же центру, который использует интегрирование.
Private Function BuildPath(ByVal loads As CSectionLoadState, ByVal loadPath As String, _
        Optional ByRef noDirectionReason As String = vbNullString) As CLoadPathVector
    Dim descriptor As New CLoadPathDescriptor
    loadPath = descriptor.NormalizeKey(loadPath)
    If Len(loadPath) = 0 Then Err.Raise 5, "CCoupledSectionAnalysis", descriptor.LoadPathChoices
    If descriptor.IsSingleSpatialKey(loadPath) Then
        descriptor.InitializeFromLoadState loadPath, loads
        noDirectionReason = descriptor.MissingSpatialLoadReason
        Set BuildPath = descriptor.SpatialVectorAboutPoint(loads, mCenterX, mCenterY)
        Exit Function
    End If

    Dim target() As Double, result As New CLoadPathVector
    loads.LoadAboutPoint mCenterX, mCenterY, target
    If loadPath = "Auto" Or loadPath = "LambdaAll6" Then
        result.Initialize 0#, target(1), 0#, target(2), 0#, target(3), 0#, target(4), 0#, target(5), 0#, target(6)
    Else
        Dim mxOffset As Double, myOffset As Double
        descriptor.InitializeFromLoadState loadPath, loads
        mxOffset = descriptor.MxOffset - descriptor.NOffset * mCenterY
        myOffset = descriptor.MyOffset - descriptor.NOffset * mCenterX
        result.Initialize descriptor.NOffset, descriptor.NBase, mxOffset, target(2) - mxOffset, _
            myOffset, target(3) - myOffset, target(4), 0#, target(5), 0#, target(6), 0#
    End If
    Set BuildPath = result
End Function

' Возвращает исправляемую пользователем причину отсутствия связанного решения.
Private Function ConfigurationError(ByVal reason As String) As CResultMeta
    Dim meta As New CResultMeta
    meta.SetResult rsInvalidConfiguration, rcInvalidConfiguration, rkDirectState, reason
    Set ConfigurationError = meta
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Ведёт отдельную траекторию от нулевой истории и интегрирует принятые
' напряжения относительно точки приложения нагрузки, в Н и Н·мм.
Friend Sub TestIntegratedState(ByVal loads As CSectionLoadState, ByVal spec As CMaterialModelSpec, _
        ByVal factor As Double, ByRef resultants() As Double, ByRef shear() As Double, ByRef activePoints As Long)
    Dim concrete As New CCoupledConcreteResponse, response As New CCoupledReinforcedResponse
    Dim fc As Double, ft As Double, meta As CResultMeta, models As Object
    Set meta = PrepareResponse(spec, concrete, response, fc, ft, models)
    If meta.InternalStatus <> rsUnset Then Err.Raise 5, "TestIntegratedState", meta.ResultComment

    Dim solver As CCoupledNewtonSolver, continuation As New CCoupledLoadContinuation, path As CLoadPathVector
    Set path = BuildPath(loads, "Auto")
    Set solver = CreateNewtonSolver(response, fc, ft, models)
    continuation.Initialize solver, path, INITIAL_STEPS, STEP_ATTEMPTS, mMaxIterations, STRESS_TOLERANCE, RELATIVE_TOLERANCE
    If Not continuation.Advance(factor) Then
        Set meta = continuation.ResultMeta
        Err.Raise 5, "TestIntegratedState", meta.ResultComment & " " & meta.DiagnosticDetails
    End If
    If continuation.AcceptedLambda <> factor Or response.AcceptedSteelUtilization > 1# Then _
        Err.Raise 5, "TestIntegratedState", "Целевая нагрузка не подтверждена при допустимой деформации арматуры."

    Dim bulkCount As Long: bulkCount = mOperators.ConcretePointCount
    ReDim resultants(1 To 6): ReDim shear(1 To 2, 1 To bulkCount): activePoints = 0
    Dim point As Long, x As Double, y As Double, weight As Double, normal As Double, tx As Double, ty As Double
    Dim stress(1 To 6) As Double, projection(1 To 6) As Double, steel As Double, tension As Double, compression As Double
    For point = 1 To mOperators.PointCount
        mOperators.Point point, x, y, weight
        concrete.CopyAcceptedStress point, stress
        normal = stress(3): tx = stress(5): ty = stress(6)
        If point <= bulkCount Then
            shear(1, point) = tx: shear(2, point) = ty
            concrete.AcceptedPointDamage point, tension, compression
            If tension > 0# Or compression > 0# Then activePoints = activePoints + 1
        Else
            mOperators.SteelProjection point, projection
            steel = response.AcceptedSteelStress(point)
            normal = steel * projection(3) - normal ' Арматура заменяет бетон в этой точке.
            tx = steel * projection(5) - tx: ty = steel * projection(6) - ty
        End If
        x = x - loads.LoadReferenceX: y = y - loads.LoadReferenceY
        resultants(1) = resultants(1) + weight * normal
        resultants(2) = resultants(2) + weight * y * normal
        resultants(3) = resultants(3) + weight * x * normal ' Рабочий знак My соответствует +x·σz.
        resultants(4) = resultants(4) + weight * (x * ty - y * tx)
        resultants(5) = resultants(5) + weight * tx
        resultants(6) = resultants(6) + weight * ty
    Next point
End Sub
