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
    Set mSteelReferences = CreateObject("Scripting.Dictionary"): mSteelReferences.CompareMode = vbTextCompare
    Dim i As Long
    For i = 1 To section.RebarCount
        RegisterSteelClass section.RebarSteelClass(i), "rngNDMSectionGeometry, стержень " & section.RebarID(i)
    Next i
    For i = 1 To transverse.Count
        RegisterSteelClass transverse.SteelClass(i), transverse.SourceReference(i)
    Next i
    Set mOperators = Nothing
    Dim area As Double
    section.ConcreteMeshCentroid mCenterX, mCenterY, area
End Sub

' Решает в выбранной точке траектории и возвращает только подтверждённое
' состояние заданных усилий. Численный отказ не публикует итоги промежуточной ступени.
Friend Function SolveTarget(ByVal loads As CSectionLoadState, ByVal loadPath As String, _
        ByVal spec As CMaterialModelSpec, ByRef field() As Double, ByRef resultants() As Double, _
        Optional ByVal captureFields As Boolean = False, Optional ByRef acceptedState As CSectionStateResult = Nothing) As CResultMeta
    Erase field: Erase resultants
    Set acceptedState = Nothing
    Dim meta As New CResultMeta
    If mSection.Revision <> mRevision Or mProvider.Revision <> mProviderRevision Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Повторно подготовьте связанный расчёт после изменения сечения или материалов."

    If spec.ConcreteTension <> ctkUseDiagram Then
        Set SolveTarget = ConfigurationError("Для связанного НДС выберите UseDiagram для растянутого бетона в модели прочности профиля.")
        Exit Function
    End If
    If spec.SteelDiagram <> sdkTwoLine And (mSection.RebarCount > 0 Or mTransverse.Count > 0) Then
        Set SolveTarget = ConfigurationError("Для связанного НДС выберите TwoLine для стали в модели прочности профиля.")
        Exit Function
    End If

    On Error GoTo InvalidMaterialOrGeometry
    Dim models As Object: Set models = SteelModels(spec)
    PrepareOperators spec, models
    Set meta = mOperators.ResultMeta
    If meta.InternalStatus <> rsUnset Then Set SolveTarget = meta: Exit Function
    Dim concrete As New CCoupledConcreteResponse, fc As Double, ft As Double
    fc = mProvider.ConcreteCompressionResistanceFromSpec(spec)
    ft = mProvider.ConcreteTensionResistanceFromSpec(spec)
    concrete.Initialize mOperators.PointCount, mEb, mPoisson, mYieldRatio * fc, fc, ft, mPeakStrain, mGt, mGc, mLength
    Dim response As New CCoupledReinforcedResponse
    response.Initialize mOperators, concrete, models
    On Error GoTo 0

    Dim solver As New CCoupledNewtonSolver, continuation As New CCoupledLoadContinuation
    solver.Initialize mOperators, response, mEb
    continuation.Initialize solver, BuildPath(loads, loadPath), INITIAL_STEPS, STEP_ATTEMPTS, mMaxIterations, _
        STRESS_TOLERANCE, RELATIVE_TOLERANCE, captureFields
    If continuation.Advance(1#) Then
        continuation.CopyAccepted field, resultants
        Set meta = continuation.ResultMeta
        If response.AcceptedSteelUtilization > 1# Then
            meta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkDirectState, _
                "Связанное равновесие найдено с превышением предельной деформации арматуры.", meta.DiagnosticDetails
        Else
            meta.ResultComment = "Связанное равновесие подтверждено по шести усилиям."
        End If
        If captureFields Then Set acceptedState = CaptureAcceptedState(loads, spec, continuation, response, field, resultants, fc, ft, meta)
    Else
        Set meta = continuation.ResultMeta
    End If
    Set SolveTarget = meta
    Exit Function

InvalidMaterialOrGeometry:
    Dim number As Long, reason As String, source As String
    number = Err.Number: reason = Err.Description: source = Err.Source
    If number = 5 Or (number >= vbObjectError + 1776 And number <= vbObjectError + 1783) Then
        Set SolveTarget = ConfigurationError(reason)
    Else
        Err.Raise number, source, reason
    End If
End Function

' Сохраняет осевое и пространственное поля принятой траектории. Отказ
' отдельной точки вывода сохраняет подтверждённое интегральное равновесие.
Private Function CaptureAcceptedState(ByVal loads As CSectionLoadState, ByVal spec As CMaterialModelSpec, _
        ByVal continuation As CCoupledLoadContinuation, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef field() As Double, ByRef resultants() As Double, ByVal fc As Double, ByVal ft As Double, _
        ByVal meta As CResultMeta) As CSectionStateResult
    Dim observer As New CCoupledConcreteResponse, fields As New CCoupledFieldSnapshot, reason As String
    observer.Initialize mSection.ConcreteCount, mEb, mPoisson, mYieldRatio * fc, fc, ft, mPeakStrain, mGt, mGc, mLength
    If Not fields.Capture(mSection, mOperators, continuation, observer, reinforcement, reason) Then
        If meta.InternalStatus = rsSuccess Then
            meta.SetResult rsSuccessWithWarning, rcNone, rkDirectState, meta.ResultComment & " " & reason, meta.DiagnosticDetails
        Else
            meta.ResultComment = meta.ResultComment & " " & reason
        End If
        Exit Function
    End If
    Dim eps0 As Double, kx As Double, ky As Double, twist As Double, state As New CSectionStateResult
    mOperators.GeneralizedDeformation field, eps0, kx, ky, twist
    state.InitializeFromCoupled mSection, spec, loads, resultants, eps0, kx, ky, twist, fields, meta
    Set CaptureAcceptedState = state
End Function

' Готовит операторы с модулями фактически назначенных моделей. Профили
' с одним набором материалов переиспользуют готовые операторы запуска.
Private Sub PrepareOperators(ByVal spec As CMaterialModelSpec, ByVal models As Object)
    Dim key As String: key = spec.ValueSetText
    mEb = mProvider.ConcreteElasticModulusFromSpec(spec)
    If mOperatorCache.Exists(key) Then
        Set mOperators = mOperatorCache(key)
    Else
        Dim prepared As New CCoupledElasticState
        prepared.PrepareModel mSection, mEb, mPoisson, 1, 1, 5, mTransverse, steelModels:=models
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

' Для Auto пропорционально наращивает шесть усилий. Явные траектории
' сохраняют правило N/Mx/My, а T, Qx и Qy относятся к преднагружению.
' Все моменты переносит к тому же центру, который использует интегрирование.
Private Function BuildPath(ByVal loads As CSectionLoadState, ByVal loadPath As String) As CLoadPathVector
    Dim target() As Double, result As New CLoadPathVector
    loads.LoadAboutPoint mCenterX, mCenterY, target
    If loadPath = "Auto" Then
        result.Initialize 0#, target(1), 0#, target(2), 0#, target(3), 0#, target(4), 0#, target(5), 0#, target(6)
    Else
        Dim descriptor As New CLoadPathDescriptor, mxOffset As Double, myOffset As Double
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
