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
Private mEb As Double, mEs As Double, mCenterX As Double, mCenterY As Double

' Принимает калибровку в мм, МПа и Н/мм, а размеры деформаций в долях единицы.
' Геометрия и диаграммы остаются неизменными между строками текущего пакета.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal transverse As CTransverseRebarLayout, _
        ByVal provider As CMaterialModelProvider, ByVal poisson As Double, ByVal gt As Double, ByVal gc As Double, _
        ByVal localizationLength As Double, ByVal yieldRatio As Double, ByVal peakStrain As Double, _
        ByVal steelClass As String, ByVal maxIterations As Long)
    If section Is Nothing Or transverse Is Nothing Or provider Is Nothing Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Передайте сечение, поперечное армирование и параметры материалов."
    If Not transverse.IsSpecified Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "В блоке Transverse задайте поперечную арматуру или выберите «Нет»."
    If poisson < 0# Or poisson >= 0.5 Or gt <= 0# Or gc <= 0# Or localizationLength <= 0# Or _
            yieldRatio <= 0# Or yieldRatio >= 1# Or peakStrain <= 0# Or maxIterations < 1 Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Проверьте коэффициент Пуассона, энергии, характерную длину и параметры пика бетона."
    If Len(Trim$(steelClass)) = 0 And (section.RebarCount > 0 Or transverse.Count > 0) Then _
        Err.Raise 5, "CCoupledSectionAnalysis", "Укажите класс стали, которому соответствуют сопротивления и деформации на Config."

    Set mSection = section: Set mTransverse = transverse: Set mProvider = provider
    mRevision = section.Revision: mProviderRevision = provider.Revision
    mPoisson = poisson: mGt = gt: mGc = gc: mLength = localizationLength
    mYieldRatio = yieldRatio: mPeakStrain = peakStrain
    mSteelClass = Trim$(steelClass): mMaxIterations = maxIterations
    Set mOperators = Nothing
    Dim area As Double
    section.ConcreteMeshCentroid mCenterX, mCenterY, area
End Sub

' Решает в выбранной точке траектории и возвращает только подтверждённое
' состояние заданных усилий. Численный отказ не публикует итоги промежуточной ступени.
Friend Function SolveTarget(ByVal loads As CSectionLoadState, ByVal loadPath As String, _
        ByVal spec As CMaterialModelSpec, ByRef field() As Double, ByRef resultants() As Double) As CResultMeta
    Erase field: Erase resultants
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
    PrepareOperators spec
    Set meta = mOperators.ResultMeta
    If meta.InternalStatus <> rsUnset Then Set SolveTarget = meta: Exit Function
    Dim concrete As New CCoupledConcreteResponse, fc As Double, ft As Double
    fc = mProvider.ConcreteCompressionResistanceFromSpec(spec)
    ft = mProvider.ConcreteTensionResistanceFromSpec(spec)
    concrete.Initialize mOperators.PointCount, mEb, mPoisson, mYieldRatio * fc, fc, ft, mPeakStrain, mGt, mGc, mLength
    Dim response As New CCoupledReinforcedResponse
    response.Initialize mOperators, concrete, SteelModels(spec)
    On Error GoTo 0

    Dim solver As New CCoupledNewtonSolver, continuation As New CCoupledLoadContinuation
    solver.Initialize mOperators, response, mEb
    continuation.Initialize solver, BuildPath(loads, loadPath), INITIAL_STEPS, STEP_ATTEMPTS, mMaxIterations, _
        STRESS_TOLERANCE, RELATIVE_TOLERANCE
    If continuation.Advance(1#) Then
        continuation.CopyAccepted field, resultants
        Set meta = continuation.ResultMeta
        If response.AcceptedSteelUtilization > 1# Then
            meta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkDirectState, _
                "Связанное равновесие найдено с превышением предельной деформации арматуры.", meta.DiagnosticDetails
        Else
            meta.ResultComment = "Связанное равновесие подтверждено по шести усилиям."
        End If
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

' Подготавливает общий базис и интегрирование. Модули входят в ключ
' операторов, поэтому смена набора профиля требует точного совпадения.
Private Sub PrepareOperators(ByVal spec As CMaterialModelSpec)
    Dim eb As Double, es As Double
    eb = mProvider.ConcreteElasticModulusFromSpec(spec)
    es = mProvider.SteelElasticModulusFromSpec(spec)
    If Not mOperators Is Nothing And eb = mEb And es = mEs Then Exit Sub
    Dim prepared As New CCoupledElasticState
    prepared.PrepareModel mSection, eb, mPoisson, 1, 1, 5, mTransverse, es
    Set mOperators = prepared: mEb = eb: mEs = es
End Sub

' Сопоставляет единственной диаграмме Config только явно заданный класс стали.
' Другие классы не получают сопротивления от этой диаграммы.
Private Function SteelModels(ByVal spec As CMaterialModelSpec) As Object
    Dim models As Object: Set models = CreateObject("Scripting.Dictionary")
    models.CompareMode = vbTextCompare
    If mOperators.PointCount > mOperators.ConcretePointCount Then
        Dim model As New CCoupledSteelModel, slope As Double
        slope = mProvider.PlateauSlopeFactor * mEs
        model.Initialize mEs, mProvider.SteelTensionResistanceFromSpec(spec), _
            mProvider.SteelCompressionResistanceFromSpec(spec), slope, slope, _
            mProvider.SteelTensionLimitFromSpec(spec), -mProvider.SteelCompressionLimitFromSpec(spec) ' Поставщик возвращает сжатие со знаком минус, закон стали принимает модуль деформации.
        models.Add mSteelClass, model
    End If
    Set SteelModels = models
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
