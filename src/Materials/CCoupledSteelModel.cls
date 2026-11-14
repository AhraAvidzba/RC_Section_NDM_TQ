VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledSteelModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Одноосная сталь с кинематическим упрочнением
' ==========================================================================
' Двухлинейная огибающая рассчитана из заданных модулей и сопротивлений.
' Эффективное напряжение ограничено интервалом [-Rsc, Rs] относительно
' обратного напряжения. Вызов не изменяет переданную историю. Ограничения
' деформаций задаются отдельно и применяются только к принятому состоянию.

Private mElasticModulus As Double
Private mTensionYield As Double, mCompressionYield As Double
Private mTensionHardening As Double, mCompressionHardening As Double
Private mTensionLimit As Double, mCompressionLimit As Double

' Задаёт двухлинейную диаграмму с одинаковым модулем разгрузки для обоих знаков.
' Модули и напряжения в МПа. Предельные деформации передаются положительными.
Friend Sub Initialize(ByVal elasticModulus As Double, ByVal tensionYield As Double, _
        ByVal compressionYield As Double, ByVal tensionSlope As Double, ByVal compressionSlope As Double, _
        ByVal tensionLimit As Double, ByVal compressionLimit As Double)
    If elasticModulus <= 0# Then InputError "Для арматуры задайте положительный модуль упругости."
    If tensionYield <= 0# Or compressionYield <= 0# Then InputError "Сопротивления арматуры должны быть положительными."
    If tensionSlope < 0# Or compressionSlope < 0# Or tensionSlope >= elasticModulus Or compressionSlope >= elasticModulus Then
        InputError "Касательный модуль упрочнения арматуры должен быть в диапазоне от нуля до Es."
    End If
    If tensionLimit <= tensionYield / elasticModulus Or compressionLimit <= compressionYield / elasticModulus Then
        InputError "Предельная деформация арматуры должна превышать деформацию начала текучести."
    End If

    mElasticModulus = elasticModulus
    mTensionYield = tensionYield: mCompressionYield = compressionYield
    mTensionHardening = tensionSlope / (1# - tensionSlope / elasticModulus)
    mCompressionHardening = compressionSlope / (1# - compressionSlope / elasticModulus)
    mTensionLimit = tensionLimit: mCompressionLimit = compressionLimit
End Sub

' Возвращает модуль материала для подготовки упругих операторов класса.
' До инициализации модуль равен нулю.
Friend Property Get ElasticModulus() As Double
    ElasticModulus = mElasticModulus
End Property

' Обновляет пробную точку заданной деформации. Подтверждённые пластическая
' деформация и обратное напряжение передаются по значению для rollback.
Friend Sub EvaluateTrial(ByVal strain As Double, ByVal committedPlastic As Double, ByVal committedBackstress As Double, _
        ByRef stress As Double, ByRef tangent As Double, ByRef trialPlastic As Double, ByRef trialBackstress As Double)
    If mElasticModulus = 0# Then InputError "Модель арматуры не подготовлена."
    stress = mElasticModulus * (strain - committedPlastic)
    tangent = mElasticModulus
    trialPlastic = committedPlastic
    trialBackstress = committedBackstress
    Dim relative As Double, signedExcess As Double, hardening As Double
    relative = stress - committedBackstress
    If relative > mTensionYield Then
        signedExcess = relative - mTensionYield
        hardening = mTensionHardening
    ElseIf relative < -mCompressionYield Then
        signedExcess = relative + mCompressionYield
        hardening = mCompressionHardening
    Else
        Exit Sub
    End If
    Dim increment As Double
    increment = signedExcess / (mElasticModulus + hardening)
    trialPlastic = committedPlastic + increment
    trialBackstress = committedBackstress + hardening * increment
    stress = stress - mElasticModulus * increment
    tangent = hardening / (1# + hardening / mElasticModulus)
End Sub

' Возвращает отношение деформации к допустимому пределу своего знака.
Friend Function StrainUtilization(ByVal strain As Double) As Double
    If mElasticModulus = 0# Then InputError "Модель арматуры не подготовлена."
    If strain >= 0# Then
        StrainUtilization = strain / mTensionLimit
    Else
        StrainUtilization = -strain / mCompressionLimit
    End If
End Function

' Сообщает проблему исходных характеристик арматуры с единым источником.
Private Sub InputError(ByVal reason As String)
    Err.Raise vbObjectError + 1780, "CCoupledSteelModel", reason
End Sub
