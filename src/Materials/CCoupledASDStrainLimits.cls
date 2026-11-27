VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledASDStrainLimits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Явные конечные ограничения эквивалентных полных деформаций ASD
' ==========================================================================
' Положительные пределы задаёт вызывающий слой, без неявного переноса из СП.
' Неизменяемый объект проверяет материальную точку и локализует пересечение
' на заданном тензорном отрезке. Это не поиск равновесия или предела сечения.

Private mCompression As Double, mTension As Double
Private mTensionEnabled As Boolean, mReady As Boolean

' Принимает конечные положительные пределы в долях единицы.
' Значение отключённого растянутого ограничения не читается и не проверяется.
Friend Sub Initialize(ByVal compressionLimit As Double, ByVal tensionLimit As Double, ByVal tensionEnabled As Boolean)
    If mReady Then InvalidInput "Ограничения ASD неизменяемы. Создайте новый набор."
    If Not (compressionLimit > 0#) Or compressionLimit > 1E+100 Then _
        InvalidInput "Задайте положительный конечный предел эквивалентного сжатия ASD."
    If tensionEnabled Then
        If Not (tensionLimit > 0#) Or tensionLimit > 1E+100 Then _
            InvalidInput "Задайте положительный конечный предел эквивалентного растяжения ASD."
        mTension = tensionLimit
    End If
    mCompression = compressionLimit: mTensionEnabled = tensionEnabled: mReady = True
End Sub

' Проверяет принятую точку, не изменяя историю. False означает неприменимость к её закону.
Friend Function TryEvaluateAccepted(ByVal concrete As CCoupledConcreteResponse, ByVal point As Long, _
        ByRef utilization As Double, ByRef criticalStrain As Double, ByRef compression As Boolean) As Boolean
    ValidateResponse concrete
    utilization = 0#: criticalStrain = 0#: compression = True
    Dim tension As Double, compressive As Double
    If Not concrete.TryReadAcceptedASDStrains(point, tension, compressive) Then Exit Function
    EvaluateMeasures tension, compressive, utilization, criticalStrain, compression
    TryEvaluateAccepted = True
End Function

' Проверяет кандидата от принятой истории; равенство единице относится к допустимой границе.
Friend Function TryEvaluateTrial(ByVal concrete As CCoupledConcreteResponse, ByVal point As Long, _
        ByRef strain() As Double, ByRef utilization As Double, ByRef criticalStrain As Double, _
        ByRef compression As Boolean) As Boolean
    ValidateResponse concrete
    utilization = 0#: criticalStrain = 0#: compression = True
    Dim tension As Double, compressive As Double
    If Not concrete.TryReadASDStrains(point, strain, tension, compressive) Then Exit Function
    EvaluateMeasures tension, compressive, utilization, criticalStrain, compression
    TryEvaluateTrial = True
End Function

' Локализует обнаруженное пересечение на линейном отрезке полных тензоров от неизменной прошлой истории.
' Возвращает параметры отрезка: нижний допустим, верхний достигает или превышает границу.
' Отрезок не считается равновесным путём. Первое разрушение всего сечения этим не подтверждается.
Friend Function TryIntersectTrialSegment(ByVal concrete As CCoupledConcreteResponse, ByVal point As Long, _
        ByRef lowerStrain() As Double, ByRef upperStrain() As Double, ByVal parameterTolerance As Double, _
        ByRef lower As Double, ByRef upper As Double) As Boolean
    ValidateResponse concrete
    lower = 0#: upper = 0#
    If Not (parameterTolerance > 0# And parameterTolerance < 1#) Then _
        InvalidInput "Точность параметра отрезка ASD должна быть больше нуля и меньше единицы."
    Dim ratio As Double, measure As Double, compression As Boolean
    If Not TryEvaluateTrial(concrete, point, lowerStrain, ratio, measure, compression) Then Exit Function
    If ratio > 1# Then Exit Function
    If ratio = 1# Then TryIntersectTrialSegment = True: Exit Function
    If Not TryEvaluateTrial(concrete, point, upperStrain, ratio, measure, compression) Then Exit Function
    If ratio < 1# Then Exit Function
    upper = 1#
    If ratio = 1# Then lower = 1#: TryIntersectTrialSegment = True: Exit Function
    Dim current(1 To 6) As Double, middle As Double, iteration As Long, component As Long
    For iteration = 1 To 64
        If upper - lower <= parameterTolerance Then TryIntersectTrialSegment = True: Exit Function
        middle = lower + (upper - lower) / 2#
        If middle = lower Or middle = upper Then Exit Function ' Машинная граница точности не выдаётся за достигнутую пользовательскую точность.
        For component = 1 To 6
            current(component) = lowerStrain(component) + middle * (upperStrain(component) - lowerStrain(component))
        Next component
        If Not TryEvaluateTrial(concrete, point, current, ratio, measure, compression) Then Exit Function
        If ratio <= 1# Then lower = middle Else upper = middle
    Next iteration
    TryIntersectTrialSegment = (upper - lower <= parameterTolerance)
End Function

' Выбирает наибольшее использование; при равенстве сохраняет сжатие для детерминированного отчёта.
Private Sub EvaluateMeasures(ByVal tension As Double, ByVal compressive As Double, ByRef utilization As Double, _
        ByRef criticalStrain As Double, ByRef compression As Boolean)
    utilization = compressive / mCompression: criticalStrain = compressive: compression = True
    If mTensionEnabled Then
        Dim ratio As Double: ratio = tension / mTension
        If ratio > utilization Then utilization = ratio: criticalStrain = tension: compression = False
    End If
End Sub

' Отделяет ошибку неподготовленного контекста от неприменимости данного закона материала.
Private Sub ValidateResponse(ByVal concrete As CCoupledConcreteResponse)
    If Not mReady Then InvalidInput "Сначала задайте явные пределы эквивалентных деформаций ASD."
    If concrete Is Nothing Then InvalidInput "Передайте подготовленный отклик бетона для проверки ограничений ASD."
End Sub

' Передаёт понятную ошибку ввода вызывающему слою без изменения статусов расчётного состояния.
Private Sub InvalidInput(ByVal reason As String)
    Err.Raise 5, "CCoupledASDStrainLimits", reason
End Sub
