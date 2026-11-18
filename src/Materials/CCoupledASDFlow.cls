VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledASDFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Спектральное пластическое повреждение бетона по подготовленным ветвям ASD
' ==========================================================================
' Математический закон Petracca и соавт. (2023), невязкая изотропная постановка
' ASDConcrete3D с cdf=0 и без фиксированных плоскостей трещин. Рабочий шаг
' дополнительно замыкает историю при фиксированной деформации; это явная
' интеграция проекта, а не единственное обновление внешней библиотеки.
' Таблицы неизменяемы, состояния принадлежат вызывающему материалу.

Private Const ALPHA As Double = 4# / 33#
Private Const GAMMA As Double = 3#
Private Const CLOSURE_TOLERANCE As Double = 0.000000001 ' МПа для номинального и эффективного напряжения и E*координат.
Private Const CLOSURE_ITERATIONS As Long = 1024 ' Повреждённая разгрузка допускает медленное замыкание; допуск напряжений неизменен.
Private mTension As CCoupledConcreteBackbone, mCompression As CCoupledConcreteBackbone
Private mModulus As Double, mPoisson As Double, mShear As Double, mLame As Double
Private mStrengthRatio As Double, mBeta As Double, mReady As Boolean
Private mTangent As CCoupledASDTangent

' Подключает две подготовленные ветви и согласованные упругие параметры один раз.
Friend Sub Initialize(ByVal tension As CCoupledConcreteBackbone, ByVal compression As CCoupledConcreteBackbone, _
        ByVal modulus As Double, ByVal poisson As Double)
    If mReady Then Err.Raise 5, "CCoupledASDFlow", "Подготовленный спектральный закон неизменяем. Создайте новый материал."
    If tension Is Nothing Or compression Is Nothing Then Err.Raise 5, "CCoupledASDFlow", "Подготовьте обе ветви спектрального бетона."
    If Not tension.IsReady Or Not compression.IsReady Or Not (modulus > 0#) Or poisson <= -1# Or poisson >= 0.5 Then _
        Err.Raise 5, "CCoupledASDFlow", "Проверьте готовность ветвей, модуль и коэффициент Пуассона спектрального бетона."
    If compression.PeakStress <= tension.PeakStress Then Err.Raise 5, "CCoupledASDFlow", "Прочность бетона на сжатие должна превышать прочность на растяжение."
    Set mTension = tension: Set mCompression = compression
    mModulus = modulus: mPoisson = poisson
    mShear = modulus / (2# * (1# + poisson)): mLame = 2# * mShear * poisson / (1# - 2# * poisson)
    mStrengthRatio = GeomMax(5#, compression.PeakStress / tension.PeakStress)
    mBeta = (1# - ALPHA) * mStrengthRatio - (1# + ALPHA)
    Set mTangent = New CCoupledASDTangent
    mTangent.Initialize tension, compression, modulus, poisson, ALPHA, GAMMA, mStrengthRatio, CLOSURE_TOLERANCE, CLOSURE_ITERATIONS
    mReady = True
End Sub

' Возвращает модуль подготовленного закона, МПа.
Friend Property Get Modulus() As Double
    Modulus = mModulus
End Property

' Возвращает коэффициент Пуассона тех же подготовленных упругих операторов.
Friend Property Get Poisson() As Double
    Poisson = mPoisson
End Property

' Возвращает явное обозначение фактически применённого алгоритма интеграции.
Friend Property Get LawIdentifier() As String
    LawIdentifier = "ASD-Closed-Cdf0"
End Property

' Возвращает номинальную прочность выбранной ветви, МПа.
Friend Function PeakStress(ByVal compression As Boolean) As Double
    If compression Then PeakStress = mCompression.PeakStress Else PeakStress = mTension.PeakStress
End Function

' Возвращает алгоритмическую касательную той же замкнутой пробы без изменения напряжений или истории.
' False сохраняет возможность разностного вычисления при кратных корнях и негладкой границе.
Friend Function TryTangent(ByRef strain() As Double, ByRef oldStrain() As Double, ByRef oldEffective() As Double, _
        ByVal oldT As Double, ByVal oldC As Double, ByRef stress() As Double, ByRef effective() As Double, _
        ByVal newT As Double, ByVal newC As Double, ByRef tangent() As Double) As Boolean
    If mTangent Is Nothing Then Exit Function
    TryTangent = mTangent.TryBuild(strain, oldStrain, oldEffective, oldT, oldC, stress, effective, newT, newC, tangent)
End Function

' Выполняет замкнутую пробу от переданного принятого состояния без его изменения.
' Отказ означает численную неудачу материала, а не физическое разрушение.
Friend Function TryStep(ByRef strain() As Double, ByRef oldStrain() As Double, ByRef oldEffective() As Double, _
        ByVal oldT As Double, ByVal oldC As Double, ByRef newT As Double, ByRef newC As Double, _
        ByRef effective() As Double, ByRef nominal() As Double, ByRef plastic() As Double, _
        ByRef damageT As Double, ByRef damageC As Double) As Boolean
    On Error GoTo ArithmeticFailed
    Dim priorEffective(1 To 6) As Double, priorNominal(1 To 6) As Double
    Dim priorT As Double, priorC As Double, maximum As Double, i As Long, iteration As Long
    SingleStep strain, oldStrain, oldEffective, oldT, oldC, newT, newC, effective, nominal, plastic, damageT, damageC
    For iteration = 1 To CLOSURE_ITERATIONS
        priorT = newT: priorC = newC
        For i = 1 To 6: priorEffective(i) = effective(i): priorNominal(i) = nominal(i): Next i
        SingleStep strain, strain, priorEffective, priorT, priorC, newT, newC, effective, nominal, plastic, damageT, damageC
        maximum = mModulus * GeomMax(Abs(newT - priorT), Abs(newC - priorC))
        For i = 1 To 6
            maximum = GeomMax(maximum, Abs(effective(i) - priorEffective(i)))
            maximum = GeomMax(maximum, Abs(nominal(i) - priorNominal(i)))
        Next i
        If maximum <= CLOSURE_TOLERANCE Then TryStep = True: Exit Function
    Next iteration
    Exit Function
ArithmeticFailed:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
End Function

' Обновляет один спектральный прогноз; все координаты и напряжения относятся к одной пробе.
Private Sub SingleStep(ByRef strain() As Double, ByRef oldStrain() As Double, ByRef oldEffective() As Double, _
        ByVal oldT As Double, ByVal oldC As Double, ByRef newT As Double, ByRef newC As Double, _
        ByRef effective() As Double, ByRef nominal() As Double, ByRef plastic() As Double, _
        ByRef damageT As Double, ByRef damageC As Double)
    Dim trial(1 To 6) As Double, increment(1 To 6) As Double, volume As Double, i As Long
    For i = 1 To 6: increment(i) = strain(i) - oldStrain(i): Next i
    volume = mLame * (increment(1) + increment(2) + increment(3))
    For i = 1 To 3: trial(i) = oldEffective(i) + 2# * mShear * increment(i) + volume: Next i
    For i = 4 To 6: trial(i) = oldEffective(i) + mShear * increment(i): Next i
    Dim trace As Double, q As Double, high As Double, middle As Double, low As Double
    CoupledStressInvariants trial, trace, q, high, middle, low
    Dim excursionT As Double, excursionC As Double
    If high >= mTension.StressFloor Then _
        excursionT = (ALPHA * trace + q + mBeta * high) / ((1# - ALPHA) * mModulus * mStrengthRatio)
    If low < 0# Then
        Dim a As Double, b As Double, negativeTrace As Double, negativeQ As Double
        a = GeomMin(0#, high): b = GeomMin(0#, middle): negativeTrace = a + b + low
        negativeQ = Sqr(0.5 * ((a - b) ^ 2 + (b - low) ^ 2 + (low - a) ^ 2))
        excursionC = (ALPHA * negativeTrace + negativeQ + GAMMA * a) / ((1# - ALPHA) * mModulus)
    End If
    Dim scaleT As Double, scaleC As Double, nominalT As Double, nominalC As Double
    BranchStep mTension, oldT, excursionT, newT, nominalT, scaleT, damageT
    BranchStep mCompression, oldC, excursionC, newC, nominalC, scaleC, damageC
    ScaleSigns trial, high, middle, low, scaleT, scaleC, effective
    ScaleSigns trial, high, middle, low, nominalT, nominalC, nominal
    volume = mPoisson * (effective(1) + effective(2) + effective(3))
    For i = 1 To 3: plastic(i) = strain(i) - ((1# + mPoisson) * effective(i) - volume) / mModulus: Next i
    For i = 4 To 6: plastic(i) = strain(i) - effective(i) / mShear: Next i
End Sub

' Продвигает необратимую координату и вычисляет номинальное и эффективное масштабирование знаковой части.
Private Sub BranchStep(ByVal branch As CCoupledConcreteBackbone, ByVal previous As Double, ByVal excursion As Double, _
        ByRef current As Double, ByRef nominalScale As Double, ByRef effectiveScale As Double, ByRef damage As Double)
    Dim nominal As Double, effective As Double, priorPlastic As Double, denominator As Double
    branch.Stresses previous, nominal, effective
    priorPlastic = previous - effective / mModulus
    current = GeomMax(previous, excursion + priorPlastic)
    branch.Stresses current, nominal, effective
    denominator = mModulus * (current - priorPlastic)
    nominalScale = nominal / denominator: effectiveScale = effective / denominator
    damage = 1# - nominal / effective
End Sub

' Собирает знаковые части без собственных векторов; проектор отдельного знака допускает кратность двух остальных значений.
Private Sub ScaleSigns(ByRef source() As Double, ByVal high As Double, ByVal middle As Double, ByVal low As Double, _
        ByVal positiveScale As Double, ByVal negativeScale As Double, ByRef result() As Double)
    Dim common As Double, i As Long
    If low >= 0# Or high <= 0# Or positiveScale = negativeScale Then
        If high <= 0# Then common = negativeScale Else common = positiveScale
        For i = 1 To 6: result(i) = common * source(i): Next i
        Exit Sub
    End If
    Dim isolated As Double, paired As Double, factor As Double
    If middle >= 0# Then
        isolated = low: paired = high: common = positiveScale: factor = negativeScale - positiveScale
    Else
        isolated = high: paired = low: common = negativeScale: factor = positiveScale - negativeScale
    End If
    factor = factor * (isolated / (isolated - middle)) / (isolated - paired)
    Dim x As Double, y As Double, z As Double, h As Double, g As Double, f As Double, sum As Double
    x = source(1): y = source(2): z = source(3): h = source(4): g = source(5): f = source(6): sum = middle + paired
    result(1) = common * x + factor * ((x - middle) * (x - paired) + h * h + g * g)
    result(2) = common * y + factor * ((y - middle) * (y - paired) + h * h + f * f)
    result(3) = common * z + factor * ((z - middle) * (z - paired) + g * g + f * f)
    result(4) = common * h + factor * (h * (x + y - sum) + g * f)
    result(5) = common * g + factor * (g * (x + z - sum) + h * f)
    result(6) = common * f + factor * (f * (y + z - sum) + h * g)
End Sub

' Копирует фактически заданную номинальную диаграмму без преобразования её к другой модели.
Friend Sub CopyCurve(ByVal compression As Boolean, ByRef strain() As Double, ByRef stress() As Double)
    If compression Then mCompression.CopyCurve -1#, strain, stress Else mTension.CopyCurve 1#, strain, stress
End Sub
