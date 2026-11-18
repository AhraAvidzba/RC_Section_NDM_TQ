VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledConcreteBackbone"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемая подготовленная ветвь спектрального бетона
' ==========================================================================
' Координата ветви безразмерная, номинальное и эффективное напряжения в МПа.
' Для обоих знаков передаются положительные ординаты. Подготовка и калибровка
' таблиц выполняются до ядра; этот класс не меняет их физические параметры.
' Табличный хвост и численный нижний уровень не являются критерием разрушения.

Private mStrain() As Double, mNominal() As Double, mEffective() As Double
Private mCount As Long, mFloor As Double, mSearchTolerance As Double, mPeak As Double
Private mReady As Boolean

' Копирует и проверяет полный набор точек один раз, не читая Excel или файлы.
' Нулевая первая точка обязательна; эффективное напряжение не ниже номинального.
Friend Sub Initialize(ByRef strain() As Double, ByRef nominal() As Double, ByRef effective() As Double, _
        ByVal pointCount As Long, ByVal stressFloor As Double)
    If mReady Then Err.Raise 5, "CCoupledConcreteBackbone", "Подготовленная ветвь бетона неизменяема. Создайте новую ветвь."
    If pointCount < 2 Or Not (stressFloor > 0#) Then InvalidInput "Нужны две или более точки и положительный нижний уровень напряжения."
    ValidateArray strain, pointCount: ValidateArray nominal, pointCount: ValidateArray effective, pointCount
    If strain(1) <> 0# Or nominal(1) <> 0# Or effective(1) <> 0# Then InvalidInput "Табличная ветвь бетона должна начинаться в нуле."
    Dim i As Long, spacing As Double, smallest As Double, peak As Double
    For i = 1 To pointCount
        If Not (strain(i) >= 0#) Or Not (nominal(i) >= 0#) Or Not (effective(i) >= nominal(i)) Then _
            InvalidInput "Проверьте неотрицательные координаты ветви и эффективные напряжения."
        If nominal(i) > peak Then peak = nominal(i)
        If i > 1 Then
            spacing = strain(i) - strain(i - 1)
            If spacing <= 0# Then InvalidInput "Координаты ветви бетона должны строго возрастать."
            If smallest = 0# Or spacing < smallest Then smallest = spacing
        End If
    Next i
    If peak <= stressFloor Then InvalidInput "Прочность табличной ветви должна превышать её численный нижний уровень."
    ReDim mStrain(1 To pointCount): ReDim mNominal(1 To pointCount): ReDim mEffective(1 To pointCount)
    For i = 1 To pointCount
        mStrain(i) = strain(i): mNominal(i) = nominal(i): mEffective(i) = effective(i)
    Next i
    mCount = pointCount: mFloor = stressFloor: mPeak = peak
    mSearchTolerance = smallest * 0.000001 ' Допуск поиска относится только к выбору соседнего табличного отрезка.
    mReady = True
End Sub

' Проверяет только используемые индексы; дополнительные элементы массива не читаются.
Private Sub ValidateArray(ByRef values() As Double, ByVal count As Long)
    On Error GoTo InvalidArray
    If LBound(values) <= 1 And UBound(values) >= count Then Exit Sub
InvalidArray:
    On Error GoTo 0
    InvalidInput "Массив ветви бетона не содержит все необходимые точки."
End Sub

' Возвращает признак завершённой неизменяемой подготовки.
Friend Property Get IsReady() As Boolean
    IsReady = mReady
End Property

' Возвращает максимальное фактическое номинальное напряжение ветви, МПа.
Friend Property Get PeakStress() As Double
    PeakStress = mPeak
End Property

' Возвращает подготовленный численный уровень, не изменяющий физическую прочность.
Friend Property Get StressFloor() As Double
    StressFloor = mFloor
End Property

' Интерполирует обе ординаты одним поиском. За хвостом продолжает только неотрицательный наклон.
Friend Sub Stresses(ByVal coordinate As Double, ByRef nominal As Double, ByRef effective As Double)
    Dim left As Long, right As Long, middle As Long, fraction As Double, width As Double
    left = 2: right = mCount
    Do While left < right
        middle = (left + right) \ 2
        If coordinate <= mStrain(middle) + mSearchTolerance Then right = middle Else left = middle + 1
    Loop
    width = mStrain(right) - mStrain(right - 1)
    If coordinate <= mStrain(right) + mSearchTolerance Then
        fraction = (coordinate - mStrain(right - 1)) / width
        nominal = mNominal(right - 1) + fraction * (mNominal(right) - mNominal(right - 1))
        effective = mEffective(right - 1) + fraction * (mEffective(right) - mEffective(right - 1))
    Else
        fraction = (coordinate - mStrain(right)) / width
        nominal = mNominal(right) + fraction * GeomMax(0#, mNominal(right) - mNominal(right - 1))
        effective = mEffective(right) + fraction * GeomMax(0#, mEffective(right) - mEffective(right - 1))
    End If
    nominal = GeomMax(mFloor, nominal): effective = GeomMax(mFloor, effective)
End Sub

' Копирует действительные узлы номинальной диаграммы в требуемом знаке.
Friend Sub CopyCurve(ByVal direction As Double, ByRef strain() As Double, ByRef stress() As Double)
    ReDim strain(1 To mCount): ReDim stress(1 To mCount)
    Dim i As Long
    For i = 1 To mCount
        strain(i) = direction * mStrain(i): stress(i) = direction * mNominal(i)
    Next i
End Sub

' Передаёт понятную причину дефектных подготовленных данных вызывающему слою.
Private Sub InvalidInput(ByVal reason As String)
    Err.Raise 5, "CCoupledConcreteBackbone", reason
End Sub
