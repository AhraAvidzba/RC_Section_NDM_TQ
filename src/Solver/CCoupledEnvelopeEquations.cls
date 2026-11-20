VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledEnvelopeEquations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Монотонная координата огибающих в общей системе равновесия
' ==========================================================================
' Координата H, МПа, усредняет потреблённые энергии одноосных огибающих CDP
' по бетонной квадратуре трёх станций. Это численная мера продолжения,
' не полная термодинамическая диссипация и не критерий разрушения.
' Геометрические веса готовятся один раз. Изменяемые состояния не кэшируются.
Implements ICoupledEquationSystem

Private mOwner As CCoupledNewtonSolver, mSection As CCoupledElasticState, mRaw As ICoupledEquationSystem
Private mCentral As CCoupledReinforcedResponse, mMinus As CCoupledReinforcedResponse, mPlus As CCoupledReinforcedResponse
Private mArea() As Double, mTotalArea As Double, mHalfStep As Double, mRowScale As Double
Private mContext As Long, mRevision As Long, mTarget As Double
Private mAllowJacobian As Boolean

' Закрепляет подготовленные операторы, три отклика и веса; истории материалов остаются у владельцев отклика.
Friend Sub Initialize(ByVal owner As CCoupledNewtonSolver, ByVal section As CCoupledElasticState, _
        ByVal central As CCoupledReinforcedResponse, ByVal minus As CCoupledReinforcedResponse, _
        ByVal plus As CCoupledReinforcedResponse, ByVal halfStep As Double, ByVal rowScale As Double)
    If Not mSection Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Подготовленная материальная координата неизменяема."
    If section Is Nothing Or central Is Nothing Or minus Is Nothing Or plus Is Nothing Then _
        Err.Raise 5, "CCoupledEnvelopeEquations", "Для материальной координаты нужны подготовленное сечение и три отклика."
    If halfStep <= 0# Or rowScale <= 0# Then Err.Raise 5, "CCoupledEnvelopeEquations", "Масштабы материальной координаты должны быть положительными."
    If section.ConcretePointCount < 1 Then Err.Raise 5, "CCoupledEnvelopeEquations", "В сечении нет бетонной квадратуры для материальной координаты."
    Set mOwner = owner: Set mRaw = owner: Set mSection = section
    Set mCentral = central: Set mMinus = minus: Set mPlus = plus
    mHalfStep = halfStep: mRowScale = rowScale: mTotalArea = 0#
    ReDim mArea(1 To section.ConcretePointCount)
    Dim point As Long, x As Double, y As Double
    For point = 1 To UBound(mArea)
        section.Point point, x, y, mArea(point): mTotalArea = mTotalArea + mArea(point)
    Next point
    If mTotalArea <= 0# Then Err.Raise 5, "CCoupledEnvelopeEquations", "Суммарная площадь бетонной квадратуры должна быть положительной."
End Sub

' Привязывает один уровень H к точной личности читающих уравнений и принятой ревизии.
Friend Sub OpenTarget(ByVal target As Double, Optional ByVal allowJacobian As Boolean = False)
    If mOwner Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    If mContext <> 0 Then Err.Raise 5, "CCoupledEnvelopeEquations", "Для другого уровня или контекста создайте новую материальную координату."
    If target < 0# Or mOwner.EquationContext = 0 Then Err.Raise 5, "CCoupledEnvelopeEquations", "Нет действующего контекста или задан отрицательный уровень H."
    mContext = mOwner.EquationContext: mRevision = mOwner.AcceptedRevision: mTarget = target
    mAllowJacobian = allowJacobian
End Sub

' Создаёт отдельную неизменную цель; завершённый контекст другой цели не становится действующим повторно.
Friend Function AtTarget(ByVal owner As CCoupledNewtonSolver, ByVal target As Double, _
        Optional ByVal allowJacobian As Boolean = False) As CCoupledEnvelopeEquations
    Dim result As New CCoupledEnvelopeEquations
    result.Initialize owner, mSection, mCentral, mMinus, mPlus, mHalfStep, mRowScale
    result.OpenTarget target, allowJacobian
    Set AtTarget = result
End Function

' Возвращает фиксированный перевод отклонения H в масштаб силовых строк, без изменения физических параметров.
Friend Property Get RowScale() As Double
    RowScale = mRowScale
End Property

' Суммирует действующий принятый либо полностью подготовленный пробный буфер без повторного возврата материала.
Friend Function TryBufferedCoordinate(ByVal accepted As Boolean, ByRef value As Double) As Boolean
    If mSection Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    Dim point As Long, side As Long, pointValue As Double, response As CCoupledReinforcedResponse
    value = 0#
    For point = 1 To UBound(mArea)
        For side = -1 To 1
            Set response = ResponseAt(side)
            If Not response.TryReadBufferedEnvelope(point, accepted, pointValue) Then Exit Function
            value = value + mArea(point) * pointValue
        Next side
    Next point
    value = value / (3# * mTotalArea): TryBufferedCoordinate = True
End Function

' Повторяет H в произвольном поле от текущих принятых историй, не принимая и не записывая пробу.
Friend Function TryCoordinate(ByRef field() As Double, ByRef value As Double) As Boolean
    If mSection Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    Dim point As Long, side As Long, component As Long, pointValue As Double, response As CCoupledReinforcedResponse
    Dim strain(1 To 6) As Double, slope(1 To 6) As Double, station(1 To 6) As Double
    value = 0#
    For point = 1 To UBound(mArea)
        mSection.PointStrain point, field, strain: mSection.AxialGradientAtPoint point, field, slope
        For side = -1 To 1
            For component = 1 To 6: station(component) = strain(component) + CDbl(side) * mHalfStep * slope(component): Next component
            Set response = ResponseAt(side)
            If Not response.TryReadEnvelope(point, station, pointValue) Then Exit Function
            value = value + mArea(point) * pointValue
        Next side
    Next point
    value = value / (3# * mTotalArea): TryCoordinate = True
End Function

' Считает производную H по физическому полю через шесть независимых деформаций каждой бетонной точки.
' История и пробные буферы не меняются; результат передаётся только после успешного обхода всех точек трёх станций.
Friend Function TryCoordinateGradient(ByRef field() As Double, ByRef gradient() As Double) As Boolean
    Const RELATIVE_STEP As Double = 0.000001
    Const STRAIN_SCALE As Double = 0.00001
    If mSection Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    If LBound(field) <> 1 Or UBound(field) <> mSection.DegreeCount Then _
        Err.Raise 5, "CCoupledEnvelopeEquations", "Размер поля не совпадает с материальной координатой."
    Dim coefficients() As Double, result() As Double
    mSection.CopySpatialStrainCoefficients mHalfStep, coefficients
    ReDim result(1 To mSection.DegreeCount)
    Dim lower(1 To 6) As Long, upper(1 To 6) As Long
    lower(1) = 1: upper(1) = 1: lower(2) = 2: upper(2) = 2
    lower(3) = 4: upper(3) = 5: lower(4) = 3: upper(4) = 3
    lower(5) = 6: upper(5) = 7: lower(6) = 8: upper(6) = 9
    Dim point As Long, side As Long, component As Long, degree As Long, offset As Long
    Dim strain(1 To 6) As Double, slope(1 To 6) As Double, station(1 To 6) As Double, shifted(1 To 6) As Double
    Dim response As CCoupledReinforcedResponse, stepSize As Double, plus As Double, minus As Double
    Dim positive As Double, negative As Double, derivative As Double, weight As Double, coefficient As Double
    For point = 1 To UBound(mArea)
        offset = 9 * (point - 1): weight = mArea(point) / (3# * mTotalArea)
        mSection.PointStrain point, field, strain: mSection.AxialGradientAtPoint point, field, slope
        For side = -1 To 1
            Set response = ResponseAt(side)
            For component = 1 To 6
                station(component) = strain(component) + CDbl(side) * mHalfStep * slope(component)
                shifted(component) = station(component)
            Next component
            For component = 1 To 6
                stepSize = RELATIVE_STEP * GeomMax(STRAIN_SCALE, Abs(station(component)))
                positive = station(component) + stepSize: negative = station(component) - stepSize
                shifted(component) = positive
                If Not response.TryReadEnvelope(point, shifted, plus) Then Exit Function
                shifted(component) = negative
                If Not response.TryReadEnvelope(point, shifted, minus) Then Exit Function
                shifted(component) = station(component)
                derivative = (plus - minus) / (positive - negative) ' Касательные деформации инженерные; их преобразование в тензор выполняет отклик материала.
                For degree = 1 To UBound(result)
                    If side < 0 Then
                        coefficient = coefficients(degree, offset + lower(component))
                    ElseIf side > 0 Then
                        coefficient = coefficients(degree, offset + upper(component))
                    Else
                        coefficient = 0.5 * (coefficients(degree, offset + lower(component)) + coefficients(degree, offset + upper(component)))
                    End If
                    result(degree) = result(degree) + weight * derivative * coefficient
                Next degree
            Next component
        Next side
    Next point
    gradient = result: TryCoordinateGradient = True
End Function

' Сохраняет силовые уравнения и заменяет последнюю строку условием H.
' По запросу цели матрица использует уточнённую ветвесогласованную касательную владельца и производную H.
Private Function ICoupledEquationSystem_TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
    If mOwner Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    If Not mOwner.MatchesEquationContext(mContext, mRevision) Then Err.Raise 5, "CCoupledEnvelopeEquations", "История или контекст материальной координаты изменились."
    If withJacobian And Not mAllowJacobian Then Err.Raise 5, "CCoupledEnvelopeEquations", "Для этой цели материальной координаты не запрошена собственная касательная."
    If withJacobian Then
        If Not mOwner.TryReadBranchLinearization(coordinates, residual, jacobian) Then Exit Function
    Else
        If Not mRaw.TryEvaluate(coordinates, False, residual, jacobian) Then Exit Function
    End If
    Dim field() As Double, factor As Double, value As Double
    mOwner.ReadEquationField coordinates, field, factor
    If Not TryCoordinate(field, value) Then Exit Function
    residual(UBound(residual)) = (value - mTarget) * mRowScale
    If withJacobian Then
        Dim gradient() As Double, degree As Long, size As Long
        If Not TryCoordinateGradient(field, gradient) Then Exit Function
        size = UBound(residual)
        If UBound(gradient) <> size - 1 Then Err.Raise 5, "CCoupledEnvelopeEquations", "Размер производной H не совпадает с системой равновесия."
        For degree = 1 To UBound(gradient)
            jacobian(degree, size) = gradient(degree) * mRowScale * mOwner.EquationCoordinateScale(degree)
        Next degree
        jacobian(size, size) = 0# ' Материальная координата зависит от поля и принятой истории, но не от множителя нагрузки отдельно.
    End If
    ICoupledEquationSystem_TryEvaluate = True
End Function

' Выбирает ту же станцию, знак которой применяется в геометрическом полушаге сечения.
Private Function ResponseAt(ByVal side As Long) As CCoupledReinforcedResponse
    If side < 0 Then
        Set ResponseAt = mMinus
    ElseIf side > 0 Then
        Set ResponseAt = mPlus
    Else
        Set ResponseAt = mCentral
    End If
End Function
