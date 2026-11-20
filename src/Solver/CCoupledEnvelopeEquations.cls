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
Friend Sub OpenTarget(ByVal target As Double)
    If mOwner Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    If mContext <> 0 Then Err.Raise 5, "CCoupledEnvelopeEquations", "Для другого уровня или контекста создайте новую материальную координату."
    If target < 0# Or mOwner.EquationContext = 0 Then Err.Raise 5, "CCoupledEnvelopeEquations", "Нет действующего контекста или задан отрицательный уровень H."
    mContext = mOwner.EquationContext: mRevision = mOwner.AcceptedRevision: mTarget = target
End Sub

' Создаёт отдельную неизменную цель; завершённый контекст другой цели не становится действующим повторно.
Friend Function AtTarget(ByVal owner As CCoupledNewtonSolver, ByVal target As Double) As CCoupledEnvelopeEquations
    Dim result As New CCoupledEnvelopeEquations
    result.Initialize owner, mSection, mCentral, mMinus, mPlus, mHalfStep, mRowScale
    result.OpenTarget target
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

' Сохраняет силовые уравнения и заменяет последнюю управляющую строку; матрица строится общим разностным решателем.
Private Function ICoupledEquationSystem_TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
    If mOwner Is Nothing Then Err.Raise 5, "CCoupledEnvelopeEquations", "Материальная координата не подготовлена."
    If Not mOwner.MatchesEquationContext(mContext, mRevision) Then Err.Raise 5, "CCoupledEnvelopeEquations", "История или контекст материальной координаты изменились."
    If withJacobian Then Err.Raise 5, "CCoupledEnvelopeEquations", "Для материальной координаты требуется полная разностная матрица."
    If Not mRaw.TryEvaluate(coordinates, False, residual, jacobian) Then Exit Function
    Dim field() As Double, factor As Double, value As Double
    mOwner.ReadEquationField coordinates, field, factor
    If Not TryCoordinate(field, value) Then Exit Function
    residual(UBound(residual)) = (value - mTarget) * mRowScale
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
