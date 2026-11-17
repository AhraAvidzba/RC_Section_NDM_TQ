VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledSteelCorrection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Поправка Ньютона при смене ветвей стали на продольных станциях
' ==========================================================================
' Действует для силового шага с тремя независимыми историями. Бетонная
' касательная фиксирована, арматура пересматривается от принятой истории.

Private Const COMPONENT_COUNT As Long = 6
Private Const BRANCH_ATTEMPTS As Long = 8
Private Const FIELD_PERTURBATION As Double = 0.000000002 ' Безразмерное возмущение координаты поля.

Private Type TSteelStation
    AxialStrain As Double
    Stress As Double
    Tangent As Double
    Branch As Long
End Type

Private mSection As CCoupledElasticState
Private mResponse(-1 To 1) As CCoupledReinforcedResponse
Private mHalfStep As Double, mFirstSteel As Long, mPointCount As Long

' Запоминает операторы и три независимые материальные истории силового шага.
Friend Sub Initialize(ByVal section As CCoupledElasticState, ByVal center As CCoupledReinforcedResponse, _
        ByVal positive As CCoupledReinforcedResponse, ByVal negative As CCoupledReinforcedResponse, ByVal halfStep As Double)
    Set mSection = section
    Set mResponse(0) = center: Set mResponse(1) = positive: Set mResponse(-1) = negative
    mHalfStep = halfStep
    mFirstSteel = section.ConcretePointCount + 1: mPointCount = section.PointCount
End Sub

' Уточняет направление на силовом шаге. При отсутствии смены ветвей
' сохраняет штатную поправку. Сходимость уточнённого направления проверяет
' вызывающий по полному отклику бетона и стали и по всем шести усилиям.
Friend Function Refine(ByRef field() As Double, ByRef matrix() As Double, ByRef residual() As Double, _
        ByRef correction() As Double, Optional ByVal concreteMatrix As Boolean = False) As Long
    Dim first As Long, last As Long, size As Long
    first = mFirstSteel: last = mPointCount: size = UBound(residual)
    Dim base() As TSteelStation
    ReDim base(first To last, -1 To 1)
    Dim point As Long, side As Long, i As Long, iteration As Long, position As Long
    Dim mask As String, visited(0 To BRANCH_ATTEMPTS) As String
    Dim strain(1 To COMPONENT_COUNT) As Double
    Dim gradient(1 To COMPONENT_COUNT) As Double, stationStrain(1 To COMPONENT_COUNT) As Double
    mask = String$(3 * (last - first + 1), "0")
    For point = first To last
        mSection.PointStrain point, field, strain
        mSection.AxialGradientAtPoint point, field, gradient
        For side = -1 To 1
            For i = 1 To COMPONENT_COUNT: stationStrain(i) = strain(i) + side * mHalfStep * gradient(i): Next i
            ReadBranch side, point, stationStrain, base(point, side)
            position = position + 1
            Mid$(mask, position, 1) = Chr$(49 + base(point, side).Branch)
        Next side
    Next point
    visited(0) = mask

    Dim candidate() As Double, adjusted() As Double, linearized() As Double, linear As CCoupledLinearSystem
    Dim changed As Boolean, pointChanged As Boolean, predicted As TSteelStation
    Dim intercept(-1 To 1) As Double, slope(-1 To 1) As Double
    If concreteMatrix Then
        adjusted = residual: linearized = matrix
        For point = first To last
            For side = -1 To 1: slope(side) = base(point, side).Tangent: Next side
            AccumulateSteelLaw point, intercept, slope, adjusted, linearized
        Next point
        Set linear = New CCoupledLinearSystem
        For i = 1 To size: adjusted(i) = -adjusted(i): Next i
        If Not linear.Solve(linearized, adjusted, size) Then Exit Function
        correction = adjusted
    End If
    ReDim candidate(1 To size)
    For iteration = 1 To BRANCH_ATTEMPTS
        For i = 1 To size: candidate(i) = field(i) + correction(i): Next i
        changed = False: position = 0
        For point = first To last
            mSection.PointStrain point, candidate, strain
            mSection.AxialGradientAtPoint point, candidate, gradient
            pointChanged = False
            For side = -1 To 1
                For i = 1 To COMPONENT_COUNT: stationStrain(i) = strain(i) + side * mHalfStep * gradient(i): Next i
                ReadBranch side, point, stationStrain, predicted
                position = position + 1
                Mid$(mask, position, 1) = Chr$(49 + predicted.Branch)
                intercept(side) = 0#: slope(side) = 0#
                If base(point, side).Branch <> predicted.Branch Then
                    slope(side) = predicted.Tangent - base(point, side).Tangent
                    intercept(side) = predicted.Stress - base(point, side).Stress - _
                        predicted.Tangent * (predicted.AxialStrain - base(point, side).AxialStrain)
                    pointChanged = True ' Сдвигает свободный член к выбранному участку диаграммы.
                End If
                If concreteMatrix Then slope(side) = predicted.Tangent
            Next side
            If pointChanged Or concreteMatrix Then
                If Not changed Then
                    adjusted = residual: linearized = matrix
                    changed = True
                End If
                AccumulateSteelLaw point, intercept, slope, adjusted, linearized
            End If
        Next point
        If Not changed Then Exit Function
        For i = 0 To iteration - 1
            If visited(i) = mask Then Exit Function ' Остаток проверяется на последнем новом направлении.
        Next i
        visited(iteration) = mask
        If linear Is Nothing Then Set linear = New CCoupledLinearSystem
        For i = 1 To size: adjusted(i) = -adjusted(i): Next i
        If Not linear.Solve(linearized, adjusted, size) Then Exit Function
        correction = adjusted
        Refine = Refine + 1
    Next iteration
End Function

' Проецирует аффинную поправку трёх участков стали теми же операторами равновесия.
Private Sub AccumulateSteelLaw(ByVal point As Long, _
        ByRef intercept() As Double, ByRef slope() As Double, ByRef residual() As Double, ByRef matrix() As Double)
    Dim i As Long, j As Long, projection As Double
    Dim direction(1 To COMPONENT_COUNT) As Double, stress(1 To COMPONENT_COUNT) As Double
    Dim stressGradient(1 To COMPONENT_COUNT) As Double, tangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
    Dim strainTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double
    Dim slopeTangent(1 To COMPONENT_COUNT, 1 To COMPONENT_COUNT) As Double, unusedResultants(1 To COMPONENT_COUNT) As Double
    mSection.SteelProjection point, direction
    For j = 1 To COMPONENT_COUNT
        stress(j) = intercept(0) * direction(j)
        stressGradient(j) = (intercept(1) - intercept(-1)) * direction(j) / (2# * mHalfStep)
        For i = 1 To COMPONENT_COUNT
            projection = direction(i) * direction(j)
            tangent(i, j) = projection * slope(0)
            strainTangent(i, j) = projection * (slope(1) - slope(-1)) / (2# * mHalfStep)
            slopeTangent(i, j) = projection * 0.5 * (slope(1) + slope(-1))
        Next i
    Next j
    mSection.AddSpatialPointEquilibrium point, stress, tangent, stressGradient, strainTangent, _
        slopeTangent, residual, matrix, True, unusedResultants
End Sub

' Заполняет копию скалярного отклика и знаковый номер участка диаграммы.
Private Sub ReadBranch(ByVal side As Long, ByVal point As Long, _
        ByRef strain() As Double, ByRef station As TSteelStation)
    station.Branch = mResponse(side).ReadSteelBranch(point, strain, station.AxialStrain, station.Stress, station.Tangent)
End Sub

' Дифференцирует бетонный остаток по полю с сохранением принятой истории.
' Связи входят в разностную матрицу, стальные участки добавляет Refine.
Friend Function TryConcreteTangent(ByRef field() As Double, ByRef external() As Double, _
        ByRef matrix() As Double, ByRef evaluations As Long) As Boolean
    Dim size As Long, row As Long, col As Long
    Dim perturbed() As Double, positive() As Double, negative() As Double
    size = UBound(field): perturbed = field
    ReDim matrix(1 To size, 1 To size)
    For col = 1 To size
        perturbed(col) = field(col) + FIELD_PERTURBATION
        evaluations = evaluations + 1
        If Not ConcreteResidual(perturbed, external, positive) Then Exit Function
        perturbed(col) = field(col) - FIELD_PERTURBATION
        evaluations = evaluations + 1
        If Not ConcreteResidual(perturbed, external, negative) Then Exit Function
        For row = 1 To size: matrix(col, row) = (positive(row) - negative(row)) / (2# * FIELD_PERTURBATION): Next row
        perturbed(col) = field(col)
    Next col
    TryConcreteTangent = True
End Function

' Собирает бетонные напряжения на трёх станциях без записи пробных буферов.
' Для разностной касательной физические усилия на выходе не используются.
Private Function ConcreteResidual(ByRef field() As Double, ByRef external() As Double, ByRef residual() As Double) As Boolean
    Dim strain(1 To COMPONENT_COUNT) As Double, axialGradient(1 To COMPONENT_COUNT) As Double
    Dim stationStrain(1 To COMPONENT_COUNT) As Double, stress(1 To COMPONENT_COUNT) As Double
    Dim plusStress(1 To COMPONENT_COUNT) As Double, minusStress(1 To COMPONENT_COUNT) As Double
    Dim stressGradient(1 To COMPONENT_COUNT) As Double, unusedResultants(1 To COMPONENT_COUNT) As Double
    Dim unusedMatrix() As Double, unusedTangent() As Double, point As Long, i As Long
    mSection.StartEquations field, external, residual, unusedMatrix, False
    For point = 1 To mPointCount
        mSection.PointStrain point, field, strain
        mSection.AxialGradientAtPoint point, field, axialGradient
        If mResponse(0).ReadConcreteStress(point, strain, stress) <> sfcNone Then Exit Function
        For i = 1 To COMPONENT_COUNT: stationStrain(i) = strain(i) + mHalfStep * axialGradient(i): Next i
        If mResponse(1).ReadConcreteStress(point, stationStrain, plusStress) <> sfcNone Then Exit Function
        For i = 1 To COMPONENT_COUNT: stationStrain(i) = strain(i) - mHalfStep * axialGradient(i): Next i
        If mResponse(-1).ReadConcreteStress(point, stationStrain, minusStress) <> sfcNone Then Exit Function
        For i = 1 To COMPONENT_COUNT: stressGradient(i) = (plusStress(i) - minusStress(i)) / (2# * mHalfStep): Next i
        mSection.AddSpatialPointEquilibrium point, stress, unusedTangent, stressGradient, unusedTangent, _
            unusedTangent, residual, unusedMatrix, False, unusedResultants
    Next point
    ConcreteResidual = True
End Function
