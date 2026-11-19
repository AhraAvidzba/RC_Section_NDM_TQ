VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledMaterialConstraint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Материальная координата в неизменной истории связанного сечения
' ==========================================================================
' Пробная поверхность CDP или главное напряжение ASD читаются существующим
' откликом. Геометрия, точка, станция, масштабы и уровень задаются явно.
' Изменение истории или контекста владельца делает координату непригодной.
Implements ICoupledScalarConstraint

Private mOwner As CCoupledNewtonSolver, mSection As CCoupledElasticState
Private mResponse As CCoupledReinforcedResponse
Private mPoint As Long, mPrincipal As Long, mCount As Long, mRevision As Long, mContext As Long
Private mStation As Double, mTarget As Double, mSeed() As Double, mScales() As Double, mBasis() As Double

' Готовит независимые копии масштабов и геометрические производные один раз на действующий контекст решателя.
Friend Sub Initialize(ByVal owner As CCoupledNewtonSolver, ByVal section As CCoupledElasticState, _
        ByVal response As CCoupledReinforcedResponse, ByVal point As Long, ByVal station As Double, _
        ByVal principal As Long, ByRef seed() As Double, ByRef scales() As Double, ByVal target As Double)
    If Not mOwner Is Nothing Then Err.Raise 5, "CCoupledMaterialConstraint", "Материальная координата уже подготовлена."
    If owner Is Nothing Or section Is Nothing Or response Is Nothing Then _
        Err.Raise 5, "CCoupledMaterialConstraint", "Для материальной координаты нужны решатель, геометрия и отклик."
    If point < 1 Or point > section.PointCount Or principal < 0 Or principal > 3 Then _
        Err.Raise 5, "CCoupledMaterialConstraint", "Точка или вид материальной границы не принадлежат подготовленной модели."
    mCount = section.DegreeCount + 1
    If LBound(seed) <> 1 Or UBound(seed) <> mCount Or LBound(scales) <> 1 Or UBound(scales) <> mCount Then _
        Err.Raise 5, "CCoupledMaterialConstraint", "Размеры исходного поля и масштабов материальной координаты не согласованы."
    Dim degree As Long, component As Long, unit() As Double, strain(1 To 6) As Double, slope(1 To 6) As Double
    For degree = 1 To mCount
        If scales(degree) <= 0# Then Err.Raise 5, "CCoupledMaterialConstraint", "Масштабы материальной координаты должны быть положительными."
    Next degree
    mContext = owner.EquationContext: mRevision = owner.AcceptedRevision
    If mContext = 0 Then Err.Raise 5, "CCoupledMaterialConstraint", "Нет действующего контекста уравнений для материальной координаты."
    Set mOwner = owner: Set mSection = section: Set mResponse = response
    mPoint = point: mStation = station: mPrincipal = principal: mTarget = target
    mSeed = seed: mScales = scales
    ReDim mBasis(1 To mCount, 1 To 6): ReDim unit(1 To mCount - 1)
    For degree = 1 To mCount - 1
        unit(degree) = 1#
        mSection.PointStrain mPoint, unit, strain: mSection.AxialGradientAtPoint mPoint, unit, slope
        For component = 1 To 6: mBasis(degree, component) = (strain(component) + mStation * slope(component)) * mScales(degree): Next component
        unit(degree) = 0#
    Next degree
End Sub

' Читает границу в действительном физическом поле; принятая история и исходные координаты не меняются.
Private Function ICoupledScalarConstraint_TryEvaluate(ByRef coordinates() As Double, ByVal withGradient As Boolean, _
        ByRef value As Double, ByRef gradient() As Double) As Boolean
    If mOwner Is Nothing Then Err.Raise 5, "CCoupledMaterialConstraint", "Материальная координата не подготовлена."
    If Not mOwner.MatchesEquationContext(mContext, mRevision) Then _
        Err.Raise 5, "CCoupledMaterialConstraint", "История или контекст материальной координаты изменились."
    If LBound(coordinates) <> 1 Or UBound(coordinates) <> mCount Then _
        Err.Raise 5, "CCoupledMaterialConstraint", "Размер пробы материальной координаты не совпадает с подготовленным полем."
    Dim field() As Double, strain(1 To 6) As Double, slope(1 To 6) As Double, materialGradient(1 To 6) As Double
    Dim degree As Long, component As Long, read As Boolean
    ReDim field(1 To mCount - 1)
    For degree = 1 To mCount - 1: field(degree) = mSeed(degree) + mScales(degree) * coordinates(degree): Next degree
    mSection.PointStrain mPoint, field, strain: mSection.AxialGradientAtPoint mPoint, field, slope
    For component = 1 To 6: strain(component) = strain(component) + mStation * slope(component): Next component
    If mPrincipal = 0 Then
        read = mResponse.TryReadTrialYield(mPoint, strain, value, materialGradient, withGradient)
    Else
        read = mResponse.TryReadPrincipalBoundary(mPoint, strain, mPrincipal, value, materialGradient, withGradient)
    End If
    If Not read Then Exit Function
    value = value - mTarget
    If withGradient Then
        ReDim gradient(1 To mCount)
        For degree = 1 To mCount - 1
            For component = 1 To 6: gradient(degree) = gradient(degree) + materialGradient(component) * mBasis(degree, component): Next component
        Next degree
    End If
    ICoupledScalarConstraint_TryEvaluate = True
End Function
