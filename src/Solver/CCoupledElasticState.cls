VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledElasticState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Упругое состояние сечения при шести усилиях
' ==========================================================================
' Реализует слабое равновесие локальной модели Bairán для однородной
' изотропной среды. Используется для верификации кинематики и границ.
' Прочность, армирование и история материала определяются отдельно.

Private Const PLANE_DOF As Long = 4
Private Const STRESS_COMPONENTS As Long = 6
Private Const SHAPE_COUNT As Long = 16
Private Const GAUGES As Long = 4

Private mQuadrature As CCoupledSectionQuadrature
Private mDof As Long
Private mGradientBase As Long
Private mGaugeBase As Long
Private mPointCount As Long
Private mSupport() As Long
Private mSectionDof() As Long
Private mX() As Double, mY() As Double, mWeight() As Double
Private mOperators() As Double ' A(компонента, локальная степень свободы, точка).
Private mShape() As Double ' Значение функции перемещений для правой части межволоконного равновесия.
Private mElastic(1 To 6, 1 To 6) As Double
Private mMatrix() As Double
Private mLength As Double, mArea As Double
Private mMeta As CResultMeta

' Подготавливает упругую систему и её геометрические операторы один раз.
' Координаты, обобщённые перемещения и силы нормированы по длине и площади.
Public Sub PrepareModel(ByVal section As CSectionModel, ByVal modulus As Double, ByVal poisson As Double, _
        ByVal spansX As Long, ByVal spansY As Long, ByVal pointsPerSpan As Long)
    Set mMeta = New CResultMeta
    mArea = 0#
    If modulus <= 0# Or poisson <= -1# Or poisson >= 0.5 Then
        Err.Raise vbObjectError + 1776, "CCoupledElasticState", _
            "Для упругого расчёта требуется E > 0 и −1 < ν < 0,5."
    End If
    Set mQuadrature = New CCoupledSectionQuadrature
    mQuadrature.Initialize section, spansX, spansY
    mLength = mQuadrature.Frame.LengthScale
    SetElasticTensor modulus, poisson
    Dim indices() As Long, dx() As Double, dy() As Double
    mQuadrature.Prepare pointsPerSpan, mX, mY, mWeight, indices, mShape, dx, dy
    mPointCount = UBound(mWeight)
    PrepareOperators indices, dx, dy
    Assemble
End Sub

' Решает при заданных усилиях относительно численного начала.
' Усилия и напряжения записывает после проверки невязки исходных уравнений.
Public Function Solve(ByVal n As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal torque As Double, ByVal qx As Double, ByVal qy As Double, _
        ByRef deformation() As Double, ByRef recovered() As Double, _
        ByRef pointStress() As Double) As Boolean
    mMeta.Clear
    Erase deformation: Erase recovered: Erase pointStress
    Dim matrix() As Double, rhs() As Double, applied() As Double, i As Long
    matrix = mMatrix
    ReDim rhs(1 To mDof)
    rhs(1) = n / mArea
    rhs(2) = mx / mArea / mLength
    rhs(3) = my / mArea / mLength
    rhs(4) = torque / mArea / mLength
    rhs(mGradientBase + 2) = qy / mArea
    rhs(mGradientBase + 3) = qx / mArea
    applied = rhs
    Dim system As New CCoupledLinearSystem
    If Not system.Solve(matrix, rhs, mDof) Then
        mMeta.SetSolverFailure system.FailureCode, rkDirectState, _
            "Не удалось решить упругие уравнения связанного сечения. Проверьте разрешение базиса и геометрию.", True
        Exit Function
    End If
    If Not EquilibriumVerified(rhs, applied) Then
        mMeta.SetSolverFailure sfcNumericalFailure, rkDirectState, _
            "Упругие поля связанного сечения не прошли проверку внутреннего равновесия.", True
        Exit Function
    End If
    RecoverFields rhs, recovered, pointStress
    If Not ResultantsVerified(recovered, n, mx, my, torque, qx, qy) Then
        Erase recovered: Erase pointStress
        mMeta.SetSolverFailure sfcNumericalFailure, rkDirectState, _
            "Интегралы напряжений связанного сечения не совпали с заданными усилиями.", True
        Exit Function
    End If
    ReDim deformation(1 To 4)
    deformation(1) = rhs(1)
    For i = 2 To 4
        deformation(i) = rhs(i) / mLength
    Next i
    mMeta.SetResult rsSuccess, rcNone, rkDirectState
    Solve = True
End Function

' Возвращает координаты расчётной точки в исходных осях и её вес, мм².
Friend Sub Point(ByVal index As Long, ByRef x As Double, ByRef y As Double, ByRef area As Double)
    x = mX(index) * mLength + mQuadrature.Frame.OriginX
    y = mY(index) * mLength + mQuadrature.Frame.OriginY
    area = mWeight(index) * mArea
End Sub

' Возвращает структурированную причину результата в общем формате проекта.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Вычисляет матрицу закона Гука для инженерных угловых деформаций.
Private Sub SetElasticTensor(ByVal modulus As Double, ByVal poisson As Double)
    Dim shear As Double, lame As Double, i As Long, j As Long
    shear = modulus / (2# * (1# + poisson))
    lame = 2# * shear * poisson / (1# - 2# * poisson)
    For i = 1 To 3
        For j = 1 To 3
            mElastic(i, j) = lame
        Next j
        mElastic(i, i) = lame + 2# * shear
        mElastic(i + 3, i + 3) = shear
    Next i
End Sub

' Нумерует базис с ненулевой опорой и сохраняет его локальные производные.
' Физические градиенты переводятся к коэффициентам u/L, v/L и w/L.
Private Sub PrepareOperators(ByRef indices() As Long, ByRef dx() As Double, ByRef dy() As Double)
    ReDim mSupport(1 To mQuadrature.BasisCount)
    Dim i As Long, k As Long, globalIndex As Long, active As Long
    For i = 1 To mPointCount
        mArea = mArea + mWeight(i)
        For k = 1 To SHAPE_COUNT
            globalIndex = indices(k, i)
            If mSupport(globalIndex) = 0 Then
                If mShape(k, i) <> 0# Or dx(k, i) <> 0# Or dy(k, i) <> 0# Then mSupport(globalIndex) = -1
            End If
        Next k
    Next i
    For globalIndex = 1 To UBound(mSupport)
        If mSupport(globalIndex) = -1 Then
            active = active + 1
            mSupport(globalIndex) = active
        End If
    Next globalIndex
    mGradientBase = PLANE_DOF + 3 * active
    mGaugeBase = mGradientBase + PLANE_DOF
    mDof = mGaugeBase + GAUGES

    Dim localCount As Long
    localCount = PLANE_DOF + 3 * SHAPE_COUNT
    ReDim mSectionDof(1 To localCount, 1 To mPointCount)
    ReDim mOperators(1 To STRESS_COMPONENTS, 1 To localCount, 1 To mPointCount)
    Dim u As Long, v As Long, w As Long
    Dim gX As Double, gY As Double
    For i = 1 To mPointCount
        mWeight(i) = mWeight(i) / mArea
        mX(i) = mX(i) / mLength
        mY(i) = mY(i) / mLength
        For k = 1 To 4
            mSectionDof(k, i) = k
        Next k
        mOperators(3, 1, i) = 1#
        mOperators(3, 2, i) = mY(i)
        mOperators(3, 3, i) = mX(i)
        mOperators(5, 4, i) = -mY(i)
        mOperators(6, 4, i) = mX(i)
        For k = 1 To SHAPE_COUNT
            globalIndex = mSupport(indices(k, i))
            If globalIndex > 0 Then
                u = PLANE_DOF + 3 * k - 2
                v = u + 1: w = u + 2
                mSectionDof(u, i) = PLANE_DOF + 3 * globalIndex - 2
                mSectionDof(v, i) = mSectionDof(u, i) + 1
                mSectionDof(w, i) = mSectionDof(u, i) + 2
                gX = dx(k, i) * mLength: gY = dy(k, i) * mLength
                mOperators(1, u, i) = gX: mOperators(4, u, i) = gY
                mOperators(2, v, i) = gY: mOperators(4, v, i) = gX
                mOperators(5, w, i) = gX: mOperators(6, w, i) = gY
            End If
        Next k
    Next i
End Sub

' Собирает интегралы B^T D B, уравнения производных усилий и четыре условия
' ортогональности. Члены σ'zx, σ'zy, σ'zz берутся из D Bp d, где d = L dq/dz.
Private Sub Assemble()
    ReDim mMatrix(1 To mDof, 1 To mDof)
    Dim i As Long, a As Long, b As Long, j As Long, k As Long
    Dim row As Long, col As Long, localCount As Long, component As Long
    Dim product As Double, weighted() As Double
    localCount = UBound(mSectionDof, 1)
    ReDim weighted(1 To STRESS_COMPONENTS, 1 To localCount)
    For i = 1 To mPointCount
        For b = 1 To localCount
            If mSectionDof(b, i) > 0 Then
                For j = 1 To STRESS_COMPONENTS
                    product = 0#
                    For k = 1 To STRESS_COMPONENTS
                        product = product + mElastic(j, k) * mOperators(k, b, i)
                    Next k
                    weighted(j, b) = product * mWeight(i)
                Next j
            End If
        Next b
        For a = 1 To localCount
            row = mSectionDof(a, i)
            If row > 0 Then
                For b = 1 To localCount
                    col = mSectionDof(b, i)
                    If col > 0 Then
                        product = 0#
                        For j = 1 To STRESS_COMPONENTS
                            product = product + mOperators(j, a, i) * weighted(j, b)
                        Next j
                        mMatrix(col, row) = mMatrix(col, row) + product
                        If a <= PLANE_DOF And b <= PLANE_DOF Then
                            mMatrix(mGradientBase + b, mGradientBase + a) = _
                                mMatrix(mGradientBase + b, mGradientBase + a) + product
                        End If
                    End If
                Next b
                If a > PLANE_DOF Then
                    k = (a - PLANE_DOF - 1) \ 3 + 1
                    component = (a - PLANE_DOF - 1) Mod 3 + 1
                    AssembleTraction i, k, component, row, weighted
                End If
            End If
        Next a
    Next i
End Sub

' Добавляет источник продольного градиента и условия отсутствия жёсткого
' перемещения по трём осям и поворота вокруг продольной оси.
Private Sub AssembleTraction(ByVal point As Long, ByVal k As Long, ByVal component As Long, _
        ByVal row As Long, ByRef weighted() As Double)
    Dim stressIndex As Long, b As Long, factor As Double
    Select Case component
        Case 1: stressIndex = 5
        Case 2: stressIndex = 6
        Case 3: stressIndex = 3
    End Select
    For b = 1 To PLANE_DOF
        mMatrix(mGradientBase + b, row) = mMatrix(mGradientBase + b, row) - _
            mShape(k, point) * weighted(stressIndex, b)
    Next b
    factor = mShape(k, point) * mWeight(point)
    AddGauge row, mGaugeBase + component, factor
    If component = 1 Then AddGauge row, mGaugeBase + 4, -mY(point) * factor
    If component = 2 Then AddGauge row, mGaugeBase + 4, mX(point) * factor
End Sub

' Вводит множитель точного ограничения без штрафной жёсткости.
Private Sub AddGauge(ByVal row As Long, ByVal gauge As Long, ByVal factor As Double)
    mMatrix(gauge, row) = mMatrix(gauge, row) + factor
    mMatrix(row, gauge) = mMatrix(row, gauge) + factor
End Sub

' Проверяет исходные уравнения, включая ортогональность и слабое равновесие.
Private Function EquilibriumVerified(ByRef solution() As Double, ByRef applied() As Double) As Boolean
    Dim i As Long, j As Long, term As Double, residual As Double, equationScale As Double
    For i = 1 To mDof
        residual = -applied(i): equationScale = Abs(applied(i))
        For j = 1 To mDof
            term = mMatrix(j, i) * solution(j)
            residual = residual + term
            equationScale = equationScale + Abs(term)
        Next j
        If Abs(residual) > 1E-10 * equationScale + 1E-13 Then Exit Function
    Next i
    EquilibriumVerified = True
End Function

' Независимо от строк матрицы интегрирует индикаторы полного силового винта.
Private Sub RecoverFields(ByRef solution() As Double, ByRef recovered() As Double, ByRef stress() As Double)
    ReDim recovered(1 To 6)
    ReDim stress(1 To STRESS_COMPONENTS, 1 To mPointCount)
    Dim i As Long, j As Long, k As Long, a As Long, degree As Long
    Dim strain(1 To STRESS_COMPONENTS) As Double, contribution As Double, area As Double
    For i = 1 To mPointCount
        For j = 1 To STRESS_COMPONENTS
            strain(j) = 0#
        Next j
        For a = 1 To UBound(mSectionDof, 1)
            degree = mSectionDof(a, i)
            If degree > 0 Then
                contribution = solution(degree)
                For j = 1 To STRESS_COMPONENTS
                    strain(j) = strain(j) + mOperators(j, a, i) * contribution
                Next j
            End If
        Next a
        For j = 1 To STRESS_COMPONENTS
            For k = 1 To STRESS_COMPONENTS
                stress(j, i) = stress(j, i) + mElastic(j, k) * strain(k)
            Next k
        Next j
        area = mArea * mWeight(i)
        recovered(1) = recovered(1) + area * stress(3, i)
        recovered(2) = recovered(2) + area * stress(3, i) * mY(i) * mLength
        recovered(3) = recovered(3) + area * stress(3, i) * mX(i) * mLength
        recovered(4) = recovered(4) + area * (mX(i) * stress(6, i) - mY(i) * stress(5, i)) * mLength
        recovered(5) = recovered(5) + area * stress(5, i)
        recovered(6) = recovered(6) + area * stress(6, i)
    Next i
End Sub

' Сверяет компоненты с общим приводимым масштабом сил и плечом mLength.
Private Function ResultantsVerified(ByRef recovered() As Double, ByVal n As Double, ByVal mx As Double, _
        ByVal my As Double, ByVal torque As Double, ByVal qx As Double, ByVal qy As Double) As Boolean
    Dim forceScale As Double, momentScale As Double
    forceScale = Abs(n) + Abs(qx) + Abs(qy) + (Abs(mx) + Abs(my) + Abs(torque)) / mLength
    forceScale = 1E-8 * forceScale + 1E-7
    momentScale = forceScale * mLength
    ResultantsVerified = Abs(recovered(1) - n) <= forceScale And Abs(recovered(2) - mx) <= momentScale And _
        Abs(recovered(3) - my) <= momentScale And Abs(recovered(4) - torque) <= momentScale And _
        Abs(recovered(5) - qx) <= forceScale And Abs(recovered(6) - qy) <= forceScale
End Function
