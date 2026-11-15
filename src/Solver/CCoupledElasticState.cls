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
' Арматура даёт направленный вклад с вычитанием объёма замещённого бетона.
' Прочность и история материала определяются отдельно.

Private Const PLANE_DOF As Long = 4
Private Const STRESS_COMPONENTS As Long = 6
Private Const SHAPE_COUNT As Long = 16
Private Const GAUGES As Long = 4
Private Const LOCAL_DOF As Long = PLANE_DOF + 3 * SHAPE_COUNT

Private mQuadrature As CCoupledSectionQuadrature
Private mDof As Long
Private mGradientBase As Long
Private mGaugeBase As Long
Private mPointCount As Long
Private mConcreteCount As Long
Private mSteelModulus() As Double, mSteelDirection() As Double
Private mSteelClass() As String
Private mTransverseGroup() As Long ' Номер группы в блоке Transverse, ноль у продольных стержней.
Private mSupport() As Long
Private mSectionDof() As Long
Private mX() As Double, mY() As Double, mWeight() As Double
Private mFirstComponent(1 To LOCAL_DOF) As Long, mSecondComponent(1 To LOCAL_DOF) As Long
Private mOperators() As Double ' Не более двух ненулевых коэффициентов B на локальную степень свободы.
Private mShape() As Double ' Значение функции перемещений для правой части межволоконного равновесия.
Private mElastic(1 To 6, 1 To 6) As Double
Private mMatrix() As Double
Private mSystem As CCoupledLinearSystem
Private mLength As Double, mArea As Double
Private mMeta As CResultMeta

' Подготавливает упругую систему и её геометрические операторы один раз.
' Координаты, обобщённые перемещения и силы нормированы по длине и площади.
Public Sub PrepareModel(ByVal section As CSectionModel, ByVal modulus As Double, ByVal poisson As Double, _
        ByVal spansX As Long, ByVal spansY As Long, ByVal pointsPerSpan As Long, _
        Optional ByVal transverse As CTransverseRebarLayout = Nothing, _
        Optional ByVal steelModulus As Double = 0#, Optional ByVal steelModels As Object = Nothing, _
        Optional ByVal importedGrade As String = vbNullString)
    Set mMeta = New CResultMeta
    mArea = 0#
    If modulus <= 0# Or poisson <= -1# Or poisson >= 0.5 Then
        Err.Raise vbObjectError + 1776, "CCoupledElasticState", _
            "Для упругого расчёта задайте E больше нуля и коэффициент Пуассона между -1 и 0,5."
    End If
    Set mQuadrature = New CCoupledSectionQuadrature
    mQuadrature.Initialize section, spansX, spansY
    mLength = mQuadrature.Frame.LengthScale
    CoupledElasticTensor modulus, poisson, mElastic
    Dim indices() As Long, dx() As Double, dy() As Double
    mQuadrature.Prepare pointsPerSpan, mX, mY, mWeight, indices, mShape, dx, dy
    mConcreteCount = UBound(mWeight)
    mPointCount = mConcreteCount
    PrepareReinforcement section, transverse, steelModulus, steelModels, importedGrade, indices, dx, dy
    PrepareOperators indices, dx, dy
    Assemble
    Set mSystem = New CCoupledLinearSystem
    If Not mSystem.Factor(mMatrix, mDof) Then
        mMeta.SetSolverFailure mSystem.FailureCode, rkDirectState, _
            "Не удалось подготовить линейную систему связанного сечения.", True
    End If
End Sub

' Решает при заданных усилиях относительно численного начала.
' Усилия и напряжения записывает после проверки невязки исходных уравнений.
Public Function Solve(ByVal n As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal torque As Double, ByVal qx As Double, ByVal qy As Double, _
        ByRef deformation() As Double, ByRef recovered() As Double, _
        ByRef pointStress() As Double, Optional ByRef solvedField As Variant) As Boolean
    mMeta.Clear
    Erase deformation: Erase recovered: Erase pointStress
    If Not IsMissing(solvedField) Then solvedField = Empty
    Dim rhs() As Double, applied() As Double, i As Long
    LoadEquations n, mx, my, torque, qx, qy, rhs
    applied = rhs
    If Not mSystem.SolveFactored(rhs) Then
        mMeta.SetSolverFailure mSystem.FailureCode, rkDirectState, _
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
    If Not IsMissing(solvedField) Then solvedField = rhs ' Передаётся только подтверждённое поле того же сечения.
    mMeta.SetResult rsSuccess, rcNone, rkDirectState
    Solve = True
End Function

' Возвращает координаты точки в исходных осях и её вес, мм².
' После бетонных точек идут поправки замещения бетона объёмом арматуры.
Friend Sub Point(ByVal index As Long, ByRef x As Double, ByRef y As Double, ByRef area As Double)
    x = mX(index) * mLength + mQuadrature.Frame.OriginX
    y = mY(index) * mLength + mQuadrature.Frame.OriginY
    area = mWeight(index) * mArea
End Sub

' Возвращает исходную группу поперечной арматуры для материальной точки.
Friend Property Get TransverseGroupAt(ByVal point As Long) As Long
    TransverseGroupAt = mTransverseGroup(point)
End Property

' Возвращает структурированную причину результата в общем формате проекта.
Friend Property Get ResultMeta() As CResultMeta
    Set ResultMeta = mMeta.Clone
End Property

' Добавляет точки продольных стержней из CSectionModel и точки хомутов.
' Вес каждой точки стали одновременно вычитает бетон в той же кинематике.
Private Sub PrepareReinforcement(ByVal section As CSectionModel, ByVal transverse As CTransverseRebarLayout, _
        ByVal steelModulus As Double, ByVal steelModels As Object, ByVal importedGrade As String, _
        ByRef indices() As Long, ByRef dx() As Double, ByRef dy() As Double)
    Erase mSteelModulus: Erase mSteelDirection: Erase mSteelClass: Erase mTransverseGroup
    Dim bars As New CTransverseRebarQuadrature
    If Not transverse Is Nothing Then bars.Prepare transverse, mQuadrature
    Dim last As Long, i As Long, point As Long, k As Long
    last = mConcreteCount + section.RebarCount + bars.Count
    If last = mConcreteCount Then Exit Sub
    If steelModels Is Nothing And steelModulus <= 0# Then Err.Raise vbObjectError + 1779, "CCoupledElasticState", _
        "Для упругой проверки армированного сечения задайте положительный модуль стали."
    ReDim Preserve mX(1 To last): ReDim Preserve mY(1 To last): ReDim Preserve mWeight(1 To last)
    ReDim Preserve indices(1 To SHAPE_COUNT, 1 To last): ReDim Preserve mShape(1 To SHAPE_COUNT, 1 To last)
    ReDim Preserve dx(1 To SHAPE_COUNT, 1 To last): ReDim Preserve dy(1 To SHAPE_COUNT, 1 To last)
    ReDim mSteelModulus(mConcreteCount + 1 To last)
    ReDim mSteelClass(mConcreteCount + 1 To last)
    ReDim mTransverseGroup(mConcreteCount + 1 To last)
    ReDim mSteelDirection(1 To STRESS_COMPONENTS, mConcreteCount + 1 To last)

    Dim frame As CSectionNumericalFrame, region As CConcreteRegion, query As New CSectionGeometryQuery
    Set frame = mQuadrature.Frame
    Set region = mQuadrature.Region
    Dim x As Double, y As Double, area As Double, diameter As Double, gradeModel As CCoupledSteelModel
    Dim txx As Double, tyy As Double, txy As Double, group As Long
    Dim localIndices(1 To SHAPE_COUNT) As Long, localValues(1 To SHAPE_COUNT) As Double
    Dim localX(1 To SHAPE_COUNT) As Double, localY(1 To SHAPE_COUNT) As Double
    For point = mConcreteCount + 1 To last
        i = point - mConcreteCount
        If i <= section.RebarCount Then
            section.RebarElementForSolver i, x, y, area
            diameter = section.RebarDiameter(i)
            If Not query.ContainsPoint(region, x, y, False) Or _
                    query.BoundaryDistance(region, x, y) < diameter / 2# - GEOM_BOUNDARY_TOLERANCE Then
                Err.Raise vbObjectError + 1779, "CCoupledElasticState", _
                    "Продольный стержень " & section.RebarID(i) & " не помещается в области бетона."
            End If
            x = x - frame.OriginX: y = y - frame.OriginY
            mSteelDirection(3, point) = 1#
            mSteelClass(point) = section.RebarSteelClassForMaterial(i, importedGrade)
        Else
            bars.Point i - section.RebarCount, x, y, area, txx, tyy, txy, group
            mSteelDirection(1, point) = txx
            mSteelDirection(2, point) = tyy
            mSteelDirection(4, point) = txy
            mSteelClass(point) = transverse.SteelClass(group)
            mTransverseGroup(point) = group
        End If
        mX(point) = x: mY(point) = y: mWeight(point) = area
        If steelModels Is Nothing Then
            mSteelModulus(point) = steelModulus
        Else
            If Not steelModels.Exists(mSteelClass(point)) Then Err.Raise vbObjectError + 1779, "CCoupledElasticState", _
                "Для класса «" & mSteelClass(point) & "» не назначена модель арматуры."
            Set gradeModel = steelModels.Item(mSteelClass(point))
            mSteelModulus(point) = gradeModel.ElasticModulus
            If mSteelModulus(point) <= 0# Then Err.Raise vbObjectError + 1779, "CCoupledElasticState", _
                "Для класса «" & mSteelClass(point) & "» не задан положительный модуль Es."
        End If
        mQuadrature.EvaluateShape x, y, localIndices, localValues, localX, localY
        For k = 1 To SHAPE_COUNT
            indices(k, point) = localIndices(k)
            mShape(k, point) = localValues(k)
            dx(k, point) = localX(k)
            dy(k, point) = localY(k)
        Next k
    Next point
    mPointCount = last
End Sub

' Подставляет в рабочую матрицу закон Гука либо приращение материала
' от замены объёма бетона одноосной арматурой. Бетон учитывается один раз.
Private Sub PointTangent(ByVal point As Long, ByRef tangent() As Double)
    Dim j As Long, k As Long
    If point <= mConcreteCount Then
        For j = 1 To STRESS_COMPONENTS
            For k = 1 To STRESS_COMPONENTS
                tangent(j, k) = mElastic(j, k)
            Next k
        Next j
    Else
        For j = 1 To STRESS_COMPONENTS
            For k = 1 To STRESS_COMPONENTS
                tangent(j, k) = mSteelModulus(point) * mSteelDirection(j, point) * mSteelDirection(k, point) - mElastic(j, k)
            Next k
        Next j
    End If
End Sub

' Нумерует базис с ненулевой опорой и сохраняет его локальные производные.
' Физические градиенты переводятся к коэффициентам u/L, v/L и w/L.
Private Sub PrepareOperators(ByRef indices() As Long, ByRef dx() As Double, ByRef dy() As Double)
    ReDim mSupport(1 To mQuadrature.BasisCount)
    Dim i As Long, k As Long, globalIndex As Long, active As Long
    For i = 1 To mPointCount
        If i <= mConcreteCount Then mArea = mArea + mWeight(i)
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

    ReDim mSectionDof(1 To LOCAL_DOF, 1 To mPointCount)
    ReDim mOperators(1 To 2, 1 To LOCAL_DOF, 1 To mPointCount)
    Dim u As Long, v As Long, w As Long
    Dim gX As Double, gY As Double
    For k = 1 To 3
        mFirstComponent(k) = 3: mSecondComponent(k) = 4 ' Продольная деформация. Второй коэффициент равен нулю.
    Next k
    mFirstComponent(4) = 5: mSecondComponent(4) = 6 ' Кручение входит в два продольных сдвига.
    For u = PLANE_DOF + 1 To LOCAL_DOF Step 3
        mFirstComponent(u) = 1: mSecondComponent(u) = 4
        mFirstComponent(u + 1) = 2: mSecondComponent(u + 1) = 4
        mFirstComponent(u + 2) = 5: mSecondComponent(u + 2) = 6
    Next u
    For i = 1 To mPointCount
        mWeight(i) = mWeight(i) / mArea
        mX(i) = mX(i) / mLength
        mY(i) = mY(i) / mLength
        For k = 1 To 4
            mSectionDof(k, i) = k
        Next k
        mOperators(1, 1, i) = 1#
        mOperators(1, 2, i) = mY(i)
        mOperators(1, 3, i) = mX(i)
        mOperators(1, 4, i) = -mY(i)
        mOperators(2, 4, i) = mX(i)
        For k = 1 To SHAPE_COUNT
            globalIndex = mSupport(indices(k, i))
            If globalIndex > 0 Then
                u = PLANE_DOF + 3 * k - 2
                v = u + 1: w = u + 2
                mSectionDof(u, i) = PLANE_DOF + 3 * globalIndex - 2
                mSectionDof(v, i) = mSectionDof(u, i) + 1
                mSectionDof(w, i) = mSectionDof(u, i) + 2
                gX = dx(k, i) * mLength: gY = dy(k, i) * mLength
                mOperators(1, u, i) = gX: mOperators(2, u, i) = gY
                mOperators(1, v, i) = gY: mOperators(2, v, i) = gX
                mOperators(1, w, i) = gX: mOperators(2, w, i) = gY
            End If
        Next k
    Next i
End Sub

' Собирает интегралы B^T D B, уравнения производных усилий и четыре условия
' ортогональности. Члены σ'zx, σ'zy, σ'zz берутся из D Bp d, где d = L dq/dz.
Private Sub Assemble()
    ReDim mMatrix(1 To mDof, 1 To mDof)
    Dim point As Long, a As Long, degree As Long, axis As Long, shapeValue As Double
    Dim tangent(1 To STRESS_COMPONENTS, 1 To STRESS_COMPONENTS) As Double
    For point = 1 To mPointCount
        PointTangent point, tangent
        AddPointTangent point, tangent, mMatrix
        If point <= mConcreteCount Then
            For a = PLANE_DOF + 1 To UBound(mSectionDof, 1)
                degree = mSectionDof(a, point)
                If degree > 0 Then
                    axis = (a - PLANE_DOF - 1) Mod 3 + 1
                    shapeValue = mShape((a - PLANE_DOF - 1) \ 3 + 1, point) * mWeight(point)
                    AddGauge degree, mGaugeBase + axis, shapeValue
                    If axis = 1 Then AddGauge degree, mGaugeBase + 4, -mY(point) * shapeValue
                    If axis = 2 Then AddGauge degree, mGaugeBase + 4, mX(point) * shapeValue
                End If
            Next a
        End If
    Next point
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
    Dim point As Long, j As Long, k As Long
    Dim strain(1 To STRESS_COMPONENTS) As Double, localStress(1 To STRESS_COMPONENTS) As Double
    Dim tangent(1 To STRESS_COMPONENTS, 1 To STRESS_COMPONENTS) As Double
    For point = 1 To mPointCount
        PointTangent point, tangent
        PointStrain point, solution, strain
        For j = 1 To STRESS_COMPONENTS
            localStress(j) = 0#
            For k = 1 To STRESS_COMPONENTS
                localStress(j) = localStress(j) + tangent(j, k) * strain(k)
            Next k
            stress(j, point) = localStress(j)
        Next j
        AccumulateResultants point, localStress, recovered
    Next point
End Sub

' Возвращает полную размерность уравнений до конденсации.
Friend Property Get DegreeCount() As Long
    DegreeCount = mDof
End Property

' Возвращает число точек интегрирования бетона и поправок от стали.
Friend Property Get PointCount() As Long
    PointCount = mPointCount
End Property

' Возвращает число точек исходного бетона до поправок замещения.
Friend Property Get ConcretePointCount() As Long
    ConcretePointCount = mConcreteCount
End Property

' Готовит правые части равновесия в масштабе напряжений.
' Моменты заданы относительно подготовленного численного начала.
Friend Sub LoadEquations(ByVal force As Double, ByVal mx As Double, ByVal my As Double, _
        ByVal torque As Double, ByVal qx As Double, ByVal qy As Double, ByRef external() As Double)
    ReDim external(1 To mDof)
    external(1) = force / mArea
    external(2) = mx / mArea / mLength
    external(3) = my / mArea / mLength
    external(4) = torque / mArea / mLength
    external(mGradientBase + 2) = qy / mArea
    external(mGradientBase + 3) = qx / mArea
End Sub

' Проверяет, нужны ли касательные столбцы для продольного градиента напряжений.
Friend Function HasAxialGradient(ByRef field() As Double) As Boolean
    Dim i As Long
    For i = mGradientBase + 1 To mGaugeBase
        If field(i) <> 0# Then
            HasAxialGradient = True
            Exit Function
        End If
    Next i
End Function

' Возвращает Bp·d для данной точки в безразмерных координатах модели.
' Из кинематики продольная производная ненулевая только в zz, xz и yz.
Friend Sub AxialGradientAtPoint(ByVal point As Long, ByRef field() As Double, ByRef gradient() As Double)
    gradient(3) = field(mGradientBase + 1) + _
        mY(point) * field(mGradientBase + 2) + mX(point) * field(mGradientBase + 3)
    gradient(5) = -mY(point) * field(mGradientBase + 4)
    gradient(6) = mX(point) * field(mGradientBase + 4)
End Sub

' Готовит линейные меры средних инженерных сдвигов по площади бетона.
' Вес бетона относится к полной области, включая места расположения стержней.
Friend Sub MeanShearProjections(ByRef alongX() As Double, ByRef alongY() As Double)
    ReDim alongX(1 To mDof), alongY(1 To mDof)
    Dim point As Long, localIndex As Long, degree As Long, weight As Double
    For point = 1 To mConcreteCount
        weight = mWeight(point)
        For localIndex = PLANE_DOF To LOCAL_DOF Step 3 ' Кручение и продольные перемещения w.
            degree = mSectionDof(localIndex, point)
            If degree > 0 Then
                alongX(degree) = alongX(degree) + weight * mOperators(1, localIndex, point)
                alongY(degree) = alongY(degree) + weight * mOperators(2, localIndex, point)
            End If
        Next localIndex
    Next point
End Sub

' Деформации материальной точки от общих деформаций и внутренних полей.
' Порядок εxx, εyy, εzz, γxy, γxz, γyz. Вызывающий переиспользует массив.
Friend Sub PointStrain(ByVal point As Long, ByRef field() As Double, ByRef strain() As Double)
    Dim localIndex As Long, degree As Long, j As Long, first As Long, second As Long, coefficient As Double
    For j = 1 To STRESS_COMPONENTS
        strain(j) = 0#
    Next j
    For localIndex = 1 To LOCAL_DOF
        degree = mSectionDof(localIndex, point)
        If degree > 0 Then
            coefficient = field(degree)
            first = mFirstComponent(localIndex): second = mSecondComponent(localIndex)
            strain(first) = strain(first) + mOperators(1, localIndex, point) * coefficient
            strain(second) = strain(second) + mOperators(2, localIndex, point) * coefficient
        End If
    Next localIndex
End Sub

' Снимает продольное перемещение w в мм в исходных координатах сечения.
' В поле передаются общие, внутренние и служебные степени свободы этой сетки.
Friend Function WarpingAt(ByVal x As Double, ByVal y As Double, ByRef field() As Double) As Double
    Dim fieldStrain(1 To STRESS_COMPONENTS) As Double, warping As Double
    KinematicsAt x, y, field, fieldStrain, warping
    WarpingAt = warping
End Function

' Определяет деформации и продольное перемещение в исходной точке x,y.
' Использует те же функции и нормирование, что операторы материальных точек.
Friend Sub KinematicsAt(ByVal x As Double, ByVal y As Double, ByRef field() As Double, _
        ByRef strain() As Double, ByRef warping As Double)
    Dim indices(1 To SHAPE_COUNT) As Long, values(1 To SHAPE_COUNT) As Double
    Dim dx(1 To SHAPE_COUNT) As Double, dy(1 To SHAPE_COUNT) As Double
    x = x - mQuadrature.Frame.OriginX: y = y - mQuadrature.Frame.OriginY
    mQuadrature.EvaluateShape x, y, indices, values, dx, dy
    strain(1) = 0#: strain(2) = 0#: strain(4) = 0#: warping = 0#
    strain(3) = field(1) + (field(2) * y + field(3) * x) / mLength
    strain(5) = -field(4) * y / mLength: strain(6) = field(4) * x / mLength
    Dim k As Long, active As Long, u As Double, v As Double, w As Double, gx As Double, gy As Double
    For k = 1 To SHAPE_COUNT
        active = mSupport(indices(k))
        If active > 0 Then
            active = PLANE_DOF + 3 * active
            u = field(active - 2): v = field(active - 1): w = field(active)
            gx = dx(k) * mLength: gy = dy(k) * mLength
            strain(1) = strain(1) + gx * u
            strain(2) = strain(2) + gy * v
            strain(4) = strain(4) + gy * u + gx * v
            strain(5) = strain(5) + gx * w
            strain(6) = strain(6) + gy * w
            warping = warping + values(k) * w
        End If
    Next k
    warping = warping * mLength ' Коэффициенты внутренних перемещений нормированы по габариту.
End Sub

' Возвращает обобщённые деформации в начале координат модели и закручивание.
Friend Sub GeneralizedDeformation(ByRef field() As Double, ByRef epsilon0 As Double, ByRef kappaX As Double, _
        ByRef kappaY As Double, ByRef twist As Double)
    kappaX = field(2) / mLength: kappaY = field(3) / mLength: twist = field(4) / mLength
    epsilon0 = mQuadrature.Frame.OriginalStrain(field(1), kappaX, kappaY)
End Sub

' Возвращает класс материала конкретного продольного или поперечного стержня.
Friend Property Get SteelClassAt(ByVal point As Long) As String
    SteelClassAt = mSteelClass(point)
End Property

' Возвращает проекцию деформации на ось стержня. Точка находится после бетона.
Friend Sub SteelProjection(ByVal point As Long, ByRef direction() As Double)
    Dim j As Long
    For j = 1 To STRESS_COMPONENTS
        direction(j) = mSteelDirection(j, point)
    Next j
End Sub

' Готовит буфер нелинейных уравнений и добавляет четыре точных ограничения.
' Материальные вклады накапливаются вызовами AddPointEquilibrium.
Friend Sub StartEquations(ByRef field() As Double, ByRef external() As Double, ByRef residual() As Double, _
        ByRef linearized() As Double, ByVal assembleMatrix As Boolean)
    ReDim residual(1 To mDof)
    If assembleMatrix Then ReDim linearized(1 To mDof, 1 To mDof)
    Dim row As Long, gauge As Long, value As Double
    For row = 1 To mDof
        residual(row) = -external(row)
    Next row
    For gauge = mGaugeBase + 1 To mDof
        For row = 1 To mGaugeBase
            value = mMatrix(gauge, row)
            If value <> 0# Then
                residual(row) = residual(row) + value * field(gauge)
                residual(gauge) = residual(gauge) + value * field(row)
                If assembleMatrix Then
                    linearized(gauge, row) = value
                    linearized(row, gauge) = value
                End If
            End If
        Next row
    Next gauge
End Sub

' Добавляет слабое равновесие для напряжения и продольного градиента D Bp d.
' D и история материала фиксируются на данной линеаризации. После каждой
' пробы вызывающий пересобирает невязку по фактическим напряжениям и D.
Friend Sub AddPointEquilibrium(ByVal point As Long, ByRef field() As Double, _
        ByRef stress() As Double, ByRef tangent() As Double, ByRef residual() As Double, _
        ByRef linearized() As Double, ByVal assembleMatrix As Boolean, ByRef resultants() As Double, _
        Optional ByVal useAxialGradient As Boolean = True)
    Dim axialGradient(1 To STRESS_COMPONENTS) As Double
    Dim weightedGradient(1 To STRESS_COMPONENTS) As Double
    Dim weightedStress(1 To STRESS_COMPONENTS) As Double
    Dim a As Long, j As Long, row As Long, first As Long, second As Long
    Dim weight As Double, value As Double, traction As Long, shapeValue As Double
    Dim coeffFirst As Double, coeffSecond As Double
    weight = mWeight(point)
    If useAxialGradient Then AxialGradientAtPoint point, field, axialGradient
    For j = 1 To STRESS_COMPONENTS
        weightedStress(j) = stress(j) * weight
        If useAxialGradient Then _
            weightedGradient(j) = weight * (tangent(j, 3) * axialGradient(3) + tangent(j, 5) * axialGradient(5) + tangent(j, 6) * axialGradient(6))
    Next j

    For a = 1 To LOCAL_DOF
        row = mSectionDof(a, point)
        If row > 0 Then
            first = mFirstComponent(a): second = mSecondComponent(a)
            coeffFirst = mOperators(1, a, point): coeffSecond = mOperators(2, a, point)
            value = coeffFirst * weightedStress(first) + coeffSecond * weightedStress(second)
            If a <= PLANE_DOF Then
                residual(mGradientBase + a) = residual(mGradientBase + a) + _
                    coeffFirst * weightedGradient(first) + coeffSecond * weightedGradient(second)
            Else
                traction = TractionComponent(a)
                shapeValue = mShape((a - PLANE_DOF - 1) \ 3 + 1, point)
                value = value - shapeValue * weightedGradient(traction)
            End If
            residual(row) = residual(row) + value
        End If
    Next a
    AccumulateResultants point, stress, resultants
    If assembleMatrix Then AddPointTangent point, tangent, linearized
End Sub

' Добавляет зависимость D Bp d от текущей деформации B q.
' Столбец первичной матрицы относится к q, а строки содержат обе проекции
' градиента напряжений, включая слабое равновесие внутренних перемещений.
Friend Sub AddPointGradientTangent(ByVal point As Long, ByRef stressGradientDerivative() As Double, _
        ByRef linearized() As Double)
    Dim weighted(1 To STRESS_COMPONENTS, 1 To LOCAL_DOF) As Double
    Dim a As Long, b As Long, row As Long, col As Long, first As Long, second As Long
    Dim value As Double, shapeValue As Double, traction As Long, coeffFirst As Double, coeffSecond As Double
    WeightTangentColumns point, stressGradientDerivative, weighted
    For a = 1 To LOCAL_DOF
        row = mSectionDof(a, point)
        If row > 0 Then
            If a <= PLANE_DOF Then
                row = mGradientBase + a
                first = mFirstComponent(a): second = mSecondComponent(a)
                coeffFirst = mOperators(1, a, point): coeffSecond = mOperators(2, a, point)
                For b = 1 To LOCAL_DOF
                    col = mSectionDof(b, point)
                    If col > 0 Then
                        value = coeffFirst * weighted(first, b) + coeffSecond * weighted(second, b)
                        linearized(col, row) = linearized(col, row) + value
                    End If
                Next b
            Else
                traction = TractionComponent(a)
                shapeValue = mShape((a - PLANE_DOF - 1) \ 3 + 1, point)
                For b = 1 To LOCAL_DOF
                    col = mSectionDof(b, point)
                    If col > 0 Then linearized(col, row) = linearized(col, row) - shapeValue * weighted(traction, b)
                Next b
            End If
        End If
    Next a
End Sub

' Добавляет блоки Bᵀ D B и производные правой части по градиенту деформации.
' Производные D по деформации при необходимости добавляет AddPointGradientTangent.
Private Sub AddPointTangent(ByVal point As Long, ByRef tangent() As Double, ByRef linearized() As Double)
    Dim weighted(1 To STRESS_COMPONENTS, 1 To LOCAL_DOF) As Double
    Dim a As Long, b As Long, row As Long, col As Long, traction As Long, first As Long, second As Long
    Dim value As Double, shapeValue As Double, coeffFirst As Double, coeffSecond As Double
    WeightTangentColumns point, tangent, weighted
    For a = 1 To LOCAL_DOF
        row = mSectionDof(a, point)
        If row > 0 Then
            first = mFirstComponent(a): second = mSecondComponent(a)
            coeffFirst = mOperators(1, a, point): coeffSecond = mOperators(2, a, point)
            For b = 1 To LOCAL_DOF
                col = mSectionDof(b, point)
                If col > 0 Then
                    value = coeffFirst * weighted(first, b) + coeffSecond * weighted(second, b)
                    linearized(col, row) = linearized(col, row) + value
                    If a <= PLANE_DOF And b <= PLANE_DOF Then
                        linearized(mGradientBase + b, mGradientBase + a) = _
                            linearized(mGradientBase + b, mGradientBase + a) + value
                    End If
                End If
            Next b
            If a > PLANE_DOF Then
                traction = TractionComponent(a)
                shapeValue = mShape((a - PLANE_DOF - 1) \ 3 + 1, point)
                For b = 1 To PLANE_DOF
                    linearized(mGradientBase + b, row) = linearized(mGradientBase + b, row) - _
                        shapeValue * weighted(traction, b)
                Next b
            End If
        End If
    Next a
End Sub

' Умножает D на два сохранённых коэффициента каждой колонки B и вес точки.
' Порядок слагаемых совпадает с порядком компонент тензора напряжений.
Private Sub WeightTangentColumns(ByVal point As Long, ByRef tangent() As Double, ByRef weighted() As Double)
    Dim b As Long, j As Long, first As Long, second As Long
    Dim weight As Double, coeffFirst As Double, coeffSecond As Double
    weight = mWeight(point)
    For b = 1 To LOCAL_DOF
        If mSectionDof(b, point) > 0 Then
            first = mFirstComponent(b): second = mSecondComponent(b)
            coeffFirst = mOperators(1, b, point): coeffSecond = mOperators(2, b, point)
            For j = 1 To STRESS_COMPONENTS
                weighted(j, b) = (tangent(j, first) * coeffFirst + tangent(j, second) * coeffSecond) * weight
            Next j
        End If
    Next b
End Sub

' Возвращает компоненту вектора напряжений на продольной площадке.
Private Function TractionComponent(ByVal localDegree As Long) As Long
    Select Case (localDegree - PLANE_DOF - 1) Mod 3
        Case 0: TractionComponent = 5
        Case 1: TractionComponent = 6
        Case 2: TractionComponent = 3
    End Select
End Function

' Суммирует шесть физических усилий относительно численного начала.
Private Sub AccumulateResultants(ByVal point As Long, ByRef stress() As Double, ByRef resultants() As Double)
    Dim area As Double, momentWeight As Double
    area = mArea * mWeight(point)
    momentWeight = area * mLength
    resultants(1) = resultants(1) + area * stress(3)
    resultants(2) = resultants(2) + momentWeight * stress(3) * mY(point)
    resultants(3) = resultants(3) + momentWeight * stress(3) * mX(point)
    resultants(4) = resultants(4) + momentWeight * (mX(point) * stress(6) - mY(point) * stress(5))
    resultants(5) = resultants(5) + area * stress(5)
    resultants(6) = resultants(6) + area * stress(6)
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

' Проверяет независимые интегралы всех шести напряжений в масштабе МПа.
' Моменты делятся на характерное плечо, все компоненты на бетонную площадь.
Friend Function ResultantResidualNorm(ByRef actual() As Double, ByRef expected() As Double) As Double
    Dim i As Long, difference As Double
    For i = 1 To 6
        difference = Abs(actual(i) - expected(i)) / mArea
        If i >= 2 And i <= 4 Then difference = difference / mLength
        If difference > ResultantResidualNorm Then ResultantResidualNorm = difference
    Next i
End Function
