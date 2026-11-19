VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledConstraintEquations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Исключение одной координаты через нелинейное условие управления
' ==========================================================================
' Первые n-1 уравнений исходного отклика сохраняются. Последняя строка
' заменяется условием однозначности исключённой численной координаты.
' Каждая проба, включая разностную, отображается на заданную поверхность.
' Перед силовой проверкой решение обязательно преобразуется TryMap.
Implements ICoupledEquationSystem

Private mEquations As ICoupledEquationSystem, mConstraint As ICoupledScalarConstraint
Private mCount As Long, mPivot As Long, mProjectionLimit As Long
Private mPivotOrigin As Double, mTolerance As Double
Private mProjectionCalls As Long, mProjectionFailures As Long

' Сохраняет неизменяемый контекст и выбирает ненулевую производную среди полевых координат; последняя координата нагрузки не исключается.
Friend Sub Initialize(ByVal equations As ICoupledEquationSystem, ByVal constraint As ICoupledScalarConstraint, _
        ByRef origin() As Double, ByVal tolerance As Double, ByVal projectionLimit As Long)
    If Not mEquations Is Nothing Then Err.Raise 5, "CCoupledConstraintEquations", "Контекст исключения координаты уже подготовлен."
    If equations Is Nothing Or constraint Is Nothing Or tolerance <= 0# Or projectionLimit < 1 Then _
        Err.Raise 5, "CCoupledConstraintEquations", "Не заданы отклик, материальная координата или точность её удержания."
    If LBound(origin) <> 1 Or UBound(origin) < 2 Then _
        Err.Raise 5, "CCoupledConstraintEquations", "Для исключения нужны полевые координаты и координата нагрузки."
    mCount = UBound(origin)
    Dim value As Double, gradient() As Double, largest As Double, i As Long
    If Not constraint.TryEvaluate(origin, True, value, gradient) Then _
        Err.Raise 5, "CCoupledConstraintEquations", "Не прочитана исходная материальная координата."
    If Not MatchesSize(gradient) Then Err.Raise 5, "CCoupledConstraintEquations", "Размер градиента материальной координаты не совпадает с системой."
    For i = 1 To mCount - 1
        If Abs(gradient(i)) > largest Then largest = Abs(gradient(i)): mPivot = i
    Next i
    If mPivot = 0 Then Err.Raise 5, "CCoupledConstraintEquations", "Материальная координата не имеет ненулевой производной по полю."
    mPivotOrigin = origin(mPivot): mTolerance = tolerance: mProjectionLimit = projectionLimit
    Set mConstraint = constraint: Set mEquations = equations
    mProjectionCalls = 1
End Sub

' Возвращает выбранный индекс; индекс зависит от фактического градиента, а не от типа сечения.
Friend Property Get Pivot() As Long
    Pivot = mPivot
End Property

' Возвращает число чтений скалярного отклика, включая первоначальный выбор координаты.
Friend Property Get ProjectionCalls() As Long
    ProjectionCalls = mProjectionCalls
End Property

' Возвращает число проб, для которых нелинейное отображение не найдено в установленном бюджете.
Friend Property Get ProjectionFailures() As Long
    ProjectionFailures = mProjectionFailures
End Property

' Восстанавливает действительные координаты, не меняя исходный массив; допускается не более projectionLimit поправок и итоговое чтение.
Friend Function TryMap(ByRef coordinates() As Double, ByRef projected() As Double) As Boolean
    If mEquations Is Nothing Then Err.Raise 5, "CCoupledConstraintEquations", "Контекст исключения координаты не подготовлен."
    If Not MatchesSize(coordinates) Then Err.Raise 5, "CCoupledConstraintEquations", "Размер пробы не совпадает с подготовленным контекстом."
    Erase projected
    Dim trial() As Double, gradient() As Double, value As Double, previous As Double, iteration As Long
    trial = coordinates: trial(mPivot) = mPivotOrigin ' Численная координата однозначности не влияет на физическое поле.
    For iteration = 0 To mProjectionLimit
        mProjectionCalls = mProjectionCalls + 1
        If Not mConstraint.TryEvaluate(trial, True, value, gradient) Then Exit For
        If Abs(value) <= mTolerance Then
            projected = trial: TryMap = True
            Exit Function
        End If
        If iteration = mProjectionLimit Then Exit For
        If Not MatchesSize(gradient) Then Err.Raise 5, "CCoupledConstraintEquations", "Размер градиента изменился во время отображения."
        If gradient(mPivot) = 0# Then Exit For
        previous = trial(mPivot): trial(mPivot) = previous - value / gradient(mPivot)
        If trial(mPivot) = previous Then Exit For
    Next iteration
    mProjectionFailures = mProjectionFailures + 1
End Function

' Передаёт исходной системе отображённую пробу; правило цепочки сохраняет первые n-1 строк при столбцовом первом индексе матрицы.
Private Function ICoupledEquationSystem_TryEvaluate(ByRef coordinates() As Double, ByVal withJacobian As Boolean, _
        ByRef residual() As Double, ByRef jacobian() As Double) As Boolean
    Erase residual
    If withJacobian Then Erase jacobian
    Dim projected() As Double
    If Not TryMap(coordinates, projected) Then Exit Function
    If Not mEquations.TryEvaluate(projected, withJacobian, residual, jacobian) Then Exit Function
    If Not MatchesSize(residual) Then Err.Raise 5, "CCoupledConstraintEquations", "Исходный отклик вернул неверный размер невязки."
    residual(mCount) = coordinates(mPivot) - mPivotOrigin
    If withJacobian Then
        If LBound(jacobian, 1) <> 1 Or UBound(jacobian, 1) <> mCount Or _
                LBound(jacobian, 2) <> 1 Or UBound(jacobian, 2) <> mCount Then _
            Err.Raise 5, "CCoupledConstraintEquations", "Исходный отклик вернул неверный размер матрицы."
        Dim value As Double, gradient() As Double, normal() As Double, i As Long, j As Long
        mProjectionCalls = mProjectionCalls + 1
        If Not mConstraint.TryEvaluate(projected, True, value, gradient) Then Exit Function
        If Not MatchesSize(gradient) Then Err.Raise 5, "CCoupledConstraintEquations", "Неверный размер градиента отображения."
        If gradient(mPivot) = 0# Then Exit Function
        ReDim normal(1 To mCount - 1)
        For i = 1 To mCount - 1: normal(i) = jacobian(mPivot, i): Next i
        For j = 1 To mCount
            For i = 1 To mCount - 1
                If j = mPivot Then
                    jacobian(j, i) = 0#
                Else
                    jacobian(j, i) = jacobian(j, i) - normal(i) * gradient(j) / gradient(mPivot)
                End If
            Next i
            jacobian(j, mCount) = 0#
        Next j
        jacobian(mPivot, mCount) = 1#
    End If
    ICoupledEquationSystem_TryEvaluate = True
End Function

' Проверяет внутренний вектор, включая отсутствие невыделенного массива; неверный размер не является отказом прочности.
Private Function MatchesSize(ByRef values() As Double) As Boolean
    On Error GoTo InvalidSize
    MatchesSize = (LBound(values) = 1 And UBound(values) = mCount)
InvalidSize:
End Function
