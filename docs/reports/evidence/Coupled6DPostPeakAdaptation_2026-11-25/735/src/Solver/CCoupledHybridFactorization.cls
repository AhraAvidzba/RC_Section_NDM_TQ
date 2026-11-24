VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledHybridFactorization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' QR, dogleg и регуляризованная поправка для полного равновесия
' ==========================================================================
' Перенос квадратной ветви MINPACK/CMinpack: qrfac с необязательными перестановками,
' qform, dogleg, r1updt, r1mpyq. Масштаб координат задан снаружи (diag=1).
' Источник: literature/Coupled6D/NumericalControl/Minpack, ревизия
' 64c5635addd72fb54f44db8546c191feeb70e338; полная лицензия
' CopyrightMINPACK.txt хранится там же и сопровождает распространение.
' This product includes software developed by the University of Chicago,
' as Operator of Argonne National Laboratory.
' Регуляризованная поправка: Морэ(1977), §3,5, lmpar/qrsolv той же ревизии MINPACK.
' Численная регуляризация не меняет материал и не определяет физическое разрушение.
' Первый индекс матрицы — столбец; R упакована последовательно по строкам.

Private Const MACHINE_EPSILON As Double = 2# ^ (-52)
Private Const INVERSE_MAXIMUM As Double = (2# ^ (-1022)) / 4# ' IEEE Double: округлённое 1/DBL_MAX без потери цифр при экспорте VBE.
Private mCount As Long, mQ() As Double, mR() As Double, mQtf() As Double
Private mColumnOrder() As Long, mPivoted As Boolean

' Вычисляет QR исходной матрицы и Q^T*f, не меняя входные массивы.
' Перестановки разрешены для регуляризованного решения; dogleg и обновление ранга один требуют исходного порядка столбцов.
Friend Sub Factor(ByRef matrix() As Double, ByRef residual() As Double, ByVal count As Long, Optional ByVal pivotColumns As Boolean = False)
    mCount = count: mQ = matrix: mQtf = residual: mPivoted = pivotColumns
    ReDim mR(1 To CLng(CDbl(count) * (count + 1) / 2#))
    Dim diagonal() As Double, work() As Double
    ReDim diagonal(1 To count): ReDim work(1 To count)
    Dim i As Long, j As Long, k As Long, position As Long, total As Double, magnitude As Double, multiplier As Double
    ReDim mColumnOrder(1 To count)
    For i = 1 To count: mColumnOrder(i) = i: Next i
    For j = 1 To count
        If pivotColumns Then PivotLeastSquaresColumns j
        For i = j To count: work(i) = mQ(j, i): Next i
        magnitude = VectorNorm(work, j, count)
        If magnitude <> 0# Then
            If mQ(j, j) < 0# Then magnitude = -magnitude
            For i = j To count: mQ(j, i) = mQ(j, i) / magnitude: Next i
            mQ(j, j) = mQ(j, j) + 1#
            For k = j + 1 To count
                total = 0#
                For i = j To count: total = total + mQ(j, i) * mQ(k, i): Next i
                multiplier = total / mQ(j, j)
                For i = j To count: mQ(k, i) = mQ(k, i) - multiplier * mQ(j, i): Next i
            Next k
        End If
        diagonal(j) = -magnitude
    Next j
    For j = 1 To count
        If mQ(j, j) <> 0# Then
            total = 0#
            For i = j To count: total = total + mQ(j, i) * mQtf(i): Next i
            multiplier = -total / mQ(j, j)
            For i = j To count: mQtf(i) = mQtf(i) + mQ(j, i) * multiplier: Next i
        End If
    Next j
    For i = 1 To count
        For j = i To count
            position = position + 1
            If i = j Then mR(position) = diagonal(i) Else mR(position) = mQ(j, i)
        Next j
    Next i
    FormOrthogonal work
End Sub

' Собирает Q из отражений Хаусхолдера; верхняя часть R уже сохранена отдельно.
Private Sub FormOrthogonal(ByRef work() As Double)
    Dim i As Long, j As Long, k As Long, total As Double, multiplier As Double
    For j = 2 To mCount
        For i = 1 To j - 1: mQ(j, i) = 0#: Next i
    Next j
    For k = mCount To 1 Step -1
        For i = k To mCount: work(i) = mQ(k, i): mQ(k, i) = 0#: Next i
        mQ(k, k) = 1#
        If work(k) <> 0# Then
            For j = k To mCount
                total = 0#
                For i = k To mCount: total = total + mQ(j, i) * work(i): Next i
                multiplier = total / work(k)
                For i = k To mCount: mQ(j, i) = mQ(j, i) - multiplier * work(i): Next i
            Next j
        End If
    Next k
End Sub

' Находит положительный dogleg-вектор R*d=Q^T*f внутри заданного радиуса.
' Решатель меняет знак d для шага поиска корня; здесь сохранён знак MINPACK.
Friend Sub Dogleg(ByVal radius As Double, ByRef direction() As Double)
    If mPivoted Then Err.Raise 5, "CCoupledHybridFactorization", "Для dogleg и секущего обновления требуется QR без перестановок столбцов."
    Dim gradient() As Double, work() As Double
    ReDim direction(1 To mCount): ReDim gradient(1 To mCount): ReDim work(1 To mCount)
    Dim i As Long, j As Long, k As Long, position As Long, diagonal As Long, total As Double, divisor As Double
    diagonal = UBound(mR) + 1
    For k = 1 To mCount
        j = mCount - k + 1: diagonal = diagonal - k: position = diagonal + 1: total = 0#
        For i = j + 1 To mCount: total = total + mR(position) * direction(i): position = position + 1: Next i
        divisor = mR(diagonal)
        If divisor = 0# Then
            position = j
            For i = 1 To j
                divisor = GeomMax(divisor, Abs(mR(position))): position = position + mCount - i
            Next i
            divisor = MACHINE_EPSILON * divisor
            If divisor = 0# Then divisor = MACHINE_EPSILON
        End If
        direction(j) = (mQtf(j) - total) / divisor
    Next k
    Dim newtonNorm As Double: newtonNorm = VectorNorm(direction, 1, mCount)
    If newtonNorm <= radius Then Exit Sub
    position = 1
    For j = 1 To mCount
        For i = j To mCount
            gradient(i) = gradient(i) + mR(position) * mQtf(j): position = position + 1
        Next i
    Next j
    Dim gradientNorm As Double, optimalLength As Double, alpha As Double, magnitude As Double, rhsNorm As Double
    gradientNorm = VectorNorm(gradient, 1, mCount): alpha = radius / newtonNorm
    If gradientNorm <> 0# Then
        For j = 1 To mCount: gradient(j) = gradient(j) / gradientNorm: Next j
        position = 1
        For j = 1 To mCount
            total = 0#
            For i = j To mCount: total = total + mR(position) * gradient(i): position = position + 1: Next i
            work(j) = total
        Next j
        magnitude = VectorNorm(work, 1, mCount)
        If magnitude = 0# Then optimalLength = radius Else optimalLength = gradientNorm / magnitude / magnitude
        alpha = 0#
        If optimalLength < radius Then
            rhsNorm = VectorNorm(mQtf, 1, mCount)
            magnitude = rhsNorm / gradientNorm * (rhsNorm / newtonNorm) * (optimalLength / radius)
            magnitude = magnitude - radius / newtonNorm * (optimalLength / radius) ^ 2 + _
                Sqr((magnitude - radius / newtonNorm) ^ 2 + (1# - (radius / newtonNorm) ^ 2) * (1# - (optimalLength / radius) ^ 2))
            alpha = radius / newtonNorm * (1# - (optimalLength / radius) ^ 2) / magnitude
        End If
    End If
    magnitude = (1# - alpha) * GeomMin(optimalLength, radius)
    For j = 1 To mCount: direction(j) = magnitude * gradient(j) + alpha * direction(j): Next j
End Sub

' Вычисляет Q^T*f+R*step, то есть невязку линейной модели в ортогональных координатах.
Friend Sub ModelResidual(ByRef movement() As Double, ByRef residual() As Double)
    ReDim residual(1 To mCount)
    If mPivoted Then Err.Raise 5, "CCoupledHybridFactorization", "Для dogleg и секущего обновления требуется QR без перестановок столбцов."
    Dim i As Long, j As Long, position As Long, total As Double: position = 1
    For i = 1 To mCount
        total = 0#
        For j = i To mCount: total = total + mR(position) * movement(j): position = position + 1: Next j
        residual(i) = mQtf(i) + total
    Next i
End Sub

' Вносит секущее обновление ранга один; обе системы вращений применяются к Q и Q^T*f.
' Невязка новой точки участвует и при отвергнутом шаге, как в исходном HYBRJ.
Friend Sub Update(ByRef movement() As Double, ByRef residual() As Double, ByRef model() As Double, _
        ByVal movementNorm As Double, ByVal accepted As Boolean)
    If mPivoted Then Err.Raise 5, "CCoupledHybridFactorization", "Для dogleg и секущего обновления требуется QR без перестановок столбцов."
    Dim u() As Double, v() As Double, work() As Double
    ReDim u(1 To mCount): ReDim v(1 To mCount): ReDim work(1 To mCount)
    Dim i As Long, j As Long, total As Double
    For j = 1 To mCount
        total = 0#
        For i = 1 To mCount: total = total + mQ(j, i) * residual(i): Next i
        v(j) = (total - model(j)) / movementNorm: u(j) = movement(j) / movementNorm
        If accepted Then mQtf(j) = total
    Next j
    UpdateTriangle u, v, work
    ApplyRotations v, work
End Sub

' Обновляет упакованную R через эквивалентную нижнюю матрицу R^T исходного r1updt.
' Параметры вращений записываются поверх v и work в формате MINPACK.
Private Sub UpdateTriangle(ByRef u() As Double, ByRef v() As Double, ByRef work() As Double)
    Dim i As Long, j As Long, diagonal As Long, position As Long
    Dim cosine As Double, sine As Double, tau As Double, ratio As Double, temporary As Double
    diagonal = UBound(mR): work(mCount) = mR(diagonal)
    For j = mCount - 1 To 1 Step -1
        diagonal = diagonal - mCount + j - 1: work(j) = 0#
        If v(j) <> 0# Then
            If Abs(v(mCount)) < Abs(v(j)) Then
                ratio = v(mCount) / v(j): sine = 0.5 / Sqr(0.25 + 0.25 * ratio * ratio)
                cosine = sine * ratio: tau = 1#
                If Abs(cosine) > INVERSE_MAXIMUM Then tau = 1# / cosine ' Проверка до деления исключает переполнение Double.
            Else
                ratio = v(j) / v(mCount): cosine = 0.5 / Sqr(0.25 + 0.25 * ratio * ratio)
                sine = cosine * ratio: tau = sine
            End If
            v(mCount) = sine * v(j) + cosine * v(mCount): v(j) = tau
            position = diagonal
            For i = j To mCount
                temporary = cosine * mR(position) - sine * work(i)
                work(i) = sine * mR(position) + cosine * work(i): mR(position) = temporary: position = position + 1
            Next i
        End If
    Next j
    For i = 1 To mCount: work(i) = work(i) + v(mCount) * u(i): Next i
    For j = 1 To mCount - 1
        If work(j) <> 0# Then
            If Abs(mR(diagonal)) < Abs(work(j)) Then
                ratio = mR(diagonal) / work(j): sine = 0.5 / Sqr(0.25 + 0.25 * ratio * ratio)
                cosine = sine * ratio: tau = 1#
                If Abs(cosine) > INVERSE_MAXIMUM Then tau = 1# / cosine
            Else
                ratio = work(j) / mR(diagonal): cosine = 0.5 / Sqr(0.25 + 0.25 * ratio * ratio)
                sine = cosine * ratio: tau = sine
            End If
            position = diagonal
            For i = j To mCount
                temporary = cosine * mR(position) + sine * work(i)
                work(i) = -sine * mR(position) + cosine * work(i): mR(position) = temporary: position = position + 1
            Next i
            work(j) = tau
        End If
        diagonal = diagonal + mCount - j + 1
    Next j
    mR(diagonal) = work(mCount)
End Sub

' Применяет обе последовательности вращений Гивенса к Q и единственной строке Q^T*f.
Private Sub ApplyRotations(ByRef v() As Double, ByRef work() As Double)
    Dim j As Long
    For j = mCount - 1 To 1 Step -1: RotateColumns j, v(j), False: Next j
    For j = 1 To mCount - 1: RotateColumns j, work(j), True: Next j
End Sub

' Восстанавливает одно вращение из компактного параметра и изменяет пару столбцов.
Private Sub RotateColumns(ByVal column As Long, ByVal parameter As Double, ByVal second As Boolean)
    Dim cosine As Double, sine As Double, temporary As Double, i As Long
    If Abs(parameter) > 1# Then
        cosine = 1# / parameter: sine = Sqr(1# - cosine * cosine)
    Else
        sine = parameter: cosine = Sqr(1# - sine * sine)
    End If
    If Not second Then sine = -sine
    For i = 1 To mCount
        temporary = cosine * mQ(column, i) + sine * mQ(mCount, i)
        mQ(mCount, i) = -sine * mQ(column, i) + cosine * mQ(mCount, i): mQ(column, i) = temporary
    Next i
    temporary = cosine * mQtf(column) + sine * mQtf(mCount)
    mQtf(mCount) = -sine * mQtf(column) + cosine * mQtf(mCount): mQtf(column) = temporary
End Sub

' Находит евклидову норму диапазона без разрушительного возведения больших или малых чисел в квадрат.
' Три диапазона и IEEE-границы соответствуют сохранённому CMinpack enorm для Double.
Friend Function VectorNorm(ByRef values() As Double, ByVal first As Long, ByVal last As Long) As Double
    Const DWARF As Double = 1.82691291192569E-153
    Const GIANT As Double = 1.34078079299426E+153
    Dim largeSum As Double, middleSum As Double, smallSum As Double, largest As Double, smallestMaximum As Double
    Dim border As Double, magnitude As Double, ratio As Double, i As Long
    border = GIANT / (last - first + 1)
    For i = first To last
        magnitude = Abs(values(i))
        If magnitude <= DWARF Then
            If magnitude > smallestMaximum Then
                ratio = smallestMaximum / magnitude: smallSum = 1# + smallSum * ratio * ratio: smallestMaximum = magnitude
            ElseIf magnitude <> 0# Then
                ratio = magnitude / smallestMaximum: smallSum = smallSum + ratio * ratio
            End If
        ElseIf magnitude >= border Then
            If magnitude > largest Then
                ratio = largest / magnitude: largeSum = 1# + largeSum * ratio * ratio: largest = magnitude
            Else
                ratio = magnitude / largest: largeSum = largeSum + ratio * ratio
            End If
        Else
            middleSum = middleSum + magnitude * magnitude
        End If
    Next i
    If largeSum <> 0# Then
        VectorNorm = largest * Sqr(largeSum + (middleSum / largest) / largest)
    ElseIf middleSum <> 0# Then
        If middleSum >= smallestMaximum Then
            VectorNorm = Sqr(middleSum * (1# + (smallestMaximum / middleSum) * (smallestMaximum * smallSum)))
        Else
            VectorNorm = Sqr(smallestMaximum * ((middleSum / smallestMaximum) + (smallestMaximum * smallSum)))
        End If
    Else
        VectorNorm = smallestMaximum * Sqr(smallSum)
    End If
End Function

' ======================================================================
' Лицензия первичного MINPACK: сохраняется также в нативном экспорте VBA.
' ======================================================================
' Minpack Copyright Notice (1999) University of Chicago.  All rights reserved
'
' Redistribution and use in source and binary forms, with or
' without modification, are permitted provided that the
' following conditions are met:
'
' 1. Redistributions of source code must retain the above
' copyright notice, this list of conditions and the following
' disclaimer.
'
' 2. Redistributions in binary form must reproduce the above
' copyright notice, this list of conditions and the following
' disclaimer in the documentation and/or other materials
' provided with the distribution.
'
' 3. The end-user documentation included with the
' redistribution, if any, must include the following
' acknowledgment:
'
'    "This product includes software developed by the
'    University of Chicago, as Operator of Argonne National
'    Laboratory.
'
' Alternately, this acknowledgment may appear in the software
' itself, if and wherever such third-party acknowledgments
' normally appear.
'
' 4. WARRANTY DISCLAIMER. THE SOFTWARE IS SUPPLIED "AS IS"
' WITHOUT WARRANTY OF ANY KIND. THE COPYRIGHT HOLDER, THE
' UNITED STATES, THE UNITED STATES DEPARTMENT OF ENERGY, AND
' THEIR EMPLOYEES: (1) DISCLAIM ANY WARRANTIES, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO ANY IMPLIED WARRANTIES
' OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, TITLE
' OR NON-INFRINGEMENT, (2) DO NOT ASSUME ANY LEGAL LIABILITY
' OR RESPONSIBILITY FOR THE ACCURACY, COMPLETENESS, OR
' USEFULNESS OF THE SOFTWARE, (3) DO NOT REPRESENT THAT USE OF
' THE SOFTWARE WOULD NOT INFRINGE PRIVATELY OWNED RIGHTS, (4)
' DO NOT WARRANT THAT THE SOFTWARE WILL FUNCTION
' UNINTERRUPTED, THAT IT IS ERROR-FREE OR THAT ANY ERRORS WILL
' BE CORRECTED.
'
' 5. LIMITATION OF LIABILITY. IN NO EVENT WILL THE COPYRIGHT
' HOLDER, THE UNITED STATES, THE UNITED STATES DEPARTMENT OF
' ENERGY, OR THEIR EMPLOYEES: BE LIABLE FOR ANY INDIRECT,
' INCIDENTAL, CONSEQUENTIAL, SPECIAL OR PUNITIVE DAMAGES OF
' ANY KIND OR NATURE, INCLUDING BUT NOT LIMITED TO LOSS OF
' PROFITS OR LOSS OF DATA, FOR ANY REASON WHATSOEVER, WHETHER
' SUCH LIABILITY IS ASSERTED ON THE BASIS OF CONTRACT, TORT
' (INCLUDING NEGLIGENCE OR STRICT LIABILITY), OR OTHERWISE,
' EVEN IF ANY OF SAID PARTIES HAS BEEN WARNED OF THE
' POSSIBILITY OF SUCH LOSS OR DAMAGES.

' Выбирает наибольший оставшийся столбец QR по норме его необнулённого хвоста; предыдущие отражения сохраняются.
Private Sub PivotLeastSquaresColumns(ByVal first As Long)
    Dim column As Long, row As Long, selected As Long, magnitude As Double, largest As Double
    Dim work() As Double: ReDim work(1 To mCount)
    largest = -1#: selected = first
    For column = first To mCount
        For row = first To mCount: work(row) = mQ(column, row): Next row
        magnitude = VectorNorm(work, first, mCount)
        If magnitude > largest Then largest = magnitude: selected = column
    Next column
    If selected = first Then Exit Sub
    Dim value As Double, order As Long
    For row = 1 To mCount
        value = mQ(first, row): mQ(first, row) = mQ(selected, row): mQ(selected, row) = value
    Next row
    order = mColumnOrder(first): mColumnOrder(first) = mColumnOrder(selected): mColumnOrder(selected) = order
End Sub

' Решает min(||J*p+f||^2+parameter*||p||^2) через готовый QR и добавленные диагональные строки; parameter — численная регуляризация.
' Положительное направление p уже включает знак минус правой части; исходные R,Q^T*f и порядок столбцов не меняются.
Friend Function TryDampedStep(ByVal parameter As Double, ByRef direction() As Double, ByRef normDerivative As Double) As Boolean
    Erase direction: normDerivative = 0#
    If mCount < 1 Or parameter < 0# Then Err.Raise 5, "CCoupledHybridFactorization", "Для регуляризованной пробы нужны QR и неотрицательный параметр."
    On Error GoTo ArithmeticFailure
    Dim triangle() As Double, rhs() As Double, workRow() As Double, ordered() As Double, inverse() As Double
    ReDim triangle(1 To mCount, 1 To mCount): ReDim rhs(1 To mCount): ReDim workRow(1 To mCount)
    ReDim ordered(1 To mCount): ReDim inverse(1 To mCount)
    Dim row As Long, column As Long, position As Long, first As Long, k As Long, largest As Double
    For row = 1 To mCount
        rhs(row) = -mQtf(row)
        For column = row To mCount
            position = position + 1: triangle(column, row) = mR(position)
            largest = GeomMax(largest, Abs(mR(position)))
        Next column
    Next row
    If parameter = 0# Then
        For row = 1 To mCount
            If Abs(triangle(row, row)) <= MACHINE_EPSILON * mCount * largest Then Exit Function
        Next row
    Else
        Dim diagonal As Double, extra As Double, length As Double, cosine As Double, sine As Double, value As Double
        diagonal = Sqr(parameter)
        For first = 1 To mCount
            For column = first To mCount: workRow(column) = 0#: Next column
            workRow(first) = diagonal: extra = 0#
            For k = first To mCount
                If workRow(k) <> 0# Then
                    length = HypotenuseLeastSquares(triangle(k, k), workRow(k))
                    If length = 0# Then Exit Function
                    cosine = triangle(k, k) / length: sine = workRow(k) / length
                    triangle(k, k) = length
                    For column = k + 1 To mCount
                        value = cosine * triangle(column, k) + sine * workRow(column)
                        workRow(column) = -sine * triangle(column, k) + cosine * workRow(column)
                        triangle(column, k) = value
                    Next column
                    value = cosine * rhs(k) + sine * extra
                    extra = -sine * rhs(k) + cosine * extra: rhs(k) = value
                End If
            Next k
        Next first
    End If
    Dim total As Double
    For row = mCount To 1 Step -1
        If triangle(row, row) = 0# Then Exit Function
        total = rhs(row)
        For column = row + 1 To mCount: total = total - triangle(column, row) * ordered(column): Next column
        ordered(row) = total / triangle(row, row)
    Next row
    Dim norm As Double, inverseNorm As Double
    norm = VectorNorm(ordered, 1, mCount)
    If norm > 0# Then
        For row = 1 To mCount
            total = ordered(row)
            For column = 1 To row - 1: total = total - triangle(row, column) * inverse(column): Next column
            inverse(row) = total / triangle(row, row)
        Next row
        inverseNorm = VectorNorm(inverse, 1, mCount)
        normDerivative = -(inverseNorm / norm) * inverseNorm ' Производная ||p|| по параметру регуляризации, без формирования J^T*J.
    End If
    ReDim direction(1 To mCount)
    For column = 1 To mCount: direction(mColumnOrder(column)) = ordered(column): Next column
    TryDampedStep = True
    Exit Function
ArithmeticFailure:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
    Erase direction: normDerivative = 0#
End Function

' Выбирает регуляризацию для заданного радиуса по Морэ(1977),§5; масштаб координат должен быть включён в исходную матрицу до QR.
' Не определяет равновесие нелинейной задачи: ограничивает только одну линейную поправку в доверительной области.
Friend Function TryLeastSquaresRadius(ByVal radius As Double, ByRef parameter As Double, _
        ByRef direction() As Double, ByRef iterations As Long) As Boolean
    Erase direction: iterations = 0
    If mCount < 1 Or radius <= 0# Or parameter < 0# Then Err.Raise 5, "CCoupledHybridFactorization", "Для доверительной пробы нужны QR, положительный радиус и неотрицательная регуляризация."
    On Error GoTo ArithmeticFailure
    Dim gradient() As Double: ReDim gradient(1 To mCount)
    Dim row As Long, column As Long, position As Long
    For row = 1 To mCount
        For column = row To mCount
            position = position + 1: gradient(column) = gradient(column) + mR(position) * mQtf(row)
        Next column
    Next row
    Dim upper As Double: upper = VectorNorm(gradient, 1, mCount) / radius
    If upper = 0# Then
        ReDim direction(1 To mCount): parameter = 0#: TryLeastSquaresRadius = True
        Exit Function
    End If
    Dim derivative As Double, norm As Double, phi As Double, lower As Double, fullRank As Boolean
    fullRank = TryDampedStep(0#, direction, derivative)
    If Not fullRank Then
        If Not TryUndampedRankStep(direction) Then Exit Function
        norm = VectorNorm(direction, 1, mCount)
        If norm <= radius Then parameter = 0#: TryLeastSquaresRadius = True: Exit Function
    End If
    If fullRank Then
        norm = VectorNorm(direction, 1, mCount)
        If norm <= radius Then parameter = 0#: TryLeastSquaresRadius = True: Exit Function
        If derivative >= 0# Then Exit Function
        lower = -(norm - radius) / derivative
    End If
    If Not fullRank And parameter = 0# Then parameter = GeomMax(0.001 * upper, Sqr(lower) * Sqr(upper))
    Dim ratio As Double, attempt As Long
    For attempt = 1 To 10
        iterations = attempt
        If parameter < lower Or parameter > upper Then parameter = GeomMax(0.001 * upper, Sqr(lower) * Sqr(upper))
        If Not TryDampedStep(parameter, direction, derivative) Then Exit Function
        Dim previousPhi As Double: previousPhi = phi
        norm = VectorNorm(direction, 1, mCount): phi = norm - radius
        If lower = 0# And phi <= previousPhi And previousPhi < 0# Then TryLeastSquaresRadius = True: Exit Function
        If norm = 0# Or derivative >= 0# Then Exit Function
        If phi < 0# Then upper = parameter
        ratio = phi / derivative
        lower = GeomMax(lower, parameter - ratio)
        parameter = parameter - (norm / radius) * ratio
        If Abs(phi) < 0.01 * radius Then Exit For
    Next attempt
    If parameter < 0# Then Exit Function
    If Not TryDampedStep(parameter, direction, derivative) Then Exit Function
    norm = VectorNorm(direction, 1, mCount)
    If norm = 0# Then Exit Function
    For column = 1 To mCount: direction(column) = direction(column) * (radius / norm): Next column
    TryLeastSquaresRadius = True
    Exit Function
ArithmeticFailure:
    If Err.Number <> 6 And Err.Number <> 11 Then Err.Raise Err.Number, Err.Source, Err.Description
    Erase direction
End Function

' Вычисляет длину пары без возведения большого входного числа в квадрат.
Private Function HypotenuseLeastSquares(ByVal first As Double, ByVal second As Double) As Double
    first = Abs(first): second = Abs(second)
    If first < second Then
        If second > 0# Then HypotenuseLeastSquares = second * Sqr(1# + (first / second) ^ 2)
    ElseIf first > 0# Then
        HypotenuseLeastSquares = first * Sqr(1# + (second / first) ^ 2)
    End If
End Function


' Вычисляет базовое решение вырожденной задачи наименьших квадратов в переставленном QR по lmpar MINPACK.
' Свободные координаты равны нулю; конечное нелинейное равновесие проверяет вызывающий решатель.
Private Function TryUndampedRankStep(ByRef direction() As Double) As Boolean
    Erase direction
    If Not mPivoted Then Exit Function
    Dim triangle() As Double, ordered() As Double
    ReDim triangle(1 To mCount, 1 To mCount): ReDim ordered(1 To mCount)
    Dim row As Long, column As Long, position As Long, rank As Long, largest As Double, total As Double
    For row = 1 To mCount
        For column = row To mCount
            position = position + 1: triangle(column, row) = mR(position)
            largest = GeomMax(largest, Abs(mR(position)))
        Next column
    Next row
    rank = mCount
    For row = 1 To mCount
        If Abs(triangle(row, row)) <= MACHINE_EPSILON * mCount * largest Then rank = row - 1: Exit For
    Next row
    For row = rank To 1 Step -1
        total = -mQtf(row)
        For column = row + 1 To rank: total = total - triangle(column, row) * ordered(column): Next column
        ordered(row) = total / triangle(row, row)
    Next row
    ReDim direction(1 To mCount)
    For column = 1 To mCount: direction(mColumnOrder(column)) = ordered(column): Next column
    TryUndampedRankStep = True
End Function
