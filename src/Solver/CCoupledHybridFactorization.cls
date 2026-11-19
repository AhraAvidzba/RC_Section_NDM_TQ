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
' QR и dogleg для гибридного метода решения полного равновесия
' ==========================================================================
' Перенос квадратной ветви MINPACK/CMinpack: qrfac без перестановок,
' qform, dogleg, r1updt, r1mpyq. Масштаб координат задан снаружи (diag=1).
' Источник: literature/Coupled6D/NumericalControl/Minpack, ревизия
' 64c5635addd72fb54f44db8546c191feeb70e338; полная лицензия
' CopyrightMINPACK.txt хранится там же и сопровождает распространение.
' This product includes software developed by the University of Chicago,
' as Operator of Argonne National Laboratory.
' Первый индекс матрицы — столбец; R упакована последовательно по строкам.

Private Const MACHINE_EPSILON As Double = 2# ^ (-52)
Private Const INVERSE_MAXIMUM As Double = (2# ^ (-1022)) / 4# ' IEEE Double: округлённое 1/DBL_MAX без потери цифр при экспорте VBE.
Private mCount As Long, mQ() As Double, mR() As Double, mQtf() As Double

' Вычисляет QR исходной матрицы и Q^T*f, не меняя входные массивы.
Friend Sub Factor(ByRef matrix() As Double, ByRef residual() As Double, ByVal count As Long)
    mCount = count: mQ = matrix: mQtf = residual
    ReDim mR(1 To CLng(CDbl(count) * (count + 1) / 2#))
    Dim diagonal() As Double, work() As Double
    ReDim diagonal(1 To count): ReDim work(1 To count)
    Dim i As Long, j As Long, k As Long, position As Long, total As Double, magnitude As Double, multiplier As Double
    For j = 1 To count
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
