Attribute VB_Name = "modCoupledConcreteTensor"
Option Explicit

' ==========================================================================
' Тензорные операции для пространственного пластического бетона
' ==========================================================================
' Векторы упорядочены как xx, yy, zz, xy, xz, yz. Напряжения содержат
' физические касательные компоненты, деформации — инженерные сдвиги.
' Растяжение положительно. Используются инварианты физических напряжений.

Public Enum ECoupledTangentMode
    ctmFull = 0            ' Все шесть столбцов для решения системы поправок.
    ctmAxialGradient = 1   ' Только zz, xz, yz для продольного градиента напряжений.
    ctmStressOnly = 2      ' Проба без матрицы и с нулевым продольным градиентом.
    ctmFullBranchConsistent = 3    ' Все столбцы на той же материальной ветви, только без записи пробы.
    ctmRefinedBranchConsistent = 4 ' Согласованная касательная с малым шагом возмущения.
End Enum

' Готовит тензор изотропной упругости в инженерных компонентах.
' Вызывающий проверяет диапазон E и коэффициента Пуассона.
Public Sub CoupledElasticTensor(ByVal modulus As Double, ByVal poisson As Double, ByRef tensor() As Double)
    Dim shear As Double, lame As Double, i As Long, j As Long
    shear = modulus / (2# * (1# + poisson))
    lame = 2# * shear * poisson / (1# - 2# * poisson)
    For j = 1 To 6
        For i = 1 To 6: tensor(i, j) = 0#: Next i
    Next j
    For i = 1 To 3
        For j = 1 To 3: tensor(i, j) = lame: Next j
        tensor(i, i) = lame + 2# * shear
        tensor(i + 3, i + 3) = shear
    Next i
End Sub

' Возвращает след, эквивалентное напряжение q = sqrt(3 J2) и главные
' напряжения по убыванию. Гидростатика обрабатывается без деления на q.
Public Sub CoupledStressInvariants(ByRef stress() As Double, ByRef trace As Double, ByRef q As Double, _
        ByRef maxPrincipal As Double, ByRef midPrincipal As Double, ByRef minPrincipal As Double)
    Dim mean As Double, sx As Double, sy As Double, sz As Double, xy As Double, xz As Double, yz As Double
    trace = stress(1) + stress(2) + stress(3)
    mean = trace / 3#
    sx = stress(1) - mean: sy = stress(2) - mean: sz = stress(3) - mean
    xy = stress(4): xz = stress(5): yz = stress(6)
    Dim p As Double, determinant As Double, theta As Double
    p = Sqr((sx * sx + sy * sy + sz * sz + 2# * (xy * xy + xz * xz + yz * yz)) / 6#)
    q = 3# * p
    If p = 0# Then
        maxPrincipal = mean: midPrincipal = mean: minPrincipal = mean
        Exit Sub
    End If
    If xy = 0# And xz = 0# And yz = 0# Then
        OrderPrincipal stress(1), stress(2), stress(3), maxPrincipal, midPrincipal, minPrincipal
        Exit Sub
    End If
    sx = sx / p: sy = sy / p: sz = sz / p
    xy = xy / p: xz = xz / p: yz = yz / p
    determinant = sx * (sy * sz - yz * yz) - xy * (xy * sz - xz * yz) + xz * (xy * yz - xz * sy)
    If Abs(determinant * 0.5) > 0.99999999 Then
        StablePrincipal stress, maxPrincipal, midPrincipal, minPrincipal
        Exit Sub
    End If
    theta = ThirdAcos(determinant * 0.5)
    maxPrincipal = mean + 2# * p * Cos(theta)
    minPrincipal = mean + 2# * p * Cos(theta + 2# * GEOM_PI / 3#)
    midPrincipal = trace - maxPrincipal - minPrincipal
End Sub

' Упорядочивает три действительных значения без потери точного нуля одноосного состояния.
Private Sub OrderPrincipal(ByVal a As Double, ByVal b As Double, ByVal c As Double, _
        ByRef high As Double, ByRef middle As Double, ByRef low As Double)
    high = GeomMax(a, GeomMax(b, c)): low = GeomMin(a, GeomMin(b, c))
    If a >= GeomMin(b, c) And a <= GeomMax(b, c) Then
        middle = a
    ElseIf b >= GeomMin(a, c) And b <= GeomMax(a, c) Then
        middle = b
    Else
        middle = c
    End If
End Sub

' Вычисляет близкие корни симметричного тензора поворотами Якоби.
' В окрестности кратного корня тригонометрическая формула теряет точность
' как квадратный корень машинной ошибки и может создать ложное растяжение.
Private Sub StablePrincipal(ByRef stress() As Double, ByRef high As Double, ByRef middle As Double, ByRef low As Double)
    Dim a(1 To 3, 1 To 3) As Double, tensorScale As Double, i As Long, j As Long, k As Long, iteration As Long
    For i = 1 To 6: tensorScale = GeomMax(tensorScale, Abs(stress(i))): Next i
    For i = 1 To 3: a(i, i) = stress(i) / tensorScale: Next i
    a(1, 2) = stress(4) / tensorScale: a(1, 3) = stress(5) / tensorScale: a(2, 3) = stress(6) / tensorScale
    a(2, 1) = a(1, 2): a(3, 1) = a(1, 3): a(3, 2) = a(2, 3)
    Dim offDiagonal As Double, tau As Double, t As Double, cosine As Double, sine As Double, first As Double, second As Double
    For iteration = 1 To 32
        i = 1: j = 2: offDiagonal = Abs(a(1, 2))
        If Abs(a(1, 3)) > offDiagonal Then i = 1: j = 3: offDiagonal = Abs(a(1, 3))
        If Abs(a(2, 3)) > offDiagonal Then i = 2: j = 3: offDiagonal = Abs(a(2, 3))
        If offDiagonal <= 0.0000000000000001 Then Exit For
        tau = (a(j, j) - a(i, i)) / (2# * a(i, j))
        t = 1# / (Abs(tau) + Sqr(1# + tau * tau))
        If tau < 0# Then t = -t
        cosine = 1# / Sqr(1# + t * t): sine = t * cosine
        a(i, i) = a(i, i) - t * a(i, j): a(j, j) = a(j, j) + t * a(i, j)
        a(i, j) = 0#: a(j, i) = 0#
        For k = 1 To 3
            If k <> i And k <> j Then
                first = a(k, i): second = a(k, j)
                a(k, i) = cosine * first - sine * second: a(i, k) = a(k, i)
                a(k, j) = sine * first + cosine * second: a(j, k) = a(k, j)
            End If
        Next k
    Next iteration
    OrderPrincipal tensorScale * a(1, 1), tensorScale * a(2, 2), tensorScale * a(3, 3), high, middle, low
End Sub

' Вычисляет acos(x)/3 для спектральной формулы, ограничивая погрешность det.
Private Function ThirdAcos(ByVal argument As Double) As Double
    If argument >= 1# Then Exit Function
    If argument <= -1# Then
        ThirdAcos = GEOM_PI / 3#
    ElseIf argument = 0# Then
        ThirdAcos = GEOM_PI / 6#
    Else
        Dim angle As Double
        angle = Atn(Sqr((1# - argument) * (1# + argument)) / argument)
        If argument < 0# Then angle = angle + GEOM_PI
        ThirdAcos = angle / 3#
    End If
End Function

' Вычисляет поверхность Барселоны в эффективных напряжениях.
' Знак F проверяется с заданным сжатием cc и растяжением ct в МПа.
Public Function CoupledBarcelonaYield(ByVal trace As Double, ByVal q As Double, ByVal maxPrincipal As Double, _
        ByVal cc As Double, ByVal ct As Double, ByVal alpha As Double, ByVal gamma As Double) As Double
    Dim principalTerm As Double
    If maxPrincipal > 0# Then
        principalTerm = (cc / ct * (1# - alpha) - (1# + alpha)) * maxPrincipal
    Else
        principalTerm = gamma * maxPrincipal
    End If
    CoupledBarcelonaYield = (alpha * trace + q + principalTerm) / (1# - alpha) - cc
End Function

' Возвращает гиперболический потенциал и его градиент по напряжениям.
' Нормируемое скалярное произведение с инженерной деформацией даёт работу.
' Нулевая дилатансия и hydrostatic apex имеют нулевой градиент по пределу
' изотропной пробы; вызывающий отдельно проверяет возможность возврата.
Public Function CoupledConcretePotential(ByRef stress() As Double, ByVal trace As Double, _
        ByVal q As Double, ByVal dilation As Double, ByVal ft0 As Double, ByVal eccentricity As Double, _
        ByRef flow() As Double) As Double
    Dim offset As Double, radius As Double, mean As Double, i As Long
    offset = eccentricity * dilation * ft0
    radius = Sqr(offset * offset + (2# / 3#) * q * q)
    mean = trace / 3#
    If radius > 0# Then
        For i = 1 To 3
            flow(i) = (stress(i) - mean) / radius + dilation
            flow(i + 3) = 2# * stress(i + 3) / radius
        Next i
    Else
        For i = 1 To 3
            flow(i) = dilation
            flow(i + 3) = 0#
        Next i
    End If
    CoupledConcretePotential = radius + dilation * trace
End Function
