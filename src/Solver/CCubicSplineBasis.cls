VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCubicSplineBasis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Компактный кубический базис для поля перемещений сечения
' ==========================================================================
' Равномерные внутренние узлы, четыре совпадающих узла на каждом конце.
' В одной точке работает не более четырёх продольных функций вдоль любой
' координаты. Производная вычисляется по физической координате, 1/мм.
' Значения и производные сохраняет подготовщик операторов сечения.

Private Const DEGREE As Long = 3
Private Const ENDPOINT_TOLERANCE As Double = 0.000000000001

Private mMinimum As Double
Private mLength As Double
Private mSpanCount As Long
Private mKnots() As Double

' Задаёт ось и число интервалов между различными узлами.
' Четыре узла на концах обеспечивают интерполяцию значений на границе.
Friend Sub Initialize(ByVal minimum As Double, ByVal maximum As Double, ByVal spanCount As Long)
    If maximum <= minimum Or spanCount < 1 Then
        Err.Raise vbObjectError + 1771, "CCubicSplineBasis", _
            "Для кубических сплайнов нужны положительная длина оси и хотя бы один интервал."
    End If

    mMinimum = minimum
    mLength = maximum - minimum
    mSpanCount = spanCount
    ReDim mKnots(0 To spanCount + 2 * DEGREE)
    Dim i As Long
    For i = 1 To spanCount - 1
        mKnots(DEGREE + i) = i / spanCount
    Next i
    For i = DEGREE + spanCount To UBound(mKnots)
        mKnots(i) = 1#
    Next i
End Sub

' Возвращает число функций; их индексы начинаются с единицы.
Friend Property Get Count() As Long
    Count = mSpanCount + DEGREE
End Property

' Возвращает коэффициент точного воспроизведения координаты в кубическом базисе.
' Индекс функции начинается с единицы, три внутренних узла задают абсциссу Гревилля.
Friend Function CoordinateCoefficient(ByVal basisIndex As Long) As Double
    Dim parameter As Double
    parameter = (mKnots(basisIndex) + mKnots(basisIndex + 1) + mKnots(basisIndex + 2)) / DEGREE
    CoordinateCoefficient = mMinimum + mLength * parameter
End Function

' Возвращает начало блока из четырёх активных функций и их значения.
' Вызывающий передаёт два заранее выделенных массива с границами 1..4.
Friend Sub Evaluate(ByVal coordinate As Double, ByRef firstIndex As Long, _
        ByRef values() As Double, ByRef derivatives() As Double)
    If mSpanCount = 0 Then
        Err.Raise vbObjectError + 1772, "CCubicSplineBasis", "Кубический базис не подготовлен."
    End If
    Dim parameter As Double
    parameter = (coordinate - mMinimum) / mLength
    If parameter < -ENDPOINT_TOLERANCE Or parameter > 1# + ENDPOINT_TOLERANCE Then
        Err.Raise vbObjectError + 1773, "CCubicSplineBasis", _
            "Расчётная точка находится вне области кубического базиса."
    End If
    If parameter < 0# Then parameter = 0#
    If parameter > 1# Then parameter = 1#

    Dim span As Long
    span = CLng(Int(parameter * mSpanCount)) + DEGREE
    If parameter = 1# Then span = mSpanCount + DEGREE - 1
    firstIndex = span - DEGREE + 1

    Dim basis(0 To DEGREE) As Double
    Dim leftDistance(1 To DEGREE) As Double
    Dim rightDistance(1 To DEGREE) As Double
    Dim saved As Double
    Dim term As Double
    Dim denominator As Double
    Dim order As Long
    Dim j As Long
    Dim k As Long
    Dim node As Long
    Dim incoming As Double
    Dim outgoing As Double

    basis(0) = 1#
    For order = 1 To DEGREE
        leftDistance(order) = parameter - mKnots(span + 1 - order)
        rightDistance(order) = mKnots(span + order) - parameter
        saved = 0#
        For j = 0 To order - 1
            denominator = rightDistance(j + 1) + leftDistance(order - j)
            term = basis(j) / denominator ' Две границы рекурсии включают текущий ненулевой интервал.
            basis(j) = saved + rightDistance(j + 1) * term
            saved = leftDistance(order - j) * term
        Next j
        basis(order) = saved

        If order = DEGREE - 1 Then
            ' Производная кубических функций выражается через уже рассчитанные квадратичные.
            For k = 0 To DEGREE
                node = span - DEGREE + k
                incoming = 0#
                outgoing = 0#
                If k > 0 Then incoming = basis(k - 1) / (mKnots(node + DEGREE) - mKnots(node))
                If k < DEGREE Then outgoing = basis(k) / (mKnots(node + DEGREE + 1) - mKnots(node + 1))
                derivatives(k + 1) = (DEGREE / mLength) * (incoming - outgoing)
            Next k
        End If
    Next order
    For k = 0 To DEGREE
        values(k + 1) = basis(k)
    Next k
End Sub
