VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTorsionConcreteBalance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ===========================================================================
' Компенсирующее продольное сжатие бетона при кручении
' ===========================================================================
' Подготавливает напряжения на 1 Н растягивающей силы продольной арматуры.
' Использует центры и площади имеющихся бетонных элементов. Пригодны ячейки
' с достоверной границей, центром в заданной полосе оси хомута и проверенным
' отсутствием контакта с продольными стержнями. Это распределение по
' равновесию; сопротивление и деформации материала проверяет вызывающий код.

Private Const DETERMINANT_TOLERANCE As Double = 0.000000000001 ' Относительный допуск обусловленности ковариации.
Private Const RESULTANT_TOLERANCE As Double = 0.0000000001 ' Допуск при силе 1 Н и моментах, делённых на периметр.

Private Type TBalanceCell
    ElementIndex As Long
    Area As Double
    X As Double
    Y As Double
End Type

Private Type TBarExtent
    X As Double
    Y As Double
    Radius As Double
End Type

Private mStressPerUnitForce() As Double ' Порядок совпадает с индексами бетонных элементов модели.
Private mPrepared As Boolean

' Строит начальное сжатие при единичной силе и заданной точке её приложения.
' bandWidth — максимальное расстояние центра бетонной ячейки до оси, мм.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal hoop As CTorsionHoopGeometry, _
        ByVal bandWidth As Double, ByVal resultantX As Double, ByVal resultantY As Double)
    If mPrepared Then BalanceError "Распределение сжатия бетона уже подготовлено."
    If section Is Nothing Or hoop Is Nothing Then BalanceError "Передайте бетонное сечение и проверенную ось хомута."
    If Not IsFiniteGeometry(bandWidth) Or bandWidth <= GEOM_BOUNDARY_TOLERANCE Then _
        BalanceError "Ширина расчётной полосы бетона у хомута должна быть больше геометрического допуска."
    If Not IsFiniteGeometry(resultantX) Or Not IsFiniteGeometry(resultantY) Then _
        BalanceError "Укажите конечные координаты продольной равнодействующей."

    Dim cellCount As Long, barCount As Long, perimeter As Double
    cellCount = section.ConcreteCount: barCount = section.RebarCount: perimeter = hoop.Length
    If cellCount < 3 Then BalanceError "Для распределения сжатия нужны бетонные элементы в трёх неколлинеарных точках."
    If Not IsFiniteGeometry(perimeter) Or perimeter <= GEOM_BOUNDARY_TOLERANCE Then _
        BalanceError "Периметр оси хомута недопустим для проверки моментов."

    Dim bars() As TBarExtent, cells() As TBalanceCell
    ReadBarExtents section, barCount, bars
    ReDim cells(1 To cellCount)
    Dim activeCount As Long, totalArea As Double, meanX As Double, meanY As Double
    Dim i As Long, centerX As Double, centerY As Double, area As Double
    For i = 1 To cellCount
        section.ConcreteElementForSolver i, centerX, centerY, area
        If IsEligibleCell(section, hoop, i, centerX, centerY, bandWidth, bars, barCount) Then
            If Not IsFiniteGeometry(area) Or area <= 0# Then BalanceError "Площадь бетонного элемента " & section.ConcreteID(i) & " недопустима."
            activeCount = activeCount + 1
            With cells(activeCount)
                .ElementIndex = i: .Area = area
                .X = (centerX - resultantX) / perimeter: .Y = (centerY - resultantY) / perimeter ' В этой системе требуемые моменты равны нулю.
                totalArea = totalArea + area
                meanX = meanX + area * .X: meanY = meanY + area * .Y
            End With
        End If
    Next i
    If activeCount < 3 Then BalanceError "В расчётной полосе мало ячеек с достоверными границами вне продольной арматуры. Проверьте полосу и бетонную сетку."
    meanX = meanX / totalArea: meanY = meanY / totalArea

    Dim slopeX As Double, slopeY As Double, candidate() As Double
    SolveAffineWeights cells, activeCount, totalArea, meanX, meanY, slopeX, slopeY
    ReDim candidate(1 To cellCount)
    BuildCoefficients section, cells, activeCount, totalArea, meanX, meanY, slopeX, slopeY, candidate
    mStressPerUnitForce = candidate
    mPrepared = True
End Sub

' Возвращает независимый массив отрицательных напряжений в МПа на 1 Н.
' В ячейках вне расчётной полосы и ячейках с арматурой значения нулевые.
Friend Sub CopyStressPerUnitForce(ByRef coefficients() As Double)
    If Not mPrepared Then BalanceError "Сначала подготовьте распределение сжатия бетона."
    coefficients = mStressPerUnitForce
End Sub

' Готовит охватывающие окружности продольных стержней. При переданной
' увеличенной площади учитывается также радиус равновеликого круга.
Private Sub ReadBarExtents(ByVal section As CSectionModel, ByVal count As Long, ByRef bars() As TBarExtent)
    If count = 0 Then Exit Sub
    ReDim bars(1 To count)
    Dim i As Long, area As Double
    For i = 1 To count
        section.RebarElementForSolver i, bars(i).X, bars(i).Y, area
        bars(i).Radius = GeomMax(0.5 * section.RebarDiameter(i), Sqr(area / GEOM_PI)) + GEOM_BOUNDARY_TOLERANCE
    Next i
End Sub

' Выбирает элементы по точному расстоянию центра до оси. Контакт окружности
' стержня с охватывающим прямоугольником исключает ячейку, поэтому запас
' при сложных и повернутых границах уменьшает набор допустимого бетона.
Private Function IsEligibleCell(ByVal section As CSectionModel, ByVal hoop As CTorsionHoopGeometry, _
        ByVal index As Long, ByVal centerX As Double, ByVal centerY As Double, ByVal bandWidth As Double, _
        ByRef bars() As TBarExtent, ByVal barCount As Long) As Boolean
    Dim station As Double, distance As Double
    If Not hoop.TryProjectStation(centerX, centerY, station, distance) Then _
        BalanceError "Недопустимы координаты бетонного элемента " & section.ConcreteID(index) & "."
    If distance > bandWidth + GEOM_BOUNDARY_TOLERANCE Then Exit Function

    Dim cell As CConcreteElement: Set cell = section.ConcreteElement(index)
    If Not cell.HasExactGeometry And Not cell.HasKnownRectangle Then Exit Function
    Dim minX As Double, maxX As Double, minY As Double, maxY As Double
    section.ConcreteProjectionBounds index, 1#, 0#, minX, maxX
    section.ConcreteProjectionBounds index, 0#, 1#, minY, maxY

    Dim j As Long, dx As Double, dy As Double
    For j = 1 To barCount
        With bars(j)
            If .X + .Radius >= minX And .X - .Radius <= maxX And .Y + .Radius >= minY And .Y - .Radius <= maxY Then
                dx = GeomMax(0#, GeomMax(minX - .X, .X - maxX))
                dy = GeomMax(0#, GeomMax(minY - .Y, .Y - maxY))
                If dx * dx + dy * dy <= .Radius * .Radius Then Exit Function
            End If
        End With
    Next j
    IsEligibleCell = True
End Function

' Находит наклоны положительного аффинного распределения. При равномерном
' весе площади среднее исходное положение корректирует ковариация центров.
Private Sub SolveAffineWeights(ByRef cells() As TBalanceCell, ByVal count As Long, ByVal totalArea As Double, _
        ByVal meanX As Double, ByVal meanY As Double, ByRef slopeX As Double, ByRef slopeY As Double)
    Dim varianceX As Double, varianceY As Double, covariance As Double
    Dim i As Long, dx As Double, dy As Double, fraction As Double
    For i = 1 To count
        With cells(i)
            dx = .X - meanX: dy = .Y - meanY: fraction = .Area / totalArea
            varianceX = varianceX + fraction * dx * dx
            varianceY = varianceY + fraction * dy * dy
            covariance = covariance + fraction * dx * dy
        End With
    Next i
    Dim determinant As Double, trace As Double
    determinant = varianceX * varianceY - covariance * covariance
    trace = varianceX + varianceY
    If determinant <= DETERMINANT_TOLERANCE * trace * trace Then _
        BalanceError "Бетонные ячейки полосы не обеспечивают независимые моменты относительно двух осей."
    slopeX = (-varianceY * meanX + covariance * meanY) / determinant
    slopeY = (-varianceX * meanY + covariance * meanX) / determinant
End Sub

' Переводит неотрицательные доли единичной силы в сжимающие напряжения.
' Повторно интегрирует силу и моменты по тем же площадям и центрам.
Private Sub BuildCoefficients(ByVal section As CSectionModel, ByRef cells() As TBalanceCell, ByVal count As Long, _
        ByVal totalArea As Double, ByVal meanX As Double, ByVal meanY As Double, _
        ByVal slopeX As Double, ByVal slopeY As Double, ByRef coefficients() As Double)
    Dim i As Long, factor As Double, force As Double
    Dim axial As Double, momentX As Double, momentY As Double
    For i = 1 To count
        With cells(i)
            factor = 1# + slopeX * (.X - meanX) + slopeY * (.Y - meanY)
            If factor < 0# Then BalanceError "Для продольной равнодействующей выбранная полоса потребовала бы растяжения ячейки " & _
                section.ConcreteID(.ElementIndex) & ". Проверьте положение арматуры и расчётную полосу."
            coefficients(.ElementIndex) = -factor / totalArea
            force = coefficients(.ElementIndex) * .Area
            axial = axial + force
            momentX = momentX + force * .Y: momentY = momentY + force * .X
        End With
    Next i
    If Abs(axial + 1#) > RESULTANT_TOLERANCE Or Abs(momentX) > RESULTANT_TOLERANCE Or Abs(momentY) > RESULTANT_TOLERANCE Then _
        BalanceError "Сжатие бетонной полосы не воспроизводит заданную силу и два момента."
End Sub

' Ограничивает входы диапазоном, пригодным для квадратов и произведений.
Private Function IsFiniteGeometry(ByVal value As Double) As Boolean
    IsFiniteGeometry = (value = value And Abs(value) < 1E+100)
End Function

' Передаёт понятную причину неприменимости расчётной полосы вызывающему коду.
Private Sub BalanceError(ByVal explanation As String)
    Err.Raise vbObjectError + 5670, "CTorsionConcreteBalance", explanation
End Sub
