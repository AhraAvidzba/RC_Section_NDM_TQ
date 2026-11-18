VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledFieldSubspace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Квадратичное подпространство для вспомогательного прогноза связанного поля
' ==========================================================================
' Материал, нагрузки и принятые состояния здесь не хранятся. Отображение
' ограничивает виртуальные перемещения и остаток одинаковым оператором.
' Матрицы хранят столбец первым индексом, как CCoupledLinearSystem.

Private Const COORDINATE_COUNT As Long = 39
Private Const GAUGE_COUNT As Long = 4
Private mProjection() As Double
Private mNormal As CCoupledLinearSystem
Private mDegreeCount As Long

' Подготавливает постоянное отображение и нормальную матрицу один раз.
' Линейно зависимое подпространство недоступно, но не запрещает полный расчёт.
Friend Function Prepare(ByVal operators As CCoupledElasticState) As Boolean
    mDegreeCount = 0: Set mNormal = Nothing: Erase mProjection
    Dim projection() As Double, normal() As Double, i As Long, j As Long, row As Long
    operators.CopyQuadraticProjection projection
    ReDim normal(1 To COORDINATE_COUNT, 1 To COORDINATE_COUNT)
    For j = 1 To COORDINATE_COUNT
        For i = 1 To COORDINATE_COUNT
            For row = 1 To operators.DegreeCount
                normal(j, i) = normal(j, i) + projection(j, row) * projection(i, row)
            Next row
        Next i
    Next j
    Dim system As New CCoupledLinearSystem
    If Not system.Factor(normal, COORDINATE_COUNT) Then Exit Function
    mProjection = projection: Set mNormal = system
    mDegreeCount = operators.DegreeCount
    Prepare = True
End Function

' Возвращает размер вспомогательных уравнений, включая множители связей.
Friend Property Get Count() As Long
    Count = COORDINATE_COUNT
End Property

' Получает ортогональную проекцию начального поля без сохранения высоких мод.
Friend Function TryCoordinates(ByRef field() As Double, ByRef coordinates() As Double) As Boolean
    If mNormal Is Nothing Then Exit Function
    PullVector field, coordinates
    TryCoordinates = mNormal.SolveFactored(coordinates)
End Function

' Восстанавливает поле в исходных координатах без изменения материальных буферов.
Friend Sub Expand(ByRef coordinates() As Double, ByRef field() As Double)
    ReDim field(1 To mDegreeCount)
    Dim row As Long, i As Long
    For row = 1 To mDegreeCount
        For i = 1 To COORDINATE_COUNT
            field(row) = field(row) + mProjection(i, row) * coordinates(i)
        Next i
    Next row
End Sub

' Ограничивает остаток или линейную меру сопряжённым отображением P-transpose.
Friend Sub PullVector(ByRef full() As Double, ByRef reduced() As Double)
    ReDim reduced(1 To COORDINATE_COUNT)
    Dim row As Long, i As Long
    For row = 1 To mDegreeCount
        For i = 1 To COORDINATE_COUNT
            reduced(i) = reduced(i) + mProjection(i, row) * full(row)
        Next i
    Next row
End Sub

' Собирает P-transpose K P без симметризации исходной касательной.
Friend Sub PullMatrix(ByRef full() As Double, ByRef reduced() As Double)
    Dim intermediate() As Double, column As Long, row As Long, i As Long, j As Long
    ReDim intermediate(1 To COORDINATE_COUNT, 1 To mDegreeCount)
    ReDim reduced(1 To COORDINATE_COUNT, 1 To COORDINATE_COUNT)
    For column = 1 To mDegreeCount
        For j = 1 To COORDINATE_COUNT
            If mProjection(j, column) <> 0# Then
                For row = 1 To mDegreeCount
                    intermediate(j, row) = intermediate(j, row) + full(column, row) * mProjection(j, column)
                Next row
            End If
        Next j
    Next column
    For row = 1 To mDegreeCount
        For i = 1 To COORDINATE_COUNT
            If mProjection(i, row) <> 0# Then
                For j = 1 To COORDINATE_COUNT
                    reduced(j, i) = reduced(j, i) + mProjection(i, row) * intermediate(j, row)
                Next j
            End If
        Next i
    Next row
End Sub

' Возвращает невязку ограниченных уравнений в МПа; четыре связи масштабируются модулем.
Friend Function ResidualNorm(ByRef reduced() As Double, ByVal referenceModulus As Double) As Double
    Dim i As Long, value As Double
    For i = 1 To COORDINATE_COUNT
        value = Abs(reduced(i))
        If i > COORDINATE_COUNT - GAUGE_COUNT Then value = value * referenceModulus
        If value > ResidualNorm Then ResidualNorm = value
    Next i
End Function
