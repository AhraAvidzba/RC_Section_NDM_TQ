VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledSectionQuadrature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Квадратура связанного расчёта и локальные операторы сплайнов
' ==========================================================================
' Использует общую бетонную область и точные площади пересечений ячеек.
' Одноточечная квадратура по частям ячеек уточняется независимо от базиса.
' Данные создаются один раз на ревизию сечения и используются только на чтение.

Private Const GAUSS_ORDER As Long = 5
Private mSplineX As CCubicSplineBasis
Private mSplineY As CCubicSplineBasis
Private mRegion As CConcreteRegion
Private mGeometry As CGeometryRegion
Private mFrame As CSectionNumericalFrame

' Подготавливает область и численное начало из геометрии без повторного чтения Excel.
' Достоверные дуги и отверстия сохраняются в общих геометрических классах.
Friend Sub Initialize(ByVal section As CSectionModel, ByVal spansX As Long, ByVal spansY As Long)
    If section Is Nothing Then GeometryError "Сформируйте сечение перед связанным расчётом."
    If section.ConcreteCount = 0 Then GeometryError "В сформированном сечении нет бетона."

    Dim query As New CSectionGeometryQuery
    query.Initialize section
    Dim region As CConcreteRegion
    Set region = query.ConcreteDomain
    Dim components As Collection
    Set components = query.ConnectedRegions(region)
    If components.Count <> 1 Then
        GeometryError "Для передачи поперечных сил требуется одна связная бетонная область. Найдено частей: " & CStr(components.Count) & "."
    End If

    Dim geometry As New CGeometryRegion, frame As New CSectionNumericalFrame
    geometry.Initialize region
    frame.Initialize section
    Dim splineX As New CCubicSplineBasis, splineY As New CCubicSplineBasis
    splineX.Initialize geometry.MinX - frame.OriginX, geometry.MaxX - frame.OriginX, spansX
    splineY.Initialize geometry.MinY - frame.OriginY, geometry.MaxY - frame.OriginY, spansY

    Set mGeometry = geometry
    Set mRegion = region
    Set mFrame = frame
    Set mSplineX = splineX
    Set mSplineY = splineY
End Sub

' Создаёт точные площади и центры квадратуры, затем раскрывает только 16
' ненулевых функций и их производные. x/y приведены к численному началу.
Friend Sub Prepare(ByVal pointsPerSpan As Long, ByRef x() As Double, ByRef y() As Double, _
        ByRef area() As Double, ByRef indices() As Long, ByRef values() As Double, _
        ByRef gradientX() As Double, ByRef gradientY() As Double, Optional ByVal useFullSpanGauss5 As Boolean = False)
    If mGeometry Is Nothing Then GeometryError "Бетонная область связанного расчёта не подготовлена."
    If pointsPerSpan < 4 Then GeometryError "Требуется не менее четырёх делений на интервал кубического базиса."
    If useFullSpanGauss5 And pointsPerSpan <> GAUSS_ORDER Then _
        GeometryError "Для квадратуры Гаусса по целому интервалу нужны пять узлов по каждой оси."
    Dim mesh As New CFiberMeshBuilder
    Dim stepX As Double, stepY As Double
    stepX = mGeometry.Width / ((mSplineX.Count - 3) * CDbl(pointsPerSpan))
    stepY = mGeometry.Height / ((mSplineY.Count - 3) * CDbl(pointsPerSpan))
    mesh.BuildMesh mGeometry, stepX, stepY, cellGeometry:="Contour"
    Dim gaussX() As Double, gaussY() As Double, gaussArea() As Double
    If useFullSpanGauss5 Then PrepareFullSpanGauss mesh, stepX, stepY, gaussX, gaussY, gaussArea

    Dim count As Long
    count = mesh.FiberCount
    ReDim x(1 To count): ReDim y(1 To count): ReDim area(1 To count)
    ReDim indices(1 To 16, 1 To count): ReDim values(1 To 16, 1 To count)
    ReDim gradientX(1 To 16, 1 To count): ReDim gradientY(1 To 16, 1 To count)
    Dim i As Long, k As Long, localIndex(1 To 16) As Long
    Dim shape(1 To 16) As Double, dx(1 To 16) As Double, dy(1 To 16) As Double
    Dim totalArea As Double, worldX As Double, worldY As Double
    For i = 1 To count
        worldX = mesh.FiberX(i)
        worldY = mesh.FiberY(i)
        area(i) = mesh.FiberArea(i)
        If useFullSpanGauss5 Then
            If gaussArea(i) > 0# Then
                worldX = gaussX(i): worldY = gaussY(i): area(i) = gaussArea(i)
            End If
        End If
        If Not mGeometry.ContainsPoint(worldX, worldY) Then
            GeometryError "Центр ячейки связанного расчёта оказался вне бетона. Уточните квадратуру у границ отверстий."
        End If
        x(i) = worldX - mFrame.OriginX
        y(i) = worldY - mFrame.OriginY
        totalArea = totalArea + area(i)
        EvaluateShape x(i), y(i), localIndex, shape, dx, dy
        For k = 1 To 16
            indices(k, i) = localIndex(k)
            values(k, i) = shape(k)
            gradientX(k, i) = dx(k)
            gradientY(k, i) = dy(k)
        Next k
    Next i
    If Abs(totalArea - mRegion.Area) > 1E-7 * mRegion.Area Then _
        GeometryError "Суммарная площадь ячеек связанного расчёта не совпадает с площадью бетонной области. Проверьте контуры и разрешение квадратуры."
End Sub

' Заменяет квадратуру целого интервала только после проверки всех исходных ячеек и новых узлов.
' Пересечение любой границы, включая отверстие, сохраняет фактические контурные ячейки.
Private Sub PrepareFullSpanGauss(ByVal mesh As CFiberMeshBuilder, ByVal stepX As Double, ByVal stepY As Double, _
        ByRef gx() As Double, ByRef gy() As Double, ByRef weight() As Double)
    Const GRID_MATCH As Double = 1E-7
    Const AREA_MATCH As Double = 1E-8
    Dim spansX As Long, spansY As Long, ix As Long, iy As Long, minX As Double, minY As Double
    Dim intervalX As Double, intervalY As Double, startX As Double, startY As Double
    spansX = mSplineX.Count - 3: spansY = mSplineY.Count - 3
    minX = mGeometry.MinX: minY = mGeometry.MinY
    intervalX = mGeometry.Width / spansX: intervalY = mGeometry.Height / spansY
    Dim fullSpan() As Boolean, sourcePoint() As Long
    ReDim fullSpan(1 To spansX, 1 To spansY)
    ReDim sourcePoint(1 To GAUSS_ORDER, 1 To GAUSS_ORDER, 1 To spansX, 1 To spansY)
    Dim slotX As Long, slotY As Long, count As Long
    For iy = 1 To spansY
        startY = minY + (iy - 1) * intervalY
        For ix = 1 To spansX
            startX = minX + (ix - 1) * intervalX
            If Not mGeometry.MeshBoundaryCrossesCell(startX, startY, startX + intervalX, startY + intervalY) Then
                fullSpan(ix, iy) = mGeometry.ContainsPoint(startX + 0.5 * intervalX, startY + 0.5 * intervalY)
            End If
        Next ix
    Next iy

    Dim i As Long, gridX As Double, gridY As Double, cellX As Long, cellY As Long, cellArea As Double
    count = mesh.FiberCount: cellArea = stepX * stepY
    ReDim gx(1 To count): ReDim gy(1 To count): ReDim weight(1 To count)
    For i = 1 To count
        gridX = (mesh.FiberX(i) - minX) / stepX: gridY = (mesh.FiberY(i) - minY) / stepY
        cellX = CLng(Fix(gridX)): cellY = CLng(Fix(gridY))
        If cellX >= 0 And cellX < GAUSS_ORDER * spansX And cellY >= 0 And cellY < GAUSS_ORDER * spansY Then
            ix = 1 + cellX \ GAUSS_ORDER: iy = 1 + cellY \ GAUSS_ORDER
            If fullSpan(ix, iy) Then
                slotX = 1 + cellX Mod GAUSS_ORDER: slotY = 1 + cellY Mod GAUSS_ORDER
                If Abs(gridX - cellX - 0.5) > GRID_MATCH Or Abs(gridY - cellY - 0.5) > GRID_MATCH Or _
                        Abs(mesh.FiberArea(i) / cellArea - 1#) > AREA_MATCH Or sourcePoint(slotX, slotY, ix, iy) <> 0 Then
                    fullSpan(ix, iy) = False
                Else
                    sourcePoint(slotX, slotY, ix, iy) = i
                End If
            End If
        End If
    Next i

    Dim node(1 To GAUSS_ORDER) As Double, gaussWeight(1 To GAUSS_ORDER) As Double
    Dim spanX(1 To GAUSS_ORDER) As Double, spanY(1 To GAUSS_ORDER) As Double, intervalArea As Double
    GaussUnitRule node, gaussWeight
    intervalArea = intervalX * intervalY
    For iy = 1 To spansY
        startY = minY + (iy - 1) * intervalY
        For slotY = 1 To GAUSS_ORDER: spanY(slotY) = startY + node(slotY) * intervalY: Next slotY
        For ix = 1 To spansX
            If fullSpan(ix, iy) Then
                startX = minX + (ix - 1) * intervalX
                For slotX = 1 To GAUSS_ORDER: spanX(slotX) = startX + node(slotX) * intervalX: Next slotX
                For slotY = 1 To GAUSS_ORDER
                    For slotX = 1 To GAUSS_ORDER
                        If sourcePoint(slotX, slotY, ix, iy) = 0 Or _
                                Not mGeometry.ContainsPoint(spanX(slotX), spanY(slotY)) Then fullSpan(ix, iy) = False
                    Next slotX
                Next slotY
                If fullSpan(ix, iy) Then
                    For slotY = 1 To GAUSS_ORDER
                        For slotX = 1 To GAUSS_ORDER
                            i = sourcePoint(slotX, slotY, ix, iy)
                            gx(i) = spanX(slotX): gy(i) = spanY(slotY)
                            weight(i) = intervalArea * gaussWeight(slotX) * gaussWeight(slotY)
                        Next slotX
                    Next slotY
                End If
            End If
        Next ix
    Next iy
End Sub

' Строит пять узлов и весов на единичном интервале по точной формуле Гаусса–Лежандра.
Private Sub GaussUnitRule(ByRef node() As Double, ByRef weight() As Double)
    Dim innerNode As Double, outerNode As Double, radical As Double
    radical = 2# * Sqr(10# / 7#)
    outerNode = Sqr(5# + radical) / 3#: innerNode = Sqr(5# - radical) / 3#
    node(1) = 0.5 * (1# - outerNode): node(5) = 1# - node(1)
    node(2) = 0.5 * (1# - innerNode): node(4) = 1# - node(2)
    node(3) = 0.5
    radical = 13# * Sqr(70#)
    weight(1) = (322# - radical) / 1800#: weight(5) = weight(1)
    weight(2) = (322# + radical) / 1800#: weight(4) = weight(2)
    weight(3) = 64# / 225#
End Sub

' Вычисляет локальный тензорный базис в координатах численного начала.
' Индексы нумеруют функции по X, затем по Y; градиенты даны в 1/мм.
Friend Sub EvaluateShape(ByVal x As Double, ByVal y As Double, ByRef indices() As Long, _
        ByRef value() As Double, ByRef derivativeX() As Double, ByRef derivativeY() As Double)
    Dim shapeX(1 To 4) As Double, shapeY(1 To 4) As Double
    Dim dx(1 To 4) As Double, dy(1 To 4) As Double
    Dim firstX As Long, firstY As Long
    mSplineX.Evaluate x, firstX, shapeX, dx
    mSplineY.Evaluate y, firstY, shapeY, dy

    Dim ix As Long, iy As Long, i As Long, rowOffset As Long, countX As Long
    countX = mSplineX.Count
    For iy = 1 To 4
        rowOffset = (firstY + iy - 2) * countX + firstX - 1
        For ix = 1 To 4
            i = i + 1
            indices(i) = rowOffset + ix
            value(i) = shapeX(ix) * shapeY(iy)
            derivativeX(i) = dx(ix) * shapeY(iy)
            derivativeY(i) = shapeX(ix) * dy(iy)
        Next ix
    Next iy
End Sub

' Возвращает коэффициенты x/L и y/L в полном базисе численного начала.
' Порядок функций совпадает с EvaluateShape до исключения неактивных опор.
Friend Sub CopyCoordinateCoefficients(ByRef alongX() As Double, ByRef alongY() As Double)
    ReDim alongX(1 To BasisCount), alongY(1 To BasisCount)
    Dim ix As Long, iy As Long, index As Long
    Dim lengthScale As Double, ordinate As Double, abscissae() As Double
    lengthScale = mFrame.LengthScale
    ReDim abscissae(1 To mSplineX.Count)
    For ix = 1 To mSplineX.Count
        abscissae(ix) = mSplineX.CoordinateCoefficient(ix) / lengthScale
    Next ix
    For iy = 1 To mSplineY.Count
        ordinate = mSplineY.CoordinateCoefficient(iy) / lengthScale
        For ix = 1 To mSplineX.Count
            index = index + 1
            alongX(index) = abscissae(ix)
            alongY(index) = ordinate
        Next ix
    Next iy
End Sub

' Выражает девять тензорных квадратичных полиномов в полном кубическом
' базисе. Геометрическое исключение неактивных опор выполняет потребитель.
Friend Sub CopyQuadraticCoefficients(ByRef coefficients() As Double)
    ReDim coefficients(1 To 9, 1 To BasisCount)
    Dim ix As Long, iy As Long, a As Long, b As Long, index As Long
    Dim alongX() As Double, alongY() As Double
    For iy = 1 To mSplineY.Count
        mSplineY.QuadraticCoefficients iy, alongY
        For ix = 1 To mSplineX.Count
            index = index + 1: mSplineX.QuadraticCoefficients ix, alongX
            For b = 1 To 3
                For a = 1 To 3
                    coefficients(3 * (b - 1) + a, index) = alongX(a) * alongY(b)
                Next a
            Next b
        Next ix
    Next iy
End Sub

' Возвращает отсортированные параметры деления аналитического сегмента
' линиями внутренних узлов обеих осей квадратуры.
Friend Sub PathCuts(ByRef edge As TRegionEdge, ByRef cuts() As Double, ByRef count As Long)
    Erase cuts: count = 0
    modGeometryPrimitives.AddScalar cuts, count, 0#
    modGeometryPrimitives.AddScalar cuts, count, 1#
    Dim i As Long, spans As Long, minimum As Double, stepValue As Double
    spans = mSplineX.Count - 3
    minimum = mGeometry.MinX: stepValue = mGeometry.Width / spans
    For i = 1 To spans - 1
        modGeometryPrimitives.HalfPlaneCuts edge, 1#, 0#, minimum + stepValue * i, cuts, count
    Next i
    spans = mSplineY.Count - 3
    minimum = mGeometry.MinY: stepValue = mGeometry.Height / spans
    For i = 1 To spans - 1
        modGeometryPrimitives.HalfPlaneCuts edge, 0#, 1#, minimum + stepValue * i, cuts, count
    Next i
    modGeometryPrimitives.SortScalars cuts, count
End Sub

' Возвращает число функций до исключения неактивной опоры вне бетона.
Friend Property Get BasisCount() As Long
    BasisCount = mSplineX.Count * mSplineY.Count
End Property

' Возвращает численную систему координат исходной модели сечения.
Friend Property Get Frame() As CSectionNumericalFrame
    Set Frame = mFrame
End Property

' Даёт независимую копию бетонной области для проверки ветвей арматуры.
Friend Property Get Region() As CConcreteRegion
    Set Region = mRegion.Clone
End Property

' Сообщает о входной геометрии, по которой нельзя сформировать квадратуру.
Private Sub GeometryError(ByVal message As String)
    Err.Raise vbObjectError + 1774, "CCoupledSectionQuadrature", message
End Sub
