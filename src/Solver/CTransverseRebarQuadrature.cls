VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CTransverseRebarQuadrature"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Численная проекция ветвей поперечной арматуры
' ==========================================================================
' Для повторяющегося стержня площади As, шага s и длины dl вес As*dl/s
' имеет размерность мм2. При идеальном сцеплении деформация вдоль ветви
' равна tx^2*epsXX + ty^2*epsYY + tx*ty*gammaXY. Вес применяется также
' к объёму замещённого бетона. Стержни не заменяются суммой площадей.

Private mCount As Long, mCapacity As Long
Private mGroup() As Long
Private mX() As Double, mY() As Double, mWeight() As Double
Private mTxx() As Double, mTyy() As Double, mTxy() As Double

' Проверяет положение всей осевой линии и стержня, затем готовит квадратуру.
' При ошибке ни одна часть схемы не становится доступна решателю.
Friend Sub Prepare(ByVal layout As CTransverseRebarLayout, ByVal sectionQuadrature As CCoupledSectionQuadrature)
    On Error GoTo Failed
    mCount = 0
    If layout Is Nothing Then InputError "", "Укажите фактические хомуты или подтвердите их отсутствие."
    If Not layout.IsSpecified Then InputError "", "Укажите фактические хомуты или подтвердите их отсутствие."
    If layout.Count = 0 Then Exit Sub

    Dim region As CConcreteRegion, frame As CSectionNumericalFrame
    Dim query As New CSectionGeometryQuery
    Set region = sectionQuadrature.Region
    Set frame = sectionQuadrature.Frame
    Dim path As CGeometryPath, edge As TRegionEdge
    Dim group As Long, i As Long, clearance As Double, areaPerSpacing As Double, failedSegment As String
    For group = 1 To layout.Count
        Set path = layout.Path(group)
        clearance = layout.Diameter(group) / 2#
        If Not query.PathWithinConcrete(region, path, clearance, failedSegment) Then
            InputError layout.SourceReference(group), "Группа " & path.ID & ", ветвь " & failedSegment & _
                " выходит за бетон, пересекает отверстие или её ось ближе половины диаметра к границе."
        End If
        areaPerSpacing = GEOM_PI * clearance * clearance / layout.Spacing(group)
        For i = 1 To path.SegmentCount
            path.CopySegment i, edge
            PrepareSegment sectionQuadrature, frame, edge, group, areaPerSpacing
        Next i
    Next group
    Exit Sub
Failed:
    mCount = 0
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub

' Делит путь по линиям узлов сплайнов, затем применяет четыре точки Гаусса.
' Четырёхточечная формула точна до степени 7. Пересечения узлов
' не смешивают разные полиномы, а точность дуг и косых ветвей проверяется отдельно.
Private Sub PrepareSegment(ByVal sectionQuadrature As CCoupledSectionQuadrature, _
        ByVal frame As CSectionNumericalFrame, ByRef edge As TRegionEdge, _
        ByVal group As Long, ByVal areaPerSpacing As Double)
    Dim cuts() As Double, cutCount As Long, i As Long
    sectionQuadrature.PathCuts edge, cuts, cutCount
    Dim abscissa(1 To 4) As Double, weight(1 To 4) As Double
    abscissa(1) = -0.861136311594053: abscissa(2) = -0.339981043584856
    abscissa(3) = -abscissa(2): abscissa(4) = -abscissa(1)
    weight(1) = 0.347854845137454: weight(2) = 0.652145154862546
    weight(3) = weight(2): weight(4) = weight(1)

    Dim center As Double, halfInterval As Double, u As Double, segmentLength As Double
    Dim tx As Double, ty As Double, j As Long
    segmentLength = modGeometryPrimitives.EdgeLength(edge)
    For i = 2 To cutCount
        center = (cuts(i - 1) + cuts(i)) / 2#
        halfInterval = (cuts(i) - cuts(i - 1)) / 2#
        For j = 1 To 4
            u = center + halfInterval * abscissa(j)
            If mCount = mCapacity Then ReservePoint
            mCount = mCount + 1
            modGeometryPrimitives.EdgePoint edge, u, mX(mCount), mY(mCount)
            modGeometryPrimitives.EdgeTangent edge, u, tx, ty
            tx = tx / segmentLength: ty = ty / segmentLength
            mX(mCount) = mX(mCount) - frame.OriginX
            mY(mCount) = mY(mCount) - frame.OriginY
            mWeight(mCount) = areaPerSpacing * segmentLength * halfInterval * weight(j)
            mTxx(mCount) = tx * tx: mTyy(mCount) = ty * ty: mTxy(mCount) = tx * ty
            mGroup(mCount) = group
        Next j
    Next i
End Sub

' Выделяет место для следующего блока точек. В итерациях не вызывается.
Private Sub ReservePoint()
    If mCapacity = 0 Then
        mCapacity = 32
    Else
        mCapacity = 2 * mCapacity
    End If
    ReDim Preserve mGroup(1 To mCapacity)
    ReDim Preserve mX(1 To mCapacity)
    ReDim Preserve mY(1 To mCapacity)
    ReDim Preserve mWeight(1 To mCapacity)
    ReDim Preserve mTxx(1 To mCapacity)
    ReDim Preserve mTyy(1 To mCapacity)
    ReDim Preserve mTxy(1 To mCapacity)
End Sub

' Возвращает число точек поперечной арматуры.
Friend Property Get Count() As Long
    Count = mCount
End Property

' Возвращает координаты от численного начала, вес, проекцию деформации
' и исходную группу для подготовки матриц и связанной диагностики.
Friend Sub Point(ByVal index As Long, ByRef x As Double, ByRef y As Double, _
        ByRef weight As Double, ByRef txx As Double, ByRef tyy As Double, _
        ByRef txy As Double, ByRef group As Long)
    If index < 1 Or index > mCount Then Err.Raise 9, "CTransverseRebarQuadrature", "Нет такой точки поперечной арматуры."
    x = mX(index): y = mY(index): weight = mWeight(index)
    txx = mTxx(index): tyy = mTyy(index): txy = mTxy(index): group = mGroup(index)
End Sub

' Передаёт понятную причину с исходной строкой, если она задана читателем.
Private Sub InputError(ByVal sourceReference As String, ByVal reason As String)
    If Len(sourceReference) > 0 Then reason = reason & " " & sourceReference
    Err.Raise vbObjectError + 1778, "CTransverseRebarQuadrature", reason
End Sub
