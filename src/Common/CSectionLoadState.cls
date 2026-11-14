VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CSectionLoadState"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Нагрузка сечения и перенос моментов между точками
' ==========================================================================
' Хранит N, Mx, My, T, Qx, Qy и общую точку приведения сил. Переносит моменты
' между пользовательской точкой, произвольным расчётным центром и началом
' координат модели. Если передан CMomentZeroFilter, одинаково учитывает
' практически нулевые моменты перед прочностью, несущей способностью,
' устойчивостью и расчётом трещин.

Private mN As Double ' Продольная сила во внутренних единицах, Н.
Private mTorque As Double, mQx As Double, mQy As Double ' В пользовательской системе осей, во внутренних единицах.
Private mLoadPointMx As Double ' Mx относительно точки приложения нагрузки, Н*мм.
Private mLoadPointMy As Double ' My относительно точки приложения нагрузки, Н*мм.
Private mLoadReferenceX As Double ' X точки приложения продольной силы в координатах модели, мм.
Private mLoadReferenceY As Double ' Y точки приложения продольной силы в координатах модели, мм.
Private mMomentFilter As CMomentZeroFilter ' Общий инженерный фильтр малых моментов.
Private mConcreteDepthForMx As Double ' Бетонный габарит по Y для фильтра Mx, мм.
Private mConcreteDepthForMy As Double ' Бетонный габарит по X для фильтра My, мм.

' Инициализирует нагрузку по моментам, заданным относительно точки приложения.
Public Sub Initialize(ByVal nValue As Double, ByVal loadPointMx As Double, _
        ByVal loadPointMy As Double, ByVal loadReferenceX As Double, _
        ByVal loadReferenceY As Double, Optional ByVal momentFilter As CMomentZeroFilter = Nothing, _
        Optional ByVal concreteDepthForMx As Double = 0#, Optional ByVal concreteDepthForMy As Double = 0#)
    mN = nValue
    SetSpatialComponents 0#, 0#, 0#
    mLoadReferenceX = loadReferenceX
    mLoadReferenceY = loadReferenceY
    Set mMomentFilter = momentFilter
    mConcreteDepthForMx = concreteDepthForMx
    mConcreteDepthForMy = concreteDepthForMy
    mLoadPointMx = NormalizeMx(loadPointMx)
    mLoadPointMy = NormalizeMy(loadPointMy)
End Sub

' Добавляет кручение и поперечные силы, заданные в той же точке, что N и M.
Friend Sub SetSpatialComponents(ByVal torque As Double, ByVal qx As Double, ByVal qy As Double)
    mTorque = torque: mQx = qx: mQy = qy
End Sub

' Возвращает полный винт относительно выбранного начала сечения.
' Вектор от нового начала к точке исходной нагрузки задаёт добавку r×F.
' Рабочий знак My противоположен физической компоненте момента вокруг Y.
Friend Sub LoadAboutPoint(ByVal pointX As Double, ByVal pointY As Double, ByRef components() As Double)
    Dim dx As Double, dy As Double
    dx = mLoadReferenceX - pointX: dy = mLoadReferenceY - pointY
    ReDim components(1 To 6)
    components(1) = mN
    components(2) = MxAboutPoint(pointY)
    components(3) = MyAboutPoint(pointX)
    components(4) = mTorque + dx * mQy - dy * mQx
    components(5) = mQx
    components(6) = mQy
End Sub

' Возвращает шесть усилий равновесия из известной расчётной точки
' в пользовательскую. Фильтр нулевого момента относится только к нагрузкам.
Friend Sub MoveResultantsToLoadPoint(ByVal pointX As Double, ByVal pointY As Double, ByRef components() As Double)
    Dim dx As Double, dy As Double
    dx = mLoadReferenceX - pointX: dy = mLoadReferenceY - pointY
    components(2) = components(2) - components(1) * dy
    components(3) = components(3) - components(1) * dx
    components(4) = components(4) - dx * components(6) + dy * components(5)
End Sub

' Задаёт нагрузку по моментам относительно известной расчётной точки.
' Например, после устойчивости переводит моменты из центра приведённого сечения
' обратно в пользовательскую точку приложения нагрузки.
Public Sub InitializeFromPointMoments(ByVal nValue As Double, ByVal mxAboutPoint As Double, _
        ByVal myAboutPoint As Double, ByVal pointX As Double, ByVal pointY As Double, _
        ByVal loadReferenceX As Double, ByVal loadReferenceY As Double, _
        Optional ByVal momentFilter As CMomentZeroFilter = Nothing, _
        Optional ByVal concreteDepthForMx As Double = 0#, Optional ByVal concreteDepthForMy As Double = 0#)
    Initialize nValue, _
        mxAboutPoint + nValue * (pointY - loadReferenceY), _
        myAboutPoint + nValue * (pointX - loadReferenceX), _
        loadReferenceX, loadReferenceY, momentFilter, concreteDepthForMx, concreteDepthForMy
End Sub

' Возвращает Mx относительно точки с координатой Y = pointY.
Public Function MxAboutPoint(ByVal pointY As Double) As Double
    MxAboutPoint = NormalizeMx(mLoadPointMx + mN * (mLoadReferenceY - pointY))
End Function

' Возвращает My относительно точки с координатой X = pointX.
Public Function MyAboutPoint(ByVal pointX As Double) As Double
    MyAboutPoint = NormalizeMy(mLoadPointMy + mN * (mLoadReferenceX - pointX))
End Function

' Возвращает часть Mx, создаваемую только продольной силой относительно
' заданной точки. Пользовательские моменты в это значение не входят.
Public Function AxialMxAboutPoint(ByVal pointY As Double) As Double
    AxialMxAboutPoint = NormalizeMx(mN * (mLoadReferenceY - pointY))
End Function

' Возвращает часть My, создаваемую только продольной силой относительно
' заданной точки. Метод нужен LoadPath, где момент от N масштабируется
' вместе с N, а пользовательские Mx/My могут масштабироваться отдельно.
Public Function AxialMyAboutPoint(ByVal pointX As Double) As Double
    AxialMyAboutPoint = NormalizeMy(mN * (mLoadReferenceX - pointX))
End Function

' Проверяет, является ли растягивающая сила центральной относительно заданной точки.
' Использует допуски вызывающего расчёта с учётом масштаба проверяемых усилий.
Public Function IsCentralTensionAbout(ByVal pointX As Double, ByVal pointY As Double, _
        ByVal forceTolerance As Double, ByVal momentToleranceMx As Double, _
        ByVal momentToleranceMy As Double) As Boolean
    If Not IsTension(forceTolerance) Then Exit Function
    If Abs(MxAboutPoint(pointY)) > Abs(momentToleranceMx) Then Exit Function
    If Abs(MyAboutPoint(pointX)) > Abs(momentToleranceMy) Then Exit Function
    IsCentralTensionAbout = True
End Function

' True, если продольная сила является растягивающей во внутренней системе
' знаков расчетного ядра. Пользовательские знаки к этому моменту уже должны
' быть преобразованы CUnitSystem.
Public Function IsTension(Optional ByVal forceTolerance As Double = 0#) As Boolean
    IsTension = (mN > Abs(forceTolerance))
End Function

' True, если продольная сила является сжимающей во внутренней системе знаков.
Public Function IsCompression(Optional ByVal forceTolerance As Double = 0#) As Boolean
    IsCompression = (mN < -Abs(forceTolerance))
End Function

' Возвращает сжимающую продольную силу положительным числом.
' Если сила растягивает сечение или практически равна нулю, возвращается 0.
Public Function CompressionMagnitude(Optional ByVal forceTolerance As Double = 0#) As Double
    If IsCompression(forceTolerance) Then CompressionMagnitude = -mN
End Function

' Возвращает продольную силу в знаковой форме для формул длительности:
' сжатие положительно, растяжение отрицательно. Значение не обнуляется,
' потому что нормативная формула сама должна видеть знак длительной части.
Public Function SignedCompression() As Double
    SignedCompression = -mN
End Function

' Внутренний Mx относительно начала координат CSectionModel.
Public Property Get InternalMx() As Double
    InternalMx = MxAboutPoint(0#)
End Property

' Внутренний My относительно начала координат CSectionModel.
Public Property Get InternalMy() As Double
    InternalMy = MyAboutPoint(0#)
End Property

' Продольная сила во внутренних единицах, Н.
Public Property Get N() As Double
    N = mN
End Property

' Mx относительно точки приложения нагрузки, Н*мм.
Public Property Get LoadPointMx() As Double
    LoadPointMx = mLoadPointMx
End Property

' My относительно точки приложения нагрузки, Н*мм.
Public Property Get LoadPointMy() As Double
    LoadPointMy = mLoadPointMy
End Property

' X точки приложения продольной силы в координатах модели, мм.
Public Property Get LoadReferenceX() As Double
    LoadReferenceX = mLoadReferenceX
End Property

' Y точки приложения продольной силы в координатах модели, мм.
Public Property Get LoadReferenceY() As Double
    LoadReferenceY = mLoadReferenceY
End Property

' Применяет общий фильтр к Mx.
Private Function NormalizeMx(ByVal momentValue As Double) As Double
    If mMomentFilter Is Nothing Then
        NormalizeMx = momentValue
    Else
        NormalizeMx = mMomentFilter.NormalizeMoment(momentValue, mConcreteDepthForMx)
    End If
End Function

' Применяет общий фильтр к My.
Private Function NormalizeMy(ByVal momentValue As Double) As Double
    If mMomentFilter Is Nothing Then
        NormalizeMy = momentValue
    Else
        NormalizeMy = mMomentFilter.NormalizeMoment(momentValue, mConcreteDepthForMy)
    End If
End Function
