VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledDiscretization"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемая дискретизация пространственного поля и бетонной квадратуры
' ==========================================================================
' Интервалы относятся к кубическому базису по осям сечения. Пять узлов
' Гаусса заменяют только целые интервалы внутри бетона; у границ остаются
' центры и точные площади контурных ячеек. Физические параметры не задаёт.

Private mSpansX As Long, mSpansY As Long, mPoints As Long
Private mFullSpanGauss5 As Boolean, mReady As Boolean

' Сохраняет явные параметры один раз, чтобы кэш операторов не зависел от изменяемой настройки.
Friend Sub Initialize(ByVal spansX As Long, ByVal spansY As Long, ByVal pointsPerSpan As Long, _
        Optional ByVal useFullSpanGauss5 As Boolean = False)
    If mReady Then Err.Raise 5, "CCoupledDiscretization", "Подготовленная дискретизация неизменяема. Создайте новый набор."
    If spansX < 1 Or spansY < 1 Or pointsPerSpan < 4 Then Err.Raise 5, "CCoupledDiscretization", _
        "Задайте хотя бы один интервал базиса и не менее четырёх делений квадратуры на интервал."
    If useFullSpanGauss5 And pointsPerSpan <> 5 Then Err.Raise 5, "CCoupledDiscretization", _
        "Для квадратуры Гаусса по целому интервалу нужны пять узлов по каждой оси."
    mSpansX = spansX: mSpansY = spansY: mPoints = pointsPerSpan
    mFullSpanGauss5 = useFullSpanGauss5: mReady = True
End Sub

' Подтверждает завершённую подготовку перед подключением к расчёту.
Friend Property Get IsReady() As Boolean
    IsReady = mReady
End Property

' Возвращает число интервалов кубического базиса по X.
Friend Property Get SpansX() As Long
    SpansX = mSpansX
End Property

' Возвращает число интервалов кубического базиса по Y.
Friend Property Get SpansY() As Long
    SpansY = mSpansY
End Property

' Возвращает число делений квадратуры внутри интервала по каждой оси.
Friend Property Get PointsPerSpan() As Long
    PointsPerSpan = mPoints
End Property

' Указывает замену полных внутренних интервалов тензорным правилом Гаусса.
Friend Property Get UseFullSpanGauss5() As Boolean
    UseFullSpanGauss5 = mFullSpanGauss5
End Property
