VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CLoadPathVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Линейная зависимость усилий от множителя нагрузки
' ==========================================================================
' Хранит постоянную и масштабируемую части Target(lambda) = Offset + lambda*Base
' Вычисляет N, Mx, My, T, Qx и Qy при заданном lambda.
' Инженерный смысл траектории определяет вызывающий расчёт.

Private mNOffset As Double ' Постоянная часть N, Н.
Private mNBase As Double ' Масштабируемая часть N, Н.
Private mMxOffset As Double ' Постоянная часть Mx относительно начала модели, Н*мм.
Private mMxBase As Double ' Масштабируемая часть Mx, Н*мм.
Private mMyOffset As Double ' Постоянная часть My относительно начала модели, Н*мм.
Private mMyBase As Double ' Масштабируемая часть My, Н*мм.

Private mTorqueOffset As Double ' Постоянная часть Torque, Н*мм.
Private mTorqueBase As Double ' Масштабируемая часть Torque, Н*мм.
Private mQxOffset As Double ' Постоянная часть Qx, Н.
Private mQxBase As Double ' Масштабируемая часть Qx, Н.
Private mQyOffset As Double ' Постоянная часть Qy, Н.
Private mQyBase As Double ' Масштабируемая часть Qy, Н.

' Задает траекторию напрямую через постоянные и масштабируемые компоненты.
Public Sub Initialize(ByVal nOffset As Double, ByVal nBase As Double, _
        ByVal mxOffset As Double, ByVal mxBase As Double, _
        ByVal myOffset As Double, ByVal myBase As Double, _
        Optional ByVal torqueOffset As Double = 0#, Optional ByVal torqueBase As Double = 0#, _
        Optional ByVal qxOffset As Double = 0#, Optional ByVal qxBase As Double = 0#, _
        Optional ByVal qyOffset As Double = 0#, Optional ByVal qyBase As Double = 0#)
    mNOffset = nOffset
    mNBase = nBase
    mMxOffset = mxOffset
    mMxBase = mxBase
    mMyOffset = myOffset
    mMyBase = myBase
    mTorqueOffset = torqueOffset: mTorqueBase = torqueBase
    mQxOffset = qxOffset: mQxBase = qxBase
    mQyOffset = qyOffset: mQyBase = qyBase
End Sub

' Задаёт постоянную продольную силу и масштабируемый момент относительно выбранной точки.
' Решатель использует начало координат модели, поэтому момент от эксцентриситета N
' включается в постоянную часть Offset.
Public Sub InitializeFixedNAndMomentAboutPoint(ByVal nValue As Double, _
        ByVal momentMxAboutPoint As Double, ByVal momentMyAboutPoint As Double, _
        ByVal pointX As Double, ByVal pointY As Double)
    Initialize nValue, 0#, _
        nValue * pointY, momentMxAboutPoint, _
        nValue * pointX, momentMyAboutPoint
End Sub

' Возвращает полный расчетный вектор для заданного lambda.
Public Sub Target(ByVal lambdaValue As Double, ByRef nValue As Double, _
        ByRef mxValue As Double, ByRef myValue As Double)
    nValue = TargetN(lambdaValue)
    mxValue = TargetMx(lambdaValue)
    myValue = TargetMy(lambdaValue)
End Sub

' Возвращает N(lambda).
Public Function TargetN(ByVal lambdaValue As Double) As Double
    TargetN = mNOffset + lambdaValue * mNBase
End Function

' Возвращает Mx(lambda).
Public Function TargetMx(ByVal lambdaValue As Double) As Double
    TargetMx = mMxOffset + lambdaValue * mMxBase
End Function

' Возвращает My(lambda).
Public Function TargetMy(ByVal lambdaValue As Double) As Double
    TargetMy = mMyOffset + lambdaValue * mMyBase
End Function

' Постоянная часть N, Н.
Public Property Get NOffset() As Double
    NOffset = mNOffset
End Property

' Масштабируемая часть N, Н.
Public Property Get NBase() As Double
    NBase = mNBase
End Property

' Постоянная часть Mx относительно начала модели, Н*мм.
Public Property Get MxOffset() As Double
    MxOffset = mMxOffset
End Property

' Масштабируемая часть Mx, Н*мм.
Public Property Get MxBase() As Double
    MxBase = mMxBase
End Property

' Постоянная часть My относительно начала модели, Н*мм.
Public Property Get MyOffset() As Double
    MyOffset = mMyOffset
End Property

' Масштабируемая часть My, Н*мм.
Public Property Get MyBase() As Double
    MyBase = mMyBase
End Property

' Возвращает независимые копии постоянной части и направления в порядке
' N, Mx, My, T, Qx, Qy. Единицы внутренних сил и моментов.
Friend Sub Components(ByRef offset() As Double, ByRef direction() As Double)
    ReDim offset(1 To 6): ReDim direction(1 To 6)
    offset(1) = mNOffset: offset(2) = mMxOffset: offset(3) = mMyOffset
    offset(4) = mTorqueOffset: offset(5) = mQxOffset: offset(6) = mQyOffset
    direction(1) = mNBase: direction(2) = mMxBase: direction(3) = mMyBase
    direction(4) = mTorqueBase: direction(5) = mQxBase: direction(6) = mQyBase
End Sub

' Вычисляет шесть компонентов в заданной точке траектории.
Friend Sub TargetSix(ByVal lambdaValue As Double, ByRef load() As Double)
    Dim direction() As Double, i As Long
    Components load, direction
    For i = 1 To 6
        load(i) = load(i) + lambdaValue * direction(i)
    Next i
End Sub
