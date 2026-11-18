VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledMaterialFailure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемый контекст последнего численного отказа материальной точки
' ==========================================================================
' Это диагностические данные, не принятое НДС и не критерий разрушения.
' Деформации безразмерные, эффективные напряжения в МПа. Две координаты
' относятся к выбранному закону; для CDP они не заменяют полную историю.
' Порядок xx, yy, zz, xy, xz, yz; сдвиговые деформации инженерные.

Private mPoint As Long, mEpoch As Long, mComponent As Long, mMode As ECoupledTangentMode
Private mStage As Long, mLaw As String
Private mRequested(1 To 6) As Double, mPriorStrain(1 To 6) As Double, mPriorEffective(1 To 6) As Double
Private mTension As Double, mCompression As Double

' Сохраняет копии входа и принятого состояния только после фактического отказа.
Friend Sub Initialize(ByVal point As Long, ByVal epoch As Long, ByVal stage As Long, ByVal component As Long, _
        ByVal mode As ECoupledTangentMode, ByVal law As String, ByRef requested() As Double, _
        ByRef priorStrain() As Double, ByRef priorEffective() As Double, ByVal tension As Double, ByVal compression As Double)
    If mPoint <> 0 Then Err.Raise 5, "CCoupledMaterialFailure", "Сохранённый контекст материального отказа неизменяем."
    Dim i As Long
    For i = 1 To 6
        mRequested(i) = requested(i): mPriorStrain(i) = priorStrain(i): mPriorEffective(i) = priorEffective(i)
    Next i
    mPoint = point: mEpoch = epoch: mStage = stage: mComponent = component: mMode = mode: mLaw = law
    mTension = tension: mCompression = compression
End Sub

' Возвращает номер интеграционной точки, а не номер бетонного элемента вывода.
Friend Property Get Point() As Long
    Point = mPoint
End Property

' Идентифицирует принятую историю, от которой вычислялась неудачная проба.
Friend Property Get Epoch() As Long
    Epoch = mEpoch
End Property

' Отличает отказ обновления напряжений от отказа построения касательной.
Friend Property Get Stage() As String
    If mStage = 1 Then Stage = "Обновление напряжений" Else Stage = "Построение касательной"
End Property

' Возвращает столбец неудачной касательной; ноль относится к обновлению напряжений.
Friend Property Get Component() As Long
    Component = mComponent
End Property

' Сохраняет точный запрос режима касательной, включая согласованную ветвь.
Friend Property Get TangentMode() As ECoupledTangentMode
    TangentMode = mMode
End Property

' Указывает фактически применённый закон материала.
Friend Property Get LawIdentifier() As String
    LawIdentifier = mLaw
End Property

' Копирует только сохранённые числовые данные, не вычисляя материал повторно.
Friend Sub CopyState(ByRef requested() As Double, ByRef priorStrain() As Double, ByRef priorEffective() As Double, _
        ByRef tension As Double, ByRef compression As Double)
    ReDim requested(1 To 6): ReDim priorStrain(1 To 6): ReDim priorEffective(1 To 6)
    Dim i As Long
    For i = 1 To 6
        requested(i) = mRequested(i): priorStrain(i) = mPriorStrain(i): priorEffective(i) = mPriorEffective(i)
    Next i
    tension = mTension: compression = mCompression
End Sub
