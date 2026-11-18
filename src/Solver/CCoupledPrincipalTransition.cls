VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledPrincipalTransition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' Неизменяемый контекст локального перехода через знаковую границу бетона
' ==========================================================================
' Общая точка и входящее равновесие вычисляются от одной принятой истории.
' Исходящая сторона требует отдельного решения после принятия входящей.
' Контекст не разрешает произвольный разворот и не определяет разрушение.

Private mRevision As Long, mPoint As Long, mSide As Long, mPrincipal As Long
Private mSign As Double, mGap As Double, mReferenceIncrement As Double
Private mFaceLambda As Double, mIncomingLambda As Double
Private mOrigin() As Double, mFace() As Double, mIncoming() As Double

' Сохраняет независимые копии полей только после проверки общего и входящего равновесий.
Friend Sub Initialize(ByVal revision As Long, ByVal point As Long, ByVal side As Long, ByVal principal As Long, _
        ByVal incomingSign As Double, ByVal gap As Double, ByVal referenceIncrement As Double, _
        ByVal faceLambda As Double, ByVal incomingLambda As Double, ByRef origin() As Double, _
        ByRef face() As Double, ByRef incoming() As Double)
    If mPoint <> 0 Then Err.Raise 5, "CCoupledPrincipalTransition", "Контекст материального перехода неизменяем."
    If revision < 0 Or point < 1 Or Abs(side) > 1 Or principal < 1 Or principal > 3 Or _
            Abs(incomingSign) <> 1# Or gap <= 0# Or referenceIncrement <= 0# Or _
            faceLambda <= 0# Or incomingLambda <= 0# Then _
        Err.Raise 5, "CCoupledPrincipalTransition", "Непригодны координата или история материального перехода."
    If LBound(origin) <> 1 Or LBound(face) <> 1 Or LBound(incoming) <> 1 Or _
            UBound(origin) <> UBound(face) Or UBound(origin) <> UBound(incoming) Then _
        Err.Raise 5, "CCoupledPrincipalTransition", "Поля материального перехода должны иметь одинаковую размерность."
    mRevision = revision: mPoint = point: mSide = side: mPrincipal = principal
    mSign = incomingSign: mGap = gap: mReferenceIncrement = referenceIncrement
    mFaceLambda = faceLambda: mIncomingLambda = incomingLambda
    mOrigin = origin: mFace = face: mIncoming = incoming
End Sub

' Возвращает интеграционную точку выбранной материальной границы.
Friend Property Get Point() As Long
    Point = mPoint
End Property

' Возвращает центральную либо одну из двух продольных станций.
Friend Property Get Side() As Long
    Side = mSide
End Property

' Возвращает индекс главного напряжения прогноза от старшего к младшему.
Friend Property Get PrincipalIndex() As Long
    PrincipalIndex = mPrincipal
End Property

' Возвращает знак входящей стороны; исходящая имеет противоположный знак.
Friend Property Get IncomingSign() As Double
    IncomingSign = mSign
End Property

' Возвращает положительный зазор от нулевого главного напряжения, МПа.
Friend Property Get PrincipalGap() As Double
    PrincipalGap = mGap
End Property

' Возвращает различимое механическое приращение перед локальным событием.
Friend Property Get ReferenceIncrement() As Double
    ReferenceIncrement = mReferenceIncrement
End Property

' Проверяет исходную ревизию и точное поле без округления ключа истории.
Friend Function MatchesOrigin(ByVal revision As Long, ByRef field() As Double) As Boolean
    If mPoint = 0 Or revision <> mRevision Then Exit Function
    MatchesOrigin = SameField(field, mOrigin)
End Function

' Разрешает продолжение только после единственного принятия именно входящего поля.
Friend Function MatchesIncoming(ByVal revision As Long, ByRef field() As Double) As Boolean
    If mPoint = 0 Or revision <> mRevision + 1 Then Exit Function
    MatchesIncoming = SameField(field, mIncoming)
End Function

' Копирует общую точку, не передавая изменяемую ссылку на контекст.
Friend Sub CopyFace(ByRef field() As Double, ByRef lambda As Double)
    field = mFace: lambda = mFaceLambda
End Sub

' Копирует входящее равновесие для проверки принятого материального состояния.
Friend Sub CopyIncoming(ByRef field() As Double, ByRef lambda As Double)
    field = mIncoming: lambda = mIncomingLambda
End Sub

' Сравнивает каждую координату, включая множители кинематических связей.
Private Function SameField(ByRef first() As Double, ByRef second() As Double) As Boolean
    If LBound(first) <> LBound(second) Or UBound(first) <> UBound(second) Then Exit Function
    Dim i As Long
    For i = LBound(first) To UBound(first): If first(i) <> second(i) Then Exit Function
    Next i
    SameField = True
End Function
