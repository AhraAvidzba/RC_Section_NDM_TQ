VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CHybridDifferenceObserverTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: читающий наблюдатель разностного пересечения аналитического излома
' ==========================================================================
' Не вызывает функцию системы и не меняет переданные разностные координаты.
Implements ICoupledDifferenceObserver
Private mAccept As Boolean, mAcceptOnce As Boolean, mCalls As Long, mAccepted As Long
Private mBefore() As Double, mAfter() As Double

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сбрасывает наблюдение и разрешает все пересечения второй координаты либо только первое.
Friend Sub InitializeTest(ByVal accept As Boolean, Optional ByVal acceptOnce As Boolean = False)
    mAccept = accept: mAcceptOnce = acceptOnce: mCalls = 0: mAccepted = 0
    Erase mBefore: Erase mAfter
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число фактических проверок полностью вычисленных разностных пар.
Friend Property Get CallsTest() As Long
    CallsTest = mCalls
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число разрешённых пересечений, независимо от счётчиков решателя.
Friend Property Get AcceptedTest() As Long
    AcceptedTest = mAccepted
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сопоставляет сохранённые Double первой разрешённой пары с опубликованной диагностикой.
Friend Function MatchesFirstTest(ByRef before() As Double, ByRef after() As Double) As Boolean
    If mAccepted = 0 Then Exit Function
    Dim i As Long
    If UBound(before) <> UBound(mBefore) Or UBound(after) <> UBound(mAfter) Then Exit Function
    For i = 1 To UBound(mBefore)
        If before(i) <> mBefore(i) Or after(i) <> mAfter(i) Then Exit Function
    Next i
    MatchesFirstTest = True
End Function

' Реализация тестового интерфейса: разрешает только фактическое пересечение излома второй координаты.
Private Function ICoupledDifferenceObserver_AcceptBracket(ByRef before() As Double, ByRef after() As Double) As Boolean
    mCalls = mCalls + 1
    If Not mAccept Or (mAcceptOnce And mAccepted > 0) Then Exit Function
    If before(2) >= 0# Or after(2) <= 0# Or before(1) <> after(1) Then Exit Function
    If mAccepted = 0 Then mBefore = before: mAfter = after
    mAccepted = mAccepted + 1
    ICoupledDifferenceObserver_AcceptBracket = True
End Function
