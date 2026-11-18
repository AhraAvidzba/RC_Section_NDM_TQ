VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledProgressObserverTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: промежуточная траектория и проверка изоляции наблюдателя
' ==========================================================================
' Текстовый протокол хранит отдельные законченные снимки. При тайм-ауте
' незавершённая последняя запись не заменяет предыдущий законченный снимок.
Implements ICoupledContinuationProgress
Private mPath As String, mMutate As Boolean, mCount As Long

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Настраивает запись только в частный тестовый каталог либо проверку изменения копий.
Friend Sub TestInitialize(ByVal path As String, Optional ByVal mutateCopies As Boolean = False)
    mPath = path: mMutate = mutateCopies: mCount = 0
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает число фактически полученных контрольных снимков.
Friend Property Get TestCallCount() As Long
    TestCallCount = mCount
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Реализует рабочий контракт наблюдателя; меняет только полученные копии в проверочном режиме.
Private Sub ICoupledContinuationProgress_Report(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, ByVal meta As CResultMeta)
    mCount = mCount + 1
    If mMutate Then
        Dim changed() As Double: changed = fields(1): changed(1) = 1000000#
        fields.Remove 1: fields.Add changed
        trace.Add 1#, 1000000#
        meta.SetResult rsCheckFailed, rcPhysicalLimitExceeded, rkCapacity, "Test observer changed its copy"
        Exit Sub
    End If
    If Len(mPath) = 0 Then Exit Sub
    If fields.Count <> trace.Count Then Err.Raise 5, "CCoupledProgressObserverTest", "Не согласованы копии контрольного снимка."
    Dim handle As Integer, i As Long, j As Long, field() As Double, fraction As Double, factor As Double, line As String
    On Error GoTo Failed
    handle = FreeFile: Open mPath For Append As #handle
    Print #handle, "CHECKPOINT;" & CStr(mCount) & ";" & CStr(fields.Count) & ";" & CStr(meta.ResultCode)
    Print #handle, "DETAILS;" & Replace(Replace(meta.DiagnosticDetails, vbCr, " "), vbLf, " ")
    For i = 1 To fields.Count
        field = fields(i): trace.ReadPoint i, fraction, factor
        line = "P;" & CStr(i) & ";" & Trim$(Str$(fraction)) & ";" & Trim$(Str$(factor))
        For j = 1 To UBound(field): line = line & ";" & Trim$(Str$(field(j))): Next j
        Print #handle, line
    Next i
    Print #handle, "CHECKPOINT_END;" & CStr(mCount)
    Close #handle
    Exit Sub
Failed:
    Dim number As Long, reason As String: number = Err.Number: reason = Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    Err.Raise number, "CCoupledProgressObserverTest", "Не сохранён контрольный снимок: " & mPath & ". " & reason
End Sub
