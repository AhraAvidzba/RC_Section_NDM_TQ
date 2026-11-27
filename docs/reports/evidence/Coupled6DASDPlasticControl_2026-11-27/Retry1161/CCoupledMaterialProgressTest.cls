VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CCoupledMaterialProgressTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: точная запись принятой траектории материального ограничения
' ==========================================================================
Implements ICoupledContinuationProgress
Private mPath As String, mHeader() As Double, mOffset() As Double, mDirection() As Double
Private mCalls As Long

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует закреплённые единицы операторов и реальный путь шести усилий.
Friend Sub TestInitialize(ByVal path As String, ByRef header() As Double, ByRef offset() As Double, ByRef direction() As Double)
    mPath = path: mHeader = header: mOffset = offset: mDirection = direction
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет последний полный принятый снимок; пробные состояния в историю не добавляются.
Private Sub ICoupledContinuationProgress_Report(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace, ByVal meta As CResultMeta)
    If fields.Count <> trace.Count Then Err.Raise 5, "CCoupledMaterialProgressTest", "Не согласованы принятые поля и траектория."
    Dim handle As Integer, i As Long, field() As Double, pair(1 To 2) As Double, count(1 To 1) As Double
    On Error GoTo Failed
    mCalls = mCalls + 1
    handle = FreeFile: Open mPath For Output As #handle: Close #handle
    handle = FreeFile: Open mPath For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMMAINHIST00003": Put #handle, , magic
    TestWriteVector handle, mHeader: TestWriteVector handle, mOffset: TestWriteVector handle, mDirection
    count(1) = fields.Count: TestWriteVector handle, count
    For i = 1 To fields.Count
        trace.ReadPoint i, pair(1), pair(2): field = fields.Item(i)
        If UBound(field) <> 60 Then Err.Raise 5, "CCoupledMaterialProgressTest", "Изменился размер контрольного поля."
        TestWriteVector handle, pair: TestWriteVector handle, field
    Next i
    Close #handle: handle = 0
    FileCopy mPath, mPath & ".Accepted." & Format$(mCalls, "000000") & ".bin" ' Копия читается после закрытия собственного файла; внешнего читателя и конкуренции за доступ нет.
    Exit Sub
Failed:
    Dim reason As String: reason = Err.Source & " " & Err.Description
    On Error Resume Next
    If handle <> 0 Then Close #handle
    On Error GoTo 0
    Err.Raise 5, "CCoupledMaterialProgressTest", reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает точные Double с явным числом элементов без дескриптора VBA-массива.
Private Sub TestWriteVector(ByVal handle As Integer, ByRef values() As Double)
    Dim size As Long, i As Long, value As Double: size = UBound(values) - LBound(values) + 1
    Put #handle, , size
    For i = LBound(values) To UBound(values): value = values(i): Put #handle, , value: Next i
End Sub
