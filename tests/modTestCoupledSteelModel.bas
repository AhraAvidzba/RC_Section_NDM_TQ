Attribute VB_Name = "modTestCoupledSteelModel"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: одноосная арматура в связанной постановке
' ==========================================================================
' Проверяет аналитические ветви, разгрузку, обратную текучесть и откат.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сверяет растяжение и сжатие на асимметричной диаграмме и rollback.
Public Function RunCoupledSteelTests() As String
    On Error GoTo Failed
    Const E As Double = 200000#
    Dim steel As New CCoupledSteelModel
    steel.Initialize E, 400#, 350#, 1500#, 1000#, 0.025, 0.018

    Dim s As Double, tangent As Double, plastic As Double, backstress As Double
    Dim nextPlastic As Double, nextBackstress As Double
    Dim passed As Long, failed As Long, report As String
    steel.EvaluateTrial 0.001, 0#, 0#, s, tangent, plastic, backstress
    TestRecord Abs(s - 200#) < 1E-10 And tangent = E And plastic = 0# And backstress = 0#, _
        "steel elastic branch", passed, failed, report
    steel.EvaluateTrial 0.008, 0#, 0#, s, tangent, plastic, backstress
    TestRecord Abs(s - 409#) < 1E-10 And Abs(tangent - 1500#) < 1E-10, _
        "steel tension hardening", passed, failed, report
    TestRecord Abs(s - E * (0.008 - plastic)) < 1E-10 And Abs(s - backstress - 400#) < 1E-10, _
        "steel tension consistency", passed, failed, report

    steel.EvaluateTrial 0.007, plastic, backstress, s, tangent, nextPlastic, nextBackstress
    TestRecord Abs(s - 209#) < 1E-10 And tangent = E And nextPlastic = plastic And nextBackstress = backstress, _
        "steel elastic unloading freezes history", passed, failed, report
    steel.EvaluateTrial 0#, plastic, backstress, s, tangent, nextPlastic, nextBackstress
    TestRecord Abs(tangent - 1000#) < 1E-10 And Abs(s - nextBackstress + 350#) < 1E-10, _
        "steel reverse yield", passed, failed, report
    steel.EvaluateTrial 0.007, plastic, backstress, s, tangent, nextPlastic, nextBackstress
    TestRecord Abs(s - 209#) < 1E-10 And tangent = E And nextPlastic = plastic And nextBackstress = backstress, _
        "unaccepted reverse step leaves committed history intact", passed, failed, report

    steel.EvaluateTrial -0.008, 0#, 0#, s, tangent, nextPlastic, nextBackstress
    TestRecord Abs(s + 356.25) < 1E-10 And Abs(tangent - 1000#) < 1E-10, _
        "steel asymmetric compression hardening", passed, failed, report
    TestRecord Abs(steel.StrainUtilization(0.0125) - 0.5) < 1E-12 And _
        Abs(steel.StrainUtilization(-0.018) - 1#) < 1E-12, "steel sign dependent strain limits", passed, failed, report

    RunCoupledSteelTests = report & "TOTAL_COUPLED_STEEL: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunCoupledSteelTests = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает понятное имя и счётчики для каждого сравнения.
Private Sub TestRecord(ByVal ok As Boolean, ByVal name As String, _
        ByRef passed As Long, ByRef failed As Long, ByRef report As String)
    If ok Then
        passed = passed + 1
        report = report & "OK: " & name & vbCrLf
    Else
        failed = failed + 1
        report = report & "FAIL: " & name & vbCrLf
    End If
End Sub
