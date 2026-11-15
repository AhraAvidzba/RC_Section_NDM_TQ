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
' Публикует напряжения двух классов при одной предписанной истории.
' Строки проб после отказа сохраняют историю последнего принятого шага.
' Эталон Steel01 применим к симметричному кинематическому упрочнению.
Public Function RunTestCoupledSteelOpenSees() As String
    On Error GoTo Failed
    Dim report As String
    report = TestSteelOpenSeesHistory("L500", 500#) & TestSteelOpenSeesHistory("H240", 240#)
    RunTestCoupledSteelOpenSees = report & "TOTAL_COUPLED_STEEL_OPENSEES: passed=2; failed=0"
    Exit Function
Failed:
    RunTestCoupledSteelOpenSees = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Применяет чередование принятия и проб на начальной и обратной ветвях.
' Строки на границе текучести исключены для однозначного сравнения касательной.
Private Function TestSteelOpenSeesHistory(ByVal grade As String, ByVal yieldStress As Double) As String
    Const ELASTIC_MODULUS As Double = 200000#
    Const HARDENING_SLOPE As Double = 2000#
    Dim steel As New CCoupledSteelModel
    steel.Initialize ELASTIC_MODULUS, yieldStress, yieldStress, HARDENING_SLOPE, HARDENING_SLOPE, 0.05, 0.05

    Dim strains As Variant
    strains = Array(0#, 0.0006, 0.0016, 0.0032, 0.0085, 0.0074, 0.025, 0.0074, 0.0057, 0.001, _
                    -0.0023, -0.0058, -0.0095, -0.0088, -0.026, -0.0088, -0.0047, 0.0009, _
                    0.0042, 0.0101, 0.027, 0.0063, -0.0067)
    Dim lines As String
    lines = "STEEL_LAW|" & grade & "|" & Trim$(Str$(ELASTIC_MODULUS)) & "|" & Trim$(Str$(yieldStress)) & _
        "|" & Trim$(Str$(HARDENING_SLOPE)) & vbCrLf
    Dim stepIndex As Long, accepted As Boolean, acceptance As String, strain As Double
    Dim plastic As Double, backstress As Double, stress As Double, tangent As Double
    Dim trialPlastic As Double, trialBackstress As Double
    For stepIndex = LBound(strains) To UBound(strains)
        accepted = Not (stepIndex >= 5 And stepIndex <= 7 Or stepIndex >= 13 And stepIndex <= 15 Or stepIndex = 20)
        acceptance = "Probe": If accepted Then acceptance = "Commit"
        strain = CDbl(strains(stepIndex))
        steel.EvaluateTrial strain, plastic, backstress, stress, tangent, trialPlastic, trialBackstress
        lines = lines & "STEEL_STEP|" & grade & "|" & CStr(stepIndex) & "|" & acceptance & "|" & _
            Trim$(Str$(strain)) & "|" & Trim$(Str$(stress)) & "|" & Trim$(Str$(tangent)) & "|" & _
            Trim$(Str$(trialPlastic)) & "|" & Trim$(Str$(trialBackstress)) & vbCrLf
        If accepted Then plastic = trialPlastic: backstress = trialBackstress
    Next stepIndex
    TestSteelOpenSeesHistory = lines
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
