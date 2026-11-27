Attribute VB_Name = "modTestCoupledASDCoordinates"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: полные эквивалентные деформации и изоляция истории ASD
' ==========================================================================
' Материальные пути и результаты сохраняются для независимого нативного
' ASDConcrete3D. Совпадение мер не задаёт критерий разрушения сечения.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет чтение принятой истории и кандидата при обоих алгоритмах ASD.
Public Function RunTestCoupledASDCoordinates(ByVal folder As String) As String
    On Error GoTo Failed
    Dim report As String, passed As Long, failedCount As Long
    TestUnsupportedCoordinates passed, failedCount, report
    TestCoordinateHistory folder, True, passed, failedCount, report
    TestCoordinateHistory folder, False, passed, failedCount, report
    RunTestCoupledASDCoordinates = report & "TOTAL_COUPLED_ASD_COORDINATES_TEST: passed=" & passed & "; failed=" & failedCount
    Exit Function
Failed:
    RunTestCoupledASDCoordinates = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Готовит явно заданные ветви с пластичностью и повреждением обоих знаков.
Private Function TestCoordinateFlow(ByVal nativeStep As Boolean) As CCoupledASDFlow
    Dim tension As New CCoupledConcreteBackbone, compression As New CCoupledConcreteBackbone
    Dim e(1 To 5) As Double, s(1 To 5) As Double, effective(1 To 5) As Double
    e(2) = 0.0001: e(3) = 0.0003: e(4) = 0.001: e(5) = 0.004
    s(2) = 3#: s(3) = 2#: s(4) = 0.8: s(5) = 0.05
    effective(2) = 3#: effective(3) = 5#: effective(4) = 8#: effective(5) = 15#
    tension.Initialize e, s, effective, 5, 0.00000075 ' ASDConcrete3D: 1E-6 от минимального ненулевого изменения ординаты.
    e(2) = 0.0005: e(3) = 0.002: e(4) = 0.004: e(5) = 0.01
    s(2) = 15#: s(3) = 30#: s(4) = 20#: s(5) = 2#
    effective(2) = 15#: effective(3) = 40#: effective(4) = 60#: effective(5) = 90#
    compression.Initialize e, s, effective, 5, 0.00001 ' Нижний уровень вычисляется отдельно для каждой табличной ветви.
    Dim flow As New CCoupledASDFlow
    flow.Initialize tension, compression, 30000#, 0.2, nativeStep
    Set TestCoordinateFlow = flow
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Не позволяет выдавать пластические координаты CDP за полные деформации ASD.
Private Sub TestUnsupportedCoordinates(ByRef passed As Long, ByRef failedCount As Long, ByRef report As String)
    Dim concrete As New CCoupledConcreteResponse, strain(1 To 6) As Double
    Dim tension As Double, compression As Double, ready As Boolean
    tension = 7#: compression = 9#
    ready = concrete.TryReadAcceptedASDStrains(1, tension, compression)
    TestCheck Not ready And tension = 0# And compression = 0#, "unprepared law has no ASD coordinates", passed, failedCount, report
    concrete.Initialize 1, 30000#, 0.2, 16#, 30#, 3#, 0.002, 0.12, 80#, 125#
    tension = 7#: compression = 9#
    ready = concrete.TryReadAcceptedASDStrains(1, tension, compression)
    TestCheck Not ready And tension = 0# And compression = 0#, "CDP accepted variables are not ASD total strains", passed, failedCount, report
    tension = 7#: compression = 9#: strain(3) = 0.0002
    ready = concrete.TryReadASDStrains(1, strain, tension, compression)
    TestCheck Not ready And tension = 0# And compression = 0#, "CDP candidate variables are not ASD total strains", passed, failedCount, report
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет реальные напряжения и координаты до и после принятия либо отклонения шага.
Private Sub TestCoordinateHistory(ByVal folder As String, ByVal nativeStep As Boolean, ByRef passed As Long, _
        ByRef failedCount As Long, ByRef report As String)
    On Error GoTo Failed
    Dim inputHandle As Integer, outputHandle As Integer, line As String, row As Variant, path As String, mode As String
    Dim concrete As CCoupledConcreteResponse, response As ICoupledPointResponse, flow As CCoupledASDFlow
    Dim strain(1 To 6) As Double, probe(1 To 6) As Double, saved(1 To 6) As Double
    Dim stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim oldT As Double, oldC As Double, trialT As Double, trialC As Double, readT As Double, readC As Double
    Dim expectedT As Double, expectedC As Double, i As Long, number As Long, reason As String, ready As Boolean
    If nativeStep Then mode = "Native" Else mode = "Closed"
    Set flow = TestCoordinateFlow(nativeStep)
    inputHandle = FreeFile: Open folder & "/History.csv" For Input As #inputHandle
    outputHandle = FreeFile: Open folder & "/" & mode & ".csv" For Output As #outputHandle
    Do While Not EOF(inputHandle)
        Line Input #inputHandle, line: row = Split(line, ";")
        If UBound(row) <> 8 Then Err.Raise 5, , "Неверный размер строки материального пути."
        If path <> row(0) Then
            path = row(0)
            Set concrete = New CCoupledConcreteResponse: concrete.InitializeASD 2, flow
            Set response = concrete
            TestCheck concrete.TryReadAcceptedASDStrains(1, readT, readC) And readT = 0# And readC = 0#, _
                mode & "/" & path & " virgin history", passed, failedCount, report
            On Error Resume Next
            ready = concrete.TryReadAcceptedASDStrains(0, readT, readC): number = Err.Number
            reason = Err.Source & "|" & Err.Description: Err.Clear
            On Error GoTo Failed
            TestCheck number = vbObjectError + 1783 And reason = "CCoupledConcreteResponse|Точка эквивалентных деформаций ASD находится вне подготовленного материала.", _
                mode & " invalid accepted point rejected with material error", passed, failedCount, report
            On Error Resume Next
            ready = concrete.TryReadASDStrains(3, strain, readT, readC): number = Err.Number
            reason = Err.Source & "|" & Err.Description: Err.Clear
            On Error GoTo Failed
            TestCheck number = vbObjectError + 1783 And reason = "CCoupledConcreteResponse|Точка эквивалентных деформаций ASD находится вне подготовленного материала.", _
                mode & " invalid candidate point rejected with material error", passed, failedCount, report
        End If
        For i = 1 To 6: strain(i) = Val(row(i + 1)): saved(i) = strain(i): Next i
        ready = concrete.TryReadAcceptedASDStrains(1, oldT, oldC)
        TestCheck ready, mode & "/" & path & " accepted measure available", passed, failedCount, report
        ready = concrete.TryReadASDStrains(1, strain, trialT, trialC)
        If Not ready Then Err.Raise 5, , "Не получены полные эквивалентные деформации кандидата ASD."
        TestCheck trialT >= oldT And trialC >= oldC, mode & "/" & path & " irreversible equivalent strains", passed, failedCount, report
        For i = 1 To 6
            TestCheck strain(i) = saved(i), mode & " candidate inputHandle unchanged " & CStr(i), passed, failedCount, report
        Next i
        If response.EvaluatePoint(1, strain, stress, tangent, ctmStressOnly) <> sfcNone Then _
            Err.Raise 5, , "Материал не выполнил контрольную пробу ASD."
        For i = 1 To 6: probe(i) = strain(i) + 0.000003 * CDbl(i): Next i
        ready = concrete.TryReadASDStrains(1, probe, readT, readC)
        TestCheck ready, mode & " additional read-only candidate exists", passed, failedCount, report
        ready = concrete.TryReadAcceptedASDStrains(1, readT, readC)
        TestCheck ready And readT = oldT And readC = oldC, mode & " reads leave accepted state unchanged", passed, failedCount, report
        expectedT = oldT: expectedC = oldC
        If CLng(row(8)) = 1 Then
            response.AcceptTrial: expectedT = trialT: expectedC = trialC
        End If
        ready = concrete.TryReadAcceptedASDStrains(1, readT, readC)
        TestCheck ready And readT = expectedT And readC = expectedC, mode & " reads preserve trial buffer and accept/reject", passed, failedCount, report
        Print #outputHandle, line & ";" & Trim$(Str$(trialT)) & ";" & Trim$(Str$(trialC)) & ";" & Trim$(Str$(readT)) & ";" & Trim$(Str$(readC)) & TestVector(stress)
        readT = 99#: readC = 98#
        ready = concrete.TryReadAcceptedASDStrains(1, readT, readC)
        TestCheck ready And readT = expectedT And readC = expectedC, mode & " returned scalars do not alias history", passed, failedCount, report
    Loop
    response.ClearHistory
    ready = concrete.TryReadAcceptedASDStrains(1, readT, readC)
    TestCheck ready And readT = 0# And readC = 0#, mode & " clear resets equivalent history", passed, failedCount, report
    Close #inputHandle: inputHandle = 0: Close #outputHandle: outputHandle = 0
    Exit Sub
Failed:
    number = Err.Number: reason = Err.Description
    On Error Resume Next
    If inputHandle <> 0 Then Close #inputHandle
    If outputHandle <> 0 Then Close #outputHandle
    On Error GoTo 0
    Err.Raise number, "modTestCoupledASDCoordinates.TestCoordinateHistory", mode & "/" & path & ": " & reason
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Пишет шесть напряжений с инвариантной десятичной точкой для независимого сопоставления.
Private Function TestVector(ByRef values() As Double) As String
    Dim i As Long
    For i = 1 To 6: TestVector = TestVector & ";" & Trim$(Str$(values(i))): Next i
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает отдельные утверждения и сохраняет конкретную причину отказа.
Private Sub TestCheck(ByVal condition As Boolean, ByVal name As String, ByRef passed As Long, _
        ByRef failedCount As Long, ByRef report As String)
    If condition Then
        passed = passed + 1
    Else
        failedCount = failedCount + 1: report = report & "FAIL: " & name & vbCrLf
    End If
End Sub
