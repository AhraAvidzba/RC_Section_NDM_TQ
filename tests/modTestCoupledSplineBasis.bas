Attribute VB_Name = "modTestCoupledSplineBasis"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: кубические функции и их производные на физических координатах
' ==========================================================================
' Сравнивает с формулами Бернштейна на одном интервале и с аналитическими
' условиями воспроизведения постоянного и линейного поля на нескольких.

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет значения на обоих концах, внутреннем узле и между узлами.
Public Function RunCoupledSplineBasisTests() As String
    On Error GoTo Failed
    Dim spline As New CCubicSplineBasis
    Dim values(1 To 4) As Double
    Dim derivatives(1 To 4) As Double
    Dim first As Long
    Dim passed As Long
    Dim failures As String

    spline.Initialize 10#, 30#, 1
    spline.Evaluate 15#, first, values, derivatives
    TestCheck first = 1 And spline.Count = 4, "one span indices", passed, failures
    TestCheck Abs(values(1) - 0.421875) < 1E-14 And Abs(values(2) - 0.421875) < 1E-14 And _
        Abs(values(3) - 0.140625) < 1E-14 And Abs(values(4) - 0.015625) < 1E-14, _
        "Bernstein values", passed, failures
    TestCheck Abs(derivatives(1) + 0.084375) < 1E-14 And Abs(derivatives(2) - 0.028125) < 1E-14 And _
        Abs(derivatives(3) - 0.046875) < 1E-14 And Abs(derivatives(4) - 0.009375) < 1E-14, _
        "Bernstein physical derivatives", passed, failures

    spline.Evaluate 10#, first, values, derivatives
    TestCheck values(1) = 1# And values(4) = 0# And Abs(derivatives(1) + 0.15) < 1E-14, _
        "closed left endpoint", passed, failures
    spline.Evaluate 30#, first, values, derivatives
    TestCheck values(4) = 1# And values(1) = 0# And Abs(derivatives(4) - 0.15) < 1E-14, _
        "closed right endpoint", passed, failures

    spline.Initialize -12#, 18#, 3
    Dim x As Double, k As Long, j As Long, i As Long
    Dim leftKnot As Double, middleKnot As Double, rightKnot As Double, greville As Double
    Dim sumValue As Double, sumDerivative As Double, linearValue As Double, linearDerivative As Double
    For j = 0 To 12
        x = -12# + 2.5 * j
        spline.Evaluate x, first, values, derivatives
        sumValue = 0#: sumDerivative = 0#: linearValue = 0#: linearDerivative = 0#
        For k = 1 To 4
            i = first + k - 2
            leftKnot = TestKnot(i + 1)
            middleKnot = TestKnot(i + 2)
            rightKnot = TestKnot(i + 3)
            greville = (leftKnot + middleKnot + rightKnot) / 3# ' Коэффициент сплайна, точно представляющего x.
            sumValue = sumValue + values(k)
            sumDerivative = sumDerivative + derivatives(k)
            linearValue = linearValue + values(k) * greville
            linearDerivative = linearDerivative + derivatives(k) * greville
        Next k
        TestCheck Abs(sumValue - 1#) < 1E-13 And Abs(sumDerivative) < 1E-13, _
            "constant field " & CStr(j), passed, failures
        TestCheck Abs(linearValue - x) < 1E-12 And Abs(linearDerivative - 1#) < 1E-13, _
            "linear field " & CStr(j), passed, failures
    Next j
    RunCoupledSplineBasisTests = failures & "TOTAL_COUPLED_SPLINE_BASIS: passed=" & CStr(passed) & _
        "; failed=" & CStr(31 - passed)
    Exit Function
Failed:
    RunCoupledSplineBasisTests = "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает узел открытого вектора (-12,-12,-12,-12,-2,8,18,18,18,18).
Private Function TestKnot(ByVal index As Long) As Double
    Select Case index
        Case 0 To 3: TestKnot = -12#
        Case 4: TestKnot = -2#
        Case 5: TestKnot = 8#
        Case 6 To 9: TestKnot = 18#
    End Select
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Записывает название каждого несоответствия и считает прошедшие проверки.
Private Sub TestCheck(ByVal condition As Boolean, ByVal description As String, _
        ByRef passed As Long, ByRef failures As String)
    If condition Then
        passed = passed + 1
    Else
        failures = failures & "FAIL: " & description & vbCrLf
    End If
End Sub
