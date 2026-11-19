Attribute VB_Name = "modTestCoupledControlCondition"
Option Explicit

' ==========================================================================
' ДЛЯ ТЕСТОВ: аналитические оценки выбора координаты управления
' ==========================================================================
Private mPassedTest As Long, mFailedTest As Long
Private mReportTest As String

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет известные обратные матрицы, вырожденную координату и независимость от исходной строки управления и масштабов строк.
Public Function RunTestCoupledControlCondition() As String
    On Error GoTo Failed
    mPassedTest = 0: mFailedTest = 0: mReportTest = vbNullString
    Dim matrix(1 To 3, 1 To 3) As Double, gradient(1 To 3) As Double, score As Double, alignment As Double
    Dim selector As New CCoupledControlConditionTest, i As Long
    For i = 1 To 3: matrix(i, i) = 1#: Next i
    TestAssert selector.TryInitializeTest(matrix, 3), "Единичная матрица подготовлена"
    gradient(3) = 1#
    TestAssert selector.TryRankTest(gradient, score, alignment), "Ортогональная координата доступна"
    TestAssert Abs(score - Sqr(3#)) < 0.000000000001, "Норма обратной единичной матрицы"
    TestAssert Abs(alignment - 1#) < 0.000000000001, "Косинус нулевого направления"
    gradient(1) = 3#: gradient(3) = 4#
    TestAssert selector.TryRankTest(gradient, score, alignment), "Наклонная координата доступна"
    TestAssert Abs(score - Sqr(4.125)) < 0.000000000001, "Аналитическая норма наклонной обратной матрицы"
    TestAssert Abs(alignment - 0.8) < 0.000000000001, "Аналитический косинус наклонной координаты"
    gradient(1) = 3000#: gradient(3) = 4000#
    TestAssert selector.TryRankTest(gradient, score, alignment) And Abs(score - Sqr(4.125)) < 0.000000000001, "Масштаб градиента не меняет оценку"
    matrix(1, 3) = 1#
    TestAssert selector.TryInitializeTest(matrix, 3), "Иная исходная строка подготовлена"
    TestAssert selector.TryRankTest(gradient, score, alignment) And Abs(score - Sqr(4.125)) < 0.000000000001, "Оценка не зависит от исходного управления"
    matrix(1, 3) = 0#: matrix(1, 1) = 5#: matrix(2, 2) = -7#: matrix(3, 3) = 100#
    TestAssert selector.TryInitializeTest(matrix, 3), "Разные масштабы и знак строк подготовлены"
    TestAssert selector.TryRankTest(gradient, score, alignment) And Abs(score - Sqr(4.125)) < 0.000000000001, "Масштабы и знак строк не меняют оценку"
    gradient(3) = 0#
    TestAssert Not selector.TryRankTest(gradient, score, alignment), "Координата внутри строки равновесия не устраняет вырождение"
    gradient(1) = 0#
    TestAssert Not selector.TryRankTest(gradient, score, alignment), "Нулевой градиент исключён"
    matrix(1, 1) = 1#: matrix(2, 1) = 1#: matrix(2, 2) = 1#: matrix(3, 3) = 1#
    gradient(3) = 1#
    TestAssert selector.TryInitializeTest(matrix, 3), "Несимметричная опора подготовлена"
    TestAssert selector.TryRankTest(gradient, score, alignment) And Abs(score - Sqr(5#)) < 0.000000000001, "Точная норма несимметричной обратной матрицы"
    matrix(3, 3) = 0#
    TestAssert Not selector.TryInitializeTest(matrix, 3), "Вырожденная линейная опора отклонена"
    RunTestCoupledControlCondition = mReportTest & "TOTAL_COUPLED_CONTROL_CONDITION_TEST: passed=" & CStr(mPassedTest) & "; failed=" & CStr(mFailedTest)
    Exit Function
Failed:
    RunTestCoupledControlCondition = mReportTest & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Считает независимые аналитические утверждения и выдаёт точную причину отказа.
Private Sub TestAssert(ByVal passed As Boolean, ByVal description As String)
    If passed Then
        mPassedTest = mPassedTest + 1
        mReportTest = mReportTest & "PASS: " & description & vbCrLf
    Else
        mFailedTest = mFailedTest + 1
        mReportTest = mReportTest & "FAIL: " & description & vbCrLf
    End If
End Sub
