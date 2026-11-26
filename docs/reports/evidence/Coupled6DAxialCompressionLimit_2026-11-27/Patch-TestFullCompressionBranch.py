"""ДЛЯ ТЕСТОВ. Сохраняет всю принятую ветвь перед окончательной проверкой максимума без изменения состояния решателя."""


def TestPatchFullBranchObservation(code, canonical=False):
    """Добавляет один читающий вызов перед восстановлением кандидатного максимума."""
    before = '    If mVerifiedPeakFields Is Nothing Or mVerifiedPeakTrace Is Nothing Then Exit Function\n'
    after = before + '    modTestSelectableMaterialCore.TestCaptureFullBranch history, mTrace\n'
    if canonical:
        before, after = before.lower(), after.lower()
    assert code.count(before) == 1
    return code.replace(before, after, 1)


def TestPatchFullBranchModule(code):
    """Добавляет бинарный экспорт копий принятых полей и точных множителей в явно тестовый модуль."""
    assert 'Public Sub TestCaptureFullBranch' not in code
    return code.rstrip() + '''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет весь допустимый путь до ограничения стали; запись не коммитит и не изменяет материальную историю.
Public Sub TestCaptureFullBranch(ByVal fields As Collection, ByVal trace As CCoupledLoadTrace)
    If fields Is Nothing Or trace Is Nothing Then Err.Raise 5, , "Для проверки максимума отсутствует полный путь."
    If fields.Count <> trace.Count Then Err.Raise 5, , "Количество полей полного пути не совпадает с журналом нагрузок."
    Dim handle As Integer: handle = FreeFile
    Open mMainCapacityFolderTest & "\\Native\\FullFollowedHistory.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMFULLTRACE0001": Put #handle, , magic
    TestWriteVector handle, Array(fields.Count)
    Dim index As Long, fraction As Double, factor As Double, field() As Double
    For index = 1 To fields.Count
        trace.ReadPoint index, fraction, factor: field = fields(index)
        TestWriteVector handle, Array(fraction, factor): TestWriteVector handle, field
    Next index
    Close #handle
End Sub
'''
