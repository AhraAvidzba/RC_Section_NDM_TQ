"""ДЛЯ ТЕСТОВ. Добавляет читающий журнал составляющих ошибки максимума без изменения численных проб."""


def TestPatchPeakObservation(code, canonical=False):
    """Вставляет три наблюдательных вызова только в уточнение и завершение максимума."""
    def TestCase(value):
        """Сопоставляет исходник с приведённым регистром исполняемых имён фактического VBA."""
        return value.lower() if canonical else value

    def TestReplace(before, after):
        """Требует единственного фрагмента, чтобы не затронуть другие численные методы."""
        nonlocal code
        before, after = TestCase(before), TestCase(after)
        assert code.count(before) == 1, before
        code = code.replace(before, after, 1)

    needle = '                If refinedFields Is Nothing Then Exit Function\n'
    TestReplace(needle, '''                modTestSelectableMaterialCore.TestCapturePeakRefinement "replay_failed", _
                    Array(refinement, mProbeCount, accuracy, first.Lambda, second.Lambda, first.Uncertainty, second.Uncertainty), mProbeDetails
''' + needle)
    needle = '            retainedLeft = checkedLeft: retainedRight = checkedRight: retainedPeak = second\n'
    TestReplace(needle, '''            modTestSelectableMaterialCore.TestCapturePeakRefinement "estimate", _
                Array(refinement, mProbeCount, accuracy, tolerance, sampling, first.Lambda, second.Lambda, first.Uncertainty, second.Uncertainty, _
                    checkedLeft.Coordinate, checkedMiddle.Coordinate, checkedRight.Coordinate, mBestLambda, retainedMaximum), mProbeDetails
''' + needle)
    needle = 'Private Function CompleteVerifiedPeak(ByVal history As Collection, ByVal lastBeforeEvent As Long, ByVal eventUpper As Double) As Boolean\n'
    TestReplace(needle, needle + '''    modTestSelectableMaterialCore.TestCapturePeakRefinement "completion", _
        Array(mProbeCount, mVerifiedPeak.Lambda, mVerifiedPeakAccuracy, mLambdaTolerance, mBestLambda, lastBeforeEvent, eventUpper, Abs(CLng(mUnverifiedDecline))), mProbeDetails
''')
    return code


def TestPatchPeakModule(code):
    """Добавляет UTF-8 запись скалярного журнала в явно тестовый модуль, не читая и не меняя состояния материалов."""
    assert 'Public Sub TestCapturePeakRefinement' not in code
    return code.rstrip() + '''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет составляющие ошибки пика; значения переданы копиями и не управляют решателем.
Public Sub TestCapturePeakRefinement(ByVal stage As String, ByVal values As Variant, ByVal diagnostic As String)
    If Len(mMainCapacityFolderTest) = 0 Then Err.Raise 5, , "Не задана папка наблюдения максимума."
    Dim line As String, value As Variant: line = stage
    For Each value In values: line = line & ";" & Trim$(Str$(CDbl(value))): Next value
    line = line & ";" & Replace(Replace(diagnostic, vbCrLf, " | "), vbLf, " | ") & vbCrLf
    Dim file As String: file = mMainCapacityFolderTest & "\\Native\\PeakRefinement.Values.csv"
    Dim stream As Object: Set stream = CreateObject("ADODB.Stream")
    stream.Type = 2: stream.Charset = "utf-8": stream.Open
    If Len(Dir$(file)) > 0 Then stream.LoadFromFile file: stream.Position = stream.Size
    stream.WriteText line: stream.SaveToFile file, 2: stream.Close
End Sub
'''


def TestPatchPeakSamples(code, canonical=False):
    """Записывает только уже вычисленные точки плеч перед проверкой различимости максимума."""
    before = '    Dim peakIndex As Long, lower As Double\n'
    after = '''    For i = 0 To index
        modTestSelectableMaterialCore.TestCapturePeakRefinement "sample", _
            Array(mProbeCount, i, divisions, samples(i).Coordinate, samples(i).Lambda, samples(i).Uncertainty, _
                samples(0).Lambda, samples(index).Lambda, Abs(CLng(keepRight))), mProbeDetails
    Next i
''' + before
    if canonical:
        before, after = before.lower(), after.lower()
    assert code.count(before) == 1
    return code.replace(before, after, 1)
