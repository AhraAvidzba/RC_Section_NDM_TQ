"""ДЛЯ ТЕСТОВ. Добавляет чтение полного пути и диагностики точной пробы без изменения расчётных выражений и состояния материала."""


def TestPatchDirectRecoveryObservation(code, canonical=False):
    """Сохраняет контролируемую ветвь при отсутствии подтверждённого пика и причину отказа точного сочетания."""
    changes = [
        ('    Set controlMeta = controlled.ResultMeta\n', '''    Set controlMeta = controlled.ResultMeta
    Dim observedPeak As Double, observedAccuracy As Double, observedDesign As Boolean
    If Not branchFields Is Nothing And Not branchTrace Is Nothing Then
        If Not controlled.CopyPeak(observedPeak, observedAccuracy, observedDesign) Then _
            modTestSelectableMaterialCore.TestCaptureFullBranch branchFields, branchTrace
    End If
'''),
        ('                    If TryEquilibrium(1#, False, initial) Then\n', '''                    modTestSelectableMaterialCore.TestCapturePeakRefinement "direct_exact_start", _
                        Array(index, leftLambda, rightLambda, ratio), ""
                    If TryEquilibrium(1#, False, initial) Then
'''),
        ('                    mSolver.DiscardStep\n                End If\n            End If\n        Next index\n', '''                    modTestSelectableMaterialCore.TestCapturePeakRefinement "direct_exact_failed", _
                        Array(index, mMeta.InternalStatus, mMeta.ResultCode), mMeta.ResultComment & " | " & mMeta.DiagnosticDetails
                    mSolver.DiscardStep
                End If
            End If
        Next index
''')
    ]
    for before, after in changes:
        if canonical:
            before = before.lower()
            after = after.lower()
        assert code.count(before) == 1, before
        code = code.replace(before, after, 1)
    return code
