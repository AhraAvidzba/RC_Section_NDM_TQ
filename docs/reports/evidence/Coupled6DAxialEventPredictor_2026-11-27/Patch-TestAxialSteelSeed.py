"""ДЛЯ ТЕСТОВ. Добавляет прогноз уточнения стали по двум решённым плечам; не подменяет им равновесие."""


def TestPatchSteelSeed(code, canonical=False):
    """Меняет только начальные приближения внутри локализации; обе истории проверяются обычной пробой заново."""
    def TestCase(value):
        """Учитывает приведённый регистр кода, извлечённого из фактического VBA."""
        return value.lower() if canonical else value

    start = code.index(TestCase('Private Function RefineSteelCrossing('))
    end = code.index(TestCase('\nEnd Function'), start) + len('\nEnd Function')
    method = code[start:end]

    def TestReplace(before, after):
        """Отклоняет отсутствующий либо неоднозначный фрагмент вместо произвольного изменения метода."""
        nonlocal method
        before, after = TestCase(before), TestCase(after)
        assert method.count(before) == 1, before
        method = method.replace(before, after, 1)

    TestReplace('    Dim low As TBranchPoint, high As TBranchPoint, point As TBranchPoint\n',
                '''    Dim low As TBranchPoint, high As TBranchPoint, point As TBranchPoint
    Dim lowField() As Double, highField() As Double, guess() As Double, probeResultants() As Double
    Dim seed As Variant, seedLambda As Double, degree As Long
    ' Два решённых поля задают только прогноз; равновесие и материал проверяются от текущей принятой истории.
''')
    TestReplace('    mSolver.DiscardStep\n    low.Uncertainty',
                '    mSolver.CopyAccepted lowField: mSolver.CopyTrial highField, probeResultants\n    mSolver.DiscardStep\n    low.Uncertainty')
    TestReplace('''        If Not Probe(point, low.Lambda, keepYieldControl:=True, allowHybridRecovery:=mSolver.SupportsHybridRecovery, _
                acceptedLambda:=low.Lambda) Then Exit Function''',
                '''        ReDim guess(1 To UBound(lowField))
        For degree = 1 To UBound(lowField): guess(degree) = lowField(degree) + 0.5 * (highField(degree) - lowField(degree)): Next degree
        seed = guess: seedLambda = low.Lambda + 0.5 * (high.Lambda - low.Lambda)
        If Not Probe(point, seedLambda, keepYieldControl:=True, initialField:=seed, allowHybridRecovery:=mSolver.SupportsHybridRecovery, _
                acceptedLambda:=low.Lambda) Then Exit Function''')
    TestReplace('            high = point: eventState.ExceededRatio = ratio\n',
                '            high = point: eventState.ExceededRatio = ratio\n            mSolver.CopyTrial highField, probeResultants\n')
    TestReplace('            mSolver.CopyTrial field, actual\n            mSolver.AcceptStep\n',
                '            mSolver.CopyTrial field, actual\n            lowField = field\n            mSolver.AcceptStep\n')
    TestReplace('''            If Not Probe(high, low.Lambda, keepYieldControl:=True, allowHybridRecovery:=mSolver.SupportsHybridRecovery, _
                    acceptedLambda:=low.Lambda) Then Exit Function''',
                '''            seed = highField: seedLambda = high.Lambda
            If Not Probe(high, seedLambda, keepYieldControl:=True, initialField:=seed, allowHybridRecovery:=mSolver.SupportsHybridRecovery, _
                    acceptedLambda:=low.Lambda) Then Exit Function''')
    TestReplace('            If eventState.ExceededRatio <= 1# Then Exit Function\n',
                '            If eventState.ExceededRatio <= 1# Then Exit Function\n            mSolver.CopyTrial highField, probeResultants\n')
    return code[:start] + method + code[end:]
