"""ДЛЯ ТЕСТОВ. Уточняет равновесие пиковых проб по оценке неопределённости lambda, не ослабляя исходные критерии."""


def TestPatchPeakPrecision(code, canonical=False):
    """Добавляет масштаб более строгого допуска только для повторных плеч спектрального максимума."""
    def TestCase(value):
        """Сохраняет соответствие регистру извлечённого VBA; новые строковые ключи заданы строчными."""
        return value.lower() if canonical else value

    def TestReplace(before, after):
        """Требует ровно одно совпадение для каждой содержательной правки."""
        nonlocal code
        before, after = TestCase(before), TestCase(after)
        assert code.count(before) == 1, before
        code = code.replace(before, after, 1)

    TestReplace('''        Set sourceFields = history: Set sourceTrace = mTrace: checkedMiddle = middle: first = middle
        For refinement = 1 To STEP_LIMIT''', '''        Set sourceFields = history: Set sourceTrace = mTrace: checkedMiddle = middle: first = middle
        Dim probeToleranceScale As Double: probeToleranceScale = 1#
        Dim mayTighten As Boolean: mayTighten = mSolver.HasPrincipalBoundaries And Not mUsesEnvelopeControl
        For refinement = 1 To STEP_LIMIT''')
    TestReplace('''                    refinedFields, actual, True, retainedMaximum, refinedDesignLeft, retainedPreviousMaximum, sourceTrace) Then''',
                '''                    refinedFields, actual, True, retainedMaximum, refinedDesignLeft, retainedPreviousMaximum, sourceTrace, probeToleranceScale) Then''')
    TestReplace('''            If first.Uncertainty + second.Uncertainty >= tolerance Or mProbeCount + 4 > PROBE_LIMIT Then Exit For
            first = second''', '''            If mProbeCount + 4 > PROBE_LIMIT Then Exit For
            If first.Uncertainty + second.Uncertainty >= tolerance And Not mayTighten Then Exit For
            If mayTighten And second.Uncertainty > 0.01 * tolerance Then
                probeToleranceScale = probeToleranceScale * (0.01 * tolerance / second.Uncertainty)
                modTestSelectableMaterialCore.TestCapturePeakRefinement "precision", _
                    Array(refinement, probeToleranceScale, second.Uncertainty, 0.01 * tolerance), mProbeDetails
            End If
            first = second''')
    TestReplace('''        Optional ByRef retainedPreviousMaximum As Double = 0#, Optional ByVal sourceTrace As CCoupledLoadTrace = Nothing) As Boolean
    If Not mSolver.ReplayAcceptedFields(history, left.HistoryIndex) Then Exit Function''',
                '''        Optional ByRef retainedPreviousMaximum As Double = 0#, Optional ByVal sourceTrace As CCoupledLoadTrace = Nothing, _
        Optional ByVal probeToleranceScale As Double = 1#) As Boolean
    If Not mSolver.ReplayAcceptedFields(history, left.HistoryIndex) Then Exit Function''')
    TestReplace('''            If Not Probe(point, previousLambda, keepYieldControl:=keepRight, acceptedLambda:=previousLambda) Then Exit Function''',
                '''            If Not Probe(point, previousLambda, keepYieldControl:=keepRight, _
                    allowHybridRecovery:=(probeToleranceScale < 1# And mSolver.SupportsHybridRecovery), _
                    acceptedLambda:=previousLambda, equilibriumScale:=probeToleranceScale) Then Exit Function''')
    TestReplace('''            If Not mSolver.TrialLambdaUncertainty(mAbsTol, mRelTol, point.Uncertainty) Then''',
                '''            If Not mSolver.TrialLambdaUncertainty(mAbsTol * probeToleranceScale, mRelTol * probeToleranceScale, point.Uncertainty) Then''')
    start = code.index(TestCase('Private Function Probe('))
    end = code.index(TestCase('\nEnd Function'), start) + len('\nEnd Function')
    method = code[start:end]
    before = TestCase('        Optional ByVal robustDeparture As Boolean = False) As Boolean\n')
    after = TestCase('''        Optional ByVal robustDeparture As Boolean = False, Optional ByVal equilibriumScale As Double = 1#) As Boolean
    ' Более строгий масштаб используется только при уточнении максимума; физическая модель и конечный критерий сохраняются.
    If equilibriumScale <= 0# Or equilibriumScale > 1# Then Exit Function
    Dim probeAbs As Double, probeRel As Double
    probeAbs = mAbsTol * equilibriumScale: probeRel = mRelTol * equilibriumScale
''')
    assert method.count(before) == 1
    method = method.replace(before, after, 1)
    before = TestCase('mAbsTol, mRelTol')
    assert method.count(before) == 2
    method = method.replace(before, TestCase('probeAbs, probeRel'))
    return code[:start] + method + code[end:]
