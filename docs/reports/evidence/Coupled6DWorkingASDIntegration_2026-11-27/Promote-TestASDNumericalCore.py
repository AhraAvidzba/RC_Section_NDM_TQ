"""ДЛЯ ТЕСТОВ. Переносит четыре проверенных численных класса ASD, удаляя тестовый ввод-вывод и используя общий необязательный интерфейс наблюдения."""
from pathlib import Path
from hashlib import sha256
from difflib import unified_diff
import json, re, subprocess

repo = Path.cwd()
here = Path(__file__).resolve().parent
source = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1065ASDPureNDirectEarlyPrefix_20261127')
cold = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1066ASDPureNDirectColdRepeat_20261127')
out = Path('D:/CodexScratch/RC_Section_NDM_TQ/CleanASDWorkingPromotion_20261127')
assert not out.exists()
assert json.loads((cold / 'ColdDirectRepeatProof.json').read_bytes())['all6_cold_histories_bitexact']
assert json.loads((source / 'ASDDirectPureNExecutionProof.json').read_bytes())['native']['own_process_closed']
classes = ['CCoupledDeformationContinuation', 'CCoupledLoadContinuation', 'CCoupledNewtonSolver', 'CCoupledSectionAnalysis']
paths = [f'src/Solver/{name}.cls' for name in classes]
assert subprocess.run(['git', 'diff', '--quiet', '--', *paths]).returncode == 0
before = {name: (repo / path).read_bytes() for name, path in zip(classes, paths)}
text = {name: (source / path).read_text('utf-8-sig') for name, path in zip(classes, paths)}


def TestReplace(code, old, new):
    """Требует ровно одно совпадение каждого содержательного изменения."""
    assert code.count(old) == 1, old
    return code.replace(old, new, 1)


def TestStripCaptureCalls(code):
    """Удаляет только цельные тестовые вызовы с их продолжениями строк; численные операторы не изменяет."""
    lines = code.splitlines(keepends=True)
    result = []
    index = 0
    while index < len(lines):
        if 'modTestSelectableMaterialCore.TestCapture' in lines[index]:
            while lines[index].rstrip().endswith('_'):
                index += 1
            index += 1
            continue
        result.append(lines[index])
        index += 1
    return ''.join(result)


# Из диагностических блоков сначала исключаются оболочки, которые содержали только чтение.
code = text['CCoupledDeformationContinuation']
start = code.index('    For i = 0 To index\n        modTestSelectableMaterialCore.TestCapturePeakRefinement "sample"')
end = code.index('    Next i\n', start) + len('    Next i\n')
code = code[:start] + code[end:]
text['CCoupledDeformationContinuation'] = code
code = text['CCoupledLoadContinuation']
start = code.index('    Dim observedPeak As Double, observedAccuracy As Double, observedDesign As Boolean\n')
end = code.index('    End If\n', start) + len('    End If\n')
code = code[:start] + code[end:]
text['CCoupledLoadContinuation'] = code
for name in classes:
    text[name] = TestStripCaptureCalls(text[name])
    assert 'modTest' not in text[name] and 'TestCapture' not in text[name]
code = text['CCoupledNewtonSolver']
start = code.index("' Проверяет ограничение резерва для отдельной центральной или боковой истории, с арматурой и без неё.\n")
end = code.index('\nEnd Function', start) + len('\nEnd Function')
assert code.count('HybridResponseHasPrincipalBoundaries') == code[start:end].count('HybridResponseHasPrincipalBoundaries')
text['CCoupledNewtonSolver'] = code[:start] + code[end:]

# Более длинный бюджет и прогноз решённых полей относятся только к спектральному ASD.
code = text['CCoupledDeformationContinuation']
declaration = "Private Const PROBE_LIMIT As Long = 192 ' Общий предел проб продолжения, деления и повторного прохождения плеч."
assert code.count(declaration) == 1
code = re.sub(r'\bPROBE_LIMIT\b', 'mProbeBudget', code)
code = TestReplace(code, declaration.replace('PROBE_LIMIT', 'mProbeBudget'),
    "Private Const PROBE_LIMIT As Long = 96 ' Общий бюджет продолжения для закона без спектральных границ.\n"
    "Private Const ASD_PROBE_LIMIT As Long = 192 ' Бюджет ASD для прохождения посттрещинной ветви и локализации стали.\n"
    "Private mProbeBudget As Long ' Неизменный на одну прослеживаемую ветвь бюджет выбранного закона.")
anchor = '    Set mSolver = solver: Set mReinforcement = reinforcement\n'
code = TestReplace(code, anchor, anchor + '''    mProbeBudget = PROBE_LIMIT
    If Not mSolver Is Nothing Then
        If mSolver.HasPrincipalBoundaries Then mProbeBudget = ASD_PROBE_LIMIT
    End If
''')
code = TestReplace(code, 'conditionHybridRecovery:=(robustDeparture Or (keepYieldControl And allowHybridRecovery)),',
    'conditionHybridRecovery:=(robustDeparture Or (mSolver.HasPrincipalBoundaries And keepYieldControl And allowHybridRecovery)),')
anchor = '    Dim lowField() As Double, highField() As Double, guess() As Double, probeResultants() As Double\n'
code = TestReplace(code, anchor, anchor + '    Dim useFieldPrediction As Boolean: useFieldPrediction = mSolver.HasPrincipalBoundaries\n')
code = TestReplace(code, '    mSolver.CopyAccepted lowField: mSolver.CopyTrial highField, probeResultants\n',
    '    If useFieldPrediction Then\n        mSolver.CopyAccepted lowField: mSolver.CopyTrial highField, probeResultants\n    End If\n')
start = code.index('        ReDim guess(1 To UBound(lowField))\n')
end = code.index('        If Not UpdateTrialCoordinate(point, low.Coordinate)', start)
block = code[start:end]
code = code[:start] + '        If useFieldPrediction Then\n' + ''.join('    ' + line for line in block.splitlines(keepends=True)) + '''        Else
            If Not Probe(point, low.Lambda, keepYieldControl:=True, acceptedLambda:=low.Lambda) Then Exit Function
        End If
''' + code[end:]
assert code.count('            mSolver.CopyTrial highField, probeResultants\n') == 2
code = code.replace('            mSolver.CopyTrial highField, probeResultants\n',
    '            If useFieldPrediction Then mSolver.CopyTrial highField, probeResultants\n', 2)
code = TestReplace(code, '            lowField = field\n', '            If useFieldPrediction Then lowField = field\n')
start = code.index('            seed = highField: seedLambda = high.Lambda\n')
end = code.index('            If Not UpdateTrialCoordinate(high, low.Coordinate)', start)
block = code[start:end]
code = code[:start] + '            If useFieldPrediction Then\n' + ''.join('    ' + line for line in block.splitlines(keepends=True)) + '''            Else
                If Not Probe(high, low.Lambda, keepYieldControl:=True, acceptedLambda:=low.Lambda) Then Exit Function
            End If
''' + code[end:]
code = TestReplace(code, '    If mVerifiedPeakFields Is Nothing Or mVerifiedPeakTrace Is Nothing Then Exit Function\n',
    '    If mVerifiedPeakFields Is Nothing Or mVerifiedPeakTrace Is Nothing Then Exit Function\n    NotifyProgress\n')
anchor = "' Возвращает независимую числовую копию принятого пути; незавершённый путь не становится предельным НДС.\n"
code = TestReplace(code, anchor, '''' Передаёт завершённый допустимый путь через общий интерфейс, сохраняя независимость снимка от решателя.
Friend Sub ReportAcceptedProgress()
    NotifyProgress
End Sub

''' + anchor)
text['CCoupledDeformationContinuation'] = code

# Общий интерфейс наблюдения получает копии, а ядро не выбирает файл, Excel или другой носитель.
code = text['CCoupledLoadContinuation']
code = TestReplace(code, 'Friend Function ContinueByDeformation(ByVal maximumLambda As Double) As CCoupledDeformationContinuation',
    'Friend Function ContinueByDeformation(ByVal maximumLambda As Double, _\n        Optional ByVal progress As ICoupledContinuationProgress = Nothing) As CCoupledDeformationContinuation')
code = TestReplace(code, 'mAbsoluteTolerance, mRelativeTolerance, mAcceptedTrace\n',
    'mAbsoluteTolerance, mRelativeTolerance, mAcceptedTrace, progress\n    controlled.ReportAcceptedProgress\n')
code = TestReplace(code, 'Friend Function TryRecoverDirectByDeformation() As Boolean',
    'Friend Function TryRecoverDirectByDeformation(Optional ByVal progress As ICoupledContinuationProgress = Nothing) As Boolean')
code = TestReplace(code, 'seedResultants, 2#, 0.001, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance, seedTrace\n',
    'seedResultants, 2#, 0.001, mMaxIterations, mAbsoluteTolerance, mRelativeTolerance, seedTrace, progress\n    controlled.ReportAcceptedProgress\n')
text['CCoupledLoadContinuation'] = code
code = text['CCoupledSectionAnalysis']
anchor = '        Optional ByRef acceptedTrace As CCoupledLoadTrace = Nothing) As CResultMeta'
code = TestReplace(code, anchor, '        Optional ByRef acceptedTrace As CCoupledLoadTrace = Nothing, _\n        Optional ByVal progress As ICoupledContinuationProgress = Nothing) As CResultMeta')
code = TestReplace(code, '        STRESS_TOLERANCE, RELATIVE_TOLERANCE, (captureFields Or mCalibratedASD)\n',
    '        STRESS_TOLERANCE, RELATIVE_TOLERANCE, (captureFields Or mCalibratedASD Or Not progress Is Nothing)\n')
code = TestReplace(code, 'continuation.TryRecoverDirectByDeformation()', 'continuation.TryRecoverDirectByDeformation(progress)')
code = TestReplace(code, '    Set acceptedTrace = continuation.SavedTrace\n',
    '    Set acceptedTrace = continuation.SavedTrace\n    ReportAcceptedProgress progress, continuation.AcceptedFields, acceptedTrace, meta\n')
code = TestReplace(code, '        Optional ByRef acceptedTrace As CCoupledLoadTrace = Nothing) As CCapacityResult',
    '        Optional ByRef acceptedTrace As CCoupledLoadTrace = Nothing, _\n        Optional ByVal progress As ICoupledContinuationProgress = Nothing) As CCapacityResult')
code = TestReplace(code, 'Set controlled = branch.ContinueByDeformation(maximumLambda)',
    'Set controlled = branch.ContinueByDeformation(maximumLambda, progress)')
code = TestReplace(code, '                SaveCapacityFields capacity, solver, spec, limitHistory, response, concrete, models, limitTarget, actual, fc, ft\n',
    '                SaveCapacityFields capacity, solver, spec, limitHistory, response, concrete, models, limitTarget, actual, fc, ft\n                ReportAcceptedProgress progress, limitHistory, acceptedTrace, capacity.ResultMeta\n')
anchor = '            SaveCapacityFields capacity, solver, spec, limitHistory, response, concrete, models, limitTarget, actual, fc, ft, criticalPoint\n'
code = TestReplace(code, anchor, anchor + '            ReportAcceptedProgress progress, limitHistory, acceptedTrace, capacity.ResultMeta\n')
code = code.rstrip() + '\n' + (here / 'ASDAcceptedProgressMethod.bas.txt').read_text('utf-8-sig')
text['CCoupledSectionAnalysis'] = code
for name, path in zip(classes, paths):
    assert 'modTest' not in text[name] and 'TestCapture' not in text[name]
    assert max(map(len, text[name].splitlines())) < 1024
    assert (repo / path).read_bytes() == before[name]
out.mkdir(parents=True)
changed = {}
for name, path in zip(classes, paths):
    code = text[name]
    assert 'modTest' not in code and 'TestCapture' not in code
    assert max(map(len, code.splitlines())) < 1024
    assert (repo / path).read_bytes() == before[name]
    data = code.replace('\n', '\r\n').encode('utf-8-sig')
    backup = out / 'Before' / path
    candidate = out / 'Candidate' / path
    backup.parent.mkdir(parents=True, exist_ok=True)
    candidate.parent.mkdir(parents=True, exist_ok=True)
    backup.write_bytes(before[name])
    candidate.write_bytes(data)
    (repo / path).write_bytes(data)
    changed[name] = dict(before_sha256=sha256(before[name]).hexdigest().upper(),
        source_test_sha256=sha256((source / path).read_bytes()).hexdigest().upper(),
        after_sha256=sha256(data).hexdigest().upper())
    diff = '\n'.join(unified_diff(before[name].decode('utf-8-sig').splitlines(), code.splitlines(),
        fromfile='Before/' + path, tofile='Candidate/' + path))
    (out / (name + '.diff.txt')).write_text(diff + '\n', encoding='utf-8')
proof = dict(changed=changed, exactly_four_numerical_classes=True, no_direct_test_module_calls=True,
    observer_interface_reused='ICoupledContinuationProgress', no_kernel_file_IO_added=True,
    ASD_probe_budget=192, non_ASD_probe_budget=96,
    non_ASD_steel_refinement_uses_existing_unseeded_path=True,
    native_compilation_and_runtime_verification_pending=True, output_book_not_updated=True,
    whole_goal_completed=False)
(out / 'PromotionPreparation.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD SOURCE UPDATED', len(changed), 'CLASSES;NO TEST CALLBACKS;OPTIONAL GENERIC OBSERVER;NATIVE VERIFICATION PENDING')
