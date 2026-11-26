"""ДЛЯ ТЕСТОВ. Проверяет точность локального закона при нелинейной координате события CDP."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source):
    """Оба исходных поля и точная принятая история сохранены; ужесточаются только новые локальные пробы."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestNewtonYieldEvent.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Materials/CCoupledConcreteReturn.cls'
    code = (repo / name).read_text('utf-8-sig')
    old = '    tolerance = RELATIVE_YIELD_TOLERANCE * cc'
    assert code.count(old) == 1
    code = code.replace(old, old + '''
    If modTestCoupledElasticState.ReturnPrecisionFactorTest > 0# Then _
        tolerance = tolerance * modTestCoupledElasticState.ReturnPrecisionFactorTest ' ДЛЯ ТЕСТОВ: точность нового возврата.
''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    if Path(name).stem not in prep['modules']:
        prep['modules'].append(Path(name).stem)
    name = 'src/Materials/CCoupledConcreteResponse.cls'
    code = (root / name).read_text('utf-8-sig')
    block, = re.findall(r'(?ms)^Private Function UpdatePoint\(.*?^End Function', code)
    body = block.replace('YIELD_TOLERANCE_RATIO', 'yieldToleranceTest').replace('STATE_TOLERANCE', 'stateToleranceTest')
    marker = '    Dim elastic(1 To N_COMPONENTS) As Double, effective(1 To N_COMPONENTS) As Double'
    assert body.count(marker) == 1
    body = body.replace(marker, '''    Dim yieldToleranceTest As Double, stateToleranceTest As Double
    yieldToleranceTest = TestPrecisionYieldTolerance: stateToleranceTest = STATE_TOLERANCE
    If modTestCoupledElasticState.ReturnPrecisionFactorTest > 0# Then _
        stateToleranceTest = stateToleranceTest * modTestCoupledElasticState.ReturnPrecisionFactorTest
''' + marker, 1)
    code = code.replace(block, body, 1)
    old = '> YIELD_TOLERANCE_RATIO * cc / mPressureModulus Then Exit Function'
    assert code.count(old) == 1
    code = code.replace(old, '> TestPrecisionYieldTolerance * cc / mPressureModulus Then Exit Function', 1)
    code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет исходную точность принятого префикса и ужесточает только новые материальные пробы.
Private Function TestPrecisionYieldTolerance() As Double
    TestPrecisionYieldTolerance = YIELD_TOLERANCE_RATIO
    If modTestCoupledElasticState.ReturnPrecisionFactorTest > 0# Then _
        TestPrecisionYieldTolerance = TestPrecisionYieldTolerance * modTestCoupledElasticState.ReturnPrecisionFactorTest
End Function
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    assert code.count('Option Explicit\n') == 1
    code = code.replace('Option Explicit\n', 'Option Explicit\nPublic ReturnPrecisionFactorTest As Double\n', 1)
    block, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', code)
    old = '        If Not solver.ReplayAcceptedFields(fields, fields.Count) Then'
    assert block.count(old) == 1
    body = block.replace(old, '        ReturnPrecisionFactorTest = 0#\n' + old, 1)
    old = '        revision = solver.AcceptedRevision'
    assert body.count(old) == 1
    body = body.replace(old, '''        ReturnPrecisionFactorTest = 0.0001
        Print #handle, "LOCAL_PRECISION;" & CStr(probeCase) & ";" & TestOnsetNumber(ReturnPrecisionFactorTest)
''' + old, 1)
    body = body.replace('    Close #handle: mPolishExecutedTest = True',
                        '    ReturnPrecisionFactorTest = 0#\n    Close #handle: mPolishExecutedTest = True', 1)
    code = code.replace(block, body, 1)
    block, = re.findall(r'(?ms)^Private Sub TestExportReverseCandidate\(.*?^End Sub', code)
    old = '    For index = 1 To mReverseFieldsTest.Count'
    assert block.count(old) == 1
    body = block.replace(old, '''    Dim savedPrecisionTest As Double: savedPrecisionTest = ReturnPrecisionFactorTest
''' + old + '''
        ReturnPrecisionFactorTest = 0#
        If index = mReverseFieldsTest.Count Then ReturnPrecisionFactorTest = savedPrecisionTest
''', 1)
    assert body.count('    Next index') == 1
    body = body.replace('    Next index', '    Next index\n    ReturnPrecisionFactorTest = savedPrecisionTest', 1)
    code = code.replace(block, body, 1)
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    (root / Path(__file__).name).write_bytes(Path(__file__).read_bytes())
    prep['inputs'][Path(__file__).name] = sha256(Path(__file__).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'CompleteNewtonNonlinearYieldEventLocalPrecision',
        'scalar_yield_relative_tolerance': 1e-14, 'coupled_yield_relative_tolerance': 1e-12,
        'coupled_state_relative_tolerance': 5e-14, 'accepted_prefix_always_original_precision': True,
        'new_local_trials_precision_factor': .0001, 'global_tolerance_and_Newton_budget_unchanged': True,
        'working_runtime_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('NONLINEAR YIELD EVENT LOCAL PRECISION STUDY READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
