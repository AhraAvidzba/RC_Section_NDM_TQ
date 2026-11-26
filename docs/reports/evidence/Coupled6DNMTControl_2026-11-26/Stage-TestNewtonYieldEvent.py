"""ДЛЯ ТЕСТОВ. Решает полное равновесие при действительной нелинейной координате перехода CDP."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source):
    """Нулевая пробная поверхность является событием материала, а не критерием разрушения или пределом сечения."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestNewtonLine.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    original, = re.findall(r'(?ms)^Private Function ICoupledEquationSystem_TryEvaluate\(.*?^End Function', code)
    body = original.replace('    residual(size) = ControlResidual(field)',
                            '    residual(size) = modTestCoupledElasticState.TestReadNewtonYieldEvent(field) - mTargetControl')
    old = '        Next j\n    End If'
    assert body.count(old) == 1 and body != original
    body = body.replace(old, '''        Next j
        Dim eventGradient() As Double
        modTestCoupledElasticState.TestReadNewtonYieldEventGradient field, eventGradient
        For j = 1 To mDegreeCount: jacobian(j, size) = eventGradient(j) * mHybridScales(j): Next j
        jacobian(size, size) = 0#
    End If''', 1)
    code = code.replace(original, body, 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'src/Materials/CCoupledConcreteResponse.cls'
    code = (root / name).read_text('utf-8-sig') + '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует неизменные параметры пробной поверхности и принятую пластическую деформацию для независимого контроля события.
Friend Function ReadNewtonYieldEventStateTest(ByVal point As Long) As Variant
    ReadNewtonYieldEventStateTest = Array(mAlpha, mGamma, mHistory(point, mCommitted).Plastic(1), _
        mHistory(point, mCommitted).Plastic(2), mHistory(point, mCommitted).Plastic(3), _
        mHistory(point, mCommitted).Plastic(4), mHistory(point, mCommitted).Plastic(5), mHistory(point, mCommitted).Plastic(6), _
        mHistory(point, mCommitted).CompressionCohesion, mHistory(point, mCommitted).TensionCohesion)
End Function
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    original, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', code)
    body = original.replace('    Dim target As Double: target = coordinate + 0.5 * spacing',
                            '    Dim target As Double: target = 0#')
    old = '        revision = solver.AcceptedRevision'
    assert body.count(old) == 1 and body != original
    body = body.replace(old, old + '''
        Dim eventConcrete As CCoupledConcreteResponse: Set eventConcrete = sides(3)
        Print #handle, "EVENT_STATE;" & CStr(probeCase) & TestHardPreparedVector(eventConcrete.ReadNewtonYieldEventStateTest(25))''', 1)
    old = '        solver.ReadNewtonLineFieldTest coordinates, field, factor'
    assert body.count(old) == 1
    body = body.replace(old, old + '\n        Print #handle, "EVENT_FINAL;" & CStr(probeCase) & ";" & TestOnsetNumber(TestReadNewtonYieldEvent(field))', 1)
    code = code.replace(original, body, 1) + '\n' + (here / 'TestNativeNewtonYieldEvent.bas.txt').read_text('utf-8-sig')
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeNewtonYieldEvent.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'CompleteNewtonNonlinearYieldEvent', 'nonlinear_event_control': True,
        'event_point': 25, 'event_station_z_over_L': .001, 'event_target_MPa': 0.,
        'event_equation': 'actual trial Barcelona yield from the unchanged committed point state, MPa',
        'event_analytic_row_from_existing_material_gradient': True,
        'yield_event_is_not_a_strength_failure_or_capacity': True, 'working_runtime_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('COMPLETE NEWTON NONLINEAR YIELD EVENT READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
