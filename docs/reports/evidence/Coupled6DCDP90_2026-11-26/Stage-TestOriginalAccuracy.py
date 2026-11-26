"""ДЛЯ ТЕСТОВ. Проверяет общий адаптер с исходным рабочим законом и абсолютным допуском из точного снимка."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path
import struct


def TestStage(repo, root, source):
    """Убирает все численные перехваты материала; новый критерий явно соответствует заданному абсолютному допуску."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestGenericSequence.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Materials/CCoupledConcreteResponse.cls'
    original = (repo / name).read_text('utf-8-sig')
    current = (root / name).read_text('utf-8-sig')
    state_reader, = re.findall(r'(?ms)^Friend Function ReadNewtonYieldEventStateTest\(.*?^End Function', current)
    code = original + "\n' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Читает принятую историю для независимого контроля, без изменения рабочего закона.\n" + state_reader + '\n'
    assert 'ReturnPrecisionFactorTest' not in code
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'src/Materials/CCoupledConcreteReturn.cls'
    (root / name).write_bytes((repo / name).read_bytes())
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    raw = (root / 'Input/Reverse1.bin').read_bytes()
    assert raw[:16] == b'NDMREVTEST000001'
    offset = 20
    parameters = []
    for _ in range(2):
        count, = struct.unpack_from('<i', raw, offset)
        assert 1 <= count <= 100
        offset += 4
        parameters.append(struct.unpack_from('<' + 'd' * count, raw, offset))
        offset += 8 * count
    tolerance = parameters[1][3]
    assert 0 < tolerance < 1e-3
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    code = code.replace('ReturnPrecisionFactorTest = 0.0001', 'ReturnPrecisionFactorTest = 0#')
    block, = re.findall(r'(?ms)^Private Function TestCompleteNewtonLine\(.*?^End Function', code)
    assert block.count('0.0000000001') == 2
    body = block.replace('0.0000000001', 'mReverseOptionsTest(3)').replace('STRICT_RESIDUAL', 'CONFIGURED_ABSOLUTE_RESIDUAL')
    code = code.replace(block, body, 1)
    old = '        revision = solver.AcceptedRevision'
    assert code.count(old) == 1
    code = code.replace(old, old + '\n        Print #handle, "RESIDUAL_TOLERANCE;" & CStr(probeCase) & ";" & TestOnsetNumber(mReverseOptionsTest(3))', 1)
    block, = re.findall(r'(?ms)^Private Sub TestReturnToOrdinaryControl\(.*?^End Sub', code)
    old = '        solver.CopyTrial field, actual'
    assert block.count(old) == 1
    body = block.replace(old, '''        Dim uncertaintyTest As Double, uncertaintyReadyTest As Boolean
        uncertaintyReadyTest = solver.TrialLambdaUncertainty(mReverseOptionsTest(3), mReverseOptionsTest(4), uncertaintyTest)
        Print #report, "BRIDGE_UNCERTAINTY;" & CStr(stepIndex) & ";" & CStr(uncertaintyReadyTest) & ";" & _
            TestOnsetNumber(uncertaintyTest) & ";" & TestOnsetNumber(Abs(factor - currentFactor))
        If Not uncertaintyReadyTest Then
            solver.DiscardStep
            Print #report, "BRIDGE_STOP;" & CStr(stepIndex) & ";LambdaUncertaintyUnavailable;NotPhysicalFailure"
            Exit For
        End If
''' + old, 1)
    code = code.replace(block, body, 1)
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    (root / Path(__file__).name).write_bytes(Path(__file__).read_bytes())
    prep['inputs'][Path(__file__).name] = sha256(Path(__file__).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'GenericConstraintOriginalMaterialAccuracy',
        'configured_absolute_equilibrium_contract': True, 'requested_residual_tolerance_MPa': tolerance,
        'scalar_yield_relative_tolerance': 1e-10, 'coupled_yield_relative_tolerance': 1e-8,
        'coupled_state_relative_tolerance': 5e-10, 'new_local_trials_precision_factor': 0,
        'new_accepted_steps_keep_their_own_material_histories_at_precision_factor': 0,
        'original_material_methods_exact_except_readonly_state_export': True,
        'lambda_uncertainty_recorded_for_ordinary_bridge': True,
        'installed_workbook_changed': False, 'ordinary_solver_recovery_enabled': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('ORIGINAL MATERIAL ACCURACY WITH CONFIGURED ABSOLUTE EQUILIBRIUM CONTRACT READY', tolerance)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
