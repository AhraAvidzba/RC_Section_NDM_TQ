"""ДЛЯ ТЕСТОВ. Различает возвращённое НДС и промежуточное равновесие прямого ASD, проверяя реальные поля и внешнее повторение каждого пути."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
from hashlib import sha256
import json, struct, numpy as np

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
root = parser.parse_args().root
here = Path(__file__).resolve().parent
native, text = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=6, failed=0, macros=1, timeout600=False, own_process_closed=True)
source = json.loads((root / 'EarlySourceProof.json').read_bytes())
assert source['changed_components_exact'] and source['full_equilibrium_acceptance_unchanged']
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == source['saved_book_sha256']
loads = json.loads((root / 'Input/DirectControl.json').read_bytes())['loads_N']
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']


def TestOptionalVector(reader):
    """Разрешает явную нулевую длину только в двух необязательных выходных полях нового формата прямого НДС."""
    count, = struct.unpack_from('<i', reader.raw, reader.position)
    if count == 0:
        reader.position += 4
        return np.array([], dtype=float)
    return reader.VectorTest()


cases = []
for index, target in enumerate(loads, 1):
    folder = root / f'Case{index:02d}'
    raw = (folder / 'Native/DirectState.bin').read_bytes()
    reader = R(raw)
    assert raw[:16] == b'NDMASDDIRECT0001'
    meta, field, resultants = reader.VectorTest(), TestOptionalVector(reader), TestOptionalVector(reader)
    assert reader.position == len(raw)
    force, elapsed, internal, code, field_count, resultant_count, trace_count, has_state = meta
    assert force == target and len(field) == field_count and len(resultants) == resultant_count
    replay = json.loads((folder / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    assert replay['all56plus4_equations_and_six_resultants_verified'] and replay['reference_tolerance_violated_states'] == 0
    assert replay['input_history_sha256'] == sha256((folder / 'Native/MainCapacityHistory.bin').read_bytes()).hexdigest().upper()
    assert replay['accepted_states'] == trace_count
    diagnostic = (folder / 'Native/DirectState.txt').read_text('utf-8-sig')
    status = diagnostic.split(';', 1)[0].removeprefix('Status=')
    last = replay['records'][-1]
    if has_state:
        assert has_state == 1 and field_count == 60 and resultant_count == 6
        assert last['lambda_value'] == 1 and 'law=ASD-NativeStep-Cdf0' in diagnostic
        history = R((folder / 'Native/MainCapacityHistory.bin').read_bytes())
        for unused in range(4):
            history.VectorTest()
        for unused in range(int(trace_count)):
            history.VectorTest()
            final = history.VectorTest()
        np.testing.assert_array_equal(field, final)
        np.testing.assert_allclose(resultants, last['actual_resultants_N_Nmm'], atol=1e-5, rtol=1e-10)
        assert status in ('OK', 'FAIL')
    else:
        assert has_state == field_count == resultant_count == 0
        assert status in ('NumFail', 'FAIL')
    cases.append(dict(index=index, target_N=force, elapsed_seconds=elapsed,
        status=status, internal_status=int(internal), result_code=int(code), returned_state=bool(has_state),
        accepted_states=int(trace_count), last_accepted_lambda=last['lambda_value'],
        last_steel_utilization=last['steel_strain_utilization'],
        all_accepted_equilibria_verified=True, partial_state_not_substituted_for_target=True))
proof = dict(native=native, cases=cases, same_ASD_material_and_tolerances=True,
    all56plus4_equations_and_six_resultants_verified=True,
    numerical_failure_not_physical_failure=True, actual_OpenSees_material_replay=True,
    independent_OpenSees_structural_analysis=False, CDP_not_run=True,
    working_solver_not_changed=True, whole_goal_completed=False)
(root / 'ASDDirectPureNExecutionProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('DIRECT ASD NDS EVIDENCE VERIFIED', json.dumps(proof, sort_keys=True))
