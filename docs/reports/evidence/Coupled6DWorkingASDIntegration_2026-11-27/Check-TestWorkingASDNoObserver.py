"""ДЛЯ ТЕСТОВ. Проверяет обычный вызов перенесённого рабочего ASD по точному совпадению полей, усилий и статусов с независимо проверенным холодным расчётом."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
import json, struct, numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, default=base / 'Cold1067WorkingASDNoObserver_20261127')
root = parser.parse_args().root
mutating_observer = (root / 'ObserverMutationMode.json').exists()
reference = base / 'Cold1066ASDPureNDirectColdRepeat_20261127'
native, text = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=6, failed=0, macros=1, timeout600=False, own_process_closed=True)
source = json.loads((root / 'WorkingASDSourceProof.json').read_bytes())
assert source['four_working_classes_exact_current_worktree'] and source['no_kernel_test_exports']
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == source['saved_book_sha256']
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']


def TestOptional(reader):
    """Разрешает нулевую длину только для отсутствующего выходного НДС нового явного формата."""
    count, = struct.unpack_from('<i', reader.raw, reader.position)
    if count == 0:
        reader.position += 4
        return np.array([], dtype=float)
    return reader.VectorTest()


def TestState(folder):
    """Читает точные выходные поля и усилия; время исключается только из сравнения численного результата."""
    raw = (folder / 'Native/DirectState.bin').read_bytes()
    reader = R(raw)
    assert raw[:16] == b'NDMASDDIRECT0001'
    meta, field, resultants = reader.VectorTest(), TestOptional(reader), TestOptional(reader)
    assert reader.position == len(raw) and len(meta) == 8
    return meta, field, resultants


cases = []
for index in range(1, 7):
    folder = root / f'Case{index:02d}'
    original = reference / f'Case{index:02d}'
    meta, field, resultants = TestState(folder)
    old_meta, old_field, old_resultants = TestState(original)
    assert meta[[0, 2, 3, 4, 5, 6, 7]].tobytes() == old_meta[[0, 2, 3, 4, 5, 6, 7]].tobytes()
    assert field.tobytes() == old_field.tobytes() and resultants.tobytes() == old_resultants.tobytes(), index
    diagnostic = (folder / 'Native/DirectState.txt').read_text('utf-8-sig')
    previous = (original / 'Native/DirectState.txt').read_text('utf-8-sig')
    assert diagnostic.splitlines()[:2] == previous.splitlines()[:2]
    assert not (folder / 'Native/MainCapacityHistory.bin').exists()
    if mutating_observer:
        calls = R((folder / 'Native/ObserverCalls.bin').read_bytes())
        assert calls.raw[:16] == b'NDMOBSCALLS00001'
        count, mode = calls.VectorTest()
        assert count >= 1 and mode == 1 and calls.position == len(calls.raw)
    status = diagnostic.split(';', 1)[0].removeprefix('Status=')
    cases.append(dict(index=index, target_N=float(meta[0]), elapsed_seconds=float(meta[1]), status=status,
        returned_state=bool(meta[7]), fields_and_resultants_bitexact_reference=True,
        actual_kernel_has_no_test_exports=True))
assert [case['status'] for case in cases] == ['OK', 'OK', 'OK', 'FAIL', 'OK', 'FAIL']
proof = dict(native=native, source=source, cases=cases,
    all6_no_observer_results_bitexact_verified_reference=True,
    ordinary_working_call_verified=True, no_new_independent_equilibrium_solve_claimed=True,
    reference_full_paths_already_independently_verified=True,
    working_output_book_not_updated=True, observer_and_capacity_checks_pending=True,
    full_acceptance_not_run=True, whole_goal_completed=False)
if mutating_observer:
    proof.pop('all6_no_observer_results_bitexact_verified_reference')
    proof['all6_mutating_observer_results_bitexact_verified_reference'] = True
    proof['observer_cannot_change_solver_fields_trace_or_meta'] = True
filename = 'WorkingASDObserverMutationProof.json' if mutating_observer else 'WorkingASDNoObserverProof.json'
(root / filename).write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD OBSERVER ISOLATION VERIFIED' if mutating_observer else 'WORKING ASD NO-OBSERVER VERIFIED', json.dumps(proof, sort_keys=True))
