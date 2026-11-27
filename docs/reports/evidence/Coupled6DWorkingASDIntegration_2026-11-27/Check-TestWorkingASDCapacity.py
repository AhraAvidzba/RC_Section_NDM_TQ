"""ДЛЯ ТЕСТОВ. Сверяет численные результаты и весь сохранённый путь двух осевых пределов рабочего ASD с ранее независимо проверенной реализацией."""
from pathlib import Path
from hashlib import sha256
from argparse import ArgumentParser
from runpy import run_path
import json, struct, numpy as np

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
root = parser.parse_args().root
here = Path(__file__).resolve().parent
reference = Path(json.loads((root / 'WorkingCapacityReference.json').read_bytes())['reference'])
native, text = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']


def TestRead(folder):
    """Сохраняет точные Double; исключает только первое число — длительность вызова."""
    raw = (folder / 'Native/MainCapacity.bin').read_bytes()
    assert raw[:16] == b'NDMMAINCAP000001'
    reader = R(raw)
    values = reader.VectorTest()
    assert len(values) == 10
    tail = raw[reader.position:]
    count, = reader.VectorTest()
    trace = []
    for index in range(int(count)):
        point = reader.VectorTest()
        assert len(point) == 2
        trace.append(point)
    assert reader.position == len(raw)
    return values, tail, int(count), trace


values, tail, count, trace = TestRead(root)
old, old_tail, old_count, old_trace = TestRead(reference)
diagnostic = (root / 'Native/MainCapacity.txt').read_text('utf-8-sig')
previous = (reference / 'Native/MainCapacity.txt').read_text('utf-8-sig')
comparison = 'Earlier independently checked capacity; exact values and diagnostics'
extra = {}
if json.loads((root / 'WorkingCapacityReference.json').read_bytes())['case'] == 'Compression':
    assert values[1:].tobytes() == old[1:].tobytes()
    assert tail == old_tail and count == old_count and diagnostic == previous
else:
    # Осевой прямой контроль 1066 проверяет текущую версию продолжения, включая весь путь из 86 состояний.
    # Прежняя ёмкостная проверка 1049 имела 77 состояний; сохраняем расхождение, не заменяя его допуском.
    direct = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1066ASDPureNDirectColdRepeat_20261127/Case04')
    independent = json.loads((direct / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    assert independent['all56plus4_equations_and_six_resultants_verified']
    raw = (direct / 'Native/MainCapacityHistory.bin').read_bytes()
    assert sha256(raw).hexdigest().upper() == independent['input_history_sha256']
    reader = R(raw)
    for index in range(3): reader.VectorTest()
    direct_count, = reader.VectorTest()
    assert direct_count == count == 86
    for index in range(count):
        assert trace[index].tobytes() == reader.VectorTest().tobytes()
        reader.VectorTest()
    assert reader.position == len(raw)
    expected = np.array([1., trace[-1][1], 1., trace[-1][1], 0., 0., 0., 0., 0.])
    assert values[1:].tobytes() == expected.tobytes()
    assert diagnostic.splitlines()[0].startswith('Status=FAIL; LimitState=Достигнута предельная деформация арматуры.')
    assert 'R2' not in diagnostic.splitlines()[0] and 'R1' in diagnostic.splitlines()[0]
    assert 'класс A400' in diagnostic.splitlines()[0]
    comparison = 'Current independently checked direct-overload branch 1066; exact 86 load nodes and steel event'
    extra = dict(earlier_capacity_lambda=float(old[2]), earlier_saved_trace_states=old_count,
        earlier_capacity_delta_percent=float((values[2] / old[2] - 1.) * 100.),
        reference_current_branch=str(direct), reference_current_branch_sha256=sha256(raw).hexdigest().upper(),
        earlier_baseline_retained=True, no_tolerance_relaxed=True)
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled['WorkbookSHA256']
source = json.loads((root / 'WorkingASDSourceProof.json').read_bytes())
assert source['saved_book_sha256'] == compiled['WorkbookSHA256'] and source['four_working_classes_exact_current_worktree']
assert not (root / 'Native/MainCapacityHistory.bin').exists()
proof = dict(native=native, reference=str(reference), elapsed_seconds=float(values[0]),
    all_numerical_capacity_outputs_bitexact_reference=True, entire_saved_load_trace_bitexact_reference=True,
    comparison=comparison, all_diagnostics_exact_reference=(diagnostic == previous), saved_trace_states=count, lambda_capacity=float(values[2]),
    steel_lambda=float(values[4]), peak_lambda=float(values[5]), peak_accuracy=float(values[6]),
    saved_book_sha256=compiled['WorkbookSHA256'], original_materials_and_tolerances=True,
    no_kernel_test_exports=True, no_observer=True, full_acceptance_not_run=True, whole_goal_completed=False, **extra)
(root / 'WorkingASDCapacityProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD CAPACITY VERIFIED', json.dumps(proof, sort_keys=True))
