"""ДЛЯ ТЕСТОВ. Сверяет холодный754 с ранним745 и отделяет автоматический возврат от подтверждения физического предела."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from hashlib import sha256
import json
import re
import numpy as np

repo, here = Path.cwd(), Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold754IntegratedEarlyEnvelope_20261125'
previous = base / 'Cold745EarlyGrowingEnvelope_20261125'
parent = base / 'Cold752IntegratedGuideDirected_20261125'
helper = run_path(str(here / 'Check-TestPersistentHEvidence.py'))
regression, _ = helper['TestNative'](root, 'WorkingGuideRegressions')
assert regression == dict(passed=105, failed=0, macros=3, timeout600=False, own_process_closed=True)
native, text = helper['TestNative'](root, 'HBackendStock')
compiled_record = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled_record['WorkbookSHA256']
if native['timeout600']:
    outcome = dict(native=native, returned_workflow_result_missing=True, results_reopen_unverified=True)
else:
    saved, = re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)', text)
    (root / 'SavedWorkbook.xlsm').write_bytes(Path(saved).read_bytes())
    with ZipFile(root / 'docs/regression/Diagnostic.xlsm') as book:
        compiled = book.read('xl/vbaProject.bin')
    outcome = helper['TestStock'](repo, root, compiled)
read = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestHistory']
frames, raw = read(root / 'Native/Stock.RawHistory.bin')
old, _ = read(previous / 'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes'] == 0
for key in ('fields', 'fractions', 'lambdas'):
    assert np.array_equal(frames[0][key], old[0][key]), key
assert len(frames[0]['fields']) == 4
prefix = 0
for f in frames:
    count = 0
    for i in range(min(len(f['fields']), len(old[-1]['fields']))):
        if not all(np.array_equal(f[k][i], old[-1][k][i]) for k in ('fields', 'fractions', 'lambdas')):
            break
        count += 1
    prefix = max(prefix, count)
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledNewtonSolver.cls', 'tests/modTestCoupledElasticState.bas'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
rows = run_path(str(here / 'Validate-TestHardPrepared.py'))['TestRows'](root / 'Native/Stock.Checkpoints.csv')
details = [';'.join(row[1:]) for row in rows if row[0] == 'DETAILS']
assert details and details[0].startswith('Мера деформации: материальная координата H огибающих;')
markers = [line for line in details if 'исходная цель подтверждена после непринятого материального корня' in line]
reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']((parent / 'Native/IntegratedReturn.Result.bin').read_bytes())
returned_summary, _, _, _, returned_field, _ = (reader.VectorTest() for _ in range(6))
assert returned_summary[0] == 1
returned_matches = []
for snapshot in frames:
    for index, (field, factor) in enumerate(zip(snapshot['fields'], snapshot['lambdas'])):
        if factor == returned_summary[1] and np.array_equal(field, returned_field):
            returned_matches.append(dict(snapshot=snapshot['number'], history_index=index + 1))
frame = frames[-1]
proof = dict(regression=regression, cold=outcome, raw_history=raw, raw_frames=len(frames), accepted_states=len(frame['fields']),
    all_four_initial_force_states_exact745=True, maximum_exact_prefix_with745=prefix,
    initial_H_control_selected_before_first_controlled_probe=True, reported_original_control_return_checkpoints=len(markers),
    exact_directed752_return_present_in_accepted_history=bool(returned_matches), accepted_return_matches=returned_matches,
    missing_diagnostic_marker_alone_does_not_prove_return_not_called=True,
    last_accepted_lambda=float(frame['lambdas'][-1]), maximum_accepted_lambda=float(max(frame['lambdas'])),
    final_control_details=details[-1], material_Newton_test_module_exact752=True,
    actual_cold_all6_not_restarted_from_saved_fields=True, independent_tensor_reintegration_pending_at_cold_recording=True,
    no_physical_capacity_claimed=True, no_working_promotion=True, whole_goal_completed=False)
(root / 'IntegratedEarlyColdProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('754 COLD', 'TIMEOUT' if native['timeout600'] else outcome['capacity'], len(frame['fields']), 'STATES; MAX', proof['maximum_accepted_lambda'],
    'PREFIX745', prefix, 'ORIGINAL RETURN MARKERS', len(markers), 'NO PHYSICAL LIMIT CLAIM')
