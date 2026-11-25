"""ДЛЯ ТЕСТОВ. Проверяет холодный расчёт встроенного резерва, не подменяя численную остановку физическим пределом."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from hashlib import sha256
import json
import re
import numpy as np

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold752IntegratedGuideDirected_20261125'
parent = base / 'Cold743UnverifiedCumulativeRecovery_20261125'
helpers = run_path(str(here / 'Check-TestPersistentHEvidence.py'))
regression, _ = helpers['TestNative'](root, 'WorkingGuideRegressions')
assert regression == dict(passed=105, failed=0, macros=3, timeout600=False, own_process_closed=True)
directed = json.loads((root / 'IntegratedGuideDirectedProof.json').read_bytes())
assert directed['independent_NDS_and_original_control_verified'] and directed['total_primary_steps'] == 17
native, text = helpers['TestNative'](root, 'HBackendStock')
compiled_record = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled_record['WorkbookSHA256']
if native['timeout600']:
    outcome = dict(native=native, returned_workflow_result_missing=True, results_reopen_unverified=True)
else:
    saved, = re.findall(r'(?m)^RESULTS_SAVE_REOPEN: .*?equal=True; file=([^\r\n]+)', text)
    (root / 'SavedWorkbook.xlsm').write_bytes(Path(saved).read_bytes())
    with ZipFile(root / 'docs/regression/Diagnostic.xlsm') as book:
        compiled = book.read('xl/vbaProject.bin')
    outcome = helpers['TestStock'](repo, root, compiled)
read = run_path(str(here / 'Read-TestColdRawCapture.py'))
frames, raw = read['TestHistory'](root / 'Native/Stock.RawHistory.bin')
old, _ = read['TestHistory'](parent / 'Native/Stock.RawHistory.bin')
frame, previous = frames[-1], old[-1]
assert raw['incomplete_tail_bytes'] == 0
initial_prefix = 0
for candidate in frames:
    count = 0
    for index in range(min(len(candidate['fields']), len(previous['fields']))):
        if not all(np.array_equal(candidate[key][index], previous[key][index]) for key in ('fields', 'fractions', 'lambdas')):
            break
        count += 1
    initial_prefix = max(initial_prefix, count)
assert (root / 'src/Materials/CCoupledConcreteResponse.cls').read_bytes() == (parent / 'src/Materials/CCoupledConcreteResponse.cls').read_bytes()
rows = run_path(str(here / 'Validate-TestHardPrepared.py'))['TestRows'](root / 'Native/Stock.Checkpoints.csv')
details = [';'.join(row[1:]) for row in rows if row[0] == 'DETAILS']
markers = [line for line in details if 'исходная цель подтверждена после непринятого материального корня' in line]
proof = dict(regression=regression, directed_return=directed, cold=outcome, raw_history=raw,
    raw_frames=len(frames), accepted_states=len(frame['fields']), maximum_exact_prefix_with743=initial_prefix,
    last_accepted_lambda=float(frame['lambdas'][-1]), maximum_accepted_lambda=float(max(frame['lambdas'])),
    reported_original_control_return_checkpoints=len(markers), material_exact743=True,
    actual_cold_all6_not_restarted_from_saved_fields=True, independent_tensor_reintegration_pending=True,
    no_physical_capacity_claimed=True, no_working_promotion=True, whole_goal_completed=False)
(root / 'IntegratedMaterialColdProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('752 COLD', 'TIMEOUT' if native['timeout600'] else outcome['capacity'], len(frame['fields']), 'STATES; MAX', proof['maximum_accepted_lambda'],
      'ORIGINAL CONTROL RETURN CHECKPOINTS', len(markers), 'PHYSICAL LIMIT STILL UNPROVEN')
