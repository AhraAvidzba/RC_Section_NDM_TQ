"""ДЛЯ ТЕСТОВ. Подтверждает отклонение возврата с доказанной несогласованностью шага без потери истории и бюджета."""
from pathlib import Path
from runpy import run_path
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold757HalfGuardedMaterialReturn_20261125'
parent = base / 'Cold752IntegratedGuideDirected_20261125'
evidence = json.loads((base / 'Cold756GuideHalfConfirmation_20261125/GuideHalfConfirmationProof.json').read_bytes())
assert not evidence['locally_confirmed_by_half_step'] and evidence['half_solved']
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledDeformationContinuation.cls', 'tests/modTestCoupledElasticState.bas'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
assert (root / 'Input/Reverse1.bin').read_bytes() == (parent / 'Input/Reverse1.bin').read_bytes()
raw = (root / 'Native/IntegratedReturn.Result.bin').read_bytes()
assert raw[:16] == b'NDMINTEGRET00001'
reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader'](raw)
summary, effort, context, projection, field, actual = (reader.VectorTest() for _ in range(6))
assert reader.position == len(raw) and summary[0] == 0
assert summary[3] == summary[4] == evidence['accepted_revision_unchanged']
assert effort[0] == evidence['combined_primary_steps'] == 20 and effort[0] <= 30
assert len(projection) == len(field) == len(actual) == 1
assert not (root / 'Native/IntegratedReturn.Sequence.csv').exists()
text = (root / 'Native/IntegratedReturn.Outcome.txt').read_text('cp1251')
assert 'полный/половинный возврат: ошибка поля' in text
proof = dict(native=native, nonlocal756_candidate_rejected=True, accepted_revision_unchanged=int(summary[3]),
    combined_primary_steps=int(effort[0]), primary_limit=30, effort=effort.tolist(),
    original46_history_binary_exact752=True, material_controller_test_module_exact752=True,
    no_candidate_published=True, failure_explains_full_half_disagreement=True,
    prior_independent_tensor_error=evidence['tensor_relative_half_error'],
    prior_independent_lambda_error=evidence['lambda_half_error'],
    successful_guard_branch_specific_test_pending=True, physical_capacity_certified=False, whole_goal_completed=False)
(root / 'HalfGuardedMaterialReturnProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('757 NONLOCAL RETURN REJECTED; SAME ACCEPTED46/REVISION', int(summary[3]), 'PRIMARY20/30; POSITIVE GUARD BRANCH STILL PENDING')
