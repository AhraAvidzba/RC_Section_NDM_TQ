"""ДЛЯ ТЕСТОВ. Проверяет встроенный возврат, исходную цель, неизменность истории и общий бюджет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold752IntegratedGuideDirected_20261125'
original_root = base / 'Cold747EarlyEnvelopeFrozenReplay_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native['passed'] == 1 and native['failed'] == 0 and not native['timeout600']
assert (root / 'Input/Reverse1.bin').read_bytes() == (original_root / 'Input/Reverse1.bin').read_bytes()
original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(original['fields']) == 46
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
expected_raw = (root / 'Input/ExpectedOriginalControl.bin').read_bytes()
assert expected_raw[:16] == b'NDMDEPARTCAND001'
reader = Reader(expected_raw)
expected, accepted, guide, auxiliary_projection, original_projection = (reader.VectorTest() for _ in range(5))
assert reader.position == len(expected_raw) and np.array_equal(accepted, original['fields'][-1])
raw = (root / 'Native/IntegratedReturn.Result.bin').read_bytes()
assert raw[:16] == b'NDMINTEGRET00001'
reader = Reader(raw)
summary, effort, context, projection, field, actual = (reader.VectorTest() for _ in range(6))
assert reader.position == len(raw) and len(summary) == 6 and len(effort) == 8 and len(context) == 4
solved, factor, reference, revision_before, revision_after, uncertainty = summary
proof = dict(native=native, solved=bool(solved), returned_lambda=float(factor), reference_lambda=float(reference),
    total_primary_steps=int(effort[0]), effort=effort.tolist(), primary_limit=int(original['options'][2]),
    accepted_revision_before=int(revision_before), accepted_revision_after=int(revision_after),
    accepted_revision_unchanged=bool(revision_before == revision_after), lambda_uncertainty=float(uncertainty),
    source_states=46, actual_single_working_recovery_call=True, source_result_sha256=sha256(raw).hexdigest().upper(),
    diagnostic_sequence_not_controller_branch_admission=True, cold_capacity_not_tested=True,
    physical_capacity_certified=False, whole_goal_completed=False)
assert solved in (0, 1) and revision_before == revision_after
assert 0 <= effort[0] <= original['options'][2] == 30
assert reference == expected[1]
if solved:
    assert field.shape == projection.shape == (60,) and actual.shape == (6,)
    assert np.array_equal(projection, original_projection)
    assert context[1] == expected[6] and context[0] == factor
    assert np.isfinite(uncertainty) and uncertainty >= 0
    resolution = 8 * original['options'][4] * (1 + abs(reference))
    assert factor > reference + resolution + uncertainty
    mod = run_path(str(here / 'Validate-TestHardPrepared.py'))
    mechanics, fields = mod['TestValidate'](root, 'IntegratedReturn.Sequence')
    assert len(fields) == 47 and mechanics['details'][0] == 'DiagnosticReSolve'
    np.testing.assert_allclose(fields[:-1], original['fields'], rtol=1e-14, atol=1e-300)
    np.testing.assert_allclose(fields[-1], field, rtol=1e-14, atol=1e-300)
    rows = mod['TestRows'](root / 'Native/IntegratedReturn.Sequence.csv')
    bound = next(float(r[3]) for r in rows if r[0] == 'R' and r[1] == '47')
    control = float(projection @ field)
    assert abs(control - expected[6]) * original['calibration'][0] <= bound + 1e-9
    advance = float(original_projection @ (field - accepted))
    assert advance > 0
    proof.update(independent_NDS_and_original_control_verified=True, independent_return_equilibrium=mechanics['states'][-1],
        original_control_target=float(expected[6]), original_control_actual=control,
        original_control_residual_MPa=float(abs(control - expected[6]) * original['calibration'][0]),
        control_bound_MPa=bound, signed_advance_in_original_projection=advance,
        original_lambda_guard_verified=True, all46_original_states_verified=True,
        auxiliary_reverse_root_excluded_from_return_sequence=True)
else:
    assert len(projection) == len(field) == len(actual) == 1
    proof.update(independent_NDS_and_original_control_verified=False,
        native_outcome=(root / 'Native/IntegratedReturn.Outcome.txt').read_text('cp1251'))
(root / 'IntegratedGuideDirectedProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
assert solved == 1, 'Integrated recovery failed; preserve evidence and correct implementation, not acceptance expectations.'
print('752 INTEGRATED ORIGINAL RETURN VERIFIED; LAMBDA', factor, 'PRIMARY STEPS', effort[0], '<=30; NO PHYSICAL LIMIT CLAIM')
