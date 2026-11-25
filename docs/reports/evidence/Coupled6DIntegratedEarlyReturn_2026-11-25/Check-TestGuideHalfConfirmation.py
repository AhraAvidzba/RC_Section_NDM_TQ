"""ДЛЯ ТЕСТОВ. Отделяет равновесный возврат от подтверждённого половинным шагом локального продолжения."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold756GuideHalfConfirmation_20261125'
parent = base / 'Cold752IntegratedGuideDirected_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
raw = (root / 'Native/GuideHalf.Result.bin').read_bytes()
assert raw[:16] == b'NDMGUIDEHALF0001'
reader = Reader(raw)
summary, first, second, context, projection, accepted, full, half, actual = (reader.VectorTest() for _ in range(9))
assert reader.position == len(raw) and len(summary) == 14
full_solved, half_solved, factor, reference, half_factor, target, half_target, accepted_coordinate, rev0, rev1, budget, consistent, changed, uncertainty = summary
assert full_solved == 1 and half_solved in (0, 1) and rev0 == rev1
assert first[0] == 17 and budget == 30 - first[0] and 0 <= second[0] <= budget
assert first[0] + second[0] <= 30
expected = Reader((parent / 'Native/IntegratedReturn.Result.bin').read_bytes())
old_summary, _, _, _, old_full, _ = (expected.VectorTest() for _ in range(6))
assert np.array_equal(full, old_full) and factor == old_summary[1]
original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert np.array_equal(accepted, original['fields'][-1]) and len(original['fields']) == 46
assert half_target == .5 * (target + accepted_coordinate)
proof = dict(native=native, full_return_bitexact752=True, full_solved=True, half_solved=bool(half_solved),
    full_lambda=float(factor), reference_lambda=float(reference), half_lambda=float(half_factor),
    full_primary_steps=int(first[0]), half_primary_steps=int(second[0]), combined_primary_steps=int(first[0] + second[0]),
    primary_limit=30, accepted_revision_unchanged=int(rev0), original_projection=projection.tolist(),
    target=float(target), half_target=float(half_target), accepted_coordinate=float(accepted_coordinate),
    half_changed_control=bool(changed), half_error_tolerance=.005,
    no_physical_capacity_claimed=True, whole_goal_completed=False)
if half_solved:
    mod = run_path(str(here / 'Validate-TestHardPrepared.py'))
    mechanics, fields = mod['TestValidate'](root, 'GuideHalf.Sequence')
    assert len(fields) == 47
    np.testing.assert_allclose(fields[:-1], original['fields'], rtol=1e-14, atol=1e-300)
    np.testing.assert_allclose(fields[-1], half, rtol=1e-14, atol=1e-300)
    dim = next(row for row in mod['TestRows'](root / 'Native/GuideHalf.Sequence.csv') if row[0] == 'DIM')
    dofs, physical, points = map(int, dim[1:4])
    dz = float(dim[5])
    b = np.zeros((points, 6, dofs))
    gradient = b.copy()
    weights = np.zeros(points)
    for kind, *row in mod['TestRows'](root / 'Native/GuideHalf.Sequence.Operators.csv'):
        if kind == 'B':
            p, d = int(row[0]) - 1, int(row[1]) - 1
            b[p, :, d], gradient[p, :, d] = np.asarray(row[2:8], float), np.asarray(row[8:], float)
        elif kind == 'P':
            weights[int(row[0]) - 1] = float(row[3])
    def TestLength(delta):
        """Использует полную трёхстанционную тензорную метрику без реакций связей."""
        mechanical = np.zeros(dofs)
        mechanical[:physical] = delta[:physical]
        strains = np.asarray([np.einsum('pcd,d->pc', b + side * dz * gradient, mechanical) for side in (-1, 0, 1)])
        return float(np.sqrt(np.einsum('spc,spc,p,c->', strains, strains, weights, np.asarray([1., 1., 1., .5, .5, .5])) / (3 * sum(weights))))
    length = TestLength(full - accepted)
    error = TestLength(full - (2 * half - accepted)) / length
    lambda_error = abs(2 * half_factor - reference - factor)
    lambda_bound = max(.005 * abs(factor - reference), 8 * original['options'][4] * (1 + abs(reference)))
    residual = abs(float(projection @ half) - half_target) * original['calibration'][0]
    bound = mechanics['states'][-1]['bound_MPa']
    assert residual <= bound + 1e-9 and context[1] == half_target
    proof.update(independent_half_equilibrium=mechanics['states'][-1], half_original_control_residual_MPa=residual,
        tensor_relative_half_error=error, lambda_half_error=float(lambda_error), lambda_half_error_bound=float(lambda_bound),
        full_physical_distance=length, half_physical_distance=TestLength(half - accepted),
        locally_confirmed_by_half_step=bool(not changed and error <= .005 and lambda_error <= lambda_bound))
else:
    proof.update(locally_confirmed_by_half_step=False,
        half_failure_details=(root / 'Native/GuideHalf.Outcome.txt').read_text('cp1251'))
(root / 'GuideHalfConfirmationProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('756 HALF SOLVED', bool(half_solved), 'LOCAL HALF CONFIRMATION', proof['locally_confirmed_by_half_step'],
    'STEPS', first[0], '+', second[0], 'FIELD ERROR', proof.get('tensor_relative_half_error'),
    'LAMBDA ERROR', proof.get('lambda_half_error'), 'NO PHYSICAL LIMIT CLAIM')
