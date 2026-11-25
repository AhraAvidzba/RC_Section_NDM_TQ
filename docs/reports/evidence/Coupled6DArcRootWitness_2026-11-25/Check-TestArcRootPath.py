"""ДЛЯ ТЕСТОВ. Восстанавливает невязку по материальным тензорам и локализует нелинейность линии между принятым состоянием и корнем."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold775ArcRootPath_20261125'
parent = base / 'Cold774ArcRootDerivative_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls', 'src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(source['fields']) == 33
validate = run_path(str(here / 'Validate-TestHardPrepared.py'))
mechanics, fields = validate['TestValidate'](root, 'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields, source['fields'], rtol=1e-14, atol=1e-300)
reader = Reader((root / 'Input/ArcRoots.bin').read_bytes()); reader.position = 0
full, half, projection, context = [reader.VectorTest() for _ in range(4)]
reader = Reader((root / 'Input/ArcPathOptions.bin').read_bytes()); reader.position = 0
etas = reader.VectorTest()
reader = Reader((root / 'Native/ArcRootPath.Raw.bin').read_bytes())
assert reader.raw[:16] == b'NDMARCROOTPATH01'
info = reader.VectorTest()
revision, modulus, dofs, physical, points, reference, q, half_lambda = info
dofs, physical, points = map(int, (dofs, physical, points))
assert revision == 35 and dofs == 60 and physical == 56 and points == 43
geometry, loads, w, v, coordinates, axis, b, gradient = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv', dofs, points)
accepted = source['fields'][-1]
assert reference == source['lambdas'][-1] == context[5] and q == context[4] and half_lambda == context[1]
queries = []
maximum_reconstruction_error = 0.
for number in range(len(etas) + 1):
    meta, field, residual, stress = [reader.VectorTest() for _ in range(4)]
    assert meta[0] == number + 1 and len(field) == dofs and len(residual) == dofs + 1
    if number < len(etas):
        eta = etas[number]
        assert meta[1] == 0 and meta[2] == eta
        expected_field = accepted + eta * (half - accepted)
        expected_lambda = reference + eta * (half_lambda - reference)
    else:
        assert meta[1] == 1
        eta = None
        expected_field = full
        expected_lambda = context[0]
    np.testing.assert_allclose(field, expected_field, rtol=1e-14, atol=1e-300)
    assert meta[3] == expected_lambda
    target = projection @ field + q * expected_lambda
    assert abs(meta[4] - target) <= 1e-14
    assert len(stress) == 3 * points * 6
    stress = stress.reshape(3, points, 6)
    force = source['offset'] + expected_lambda * source['direction']
    rebuilt = geometry @ field + np.einsum('pcd,pc->d', w, stress[1])
    rebuilt += np.einsum('pcd,pc->d', v, (stress[2] - stress[0]) / .002)
    rebuilt -= loads @ force
    rebuilt[physical:] *= modulus
    gap = float(max(abs(rebuilt - residual[:-1])))
    maximum_reconstruction_error = max(maximum_reconstruction_error, gap)
    assert gap <= 1e-9
    assert abs(residual[-1]) <= 1e-9
    queries.append(dict(number=number + 1, kind=int(meta[1]), eta=None if eta is None else float(eta),
        lambda_value=float(expected_lambda), field=field, residual=residual, stress=stress,
        norm_MPa=float(max(abs(residual))), rebuilt_residual_MPa=rebuilt))
assert np.array_equal(reader.VectorTest(), [revision]) and reader.position == len(reader.raw)
zeros = [t for t in queries if t['eta'] == 0 and t['kind'] == 0]
assert len(zeros) == 2 and np.array_equal(zeros[0]['stress'], zeros[1]['stress']) and np.array_equal(zeros[0]['residual'], zeros[1]['residual'])
origin = zeros[0]
end, = [t for t in queries if t['eta'] == 1 and t['kind'] == 0]
assert end['norm_MPa'] <= 1e-10 and queries[-1]['norm_MPa'] <= 1e-10
results = []
for t in queries:
    result = {key: t[key] for key in ('number', 'kind', 'eta', 'lambda_value', 'norm_MPa')}
    if t['kind'] == 0:
        eta = t['eta']
        nonlinear_stress = t['stress'] - origin['stress'] - eta * (end['stress'] - origin['stress'])
        contribution = np.einsum('pcd,pc->pd', w, nonlinear_stress[1])
        contribution += np.einsum('pcd,pc->pd', v, (nonlinear_stress[2] - nonlinear_stress[0]) / .002)
        contribution[:, physical:] *= modulus
        nonlinear_residual = t['residual'][:-1] - origin['residual'][:-1] - eta * (end['residual'][:-1] - origin['residual'][:-1])
        error = float(max(abs(contribution.sum(axis=0) - nonlinear_residual)))
        assert error <= 1e-9
        row = int(np.argmax(abs(nonlinear_residual)))
        order = np.argsort(abs(contribution[:, row]))[::-1][:5]
        result.update(maximum_nonlinear_residual_MPa=float(abs(nonlinear_residual[row])), row=row + 1,
            point_sum_reconstruction_error_MPa=error,
            largest_point_contributions=[dict(point=int(p + 1), xy=coordinates[p, :2].tolist(), signed_MPa=float(contribution[p, row])) for p in order])
    results.append(result)
proof = dict(native=native, numerical_kernel_exact774=True, source_states=33, queries=results,
    accepted_history_independent_equilibrium=mechanics, maximum_full_residual_reconstruction_error_MPa=maximum_reconstruction_error,
    zero_stress_and_residual_repeat_bitidentical=True, accepted_revision_unchanged=int(revision),
    all_three_station_net_stresses_read_without_accepting_history=True,
    no_interpolated_point_claimed_as_equilibrium=True, no_new_NDS_search=True,
    physical_capacity_certified=False, whole_goal_completed=False)
(root / 'ArcRootPathProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
for result in results:
    print('PATH', result['number'], 'ETA', result['eta'], 'RESIDUAL', result['norm_MPa'], 'NONLINEAR', result.get('maximum_nonlinear_residual_MPa'), 'ROW', result.get('row'))
