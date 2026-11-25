"""ДЛЯ ТЕСТОВ. Проверяет нулевую физическую проекцию связей781 и подтверждает, что она сама по себе не устраняет силовую невязку."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold781GaugeProjectionWitness_20261125'
parent = base / 'Cold775ArcRootPath_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes()
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(source['fields']) == 43
validate = run_path(str(here / 'Validate-TestHardPrepared.py'))
history, fields = validate['TestValidate'](root, 'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields, source['fields'], rtol=1e-14, atol=1e-300)
g, loads, w, v, coords, axis, b, gradient = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
rr = Reader((root / 'Native/ArcRootPath.Raw.bin').read_bytes()); info = rr.VectorTest()
nodes = [[rr.VectorTest() for _ in range(4)] for _ in range(4)]
assert np.array_equal(rr.VectorTest(), [45]) and rr.position == len(rr.raw)
assert nodes[0][0][2] == nodes[2][0][2] == 0 and nodes[1][0][2] == 1 and nodes[3][0][1] == 1
assert np.array_equal(nodes[0][2], nodes[2][2]) and np.array_equal(nodes[0][3], nodes[2][3])
max_reconstruction_error = 0.
for meta, field, residual, stress in nodes:
    stress = stress.reshape(3,43,6)
    rebuilt = g @ field + np.einsum('pcd,pc->d',w,stress[1]) + np.einsum('pcd,pc->d',v,(stress[2]-stress[0])/.002)
    rebuilt -= loads @ (source['offset'] + meta[3] * source['direction'])
    rebuilt[56:] *= source['calibration'][0]
    gap = float(max(abs(rebuilt - residual[:-1]))); max_reconstruction_error=max(max_reconstruction_error,gap)
    assert gap <= 1e-9
original, projected = nodes[3], nodes[1]
delta = projected[1] - original[1]
ops = np.array([b+s*.001*gradient for s in (-1,0,1)])
strain_gap = float(max(abs(np.einsum('spcd,d->spc',ops,delta)).ravel()))
stress_gap = float(max(abs(projected[3]-original[3])))
assert strain_gap <= 1e-17 and stress_gap <= 1e-10
before = float(max(abs(original[2][56:60]))); after = float(max(abs(projected[2][56:60])))
assert after < 1e-12 and before > 1e-10
G = g[:56,56:]; mu = np.linalg.solve(G.T@G,-G.T@projected[2][:56])
predicted = projected[2].copy(); predicted[:56] += G@mu
proof = dict(native=native, source_states=43, immutable_compiled775_reused=True,
    accepted_history_independent_equilibrium=history,
    maximum_independent_residual_reconstruction_error_MPa=max_reconstruction_error,
    maximum_point_strain_change=strain_gap, maximum_net_stress_change_MPa=stress_gap,
    original_weak_norm_MPa=float(max(abs(original[2]))), projected_weak_norm_MPa=float(max(abs(projected[2]))),
    original_gauge_norm_MPa=before, projected_gauge_norm_MPa=after,
    independently_predicted_multiplier_fit_delta=mu.tolist(),
    predicted_weak_norm_after_multiplier_fit_MPa=float(max(abs(predicted))),
    geometric_projection_does_not_solve_force_equilibrium=True,
    no_new_equilibrium_search=True, physical_capacity_certified=False, whole_goal_completed=False)
(root/'GaugeProjectionWitnessProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('781 VERIFIED GAUGE', before,'->',after,'STRESS CHANGE',stress_gap,'FORCE NORM REMAINS',proof['projected_weak_norm_MPa'])
