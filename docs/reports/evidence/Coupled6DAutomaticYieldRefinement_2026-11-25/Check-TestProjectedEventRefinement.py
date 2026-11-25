"""ДЛЯ ТЕСТОВ. Сравнивает отказ780 и уточнение782 на одной побитной истории, не подменяя численное событие пределом."""
from pathlib import Path
from runpy import run_path
import json
import re
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
failed = base / 'Cold780RejectedYieldEvent_20261125'
refined = base / 'Cold782ProjectedEventRefinement_20261125'
reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
read_source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary']
native_check = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative']
validate = run_path(str(here / 'Validate-TestHardPrepared.py'))
source = read_source(failed / 'Input/Reverse1.bin')
assert len(source['fields']) == 43
for name in ('Input/Reverse1.bin', 'Input/ArcRoots.bin', 'Input/Source779Attempt013.bin', 'Input/Source779Event013.bin'):
    assert (failed / name).read_bytes() == (refined / name).read_bytes(), name
code_name = 'src/Solver/CCoupledNewtonSolver.cls'
before = (failed / code_name).read_text('utf-8-sig')
method, = re.findall(r'(?ms)^Private Function FollowYieldEventCoordinate\(.*?^End Function', before)
changed = method.replace('For attempt = 0 To 1', 'For attempt = 0 To 2', 1).replace('Or attempt = 1 Then Exit For', 'Or attempt = 2 Then Exit For', 1)
changed = changed.replace("' После застоя уточняется только разностная матрица от той же истории; остаток общего бюджета сохраняется.", "' При застое матрица уточняется до двух раз от той же истории; все попытки делят исходные остатки итераций и вычислений.", 1)
assert (refined / code_name).read_text('utf-8-sig') == before.replace(method, changed, 1)
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls', 'src/Solver/CCoupledDeformationContinuation.cls', 'tests/modTestCoupledElasticState.bas'):
    assert (failed / name).read_bytes() == (refined / name).read_bytes(), name
operators = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](refined / 'Native/EnvelopeEnergy.Accepted.Operators.csv', 60, 43)
g, loads, w, v, coordinates, axis, b, gradient = operators
op = np.array([b + s * .001 * gradient for s in (-1, 0, 1)])
metric = np.einsum('spcd,spce,p,c->de', op, op, coordinates[:, 2], np.array([1., 1., 1., .5, .5, .5])) / (3 * sum(coordinates[:, 2]))
metric[56:, :] = 0; metric[:, 56:] = 0


def TestLength(delta):
    """Измеряет физическое приращение в трёх станциях, исключая служебные множители."""
    return float(np.sqrt(max(0., delta @ metric @ delta)))


rr = reader((failed / 'Input/Source779Event013.bin').read_bytes())
source_event = [rr.VectorTest() for _ in range(11)]
alpha, gamma = source_event[-1][:2]; plastic = source_event[-1][2:8]; cc, ct = source_event[-1][8:]
modulus, poisson = source['calibration'][:2]
off = modulus * poisson / ((1 + poisson) * (1 - 2 * poisson)); shear = modulus / (2 * (1 + poisson))
elastic = np.diag([off + 2 * shear] * 3 + [shear] * 3)
elastic[:3, :3] += off * (np.ones((3, 3)) - np.eye(3))


def TestYield(field):
    """Вычисляет пробную поверхность точки39/+1 от фактически сохранённой истории779."""
    xx, yy, zz, xy, xz, yz = elastic @ ((b[38] + .001 * gradient[38]) @ field - plastic)
    tensor = np.array([[xx, xy, xz], [xy, yy, yz], [xz, yz, zz]])
    trace = xx + yy + zz; high = np.linalg.eigvalsh(tensor)[-1]
    q = np.sqrt(1.5 * np.sum((tensor - np.eye(3) * trace / 3)**2))
    term = (cc / ct * (1 - alpha) - (1 + alpha)) * high if high > 0 else gamma * high
    return float((alpha * trace + q + term) / (1 - alpha) - cc)


proofs = {}
for label, root in (('failed780', failed), ('refined782', refined)):
    native, _ = native_check(root, 'RealColdReplay')
    assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
    rr = reader((root / 'Native/ArcYieldSplit.Result.bin').read_bytes())
    info, event, effort, half_effort, restoration, face, half, projection, actual_face, actual_half = [rr.VectorTest() for _ in range(10)]
    assert rr.position == len(rr.raw) and len(event) == 77 and event[1] == 39 and event[2] == 1 and event[6] == 45
    candidate = event[17:]; distance = TestLength(candidate - source['fields'][-1])
    np.testing.assert_allclose(distance, event[8], rtol=1e-8, atol=1e-23)
    assert distance <= event[7] and abs(TestYield(candidate)) <= 2.1e-12 and abs(TestYield(candidate) - event[14]) < 2.1e-12
    gauge = float(max(abs(g[56:] @ candidate)) * modulus)
    assert bool(event[12]) == (gauge <= event[9])
    total_primary = int(effort[0] + half_effort[0] + restoration[0]); assert total_primary <= 30
    result = dict(native=native, solved=bool(info[0]), half_solved=bool(info[1]), accepted=bool(info[2]),
        candidate_lambda=float(event[15]), candidate_force_norm_MPa=float(event[11]), candidate_gauge_MPa=gauge,
        candidate_yield_MPa=TestYield(candidate), candidate_distance=distance, radius=float(event[7]),
        inner_force_bound_MPa=float(event[9]), total_primary_steps=total_primary)
    if label == 'failed780':
        assert not info[0] and event[11] > event[9] and gauge > event[9]
    else:
        assert all(info[:3] == 1) and np.array_equal(face, candidate) and restoration[0] == 0
        spacing = TestLength(face - source['fields'][-1]); error = TestLength(face - (2 * half - source['fields'][-1])) / spacing
        np.testing.assert_allclose(info[6], error, rtol=1e-6, atol=1e-14)
        assert error <= .005 and TestYield(half) < 0
        scale = np.array([1 / loads[0, 0], 1 / loads[1, 1], 1 / loads[2, 2], 1 / loads[3, 3], 1 / loads[0, 0], 1 / loads[0, 0]])
        for field, factor, actual, e in ((face, info[3], actual_face, restoration), (half, info[4], actual_half, half_effort)):
            bound = .0001 * (source['options'][3] + source['options'][4] * max(abs((loads @ (source['offset'] + factor * source['direction']))[:56])))
            assert e[-1] <= bound + 1e-14 and max(abs((actual - source['offset'] - factor * source['direction']) / scale)) <= bound + 1e-14
            assert max(abs(g[56:] @ field)) * modulus <= bound + 1e-13
        full_weak, fields = validate['TestValidate'](root, 'ArcWalk.Accepted.Sequence')
        assert len(fields) == 44
        np.testing.assert_allclose(fields[:43], source['fields'], rtol=1e-14, atol=1e-300)
        np.testing.assert_allclose(fields[43], face, rtol=1e-14, atol=1e-300)
        original_bound = event[9] / .0001
        resolution = .5 * spacing * modulus
        assert resolution > 8 * event[9] and resolution < 8 * original_bound
        uncertainty = json.loads((base / 'Cold783SmallEventCorrection_20261125/ArcRootDerivativeProof.json').read_bytes())
        # Поправка от независимо снятых матриц нужна отдельно от силовой невязки: малое число здесь не означает предел.
        result.update(tensor_half_error=error, half_increment_MPa=resolution,
            resolution_passes_original_user_bound=False, resolution_passes_actual_inner_bound=True,
            all44_states_independent_equilibrium=full_weak, independent_correction_evidence='Cold783SmallEventCorrection_20261125/ArcRootDerivativeProof.json')
        assert uncertainty['physical_capacity_certified'] is False
    proofs[label] = result
regression, _ = native_check(refined, 'WorkingGuideRegressions')
assert regression == dict(passed=105, failed=0, macros=3, timeout600=False, own_process_closed=True)
proof = dict(same43_source_states_and_actual_rejected_bracket=True, changed_only_third_difference_attempt=True,
    shared_original_primary_and_function_budgets=True, cases=proofs, targeted_regressions=regression,
    physical_capacity_certified=False, whole_goal_completed=False)
(refined / 'ProjectedEventRefinementProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('780 FAILURE VERIFIED;782 FULL/HALF EVENT AND44 WEAK/ALL6 STATES VERIFIED;105 REGRESSIONS PASS; INNER RESOLUTION SUPPORTED783')
