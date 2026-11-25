"""ДЛЯ ТЕСТОВ. Проверяет маршруты CDP-точки15, сохранность истории и соответствие штатного ответа обычному возврату."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold776ArcPointRoute_20261125'
parent = base / 'Cold775ArcRootPath_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls', 'src/Materials/CCoupledConcreteReturn.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
expected = (parent / 'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
actual = (root / 'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
assert actual.startswith(expected) and actual[len(expected):].count('Friend Function ReadOperationalPointRouteTest(') == 1
assert actual[len(expected):] == (here / 'TestPointUpdateRoute.cls.txt').read_text('utf-8-sig')
assert (root / 'Native/ArcRootPath.Raw.bin').read_bytes() == (parent / 'Native/ArcRootPath.Raw.bin').read_bytes()
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
ops = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv', 60, 43)
b, slope = ops[-2:]
raw_reader = Reader((root / 'Native/ArcRootPath.Raw.bin').read_bytes())
raw_reader.VectorTest()
queries = {}
for index in range(1, 15):
    info, field, residual, stress = [raw_reader.VectorTest() for _ in range(4)]
    assert info[0] == index
    queries[index] = dict(info=info, field=field, stress=stress.reshape(3, 43, 6))
assert np.array_equal(raw_reader.VectorTest(), [35]) and raw_reader.position == len(raw_reader.raw)


def TestState(strain, state, stress, parameters):
    """Независимо восстанавливает эффективный тензор, поверхность и номинальные напряжения по записанной истории."""
    diagonal, off, shear, alpha, gamma, fc, ft, band, rounding = parameters
    elastic = np.diag([diagonal, diagonal, diagonal, shear, shear, shear])
    elastic[:3, :3] += off * (np.ones((3, 3)) - np.eye(3))
    effective = elastic @ (strain - state[12:18])
    xx, yy, zz, xy, xz, yz = effective
    tensor = np.array([[xx, xy, xz], [xy, yy, yz], [xz, yz, zz]])
    principal = np.linalg.eigvalsh(tensor)
    trace = float(xx + yy + zz)
    dev = tensor - np.eye(3) * trace / 3
    q = float(np.sqrt(1.5 * np.sum(dev * dev)))
    r = float(np.sum(np.maximum(principal, 0)) / np.sum(abs(principal))) if np.sum(abs(principal)) else .5
    ct, cc, dt, dc = state[20:24]
    high = principal[-1]
    term = (cc / ct * (1 - alpha) - (1 + alpha)) * high if high > 0 else gamma * high
    value = (alpha * trace + q + term) / (1 - alpha) - cc
    apparent = (1 - r * dt) * (1 - dc) * effective
    gap = float(max(abs(apparent - stress)))
    return dict(effective_trace_MPa=trace, effective_q_MPa=q, effective_principal=principal.tolist(),
        apparent_reconstruction_error_MPa=gap, yield_value_MPa=float(value),
        tension_coordinate=float(state[18]), compression_coordinate=float(state[19]),
        tension_damage=float(dt), compression_damage=float(dc), tension_energy=float(state[24]), compression_energy=float(state[25]))


reader = Reader((root / 'Native/ArcPoint15.Route.bin').read_bytes())
assert reader.raw[:16] == b'NDMARCPOINT00001'
records = []
while reader.position < len(reader.raw):
    identity, strain, interface_stress = [reader.VectorTest() for _ in range(3)]
    route, work, ordinary_work, committed, updated, ordinary, stress, ordinary_stress, parameters = [reader.VectorTest() for _ in range(9)]
    index, side, point = map(int, identity)
    assert point == 15 and side in (-1, 0, 1) and 1 <= index <= 14
    assert len(route) == 6 and len(work) == len(ordinary_work) == 5 and len(parameters) == 9
    assert len(committed) == len(updated) == len(ordinary) == 26
    q = queries[index]
    np.testing.assert_allclose(strain, (b[14] + side * .001 * slope[14]) @ q['field'], rtol=1e-14, atol=1e-300)
    assert np.array_equal(interface_stress, q['stress'][side + 1, 14])
    result = dict(query=index, kind=int(q['info'][1]), eta=float(q['info'][2]), side=side, point=point,
        operational_ready=bool(route[0]), substeps=int(route[1]), operational_elastic=bool(route[2]),
        ordinary_ready=bool(route[3]), ordinary_elastic=bool(route[4]), ordinary_needs_closure=bool(route[5]),
        operational_counts=work.tolist(), ordinary_counts=ordinary_work.tolist(),
        ordinary_apex_attempted=bool(ordinary_work[1] > 70),
        accepted_history=committed.tolist())
    result['trial_yield_MPa'] = TestState(strain, committed, interface_stress, parameters)['yield_value_MPa']
    if route[0]:
        result['operational_interface_stress_max_difference_MPa'] = float(max(abs(stress - interface_stress)))
        result['operational_state'] = TestState(strain, updated, stress, parameters)
    if route[3]:
        result['ordinary_interface_stress_max_difference_MPa'] = float(max(abs(ordinary_stress - interface_stress)))
        result['ordinary_state'] = TestState(strain, ordinary, ordinary_stress, parameters)
    records.append(result)
assert len(records) == 42
for side in (-1, 0, 1):
    r = [x for x in records if x['side'] == side]
    assert len(r) == 14 and all(x['accepted_history'] == r[0]['accepted_history'] for x in r)
proof = dict(native=native, source_states=33, numerical_working_methods_exact775=True,
    all775_residual_and_stress_path_bytes_exact=True, all_three_station_accepted_histories_unchanged=True,
    records=records, new_equilibrium_search=False, physical_capacity_certified=False, whole_goal_completed=False)
(root / 'ArcPointRouteProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
for r in records:
    if r['eta'] in (.1, .5, .75, 1., 2.) or r['kind'] == 1:
        print('ROUTE', r['query'], 'ETA', r['eta'], 'SIDE', r['side'], 'SUBSTEPS', r['substeps'],
            'ORDINARY', r['ordinary_ready'], 'ITER', r['ordinary_counts'][1], 'CLOSURES', r['operational_counts'][2:4],
            'CACHE GAP', r.get('operational_interface_stress_max_difference_MPa'), 'ORD GAP', r.get('ordinary_interface_stress_max_difference_MPa'))
