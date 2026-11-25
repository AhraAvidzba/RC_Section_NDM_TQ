"""ДЛЯ ТЕСТОВ. Независимо проверяет найденную границу CDP, половинный подход и происхождение последующей дуговой истории."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
p = ArgumentParser(description=__doc__)
p.add_argument('--root', type=Path, default=base / 'Cold778ArcYieldContextFix_20261125')
args = p.parse_args(); root = args.root
parent = base / 'Cold773HalfFirstArcWalk_20261125'
witness = base / 'Cold776ArcPointRoute_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls', 'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
assert (root / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig') == (parent / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig') + (here / 'TestArcYieldBoundary.cls.txt').read_text('utf-8-sig')
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(source['fields']) == 33 and (root / 'Input/Reverse1.bin').read_bytes() == (witness / 'Input/Reverse1.bin').read_bytes()
r = Reader((root / 'Native/ArcYieldSplit.Result.bin').read_bytes())
assert r.raw[:16] == b'NDMARCYIELDSPLIT'
info, event, effort, half_effort, restoration, face, half, projection, actual_face, actual_half = [r.VectorTest() for _ in range(10)]
assert r.position == len(r.raw) and len(info) == 12 and len(event) == 7
solved, half_solved, verified, lf, lh, revision, reported_error, reported_lambda_error, reported_bound, spacing, face_target, half_target = info
assert all(flag in (0, 1) for flag in (solved, half_solved, verified))
assert revision == event[-1] == 35
assert effort[0] <= 15 and effort[0] + half_effort[0] + restoration[0] <= 30
proof = dict(native=native, source_states=33, working_methods_exact773=True,
    event_solved=bool(solved), half_solved=bool(half_solved), event_verified=bool(verified),
    station=event.tolist(), event_primary_steps=int(effort[0]), half_primary_steps=int(half_effort[0]),
    restore_primary_steps=int(restoration[0]), total_primary_steps=int(effort[0] + half_effort[0] + restoration[0]),
    accepted_revision_before_event=int(revision), event_lambda=float(lf), half_lambda=float(lh),
    no_cold_capacity_search_claimed=True, physical_capacity_certified=False, whole_goal_completed=False)
if solved:
    assert event[0] == 1 and event[1] == 15 and event[2] == 1 and event[3] < 0 < event[4] and 0 < event[5] < 1
    operators = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv', 60, 43)
    geometry, loads, w, v, coords, axis, b, gradient = operators
    physical = 56; accepted = source['fields'][-1]
    op = np.array([b + s * .001 * gradient for s in (-1, 0, 1)])
    metric = np.einsum('spcd,spce,p,c->de', op, op, coords[:, 2], np.array([1., 1., 1., .5, .5, .5])) / (3 * sum(coords[:, 2]))
    metric[physical:, :] = 0; metric[:, physical:] = 0
    def TestLength(delta):
        """Измеряет подход к событию по физической метрике трёх станций."""
        return float(np.sqrt(max(0., delta @ metric @ delta)))
    expected_spacing = TestLength(face - accepted)
    np.testing.assert_allclose(spacing, expected_spacing, rtol=1e-10, atol=1e-25)
    np.testing.assert_allclose(projection, metric @ (face - accepted) / spacing, rtol=1e-9, atol=1e-18)
    np.testing.assert_allclose(face_target, projection @ face, rtol=0, atol=1e-14)
    np.testing.assert_allclose(half_target, projection @ accepted + .5 * spacing, rtol=0, atol=1e-14)
    # Координаты истории и упругость читаются из независимого свидетеля776 той же станции.
    rr = Reader((witness / 'Native/ArcPoint15.Route.bin').read_bytes())
    while True:
        identity, strain, sigma = [rr.VectorTest() for _ in range(3)]
        vectors = [rr.VectorTest() for _ in range(9)]
        if identity[0] == 1 and identity[1] == 1:
            committed, parameters = vectors[3], vectors[8]
            break
    diagonal, off, shear, alpha, gamma, *_ = parameters
    elastic = np.diag([diagonal, diagonal, diagonal, shear, shear, shear])
    elastic[:3, :3] += off * (np.ones((3, 3)) - np.eye(3))
    def TestYield(field):
        """Независимо вычисляет пробную поверхность Барселоны от исходной принятой истории."""
        eps = (b[14] + .001 * gradient[14]) @ field
        x, y, z, xy, xz, yz = elastic @ (eps - committed[12:18])
        tensor = np.array([[x, xy, xz], [xy, y, yz], [xz, yz, z]])
        trace = x + y + z; high = np.linalg.eigvalsh(tensor)[-1]
        q = np.sqrt(1.5 * np.sum((tensor - np.eye(3) * trace / 3)**2))
        ct, cc = committed[20:22]
        term = (cc / ct * (1 - alpha) - (1 + alpha)) * high if high > 0 else gamma * high
        return float((alpha * trace + q + term) / (1 - alpha) - cc)
    face_yield = TestYield(face)
    assert abs(face_yield) <= 2.1e-12
    proof['independent_event_yield_MPa'] = face_yield
    if half_solved:
        field_error = TestLength(face - (2 * half - accepted)) / spacing
        lambda_error = abs(2 * lh - source['lambdas'][-1] - lf)
        lambda_bound = max(.005 * abs(lf - source['lambdas'][-1]), 8 * source['options'][4] * .0001 * (1 + abs(source['lambdas'][-1])))
        np.testing.assert_allclose([reported_error, reported_lambda_error, reported_bound], [field_error, lambda_error, lambda_bound], rtol=1e-6, atol=1e-14)
        proof.update(tensor_half_error=float(field_error), lambda_half_error=float(lambda_error), lambda_half_bound=float(lambda_bound),
            independent_half_yield_MPa=TestYield(half), half_arc_constraint_error_MPa=float(abs(projection @ half - half_target) * source['calibration'][0]))
        if verified:
            assert field_error <= .005 and lambda_error <= lambda_bound and TestYield(half) < 0
            assert .5 * spacing * source['calibration'][0] > 8 * (source['options'][3] + source['options'][4] * max(abs((loads @ (source['offset'] + lf * source['direction']))[:physical])))
            scales = np.array([1 / loads[0, 0], 1 / loads[1, 1], 1 / loads[2, 2], 1 / loads[3, 3], 1 / loads[0, 0], 1 / loads[0, 0]])
            for label, field, factor, actual, e in (('face', face, lf, actual_face, restoration), ('half', half, lh, actual_half, half_effort)):
                bound = .0001 * (source['options'][3] + source['options'][4] * max(abs((loads @ (source['offset'] + factor * source['direction']))[:physical])))
                raw_six = float(max(abs((actual - source['offset'] - factor * source['direction']) / scales)))
                gauge = float(max(abs(geometry[physical:, :] @ field)) * source['calibration'][0])
                assert e[-1] <= bound + 1e-14 and raw_six <= bound + 1e-14 and gauge <= bound + 1e-13
                proof[label + '_native_residual_MPa'] = float(e[-1]); proof[label + '_raw_six_residual_MPa'] = raw_six
            assert restoration[0] == 0
            frame = {key: list(source[key]) for key in ('fields', 'fractions', 'lambdas')}
            frame['fields'].append(face); frame['fractions'].append(1.); frame['lambdas'].append(lf)
            data = run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source, frame)
            (root / 'Native/PostEventHistory.bin').write_bytes(data)
            validate = run_path(str(here / 'Validate-TestHardPrepared.py'))
            mechanics, fields = validate['TestValidate'](root, 'ArcWalk.Accepted.Sequence')
            np.testing.assert_allclose(fields[:34], frame['fields'], rtol=1e-14, atol=1e-300)
            proof['event_independent_full_weak_and_six_force_equilibrium'] = mechanics['states'][33]
(root / 'ArcYieldSplitProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('YIELD EVENT', bool(solved), 'HALF', bool(half_solved), 'VERIFIED', bool(verified), 'LAMBDA', lf, 'ERROR', proof.get('tensor_half_error'))
