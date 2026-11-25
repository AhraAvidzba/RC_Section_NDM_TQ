"""ДЛЯ ТЕСТОВ. Независимо проверяет каждую принятую дуговую пару и равновесие всей автоматически полученной истории."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np
from argparse import ArgumentParser

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
p = ArgumentParser(description=__doc__)
p.add_argument('--root', type=Path, default=base / 'Cold770ConfirmedArcWalk_20261125')
p.add_argument('--material-accuracy', action='store_true')
p.add_argument('--consistent-return', action='store_true')
p.add_argument('--half-predictor', action='store_true')
p.add_argument('--yield-split', action='store_true')
p.add_argument('--source-input', type=Path, default=Path('Input/Reverse1.bin'))
p.add_argument('--prefix', type=int, default=32)
p.add_argument('--event-recovery', action='store_true')
p.add_argument('--event-refinement', action='store_true')
p.add_argument('--inner-resolution', action='store_true')
args = p.parse_args()
root = args.root
parent = base / 'Cold761SignedArcHybrid_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
assert not args.consistent_return or args.material_accuracy
for name in ('src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls', 'src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls'):
    expected = (parent / name).read_text('utf-8-sig')
    if args.yield_split and name == 'src/Solver/CCoupledNewtonSolver.cls':
        expected += (here / 'TestArcYieldBoundary.cls.txt').read_text('utf-8-sig')
    if args.event_refinement and name == 'src/Solver/CCoupledNewtonSolver.cls':
        import re
        method, = re.findall(r'(?ms)^Private Function FollowYieldEventCoordinate\(.*?^End Function', expected)
        refined = method.replace('For attempt = 0 To 1', 'For attempt = 0 To 2', 1).replace('Or attempt = 1 Then Exit For', 'Or attempt = 2 Then Exit For', 1)
        refined = refined.replace("' После застоя уточняется только разностная матрица от той же истории; остаток общего бюджета сохраняется.", "' При застое матрица уточняется до двух раз от той же истории; все попытки делят исходные остатки итераций и вычислений.", 1)
        expected = expected.replace(method, refined, 1)
    if args.material_accuracy and name == 'src/Materials/CCoupledConcreteResponse.cls':
        old = 'Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000001 '
        assert expected.count(old) == 1
        expected = expected.replace(old, 'Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000000001 ', 1)
    if args.consistent_return and name == 'src/Materials/CCoupledConcreteReturn.cls':
        old = 'Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000001'
        assert expected.count(old) == 1
        expected = expected.replace(old, 'Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000000001', 1)
    assert (root / name).read_text('utf-8-sig') == expected, name
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
read_source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary']
source = read_source(root / args.source_input)
prefix = args.prefix
assert len(source['fields']) == prefix and prefix >= 4
validate = run_path(str(here / 'Validate-TestHardPrepared.py'))
mechanics, final_fields = validate['TestValidate'](root, 'ArcWalk.Accepted.Sequence')
operator_file = root / 'Native/ArcWalk.Accepted.Sequence.Operators.csv'
rows = validate['TestRows'](root / 'Native/ArcWalk.Accepted.Sequence.csv')
dim = next(r for r in rows if r[0] == 'DIM')
dofs, physical, points = map(int, dim[1:4])
dz = float(dim[5])
operator_data = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](operator_file, dofs, points)
geometry, loads = operator_data[:2]
b = np.zeros((points, 6, dofs))
gradient = b.copy()
areas = np.zeros(points)
for kind, *r in validate['TestRows'](operator_file):
    if kind == 'B':
        point, degree = int(r[0]) - 1, int(r[1]) - 1
        b[point, :, degree], gradient[point, :, degree] = np.asarray(r[2:8], float), np.asarray(r[8:], float)
    elif kind == 'P':
        areas[int(r[0]) - 1] = float(r[3])
operators = np.asarray([b + side * dz * gradient for side in (-1, 0, 1)])
metric = np.einsum('spcd,spce,p,c->de', operators, operators, areas, np.array([1., 1., 1., .5, .5, .5])) / (3 * sum(areas))
metric[physical:, :] = 0
metric[:, physical:] = 0
scales = np.array([1 / loads[0, 0], 1 / loads[1, 1], 1 / loads[2, 2], 1 / loads[3, 3], 1 / loads[0, 0], 1 / loads[0, 0]])


def TestLength(delta):
    """Вычисляет тензорную длину на трёх станциях без координат связей."""
    return float(np.sqrt(max(0., delta @ metric @ delta)))


def TestBound(factor):
    """Возвращает исходную допустимую невязку для фактического шестикомпонентного сочетания."""
    force = source['offset'] + factor * source['direction']
    return float(source['options'][3] + source['options'][4] * max(abs((loads @ force)[:physical])))


history = list(source['fields'].copy())
factors = list(source['lambdas'].copy())
reader = Reader((root / 'Native/ArcWalk.Attempts.bin').read_bytes())
header = reader.VectorTest()
seed_calls = seed_primary = 0
events = {}
if args.event_recovery:
    assert reader.raw[:16] == b'NDMARCWALK000002'
    assert np.array_equal(header[:6], [prefix, 32, 96, 30, .005, .0001]) and len(header) == 8
    seed_calls, seed_primary = header[6:]
    assert seed_calls == 3 and seed_primary >= 0
    er = Reader((root / 'Native/ArcWalk.Events.bin').read_bytes())
    assert er.raw[:16] == b'NDMARCEVENTS0001'
    while er.position < len(er.raw):
        item = er.VectorTest(); assert len(item) == 8 and int(item[0]) not in events
        events[int(item[0])] = item
    assert er.position == len(er.raw)
else:
    assert reader.raw[:16] == b'NDMARCWALK000001'
    assert np.array_equal(header, [prefix, 32, 96, 30, .005, .0001])
records = []
while reader.position < len(reader.raw):
    records.append(reader.VectorTest())
assert 1 <= len(records) <= 32
attempts = []
event_proofs = []
previous_calls, previous_primary = seed_calls, seed_primary
for number, record in enumerate(records, 1):
    assert len(record) == 11 and record[0] == number
    r = Reader((root / f'Native/ArcWalk.{number:03d}.bin').read_bytes())
    assert r.raw[:16] == b'NDMSIGNEDARC0001'
    info, first, second, cf, ch, projection, accepted, seed, half_seed, full, half, af, ah = [r.VectorTest() for _ in range(13)]
    assert r.position == len(r.raw)
    sf, sh, lf, lh, reference, ratio, alpha, secant, radius, q, target, half_target, rev0, rev1, total, previous, full_budget = info
    assert np.array_equal(accepted, history[-1]) and reference == factors[-1] and previous == factors[-2]
    assert total == 30
    if args.half_predictor:
        assert 0 <= second[0] <= 15 and full_budget == total - second[0] and 0 <= first[0] <= full_budget
    else:
        assert 0 <= first[0] <= full_budget == 15 and 0 <= second[0] <= total - first[0]
    assert rev0 == rev1 == record[7]
    assert ratio == record[3] and 0 < ratio <= 2
    slopes = [TestLength(history[i + 1] - history[i]) / abs(factors[i + 1] - factors[i]) for i in range(len(history) - 4, len(history) - 1)]
    expected_alpha = float(np.median(slopes))
    delta = history[-1] - history[-2]
    dl = reference - previous
    expected_secant = np.sqrt(TestLength(delta)**2 + expected_alpha**2 * dl**2)
    np.testing.assert_allclose([alpha, secant, radius, q], [expected_alpha, expected_secant, ratio * expected_secant, expected_alpha**2 * dl / expected_secant], rtol=1e-9, atol=1e-24)
    np.testing.assert_allclose(projection, metric @ delta / expected_secant, rtol=1e-8, atol=1e-18)
    origin = projection @ accepted + q * reference
    np.testing.assert_allclose([target, half_target], [origin + radius, origin + .5 * radius], rtol=0, atol=1e-14)
    expected_full_seed = 2 * half[:physical] - accepted[:physical] if args.half_predictor and sh else accepted[:physical] + ratio * delta[:physical]
    np.testing.assert_allclose(seed[:physical], expected_full_seed, rtol=1e-14, atol=1e-300)
    np.testing.assert_allclose(half_seed[:physical], accepted[:physical] + .5 * ratio * delta[:physical], rtol=1e-14, atol=1e-300)
    assert np.array_equal(seed[physical:], accepted[physical:]) and np.array_equal(half_seed[physical:], accepted[physical:])
    trial_info = {}
    for label, solved, field, factor, effort, actual, context, requested, local_radius in (
        ('full', sf, full, lf, first, af, cf, target, radius),
        ('half', sh, half, lh, second, ah, ch, half_target, .5 * radius)):
        if not solved:
            assert len(field) == 1
            continue
        bound = TestBound(factor) * .0001
        assert context[0] == context[2] == factor and context[1] == requested and context[3] == q
        constraint = abs(projection @ field + q * factor - requested) * source['calibration'][0]
        raw_six = float(max(abs((actual - source['offset'] - factor * source['direction']) / scales)))
        gauge = float(max(abs(geometry[physical:, :] @ field)) * source['calibration'][0])
        assert effort[-1] <= bound + 1e-14 and constraint <= bound + 1e-13 and raw_six <= bound + 1e-14 and gauge <= bound + 1e-13
        distance = np.sqrt(TestLength(field - accepted)**2 + alpha**2 * (factor - reference)**2)
        progress = float(projection @ (field - accepted) + q * (factor - reference))
        trial_info[label] = dict(lambda_value=float(factor), native_residual_MPa=float(effort[-1]),
            raw_six_force_residual_MPa=raw_six, independently_recomputed_gauge_MPa=gauge,
            constraint_residual_MPa=float(constraint), bound_MPa=bound,
            local=bool(factor > 0 and distance <= 3 * local_radius and progress > 0))
    resolution_scale = .0001 if args.inner_resolution else 1.
    resolution = .5 * radius * source['calibration'][0] > 8 * resolution_scale * max([TestBound(reference), TestBound(lf), TestBound(lh)])
    field_error = lambda_error = -1.
    lambda_bound = 0.
    confirmed = False
    if sf and sh and resolution:
        field_error = TestLength(full - (2 * half - accepted)) / TestLength(full - accepted)
        lambda_error = abs(2 * lh - reference - lf)
        lambda_bound = max(.005 * abs(lf - reference), 8 * source['options'][4] * .0001 * (1 + abs(reference)))
        confirmed = field_error <= .005 and lambda_error <= lambda_bound and all(t['local'] for t in trial_info.values())
        np.testing.assert_allclose(record[4:7], [field_error, lambda_error, lambda_bound], rtol=1e-7, atol=1e-14)
    assert bool(record[2]) == bool(confirmed), (number, record, field_error, lambda_error, lambda_bound)
    assert record[8] == record[7] + int(confirmed)
    assert record[9] - previous_calls == 2 + int(confirmed)
    assert record[10] - previous_primary == first[0] + second[0]
    previous_calls, previous_primary = record[9:11]
    if confirmed:
        history.append(full)
        factors.append(lf)
    assert record[1] == len(history)
    attempts.append(dict(number=number, accepted=bool(confirmed), source_states=len(history) - int(confirmed),
        source_revision=int(rev0), signed_lambda_secant=float(dl), ratio=float(ratio), radius=float(radius),
            resolved_at_selected_force_tolerance=bool(resolution), resolution_scale=resolution_scale, tensor_half_error=float(field_error),
        lambda_half_error=float(lambda_error), lambda_half_bound=float(lambda_bound),
        primary_steps=int(first[0] + second[0]), trials=trial_info))
    if number in events:
        assert not confirmed
        item = events[number]
        assert item[1] == len(history)
        event, event_field, event_lambda = run_path(str(here / 'TestAutoArcEventProof.py'))['TestValidateAutoEvent'](
            root, number, source, history, factors, record,
            [info, first, second, cf, ch, projection, accepted, seed, half_seed, full, half, af, ah],
            operator_data, metric, resolution_scale=resolution_scale)
        assert item[2] == int(event['accepted']) and item[3] == event['nonlinear_calls'] and item[4] == event['primary_steps']
        previous_calls += item[3]; previous_primary += item[4]
        if event_field is not None:
            history.append(event_field); factors.append(event_lambda)
        assert item[5] == len(history) and item[6] == record[8] + int(event['accepted']) and item[7] == event_lambda
        event_proofs.append(event)
assert previous_calls <= 96
assert len(final_fields) == len(history)
np.testing.assert_allclose(final_fields, history, rtol=1e-14, atol=1e-300)
np.testing.assert_allclose([s['lambda'] for s in mechanics['states']], factors, rtol=1e-14, atol=0)
proof = dict(native=native, numerical_working_methods_exact761=not args.material_accuracy and not args.event_refinement, event_third_difference_refinement=args.event_refinement, resolution_uses_actual_inner_tolerance=args.inner_resolution, tighter_CDP_return_only=args.material_accuracy, consistent_nested_return_precision=args.consistent_return, original_prefix32_exact=prefix == 32, source_prefix_count=prefix,
    half_first_full_predictor=args.half_predictor,
    repeated_yield_event_recovery=args.event_recovery, event_proofs=event_proofs,
    initial_event_calls_included=int(seed_calls), initial_event_primary_included=int(seed_primary),
    native_pair_force_checks_and_independent_raw_six_gauges=True,
    all_accepted_states_independent_full_weak_and_six_force_verification=mechanics,
    no_independent_half_step_material_reimplementation_claimed=True,
    new_accepted_states=len(history) - prefix, accepted_states=len(history), attempts=attempts,
    nonlinear_calls=int(previous_calls), total_primary_steps=int(previous_primary),
    maximum_lambda=float(max(factors)), last_lambda=float(factors[-1]),
    accepted_declines=int(sum(np.diff(factors[prefix - 1:]) < 0)),
    stop=(root / 'Native/ArcWalk.Outcome.txt').read_text('cp1251'),
    cold_capacity_search_from_zero=False, physical_capacity_certified=False, whole_goal_completed=False)
(root / 'ConfirmedArcWalkProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('AUTOMATIC ARC WALK VERIFIED', len(history) - prefix, 'NEW STATES', len(records), 'ATTEMPTS', 'MAX', max(factors), 'NO PHYSICAL LIMIT CLAIM')
