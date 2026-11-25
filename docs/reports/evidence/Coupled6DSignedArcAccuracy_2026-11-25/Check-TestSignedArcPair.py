"""ДЛЯ ТЕСТОВ. Независимо проверяет определение знаковой псевдодуги, равновесие и локальную согласованность пары шагов."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, default=base / 'Cold760SignedArcPair_20261125')
parser.add_argument('--hybrid', action='store_true')
parser.add_argument('--tight', action='store_true')
parser.add_argument('--prefix', type=int, default=35)
args = parser.parse_args()
root, parent = args.root, base / 'Cold759HalfGuardLoadAllocationFix_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Solver/CCoupledDeformationContinuation.cls', 'src/Materials/CCoupledConcreteResponse.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
expected_newton = (parent / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
if args.hybrid:
    old = '    If Not mControlActive Or mArcLoadProjection <> 0# Or Not SupportsHybridRecovery Then Exit Function'
    new = '''    If Not mControlActive Or Not SupportsHybridRecovery Then Exit Function
    If mArcLoadProjection <> 0# And referenceLambda <> 0# Then Exit Function ' Дуговой резерв не меняет меру и не требует увеличения lambda. При referenceLambda=0 материальные резервы ниже не включаются.'''
    assert expected_newton.count(old) == 1
    expected_newton = expected_newton.replace(old, new, 1).replace(
        "' Для спектральных границ ASD и псевдодуги сохраняются их собственные правила продолжения.",
        "' Для псевдодуги сохраняется исходная плоскость без смены управляющей меры; спектральные границы ASD используют собственные правила.", 1)
assert (root / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig') == expected_newton
original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
frames, _ = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestHistory'](root / 'Input/Source752History.bin')
count = args.prefix
assert 4 <= count <= len(frames[-1]['fields']) and len(original['fields']) == count
for key in ('fields', 'fractions', 'lambdas'):
    assert np.array_equal(original[key], frames[-1][key][:count]), key
raw = (root / 'Native/SignedArc.Result.bin').read_bytes()
assert raw[:16] == b'NDMSIGNEDARC0001'
reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader'](raw)
info, first, second, full_context, half_context, projection, accepted, seed, half_seed, full, half, actual_full, actual_half = (reader.VectorTest() for _ in range(13))
assert reader.position == len(raw) and len(info) == 17
solve_factor = 1.
if args.tight:
    tolerance_reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']((root / 'Native/SignedArc.SolveTolerances.bin').read_bytes())
    tolerance_reader.position = 0
    source_abs, source_rel, local_abs, local_rel, solve_factor = tolerance_reader.VectorTest()
    assert tolerance_reader.position == len(tolerance_reader.raw)
    assert source_abs == original['options'][3] and source_rel == original['options'][4] and solve_factor == .0001
    assert local_abs == source_abs * solve_factor and local_rel == source_rel * solve_factor
sf, sh, lf, lh, reference, ratio, alpha, secant, radius, q, target, half_target, rev0, rev1, total, previous_lambda, full_budget = info
assert sf in (0, 1) and sh in (0, 1) and rev0 == rev1
assert 0 <= first[0] <= full_budget == 15 and 0 <= second[0] <= total - first[0] and total == 30
assert reference == original['lambdas'][-1] and previous_lambda == original['lambdas'][-2]
assert reference != previous_lambda and q * (reference - previous_lambda) > 0
assert np.array_equal(accepted, original['fields'][-1])
mod = run_path(str(here / 'Validate-TestHardPrepared.py'))
original_mechanics, original_fields = mod['TestValidate'](root, 'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(original_fields, original['fields'], rtol=1e-14, atol=1e-300)
rows = mod['TestRows'](root / 'Native/EnvelopeEnergy.Accepted.csv')
dim = next(r for r in rows if r[0] == 'DIM')
dofs, physical, points = map(int, dim[1:4])
dz = float(dim[5])
b = np.zeros((points, 6, dofs))
gradient = b.copy()
areas = np.zeros(points)
for kind, *r in mod['TestRows'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv'):
    if kind == 'B':
        p, d = int(r[0]) - 1, int(r[1]) - 1
        b[p, :, d], gradient[p, :, d] = np.asarray(r[2:8], float), np.asarray(r[8:], float)
    elif kind == 'P':
        areas[int(r[0]) - 1] = float(r[3])
operators = np.asarray([b + side * dz * gradient for side in (-1, 0, 1)])
metric = np.einsum('spcd,spce,p,c->de', operators, operators, areas, np.asarray([1., 1., 1., .5, .5, .5])) / (3 * sum(areas))
metric[physical:, :] = 0
metric[:, physical:] = 0
def TestLength(delta):
    """Возвращает тензорную длину без реакций связей в общей трёхстанционной метрике."""
    return float(np.sqrt(max(0., delta @ metric @ delta)))
slopes = [TestLength(original['fields'][i + 1] - original['fields'][i]) / abs(original['lambdas'][i + 1] - original['lambdas'][i]) for i in range(count - 4, count - 1)]
expected_alpha = float(np.median(slopes))
move = accepted - original['fields'][-2]
dl = reference - previous_lambda
expected_secant = np.sqrt(TestLength(move)**2 + expected_alpha**2 * dl**2)
expected_projection = metric @ move / expected_secant
expected_q = expected_alpha**2 * dl / expected_secant
np.testing.assert_allclose(projection, expected_projection, rtol=1e-10, atol=1e-20)
np.testing.assert_allclose([alpha, secant, radius, q], [expected_alpha, expected_secant, ratio * expected_secant, expected_q], rtol=1e-11, atol=1e-25)
origin = float(projection @ accepted + q * reference)
assert abs(target - (origin + radius)) <= 1e-14
assert abs(half_target - (origin + .5 * radius)) <= 1e-14
np.testing.assert_allclose(seed[:physical], accepted[:physical] + ratio * move[:physical], rtol=1e-14, atol=1e-300)
np.testing.assert_allclose(half_seed[:physical], accepted[:physical] + .5 * ratio * move[:physical], rtol=1e-14, atol=1e-300)
assert np.array_equal(seed[physical:], accepted[physical:]) and np.array_equal(half_seed[physical:], accepted[physical:])
proof = dict(native=native, source_states=count, original_source_fields_verified=True, numerical_core_exact759=not args.hybrid,
    only_761_arc_hybrid_gate_differs_from759=args.hybrid,
    original_history_after_actual_maximum34=count > 34, full_solved=bool(sf), half_solved=bool(sh), accepted_revision_unchanged=int(rev0),
    reference_lambda=float(reference), previous_lambda=float(previous_lambda), median_absolute_slope=alpha,
    signed_load_projection=q, arc_radius=radius, ratio=ratio, full_primary_steps=int(first[0]), half_primary_steps=int(second[0]),
    combined_primary_steps=int(first[0] + second[0]), primary_limit=30,
    arc_plane_and_predictors_independently_verified=True, old_scalar_control_not_requested=True,
    actual_reference_secant_retains_its_sign=True, actual_declining_reference_retains_its_sign=bool(reference < previous_lambda and q < 0), physical_capacity_certified=False, whole_goal_completed=False)
proof['local_solve_tolerance_factor'] = solve_factor
results = {}
for name, solved, field, factor, context, requested, local_radius, effort, actual in (
    ('Full', sf, full, lf, full_context, target, radius, first, actual_full),
    ('Half', sh, half, lh, half_context, half_target, .5 * radius, second, actual_half)):
    if not solved:
        assert len(field) == 1
        continue
    mechanics, fields = mod['TestValidate'](root, 'SignedArc.' + name + '.Sequence')
    assert len(fields) == count + 1
    np.testing.assert_allclose(fields[:-1], original['fields'], rtol=1e-14, atol=1e-300)
    np.testing.assert_allclose(fields[-1], field, rtol=1e-14, atol=1e-300)
    assert context[0] == context[2] == factor and context[1] == requested and context[3] == q
    residual = abs(float(projection @ field + q * factor) - requested) * original['calibration'][0]
    local_bound = mechanics['states'][-1]['bound_MPa'] * solve_factor
    assert effort[-1] <= local_bound + 1e-14
    assert residual <= local_bound + (1e-13 if args.tight else 1e-9)
    load_row = next(r for r in mod['TestRows'](root / ('Native/SignedArc.' + name + '.Sequence.Operators.csv')) if r[0] == 'L' and r[1] == '1')
    area_scale = 1 / float(load_row[2])
    moment_row = next(r for r in mod['TestRows'](root / ('Native/SignedArc.' + name + '.Sequence.Operators.csv')) if r[0] == 'L' and r[1] == '2')
    moment_scale = 1 / float(moment_row[3])
    force_scales = np.asarray([area_scale, moment_scale, moment_scale, moment_scale, area_scale, area_scale])
    raw_force_norm = float(np.max(np.abs((actual - original['offset'] - factor * original['direction']) / force_scales)))
    assert raw_force_norm <= local_bound + 1e-14
    distance = np.sqrt(TestLength(field - accepted)**2 + alpha**2 * (factor - reference)**2)
    progress = float(projection @ (field - accepted) + q * (factor - reference))
    results[name] = dict(independent_equilibrium=mechanics['states'][-1], arc_constraint_residual_MPa=residual,
        local_solve_bound_MPa=local_bound, native_full_residual_raw_MPa=float(effort[-1]), raw_six_force_residual_MPa=raw_force_norm,
        csv_weak_residual_meets_local_tight_bound=bool(mechanics['states'][-1]['residual_MPa'] <= local_bound),
        augmented_distance=float(distance), radius=float(local_radius), augmented_progress=progress,
        within_existing_three_radius_neighbourhood=bool(factor > 0 and distance <= 3 * local_radius and progress > 0),
        spatial_advance=float(move @ metric @ (field - accepted) / TestLength(move)))
proof['trials'] = results
maximum_tolerance = max([original_mechanics['states'][-1]['bound_MPa']] + [v['independent_equilibrium']['bound_MPa'] for v in results.values()])
resolved_half_arc = .5 * radius * original['calibration'][0] > 8 * maximum_tolerance
proof.update(half_arc_increment_MPa=float(.5 * radius * original['calibration'][0]),
    half_arc_exceeds_eight_equilibrium_tolerances=bool(resolved_half_arc))
if sf and sh:
    tensor_error = TestLength(full - (2 * half - accepted)) / TestLength(full - accepted)
    lambda_error = abs(2 * lh - reference - lf)
    lambda_bound = max(.005 * abs(lf - reference), 8 * original['options'][4] * solve_factor * (1 + abs(reference)))
    consistent = resolved_half_arc and tensor_error <= .005 and lambda_error <= lambda_bound and all(v['within_existing_three_radius_neighbourhood'] for v in results.values())
    proof.update(tensor_relative_half_error=tensor_error, lambda_half_error=float(lambda_error), lambda_half_error_bound=float(lambda_bound),
        locally_confirmed_by_half_step=bool(consistent))
else:
    proof.update(locally_confirmed_by_half_step=False)
proof['native_details'] = (root / 'Native/SignedArc.Outcome.txt').read_text('cp1251')
(root / 'SignedArcPairProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('SIGNED ARC PAIR', bool(sf), bool(sh), 'LOCALLY CONFIRMED', proof['locally_confirmed_by_half_step'],
    'STEPS', int(first[0] + second[0]), 'FIELD ERROR', proof.get('tensor_relative_half_error'), 'NO PHYSICAL LIMIT CLAIM')
