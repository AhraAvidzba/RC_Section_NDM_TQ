"""ДЛЯ ТЕСТОВ. Измеряет приближение фактической невязки к нулевому приращению, не объявляя численный предел физическим."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold765AcceptedMapContinuity_20261125'
parent = base / 'Cold761SignedArcHybrid_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls', 'src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
raw = (root / 'Native/MapContinuity.Result.bin').read_bytes()
assert raw[:16] == b'NDMMAPLIMIT00001'
reader = Reader(raw)
info, accepted, previous, external = (reader.VectorTest() for _ in range(4))
factor, revision, modulus, physical, dofs = info
physical, dofs = int(physical), int(dofs)
cases, = reader.ReadTest('<i')
option_reader = Reader((root / 'Input/MapContinuityOptions.bin').read_bytes())
option_reader.position = 0
etas = option_reader.VectorTest()
assert cases == len(etas) == 20
original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(original['fields']) == 35 and factor == original['lambdas'][-1]
assert np.array_equal(accepted, original['fields'][-1]) and np.array_equal(previous, original['fields'][-2])
samples = []
for index in range(cases):
    metadata, field, weak = (reader.VectorTest() for _ in range(3))
    eta, ready, changed = metadata
    assert eta == etas[index] and ready in (0, 1) and field.shape == (dofs,)
    assert changed == np.count_nonzero(field[:physical] != accepted[:physical])
    expected = accepted.copy()
    expected[:physical] += eta * (accepted[:physical] - previous[:physical])
    np.testing.assert_allclose(field, expected, rtol=4 * np.finfo(float).eps, atol=1e-300)
    assert np.array_equal(field[physical:], accepted[physical:])
    assert weak.shape == ((dofs,) if ready else (1,))
    samples.append(dict(index=index + 1, eta=float(eta), ready=bool(ready), changed=int(changed), field=field, weak=weak))
last_revision, = reader.VectorTest()
assert reader.position == len(raw) and last_revision == revision
first, last = samples[0], samples[-1]
assert first['eta'] == last['eta'] == 0 and first['ready'] and last['ready']
base_weak = first['weak']
scales = np.ones(dofs)
scales[physical:] = modulus
base_residual = float(np.max(abs((base_weak - external) * scales)))
bound = float(original['options'][3] + original['options'][4] * np.max(abs(external[:physical])))
roundoff_estimate = float(128 * np.finfo(float).eps * max(1., np.max(abs(base_weak * scales))))
rows = []
for sample in samples:
    value = dict(index=sample['index'], eta=sample['eta'], ready=sample['ready'], changed_coordinates=sample['changed'])
    if sample['ready']:
        delta = (sample['weak'] - base_weak) * scales
        row = int(np.argmax(abs(delta)))
        value.update(maximum_weak_change_MPa=float(abs(delta[row])), largest_change_row=row + 1,
            maximum_field_change=float(np.max(abs(sample['field'][:physical] - accepted[:physical]))))
    rows.append(value)
limits = {}
for label, sign in (('positive', 1), ('negative', -1)):
    tail = [next(row for row in rows if row['eta'] == sign * 10.**(-k)) for k in (6, 7, 8)]
    if all(r['ready'] for r in tail):
        values = [r['maximum_weak_change_MPa'] for r in tail]
        ratios = [values[i + 1] / values[i] if values[i] else None for i in (0, 1)]
        limits[label] = dict(samples=tail, successive_change_ratios=ratios,
            observed_near_zero_plateau=bool(all(r is not None and r > .8 for r in ratios) and values[-1] > max(1e-12, 128 * roundoff_estimate)))
    else:
        limits[label] = dict(samples=tail, missing_near_zero_response=True)
proof = dict(native=native, source_states=35, numerical_core_exact761=True, accepted_revision_unchanged=int(revision),
    original_equilibrium_bound_MPa=bound, zero_increment_residual_MPa=base_residual,
    zero_increment_passes_original_weak_equilibrium=base_residual <= bound,
    first_last_zero_weak_bitidentical=bool(np.array_equal(first['weak'], last['weak'])),
    floating_roundoff_scale_estimate_MPa=roundoff_estimate, samples=rows, near_zero=limits,
    no_new_equilibrium_search=True, no_numerical_or_physical_capacity_claimed=True, whole_goal_completed=False)
(root / 'AcceptedMapContinuityProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
assert proof['first_last_zero_weak_bitidentical'], 'Read-only probes changed the exact-zero response; inspect before further continuation.'
print('765 ZERO WEAK RESIDUAL', base_residual, 'BOUND', bound, 'ZERO REPEAT BITEXACT')
for label, value in limits.items():
    print(label, value)
