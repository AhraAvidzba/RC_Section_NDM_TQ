"""ДЛЯ ТЕСТОВ. Подтверждает успешную ветвь половинной проверки и точное восстановление полного all6-кандидата."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, default=base / 'Cold759HalfGuardLoadAllocationFix_20261125')
args = parser.parse_args()
root = args.root
parent = base / 'Cold757HalfGuardedMaterialReturn_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
old_code = (parent / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
new_code = (root / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
allocation_fixed = not new_code.startswith(old_code)
if allocation_fixed:
    old = 'Dim tensorError As Double, lambdaError As Double, length As Double, lambdaBound As Double, load() As Double, tolerance As Double'
    new = 'Dim tensorError As Double, lambdaError As Double, length As Double, lambdaBound As Double, load(1 To COMPONENT_COUNT) As Double, tolerance As Double'
    assert old_code.count(old) == 1
    assert new_code.startswith(old_code.replace(old, new, 1))
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
raw = (root / 'Native/PositiveHalf.Result.bin').read_bytes()
assert raw[:16] == b'NDMPOSHALF000001'
reader = Reader(raw)
summary, before, after, context, projection, full, restored, actual = (reader.VectorTest() for _ in range(8))
assert reader.position == len(raw) and len(summary) == 6
solved, reference, factor, target, rev0, rev1 = summary
assert solved == 1 and reference == .25 and factor == .5 and rev0 == rev1
assert 0 <= before[0] <= after[0] <= 30
assert np.array_equal(full, restored) and context[0] == factor and context[1] == target
assert context[2] == factor and context[3] == 0
mod = run_path(str(here / 'Validate-TestHardPrepared.py'))
mechanics, fields = mod['TestValidate'](root, 'PositiveHalf.Sequence')
assert len(fields) == 2
np.testing.assert_allclose(fields[-1], restored, rtol=1e-14, atol=1e-300)
bound = mechanics['states'][-1]['bound_MPa']
calibration = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')['calibration']
control_residual = abs(float(projection @ restored) - target) * calibration[0]
assert control_residual <= bound + 1e-9
proof = dict(native=native, actual_elastic_all6_prefix=True, reference_lambda=float(reference), full_lambda=float(factor),
    guarded_half_check_returned_true=True, full_candidate_bitexact_restored=True, exact_original_target_restored=True,
    accepted_revision_unchanged=int(rev0), primary_steps_before=int(before[0]), primary_steps_after=int(after[0]), primary_limit=30,
    working_methods_exact757_with_test_wrapper_appended_only=not allocation_fixed,
    only_working_difference_is_fixed_six_component_load_array=allocation_fixed, independent_full_equilibrium=mechanics['states'][-1],
    independent_full_control_residual_MPa=control_residual, no_new_capacity_search=True,
    positive_half_state_not_separately_exported=True, physical_capacity_certified=False, whole_goal_completed=False)
(root / 'PositiveHalfGuardProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('POSITIVE HALF GUARD PASS; EXACT FULL CANDIDATE/TARGET RESTORED; STEPS', before[0], '->', after[0], '<=30; NO CAPACITY CLAIM')
