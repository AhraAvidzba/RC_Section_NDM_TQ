"""ДЛЯ ТЕСТОВ. Независимо интегрирует все6 усилий и слабое равновесие фактической холодной истории752."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold753IntegratedColdFrozenReplay_20261125'
source = base / 'Cold752IntegratedGuideDirected_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native['passed'] == 1 and native['failed'] == 0 and not native['timeout600']
assert (root / 'Input/Source752History.bin').read_bytes() == (source / 'Native/Stock.RawHistory.bin').read_bytes()
frames, raw = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestHistory'](root / 'Input/Source752History.bin')
frame = frames[-1]
binary = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
for key in ('fields', 'fractions', 'lambdas'):
    assert np.array_equal(binary[key], frame[key]), key
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledNewtonSolver.cls',
             'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root / name).read_bytes() == (source / name).read_bytes(), name
mechanics, fields = run_path(str(here / 'Validate-TestHardPrepared.py'))['TestValidate'](root, 'EnvelopeEnergy.Accepted')
assert len(fields) == len(frame['fields'])
np.testing.assert_allclose(fields, frame['fields'], rtol=1e-14, atol=1e-300)
for state, factor in zip(mechanics['states'], frame['lambdas']):
    assert abs(state['lambda'] - factor) <= 1e-14 * (1 + abs(factor))
proof = dict(native=native, raw_history=raw, states=len(fields), source_fields_and_lambdas_binary_exact=True,
    material_Newton_and_controller_exact752=True, independent_force_and_weak_equilibrium=mechanics,
    no_independent_constitutive_model_claimed=True, no_new_capacity_search=True,
    physical_capacity_certified=False, whole_goal_completed=False)
(root / 'IntegratedColdReplayProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('753 ALL', len(fields), 'COLD752 STATES: INDEPENDENT ALL6/WEAK/GAUGE CHECKED; NO PHYSICAL LIMIT CLAIM')
