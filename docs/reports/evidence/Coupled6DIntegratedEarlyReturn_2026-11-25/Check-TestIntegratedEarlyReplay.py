"""ДЛЯ ТЕСТОВ. Независимо проверяет точную холодную ветвь754 и не объявляет равновесие физическим пределом."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold755IntegratedEarlyFrozenReplay_20261125'
source = base / 'Cold754IntegratedEarlyEnvelope_20261125'
parent = base / 'Cold753IntegratedColdFrozenReplay_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
assert (root / 'docs/regression/Diagnostic.xlsm').read_bytes() == (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert (root / 'Input/Source754History.bin').read_bytes() == (source / 'Native/Stock.RawHistory.bin').read_bytes()
frames, raw = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestHistory'](root / 'Input/Source754History.bin')
frame = frames[-1]
binary = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
for key in ('fields', 'fractions', 'lambdas'):
    assert np.array_equal(binary[key], frame[key]), key
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledNewtonSolver.cls',
             'src/Materials/CCoupledReinforcedResponse.cls', 'src/Solver/CCoupledElasticState.cls'):
    assert (root / name).read_bytes() == (source / name).read_bytes(), name
mechanics, fields = run_path(str(here / 'Validate-TestHardPrepared.py'))['TestValidate'](root, 'EnvelopeEnergy.Accepted')
assert len(fields) == len(frame['fields'])
np.testing.assert_allclose(fields, frame['fields'], rtol=1e-14, atol=1e-300)
for state, factor in zip(mechanics['states'], frame['lambdas']):
    assert abs(state['lambda'] - factor) <= 1e-14 * (1 + abs(factor))
proof = dict(native=native, raw_history=raw, states=len(fields), source_fields_and_lambdas_binary_exact=True,
    reused_compiled_book_exact753=True, material_Newton_reinforcement_operators_exact754=True,
    different_unused_controller753_not_called=True, independent_force_and_weak_equilibrium=mechanics,
    all_adjacent_spatial_increments_forward=not mechanics['reversed_fields'],
    no_independent_constitutive_model_claimed=True, no_new_capacity_search=True,
    physical_capacity_certified=False, whole_goal_completed=False)
(root / 'IntegratedEarlyReplayProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
assert not mechanics['reversed_fields'], 'Adjacent spatial reversal requires explanation and step verification before branch acceptance.'
print('755 ALL', len(fields), 'COLD754 STATES INDEPENDENT ALL6/WEAK/GAUGE AND FORWARD FIELD CHECKS PASS; NO PHYSICAL LIMIT')
