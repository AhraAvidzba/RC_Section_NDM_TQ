"""ДЛЯ ТЕСТОВ. Использует неизменную753 для чтения точной истории754; изменённый контроллер при повторе не вызывается."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
source = base / 'Cold754IntegratedEarlyEnvelope_20261125'
parent = base / 'Cold753IntegratedColdFrozenReplay_20261125'
root = base / 'Cold755IntegratedEarlyFrozenReplay_20261125'
proof = json.loads((source / 'IntegratedEarlyColdProof.json').read_bytes())
assert proof['cold']['native']['own_process_closed']
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Solver/CCoupledNewtonSolver.cls',
             'src/Materials/CCoupledReinforcedResponse.cls', 'src/Solver/CCoupledElasticState.cls'):
    assert (source / name).read_bytes() == (parent / name).read_bytes(), name
test_name = 'tests/modTestCoupledElasticState.bas'
code = (source / test_name).read_text('utf-8-sig')
assert code.count('    TestRunIntegratedMaterialReturn\n') == 1
assert code.replace('    TestRunIntegratedMaterialReturn\n', '', 1) == (parent / test_name).read_text('utf-8-sig')
frames, raw = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestHistory'](source / 'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes'] == 0
original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source / 'Input/Reverse1.bin')
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
prep = json.loads((root / 'Preparation.json').read_bytes())
values = {'Input/Reverse1.bin': run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original, frames[-1]),
    'Input/Source754History.bin': (source / 'Native/Stock.RawHistory.bin').read_bytes(),
    Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
assert (root / 'docs/regression/Diagnostic.xlsm').read_bytes() == (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope'] = dict(case='ActualCold754FrozenReplay', actual_import_this_pass=['modTestCoupledElasticState'],
    unchanged_actual255_compiled753_reused=True, fresh_compilation_claimed=False,
    material_Newton_reinforcement_operators_exact754=True, source_states=len(frames[-1]['fields']), source_raw_history_exact=True,
    replay_controller_is753_not754_and_not_called=True, no_new_NDS_or_capacity_search=True, no_working_promotion=True,
    physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('755 IMMUTABLE753 BOOK; EXACT754 MATERIAL/NEWTON; FIELDS', len(frames[-1]['fields']), 'CONTROLLER NOT CALLED')
