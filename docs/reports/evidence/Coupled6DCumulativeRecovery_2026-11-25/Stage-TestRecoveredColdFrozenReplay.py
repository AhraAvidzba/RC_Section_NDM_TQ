"""ДЛЯ ТЕСТОВ. Повторно использует неизменную742 для чтения материальной истории743 без повторной компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold744RecoveredColdFrozenReplay_20261125';parent=base/'Cold742RealColdFrozenReplay_20261125'
source=base/'Cold743UnverifiedCumulativeRecovery_20261125'
proof=json.loads((source/'CumulativeRecoveryColdProof.json').read_bytes())
assert proof['initial_before_recovery_all12_source_states_exact']
assert proof['cold743']['native']['own_process_closed']
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls'):
    assert (source/name).read_bytes()==(parent/name).read_bytes(),name
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin')
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
values={'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
    'Input/Source743History.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope'].update(case='RecoveredActualCold743FrozenReplay',source_states=len(frames[-1]['fields']),
    source_raw_history_exact=True,source_numerical_material_and_Newton_exact742=True,
    controller743_not_called_during_stress_only_replay=True,unused_replay_controller_is742_not743=True,
    unchanged_actual255_compiled_source_reused=True,fresh_compilation_claimed=False,
    no_new_NDS_or_capacity_search=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('744 IMMUTABLE742 COMPILED BOOK; EXACT743 MATERIAL HISTORY',len(frames[-1]['fields']),'FIELDS; NO CONTROLLER CALL')
