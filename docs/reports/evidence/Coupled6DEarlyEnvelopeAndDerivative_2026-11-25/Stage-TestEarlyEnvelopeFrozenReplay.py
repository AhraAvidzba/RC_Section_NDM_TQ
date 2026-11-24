"""ДЛЯ ТЕСТОВ. Готовит независимую выгрузку всех состояний раннего H-контроля745 неизменной книгой744."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold747EarlyEnvelopeFrozenReplay_20261125';parent=base/'Cold744RecoveredColdFrozenReplay_20261125'
source=base/'Cold745EarlyGrowingEnvelope_20261125'
proof=json.loads((source/'EarlyEnvelopeColdProof.json').read_bytes())
assert proof['cold745']['native']['own_process_closed'] and proof['all_four_initial_force_states_exact743']
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledNewtonSolver.cls'):
    assert (source/name).read_bytes()==(parent/name).read_bytes(),name
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin')
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
values={'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
    'Input/Source745History.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope']=dict(case='EarlyEnvelope745FrozenReplay',actual_import_this_pass=['modTestCoupledElasticState'],
    source_states=len(frames[-1]['fields']),source_raw_history_exact=True,material_and_Newton_exact745=True,
    controller_not_called_during_stress_only_replay=True,unchanged_actual255_compiled_source_reused=True,
    fresh_compilation_claimed=False,no_new_NDS_or_capacity_search=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('747 IMMUTABLE744 COMPILED BOOK; EXACT745 HISTORY',len(frames[-1]['fields']),'FIELDS; NO CONTROLLER CALL')
