"""ДЛЯ ТЕСТОВ. Переиспользует неизменяемую компиляцию788 для напряжений всей холодной истории787, без нового поиска и компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold788OperationalColdReplay_20261125';cold=base/'Cold787ArcContextOrder_20261125';root=base/'Cold790FixedContextColdReplay_20261125'
proof=json.loads((cold/'OperationalColdProof.json').read_bytes());assert proof['accepted_states']==56 and proof['cold']['native']['own_process_closed']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](cold/'Native/Stock.RawHistory.bin');assert raw['incomplete_tail_bytes']==0
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
original=read(cold/'Input/Reverse1.bin');parent_input=read(parent/'Input/Reverse1.bin')
for k in ('calibration','options','offset','direction'):assert np.array_equal(original[k],parent_input[k]),k
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
values={'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
    'Input/Source787RawHistory.bin':(cold/'Native/Stock.RawHistory.bin').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope']=dict(case='ActualCold787StressOnlyReplay',actual_import_this_pass=[],actual_component_count=256,
    immutable_compiled788_reused=True,fresh_compilation_claimed=False,source_states=56,source_cold787_exact=True,
    source787_material_Newton_arc_exact=True,no_new_NDS_search=True,controller_not_called=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('790 EXACT56 COLD787 FIELDS; COMPILED788 REUSED; STRESS ONLY; NO FRESH COMPILATION OR CAPACITY SEARCH')
