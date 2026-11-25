"""ДЛЯ ТЕСТОВ. Переиспользует неизменяемый читающий проект для точной холодной истории; неисполняемый контроллер может отличаться."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--source',default='Cold797ArcGrowthControl_20261125');p.add_argument('--root',default='Cold799ArcGrowthColdReplay_20261125')
args=p.parse_args();source=base/args.source;root=base/args.root;parent=base/'Cold798AdaptiveColdReplay_20261125'
outcome=json.loads((source/'OperationalColdProof.json').read_bytes());assert outcome['cold']['native']['own_process_closed']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin');assert raw['incomplete_tail_bytes']==0
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];original=read(source/'Input/Reverse1.bin');prepared=read(parent/'Input/Reverse1.bin')
for k in ('calibration','options','offset','direction'):assert np.array_equal(original[k],prepared[k]),k
for n in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (source/n).read_bytes()==(parent/n).read_bytes(),n
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
values={'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
    'Input/SourceColdRawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope']=dict(case='ActualAdaptiveColdStressReplayReused',actual_import_this_pass=[],actual_component_count=256,
    source_root=source.name,source_states=len(frames[-1]['fields']),all_source_cold_fields_exact=True,
    immutable_compiled798_reused=True,fresh_compilation_claimed=False,residual_material_Newton_and_arc_sources_exact=True,
    source_controller_may_differ_but_never_executed=True,no_new_NDS_search=True,controller_not_called=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'EXACT COLD STATES',len(frames[-1]['fields']),'READONLY COMPILED798 REUSED; CONTROLLER NOT CALLED')
