"""ДЛЯ ТЕСТОВ. Восстанавливает все60 принятых полей холодного785 точными Double; нового поиска НДС или предела нет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
source=base/'Cold785OperationalArcController_20261125';parent=base/'Cold787ArcContextOrder_20261125';root=base/'Cold788OperationalColdReplay_20261125'
proof=json.loads((source/'OperationalColdProof.json').read_bytes());assert proof['cold']['native']['own_process_closed']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0 and len(frames[-1]['fields'])==60
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'OperationalColdReplayCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
for call in ('    TestRunArcFailureContext\n','    TestRunOperationalArc\n'):
    assert code.count(call)==1;code=code.replace(call,'',1)
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
    'Input/Source785RawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualCold785StressOnlyReplay',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    source_states=60,cold_source_history_exact=True,source785_material_Newton_arc_exact=True,
    no_new_NDS_search=True,controller_not_called=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('788 EXACT60 COLD785 FIELDS; NATIVE STRESS-ONLY REPLAY; NO CAPACITY SEARCH')
