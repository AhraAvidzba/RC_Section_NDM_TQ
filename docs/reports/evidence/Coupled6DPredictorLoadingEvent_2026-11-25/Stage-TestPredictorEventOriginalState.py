"""ДЛЯ ТЕСТОВ. Проверяет новый поиск перехода с исходной точностью внутренних координат; изолирует рабочее исправление от отрицательного опыта804."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold803NestedAccuracyEndpoint_20261125';root=base/'Cold812PredictorEventOriginalState_20261125'
proof=json.loads((base/'Cold807PredictorLoadingEvent_20261125/ColdEndpointArcProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'OriginalStateCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes())
name='src/Solver/CCoupledArcContinuation.cls'
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values={name:(base/'Cold807PredictorLoadingEvent_20261125'/name).read_bytes(),
 'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 "-ModuleNames 'CCoupledArcContinuation'",1).replace('\n','\r\n').encode('utf-8-sig'),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
for n in ('src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
 assert (root/n).read_bytes()==(base/'Cold802NestedMaterialAccuracy_20261125'/n).read_bytes()
prep['directed_scope']=dict(case='PredictorLoadingEventWithOriginalStatePrecision',
 actual_import_this_pass=['CCoupledArcContinuation'],actual_component_count=256,
 arc_exact807=True,all_other_code_and45fields_exact803=True,
 material_classes_exact802=True,original_state_tolerance5e10=True,
 no_new_capacity_search=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('812 SAME807 EVENT FIX WITH ORIGINAL802 MATERIAL STATE ACCURACY; FIXED45 ENDPOINT ONLY')
