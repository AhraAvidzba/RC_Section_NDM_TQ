"""ДЛЯ ТЕСТОВ. На точном конце802 уточняет только замыкание внутренних координат истории бетона; не ищет предел заново."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold803NestedAccuracyEndpoint_20261125';root=base/'Cold804HistoryAccuracyEndpoint_20261125'
prior=json.loads((parent/'OperationalColdReplayProof.json').read_bytes())
assert prior['native']['own_process_closed'] and prior['source_cold_exact'] and prior['accepted_states']==45
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'HistoryAccuracyCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());name='src/Materials/CCoupledConcreteResponse.cls'
code=(root/name).read_text('utf-8-sig')
old='Private Const STATE_TOLERANCE As Double = 0.0000000005'
assert code.count(old)==1
code=code.replace(old,'Private Const STATE_TOLERANCE As Double = 0.0000000000005',1)
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
 'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 "-ModuleNames 'CCoupledConcreteResponse'",1).replace('\n','\r\n').encode('utf-8-sig'),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactCold802EndpointWithTighterInternalHistoryClosure',
 actual_import_this_pass=['CCoupledConcreteResponse'],actual_component_count=256,
 same45_cold802_fields_as803=True,only_state_tolerance_changed=True,
 state_tolerance_before=5e-10,state_tolerance_after=5e-13,
 scalar1e13_and_outer1e11_unchanged=True,physical_laws_limits_and_30_96_unchanged=True,
 all_force_gauge_field_and_lambda_acceptance_unchanged=True,
 accepted_source_fields_replayed_with_new_accuracy_not_cached=True,
 source_equilibrium_must_be_rechecked_not_assumed=True,
 no_new_cold_capacity_search=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('804 SAME45 SOURCE FIELDS; INTERNAL HISTORY CLOSURE5e-13; NEW MATERIAL REPLAY REQUIRED; NO CAPACITY SEARCH')
