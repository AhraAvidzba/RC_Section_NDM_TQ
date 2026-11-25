"""ДЛЯ ТЕСТОВ. Возвращает исходный допуск внутренних координат бетона при неизменных новых алгоритмах и строгом равновесии."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold821PriorityLoadingArc_20261125';root=base/'Cold822LoadingOriginalState_20261125'
assert json.loads((parent/'AutomaticLoadingArcProof.json').read_bytes())['pair']['verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'OriginalStateCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Materials/CCoupledConcreteResponse.cls';code=(root/name).read_text('utf-8-sig')
old='Private Const STATE_TOLERANCE As Double = 0.0000000000005';new='Private Const STATE_TOLERANCE As Double = 0.0000000005'
assert code.count(old)==1;code=code.replace(old,new,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledConcreteResponse']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'",1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalInternalStateToleranceLoadingArc',actual_import_this_pass=imports,actual_component_count=256,
    all_working_numerical_methods_exact821_except_state_tolerance=True,state_tolerance=5e-10,
    physical_material_parameters_and_force_gauge_tolerances_unchanged=True,
    common_call_allowance3_and_primary30_unchanged=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
