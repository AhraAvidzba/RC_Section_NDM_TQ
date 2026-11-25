"""ДЛЯ ТЕСТОВ. Проверяет согласованные допуски вложенного CDP на холодном контроллере800, не меняя диаграммы и приёмку."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold800PostArcAccuracy_20261125';root=base/'Cold802NestedMaterialAccuracy_20261125'
prior=json.loads((parent/'OperationalColdProof.json').read_bytes())
assert prior['cold']['native']['own_process_closed'] and prior['accepted_states']==45
assert prior['cold']['capacity']=='NumFail' and not prior['physical_capacity_certified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'NestedAccuracyCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
changes={
 'src/Materials/CCoupledConcreteResponse.cls':('Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000001 ',
 'Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000000001 '),
 'src/Materials/CCoupledConcreteReturn.cls':('Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000001',
 'Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000000001')}
for name,(old,new) in changes.items():
 code=(root/name).read_text('utf-8-sig');assert code.count(old)==1
 values[name]=code.replace(old,new,1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
assert compiler.count("-ModuleNames 'modTestCoupledElasticState'")==1
values.update({
 'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 "-ModuleNames 'CCoupledConcreteResponse','CCoupledConcreteReturn'",1).replace('\n','\r\n').encode('utf-8-sig'),
 'Run-TestPriorBranchMaterials.ps1':(base/'Cold772ArcConsistentAccuracy_20261125/Run-TestPriorBranchMaterials.ps1').read_bytes(),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
 Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ColdPostArcControllerWithConsistentNestedCDPAccuracy',
 actual_import_this_pass=['CCoupledConcreteResponse','CCoupledConcreteReturn'],actual_component_count=256,
 outer_local_yield_before=1e-8,outer_local_yield_after=1e-11,
 inner_scalar_yield_before=1e-10,inner_scalar_yield_after=1e-13,
 original_inner_outer_ratio100_preserved=True,only_two_numerical_constants_changed=True,
 Newton_arc_controller_and_tests_exact800=True,all_inputs_exact800=True,
 physical_material_laws_and_limits_unchanged=True,primary30_global96_unchanged=True,
 original_force_gauge_and_step_acceptance_unchanged=True,
 cold_recompute_required=True,no_mutable_material_state_reuse=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('802 TWO NESTED NUMERICAL TOLERANCES TIGHTER1000x; SAME800 CONTROLLER/DIAGRAMS/30/96; COLD REQUIRED')
