"""ДЛЯ ТЕСТОВ. Проверяет ещё пять уменьшений реальной пространственной поправки; силовые допуски и бюджеты не меняются."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold911NewtonRejectedCapture_20261126';root=base/'Cold912ExtendedSpatialDamping_20261126'
proof=json.loads((parent/'NewtonRejectedCaptureProof.json').read_bytes())
assert proof['baseline8_cases_and_all_root_exports_bit_identical'] and proof['recorded_rejected_directions']==2
assert all(x['probes'][-1]['governing_direction_opposes_prediction'] for x in proof['records'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ExtendedDampingCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='        If mSteelBranchBuilds > 0 Then minimumFraction = STEEL_MIN_FRACTION'
new='''        If mSteelBranchBuilds > 0 Or (mControlActive And mUseSpatialHistory) Then minimumFraction = STEEL_MIN_FRACTION ' Меньшие пробы нужны у близкой смены пространственной ветви; приёмка равновесия не меняется.'''
assert code.count(old)==1;code=code.replace(old,new,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExtendedSpatialLineSearchAfterActualCapture',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=256,minimum_spatial_control_fraction=1/4096,original_minimum_fraction=1/128,
 exactly_five_additional_halvings=True,no_force_acceptance_or_material_changes=True,
 original_primary30_global96_unchanged=True,no_working_or_release_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'FIVE SMALLER TRIALS;SAME EQUILIBRIUM ACCEPTANCE;SAME ITERATION/CALL BUDGETS')
