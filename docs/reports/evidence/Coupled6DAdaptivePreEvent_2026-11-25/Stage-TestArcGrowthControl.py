"""ДЛЯ ТЕСТОВ. Включает существующий контроль роста пространственного шага после подтверждённого дугового восстановления."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold794PreEventBudget_20261125';root=base/'Cold797ArcGrowthControl_20261125'
proof=json.loads((parent/'OperationalColdProof.json').read_bytes());assert proof['arc_recovered_directions']==1 and proof['accepted_states']==67
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ArcGrowthCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Function TryArcContinuationProbe\(.*?^End Function',code)
old='    mBranchTangentControl = True: mMaterialGrowthActive = False: Erase mMaterialGrowthProjection\n';assert method.count(old)==1
updated=method.replace(old,'''    mBranchTangentControl = True: mMaterialGrowthActive = True: mMaterialGrowthProjection = mProjection ' Рост ограничен существующей проверкой прогноза секущей и прекращается при изменении точной меры.
''',1)
code=code.replace(method,updated,1)
old="' Разрешает адаптацию только для той же точной пространственной меры, установленной проверенным CDP-переходом."
new="' Разрешает адаптацию для той же точной пространственной меры подтверждённого материального или дугового шага."
assert code.count(old)==1;code=code.replace(old,new,1)
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'",1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingSecantGrowthAfterConfirmedArc',actual_import_this_pass=['CCoupledDeformationContinuation'],actual_component_count=256,
    only_existing_growth_activation_for_exact_verified_projection=True,newton_arc_material_and_tests_exact794=True,
    existing2percent_secant_prediction_and125growth_and_larger_step_rejection_unchanged=True,
    primary30_global96_and_all_physical_acceptance_unchanged=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('797 REUSE EXISTING SECANT STEP GROWTH AFTER CONFIRMED ARC; SAME794 KERNEL/30/96/PHYSICAL CHECKS')
