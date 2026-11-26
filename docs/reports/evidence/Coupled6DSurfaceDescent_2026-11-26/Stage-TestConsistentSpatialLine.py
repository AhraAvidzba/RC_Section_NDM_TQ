"""ДЛЯ ТЕСТОВ. Согласует нижнюю долю пространственного основного поиска и резерва новой ветви, не меняя лимиты и приёмку."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold933ProspectiveExtraHalf_20261126';root=base/'Cold934ConsistentSpatialLine_20261126'
proof=json.loads((parent/'ProspectiveExtraHalfProof.json').read_bytes());assert all(not x['solved'] for x in proof['cases'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ConsistentSpatialLineCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private Const MIN_FRACTION As Double = 0.0078125';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'\nPrivate Const SPATIAL_MIN_FRACTION As Double = 0.0000152587890625 \' Единая нижняя доля управляемой пространственной пробы и резерва соседней ветви.',1)
old='If mControlActive And mUseSpatialHistory Then minimumFraction = 0.0000152587890625';assert code.count(old)==1
code=code.replace(old,'If mControlActive And mUseSpatialHistory Then minimumFraction = SPATIAL_MIN_FRACTION',1)
old='    Do While fraction >= MIN_FRACTION * 0.5 \' Более короткая проба учитывает смену ветви; приёмка по исходной норме сохраняется.'
assert code.count(old)==1
code=code.replace(old,'    Do While fraction >= SPATIAL_MIN_FRACTION \' Резерв проверяет тот же диапазон долей, что и основной пространственный шаг.',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ConsistentSpatialPrimaryAndProspectiveLine',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=257,shared_primary_and_helper_minimum_fraction=1/65536,
 additional_Newton_iterations_not_granted=True,primary30_and_original_equilibrium_unchanged=True,
 all_material_laws_exact933=True,material_history_never_accepted_by_recovery=True,
 actual_rejected_directions_recorded=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ONE SHARED SPATIAL FLOOR;SAME30;SAME MATERIAL/ACCEPTANCE')
