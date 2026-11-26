"""ДЛЯ ТЕСТОВ. Добавляет одну подтверждённо полезную долю пробы соседней ветви, сохраняя общий бюджет и исходную приёмку."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold929RecordedTailFineLine_20261126';root=base/'Cold933ProspectiveExtraHalf_20261126'
proof=json.loads((base/'Cold932SurfaceCorrectionScan_20261126/SurfaceCorrectionScanProof.json').read_bytes())
assert any(x.get('mode')=='matched_seed' and x.get('fraction')==1/256 and x.get('improves_original_merit') for x in proof['endpoint_results'])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ProspectiveExtraHalfCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
begin=code.index('Private Function TryProspectiveBranchStep(');end=code.index('End Function',begin)+len('End Function')
body=code[begin:end];anchor='    Do While fraction >= MIN_FRACTION';assert body.count(anchor)==1
body=body.replace(anchor,anchor+' * 0.5 \' Более короткая проба учитывает смену ветви; приёмка по исходной норме сохраняется.',1)
values[name]=(code[:begin]+body+code[end:]).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='MatchedProspectiveOneExtraHalving',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=257,only_change_from929_one_extra_fraction1_over256=True,
 primary_fine_line_exact929=True,primary30_iterations_unchanged=True,original_force_gauge_and_control_unchanged=True,
 material_history_not_accepted=True,record_actual_rejected_directions=True,no_release_promotion=True,
 physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ONE EXTRA HELPER HALF;SAME30;SAME ORIGINAL MERIT;NO HISTORY ACCEPTANCE')
