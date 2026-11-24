"""ДЛЯ ТЕСТОВ. Совмещает проверенную кривую материала711 с адаптивным контроллером697 для автоматического поиска из двух упругих точек."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold712PriorDilationCapacity_20261124';parent=base/'Cold711PriorDilationCurve_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PriorCapacityCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Solver/CCoupledDeformationContinuation.cls'
values[path]=(base/'Cold697PostCrackRankTrust_20261124'/path).read_bytes()
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['Run-TestRealPostCrack.ps1']=(base/'Cold697PostCrackRankTrust_20261124/Run-TestRealPostCrack.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='AutomaticCapacityWithPriorDilationMaterial',actual_import_this_pass=['CCoupledDeformationContinuation'],
    controller697_exact=True,material711_exact=True,actual_capacity_macro_default_assembled_tangent=True,
    only_two_elastic_initial_states=True,no_future_curve_or_peak_seed=True,original96_probe_limit=True,
    original30_primary_iterations=True,original_capacity_accuracy=.001,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('712 AUTOMATIC CAPACITY; DEFAULT ASSEMBLED TANGENT; EXACT697 CONTROLLER AND711 MATERIAL')
