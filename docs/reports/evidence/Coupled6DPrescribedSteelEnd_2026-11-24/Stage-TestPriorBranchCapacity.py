"""ДЛЯ ТЕСТОВ. Подключает проверенный локальный вариант722 к исправленному автоматическому контроллеру719."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold723PriorBranchCapacity_20261124';parent=base/'Cold719ScopedPeakRefinement_20261124'
material=base/'Cold722PriorBranchClosure_20261124'
proof=json.loads((material/'PriorBranchClosureProof.json').read_bytes())
assert proof['full_original21_prescribed_curve_passed'] and proof['native']['failed']==0
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PriorBranchCapacityCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes())
values={'src/Materials/CCoupledConcreteResponse.cls':(material/'src/Materials/CCoupledConcreteResponse.cls').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes()}
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='PriorBranchAutomaticCapacity',actual_import_this_pass=['CCoupledConcreteResponse'],
    source_material722_exact=True,controller_and_Newton719_exact=True,test_module_and_expectations719_exact=True,
    original96_probes_30_iterations_and_capacity_accuracy=True,full_cold_from_zero=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('723 FULL COLD AUTOMATIC CAPACITY; SAME719 CONTROLLER/TESTS; VERIFIED722 MATERIAL')
