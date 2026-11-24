"""ДЛЯ ТЕСТОВ. Возвращает исходные21 предписанные точки708 при новом прогнозе дилатансии710."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold711PriorDilationCurve_20261124';parent=base/'Cold710PriorDilationSeed_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PriorCurveCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='tests/modTestCoupledElasticState.bas'
values[path]=(base/'Cold708ExactCohesionClosure_20261124'/path).read_bytes()
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='Original21CurveWithPriorDilationSeed',actual_import_this_pass=['modTestCoupledElasticState'],
    original708_test_module_exact=True,original21_targets_and_all_reference_tolerances_exact=True,
    candidate710_material_exact=True,expected_complete_curve_assertions=194,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('711 ORIGINAL21 CURVE; EXACT708 TEST MODULE; EXACT710 MATERIAL')
