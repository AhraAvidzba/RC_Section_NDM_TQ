"""ДЛЯ ТЕСТОВ. Продолжает неизменную21-точечную кривую до заданного предела стали0,02."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold722PriorBranchClosure_20261124';root=base/'Cold724PriorBranchSteelEnd_20261124'
assert json.loads((parent/'PriorBranchClosureProof.json').read_bytes())['full_original21_prescribed_curve_passed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SteelEndCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='0.0018, 0.002, 0.0025, 0.003, 0.004)';assert code.count(anchor)==1
code=code.replace(anchor,'0.0018, 0.002, 0.0025, 0.003, 0.004, 0.006, 0.01, 0.015, 0.02)',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='PrescribedSteelEndOfPriorBranch',actual_import_this_pass=['modTestCoupledElasticState'],
    original21_targets_and_acceptance_unchanged=True,additional_targets=[.006,.01,.015,.02],
    steel_limit_unchanged=.02,material722_and_solver_exact=True,physical_capacity_certified=False,
    automatic_capacity_not_tested=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('724 ORIGINAL21 PLUS4 TARGETS TO EXISTING STEEL LIMIT .02; NO MATERIAL CHANGE')
