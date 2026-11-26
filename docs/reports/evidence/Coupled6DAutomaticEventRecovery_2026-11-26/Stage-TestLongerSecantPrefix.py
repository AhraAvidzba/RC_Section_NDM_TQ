"""ДЛЯ ТЕСТОВ. Проверяет влияние лимита отдельного корня10 вместо4 при прежнем общем бюджете30."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold966ShortSecantPrefix_20261126';root=base/'Cold967LongerSecantPrefix_20261126'
proof=json.loads((parent/'ControlledSecantPrefixProof.json').read_bytes());assert proof['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'LongerSecantCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='        If budget > 4 Then budget = 4';assert code.count(old)==1
updates[name]=code.replace(old,'        If budget > 10 Then budget = 10',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='Root10InsteadOf4WithinSame30SecantPrefix',actual_import_this_pass=['CCoupledNewtonSolver'],
    all256_other_actual_components_exact966=True,no_new_working_methods=True,only_test_root_budget4_to10=True,
    initial_fraction=1/64,original_force_material_steel_and_half_criteria_preserved=True)
prep['directed_scope'].pop('all257_actual_components_and_saved_book_bitexact965',None)
prep['directed_scope']['no_new_numerical_code_or_compilation']=False
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'TEST ROOT10 IN SAME30;OTHER256 EXACT966;ORIGINAL FORCE/HALF/STEEL UNCHANGED')
