"""ДЛЯ ТЕСТОВ. Отделяет 11 холодных ASD-проб от длительного CDP-набора без изменения численного алгоритма и его бюджетов."""
from pathlib import Path
from hashlib import sha256
import json,shutil

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold1006PureBendingCoupled_20261126';root=base/'Cold1009PureBendingASDOnly_20261126';assert not root.exists()
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp','docs/regression'):(root/directory).mkdir(parents=True,exist_ok=True)
files={'docs/regression/Diagnostic.xlsm':(parent/'docs/regression/Diagnostic.xlsm').read_bytes(),
       'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
       'CompiledSourceReuse.json':(json.dumps(dict(parent=str(parent)),indent=2)+'\n').encode(),
       'Guard-IsolatedOffice.ps1':(parent/'Guard-IsolatedOffice.ps1').read_bytes(),
       'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
       'Run-TestPureBendingCoupled.ps1':(parent/'Run-TestPureBendingCoupled.ps1').read_bytes(),
       Path(__file__).name:Path(__file__).read_bytes()}
source=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
start=source.index('Public Function TestRunPureBendingCoupled(');end=source.index('Private Function TestPureBendingCoupledCaseTest(',start)
body=source[start:end]
for old,new in [('Array(reference(1), 2#, 11#','Array(reference(1), 1#, 11#'),('For model = 1 To 2','For model = 2 To 2'),('completed=22','completed=11')]:
    assert body.count(old)==1;body=body.replace(old,new,1)
files['tests/modTestCoupledElasticState.bas']=(source[:start]+body+source[end:]).replace('\n','\r\n').encode('utf-8-sig')
spec=json.loads((parent/'Input/PureBendingModels.json').read_bytes());spec.update(model_codes=[2],separate_ASD_run_same_algorithm_and_budgets=True)
files['Input/PureBendingModels.json']=(json.dumps(spec,sort_keys=True,indent=2)+'\n').encode()
for name,data in files.items():(root/name).write_bytes(data)
prep=json.loads((parent/'Preparation.json').read_bytes());prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope'].update(case='PureBendingASDOnlyCold',actual_import_this_pass=['modTestCoupledElasticState'],material_model_count=1,
    all256_numerical_components_exact1006=True,no_ready_coupled_equilibrium_input=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1009;11 COLD ASD CASES;SAME NUMERICAL BUDGETS;CDP RUN NOT REPLACED')
