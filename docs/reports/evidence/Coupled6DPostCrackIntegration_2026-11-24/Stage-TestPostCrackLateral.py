"""ДЛЯ ТЕСТОВ. Сохраняет настоящую холодную историю11 и исследует локальный нуль бокового напряжения на отказавшем шаге."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold701PostCrackLateralRoot_20261124';parent=base/'Cold700PostCrackLocalMatrix_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'LateralCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
for label in ('Matrix','Refined'):(root/label/'Native').mkdir(parents=True,exist_ok=True)
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='        If Not solved Then GoTo CheckContinuation\n';assert code.count(anchor)==1
code=code.replace(anchor,'''        If Not solved Then
            If CDbl(strain) = 0.0008 Then TestPostCrackLateralProbe concrete(2), operators, accepted, seed, CDbl(strain), evidenceRoot, report, passed, failed
            GoTo CheckContinuation
        End If
''',1)
code+='\n'+(here/'TestPostCrackLateralProbe.bas.txt').read_text('utf-8-sig')
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
for name in (Path(__file__).name,'TestPostCrackLateralProbe.bas.txt'):values[name]=(here/name).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='RealCDPLateralRootAtFailedTarget',actual_import_this_pass=['modTestCoupledElasticState'],
    same_cold_history_and_full_failed_step=True,only_test_module_changed=True,
    local_response_read_only=True,not_a_section_capacity_result=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('LATERAL MATERIAL ROOT STAGED; ONLY TEST IMPORT; SAME COLD PREFIX AND FAILED TARGET')
