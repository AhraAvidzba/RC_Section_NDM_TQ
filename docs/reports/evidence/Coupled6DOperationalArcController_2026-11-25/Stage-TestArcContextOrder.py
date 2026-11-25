"""ДЛЯ ТЕСТОВ. Проверяет кэш причины до диагностического воспроизведения дополнительного состояния; рабочий код786 неизменен."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold786ArcFailureContext_20261125';root=base/'Cold787ArcContextOrder_20261125'
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ArcContextOrderCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());name='tests/modTestCoupledElasticState.bas'
code=(root/name).read_text('utf-8-sig')
anchor='    TestRunOperationalArc\n    TestRunArcFailureContext\n';assert code.count(anchor)==1
code=code.replace(anchor,'    TestRunArcFailureContext\n    TestRunOperationalArc\n',1)
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ArcFailureContextBeforeDiagnosticReplay',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_numerical_working_methods_exact786=True,context_test_uses_original33_history=True,diagnostic_extra_state_not_reused_as_initial_history=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('787 TEST CONTEXT BEFORE DIAGNOSTIC EXTRA-STATE REPLAY; KERNEL EXACT786')
