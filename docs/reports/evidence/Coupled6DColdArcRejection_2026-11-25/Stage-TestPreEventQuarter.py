"""ДЛЯ ТЕСТОВ. Проверяет подтверждённый промежуточный шаг перед разрешимым событием791 без изменения ядра."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold791ResolvedLoadingEvent_20261125';root=base/'Cold792PreEventQuarter_20261125'
proof=json.loads((parent/'RejectedOperationalArcProof.json').read_bytes());assert proof['event_resolved_at_inner_tolerance'] and not proof['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PreEventQuarterCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestRunOperationalArc\(\).*?^End Sub',code)
anchor='    mReverseSolverTest.DiscardStep\n    If solved Then TestExportSignedArcTrial';assert method.count(anchor)==1
updated=method.replace(anchor,'    mReverseSolverTest.DiscardStep\n    TestPreEventQuarter certificate\n    If solved Then TestExportSignedArcTrial',1)
code=code.replace(method,updated,1)+'\n'+(here/'TestPreEventQuarter.bas.txt').read_text('utf-8-sig')+'\n'
assert max(map(len,code.splitlines()))<1024
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
    'TestPreEventQuarter.bas.txt':(here/'TestPreEventQuarter.bas.txt').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ConfirmedHalfStepBeforeLoadingEvent',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_numerical_methods_exact791=True,full_and_half_event_fields_from_actual_operational_call=True,
    quarter_and_restore_share_original_event30_budget=True,same_inner_force_and_field_error_limits=True,
    no_history_acceptance=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('792 PRE-EVENT HALF/QUARTER CONFIRMATION; SAME791 KERNEL/30 EVENT BUDGET; NO RELAXED ACCEPTANCE')
