"""ДЛЯ ТЕСТОВ. Исправляет только разрешение чтения H-касательной в диагностике746, сохраняя исходный runtime-отказ."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold748ActualHJacobianPermission_20261125';parent=base/'Cold746ActualHFullDifferences_20261125'
text=(parent/'Native/RealColdReplay.txt').read_text('utf-8-sig')
closure=json.loads((parent/'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert 'Для этой цели материальной координаты не запрошена собственная касательная.' in text
assert 'RUNTIME ERROR:' in text and 'WATCHDOG_COMPLETED: exit=1;' in text
assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode']==1
assert closure['Closure']['Identity']['Valid'] and closure['Closure']['FinalState']['State']=='Absent'
failed=dict(native_runtime_failure=True,own_process_closed=True,phase='before first H matrix read',
    cause='AtTarget was called with its allowJacobian=False default while requesting a Jacobian',
    no_derivative_or_capacity_result=True,working_source_unchanged=True)
(parent/'FailedRead746Proof.json').write_text(json.dumps(failed,indent=2,sort_keys=True)+'\n',encoding='utf-8')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'HJacobianCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='    Set equations = prepared.AtTarget(mReverseSolverTest, target)'
new='    Set equations = prepared.AtTarget(mReverseSolverTest, target, allowJacobian:=True)'
assert code.count(old)==1
code=code.replace(old,new,1)
assert code.replace(new,old,1)==(parent/name).read_text('utf-8-sig')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualHFullDerivativePermissionRetry',actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_kernel_exact746=True,only_allowJacobian_query_permission_changed=True,
    source_history_states=32,source_cold_history_verified744=True,source_failure746_preserved=True,
    same_readonly_secant_ratio=1.25,same247two_sided_samples=True,
    no_new_NDS_or_capacity_search=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('748 ONLY ALLOWJACOBIAN QUERY FLAG FIXED; FAILED746 PRESERVED; ALL NUMERICAL CLASSES EXACT')
