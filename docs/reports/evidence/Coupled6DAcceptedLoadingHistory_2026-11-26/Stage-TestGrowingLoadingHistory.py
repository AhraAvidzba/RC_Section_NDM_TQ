"""ДЛЯ ТЕСТОВ. Увеличивает шаг материальной координаты только после дешёвой принятой пары; первый отказ не считается разрушением."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold984AcceptedLoadingHistory_20261126';root=base/'Cold985GrowingLoadingHistory_20261126'
proof=json.loads((parent/'AcceptedLoadingHistoryProof.json').read_bytes())
assert proof['accepted_steps']==3 and proof['all_accepted_states_idempotent']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'GrowingHistoryCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
begin=code.index('Public Function TestRunAcceptedLoadingHistory()');end=code.index('\nEnd Function',begin)
body=code[begin:end]
changes={
    'не более трёх шагов':'не более 16 шагов с удвоением при расходе не более половины бюджета',
    'scalar As ICoupledScalarConstraint, equations As ICoupledEquationSystem, value As Double, gradient() As Double, child As String':
        'scalar As ICoupledScalarConstraint, equations As ICoupledEquationSystem, value As Double, gradient() As Double, child As String, predictionScale As Double',
    'For stage = 1 To 3':'For stage = 1 To 16',
    'Or stage = 3 Then Exit For':'Or stage = 16 Then Exit For',
    '        increment = 2# * (meta(3) - meta(10))':
        '        predictionScale = 1#: If total <= CLng(meta(8)) \\ 2 Then predictionScale = 2#\n        increment = predictionScale * 2# * (meta(3) - meta(10))',
    'seed(i) = full(i) + (full(i) - accepted(i))':'seed(i) = full(i) + predictionScale * (full(i) - accepted(i))',
    'meta(4) = factor + (factor - reference)':'meta(4) = factor + predictionScale * (factor - reference)',
}
for old,new in changes.items():assert body.count(old)==1,old;body=body.replace(old,new,1)
updates={name:(code[:begin]+body+code[end:]).replace('\n','\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='GrowingNativeLoadingHistory',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=257,
    all256_numerical_components_exact984=True,only_test_walk_max16_and_growth_after_cost_le15=True,
    same_full_half_H_forward_postcommit_and_shared30_per_attempt_gates=True,no_retry_or_reset_on_failure=True,
    no_future_equilibrium_field_input=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'GROW ONLY AFTER ACCEPTED CHEAP PAIR;STOP FIRST FAILURE;CORE/CRITERIA UNCHANGED')
