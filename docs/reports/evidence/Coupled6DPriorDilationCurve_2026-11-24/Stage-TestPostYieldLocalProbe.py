"""ДЛЯ ТЕСТОВ. Разделяет локальный отказ и глобальный поиск на точной истории18 без повторения неудачного долгого шага."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold709PostYieldLocalProbe_20261124';parent=base/'Cold708ExactCohesionClosure_20261124'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'PostYieldCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
(root/'Matrix/Native').mkdir(parents=True,exist_ok=True)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
path='src/Materials/CCoupledConcreteResponse.cls';code=(root/path).read_text('utf-8-sig')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает штатный ответ одной точки со всеми резервами, не меняя принятую историю.
Friend Function ReadOperationalPointUpdateTest(ByRef strain() As Double, ByRef stress() As Double, ByRef substeps As Long) As Boolean
    Dim result As TConcreteHistory, virgin As Boolean
    ReadOperationalPointUpdateTest = UpdateWithSubsteps(strain, mHistory(1, mCommitted), result, stress, virgin, substeps)
End Function
'''
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
anchor='        If CDbl(strain) = 0.0008 Then TestCompareExactCohesionUpdate'
at=code.index(anchor)
code=code[:at]+'''        If CDbl(strain) > 0.002 Then
            TestPostCrackLateralProbe concrete(2), operators, accepted, seed, CDbl(strain), evidenceRoot, report, passed, failed
            GoTo CheckContinuation
        End If
'''+code[at:]
assert code.count('For Each divisions In Array(0, 2, 4, 8, 16, 32, 64)')==1
code=code.replace('For Each divisions In Array(0, 2, 4, 8, 16, 32, 64)','For Each divisions In Array(-1, 0, 1, 2, 4, 8, 16)',1)
code=code.replace('            If mLateralDivisionsTest <> 0 And index > 20 Then Exit For\n','',1)
anchor='    TestPostCrackLateralResponse = concrete.ReadPostCrackUpdateTest(strain, stress, mLateralDivisionsTest, mLateralSubstepsTest)'
assert code.count(anchor)==1
code=code.replace(anchor,'''    If mLateralDivisionsTest = -1 Then
        TestPostCrackLateralResponse = concrete.ReadOperationalPointUpdateTest(strain, stress, mLateralSubstepsTest)
    Else
'''+anchor+'''
    End If''',1)
values[path]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledConcreteResponse','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='PostYieldLocalResponseOnAccepted18',actual_import_this_pass=['CCoupledConcreteResponse','modTestCoupledElasticState'],
    same_first18_targets=True,long_failed_global_step_not_repeated=True,original_material_parameters_and_acceptance=True,
    sampled_operational_full_step_scalar_closure_and_fixed_substeps=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('709 LOCAL POST-YIELD DIAGNOSIS STAGED; ORIGINAL18 PREFIX; NO LONG FAILED GLOBAL STEP')
