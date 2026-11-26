"""ДЛЯ ТЕСТОВ. Записывает действительные полную и половинную пробы947 без изменения численного алгоритма или приёмки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold947WorkingSecantReturn_20261126';root=base/'Cold949WorkingGuidePair_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'WorkingGuidePairCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mMaterialGuideReturnActive As Boolean';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"\nPrivate mGuideReturnPairsTest As Collection, mGuideReturnCaptureTest As Boolean ' ДЛЯ ТЕСТОВ: снимки половинной проверки, не используемые рабочим алгоритмом.",1)
anchor='    If tensorError > YIELD_EVENT_HALF_ERROR Or lambdaError > lambdaBound Then GoTo CountHalf';assert code.count(anchor)==1
record='''    If mGuideReturnCaptureTest Then
        mGuideReturnPairsTest.Add Array(accepted, full, half, projection, Array(fullLambda, halfLambda, referenceLambda, _
            target, halfTarget, tolerance, tensorError, lambdaError, lambdaBound, length, curveScale, jointError, stepError))
    End If
'''
code=code.replace(anchor,record+anchor,1)
anchor='    TestMaterialGuideOriginalReturn = TryReturnOriginalMaterialControl(guide, guideLambda, projection, target, True, load, _'
assert code.count(anchor)==1
code=code.replace(anchor,'    If mGuideReturnPairsTest Is Nothing Then Set mGuideReturnPairsTest = New Collection\n    mGuideReturnCaptureTest = True\n'+anchor,1)
anchor='    If TestMaterialGuideOriginalReturn Then resultLambda = mTrialLambda Else DiscardStep';assert code.count(anchor)==1
code=code.replace(anchor,'    mGuideReturnCaptureTest = False\n'+anchor,1)
code=code.rstrip()+'''

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает только диагностические копии фактических пар; рабочая история и буферы материала от них не зависят.
Friend Function TestReadGuideReturnPairs() As Collection
    Set TestReadGuideReturnPairs = mGuideReturnPairsTest
End Function
'''
updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*13+'\n\nPublic Function TestRunWorkingGuideReturn()';assert code.count(anchor)==1
entry=(here/'TestRunWorkingGuidePair.bas.txt').read_text('utf-8-sig').rstrip()
updates[name]=code.replace(anchor,prefix*14+'\n\n'+entry+'\n\n\nPublic Function TestRunWorkingGuideReturn()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver','modTestCoupledElasticState' -EntryPointNames 'TestRunWorkingGuidePair'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunWorkingGuideReturn';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunWorkingGuidePair',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunWorkingGuidePair.bas.txt':(here/'TestRunWorkingGuidePair.bas.txt').read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ReadActualWorkingFullHalfPair',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    actual_component_count=257,numerical_body947_unchanged_except_passive_capture=True,
    previous_return_binary_must_match947_exactly=True,all_guards_and_total30_preserved=True,
    guide17_steps_charged=True,no_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'PASSIVE ACTUAL PAIR CAPTURE;MUST MATCH947 OUTCOME EXACTLY')
