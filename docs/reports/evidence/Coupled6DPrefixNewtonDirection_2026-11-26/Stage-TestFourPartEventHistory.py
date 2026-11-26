"""ДЛЯ ТЕСТОВ. Уточняет последовательную историю956 четырьмя подшагами с теми же конечной поверхностью и общим бюджетом30."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold956RefinedSequentialHistory_20261126';root=base/'Cold957FourPartEventHistory_20261126'
proof=json.loads((parent/'RefinedSequentialHistoryProof.json').read_bytes());assert proof['second_event_solved']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'FourPartCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
entry=(parent/'TestRunRefinedSequentialHistory.bas.txt').read_text('utf-8-sig')
entry=entry.replace('TestRunRefinedSequentialHistory','TestRunFourPartEventHistory').replace('TOTAL_REFINED_SEQUENTIAL_HISTORY_TEST','TOTAL_FOUR_PART_EVENT_HISTORY_TEST')
changes={
 '    total = CLng(context(6))':'    Dim substeps As New Collection, pathReady As Boolean: pathReady = True\n    total = CLng(context(6))',
 'For stage = 0 To 5':'For stage = 0 To 7',
 '            Case 1\n':'            Case 1 To 3\n',
 '0.5 * (accepted(i) + face(i))':'accepted(i) + 0.25 * stage * (face(i) - accepted(i))',
 'target - 0.5 * spacing':'target - (1# - 0.25 * stage) * spacing',
 '0.5 * (context(1) + context(2))':'context(1) + 0.25 * stage * (context(2) - context(1))',
 '                halfSolved = mReverseSolverTest.TryDeformationStep':'                halfSolved = False\n                If pathReady Then halfSolved = mReverseSolverTest.TryDeformationStep',
 '                solved = halfSolved: factor = halfLambda':'                pathReady = pathReady And halfSolved\n                solved = halfSolved: factor = halfLambda',
 '            Case 2\n':'            Case 4\n',
 '            Case 3, 5\n':'            Case 5, 7\n',
 '            Case 4\n                Set replay':'            Case 6\n                Set replay',
 '        If stage = 1 And halfSolved Then mReverseSolverTest.AcceptStep':'        If stage >= 1 And stage <= 3 And halfSolved Then\n            mReverseSolverTest.AcceptStep: substeps.Add half\n        End If',
 'If stage = 2 And eventSolved Then':'If stage = 4 And eventSolved Then',
 '            replay.Add half: replay.Add eventField':'            For i = 1 To substeps.Count: checked = substeps(i): replay.Add checked: Next i\n            replay.Add eventField',
 'ReplayAcceptedFields(replay, 27)':'ReplayAcceptedFields(replay, 25 + substeps.Count + 1)',
 'If stage = 3 Or stage = 5 Then':'If stage = 5 Or stage = 7 Then',
}
for old,new in changes.items():
    expected=2 if old=='If stage = 3 Or stage = 5 Then' else 1
    assert entry.count(old)==expected,(old,entry.count(old));entry=entry.replace(old,new)
entry=entry.replace('двумя последовательными материальными подшагами','четырьмя последовательными материальными подшагами')
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*19+'\n\nPublic Function TestRunRefinedSequentialHistory()';assert code.count(anchor)==1
updates[name]=code.replace(anchor,prefix*20+'\n\n'+entry.rstrip()+'\n\n\nPublic Function TestRunRefinedSequentialHistory()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunFourPartEventHistory'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunRefinedSequentialHistory';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunFourPartEventHistory',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunFourPartEventHistory.bas.txt':entry.replace('\n','\r\n').encode('utf-8-sig')})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='FourPartSequentialMaterialHistoryRefinement',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact956=True,subdivisions=4,half_and_second_solve_precision_factor=.0001,
    full_source_event_unchanged=True,original_force_acceptance_and_total30_budget_preserved=True,
    geometric_guard_unchanged=True,temporary_history_restored_twice=True,no_production_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'FOUR SEQUENTIAL STEPS;ORIGINAL ACCEPTANCE AND SHARED30 UNCHANGED')
