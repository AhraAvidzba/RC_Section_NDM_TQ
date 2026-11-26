"""ДЛЯ ТЕСТОВ. Проверяет короткие принятые ступени существующим штатным возвратом без ослабления половинного критерия и бюджета."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold957FourPartEventHistory_20261126';root=base/'Cold958BoundedIncomingSteps_20261126'
event=json.loads((base/'Cold952DetectedLoadingBoundary_20261126/DetectedLoadingBoundaryProof.json').read_bytes())
assert event['independently_equilibrated'] and event['effort'][0]==2
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'IncomingCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Input/DetectedEvent.bin').read_bytes());r.position=0
end_context,ends,_=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
context=[source['options'][3],source['options'][4],source['options'][2],8,event['lambda_value'],event['effort'][0],event['effort'][4],event['effort'][3],end_context[1]]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
entry=(here/'TestRunBoundedIncomingSteps.bas.txt').read_text('utf-8-sig').rstrip()
snapshot='''        record = 0
        For side = -1 To 1
            Select Case side
                Case -1: Set material = mReverseSidesTest(1): Set steel = mReverseSidesTest(2)
                Case 0: Set material = mReverseConcreteTest: Set steel = mReverseReinforcedTest
                Case 1: Set material = mReverseSidesTest(3): Set steel = mReverseSidesTest(4)
            End Select
            For point = 1 To mReverseOperatorsTest.PointCount
                record = record + 1: state = material.ReadNewtonYieldEventStateTest(point)
                material.AcceptedPointDamage point, tension, compression: material.CopyAcceptedStress point, stress
                bar = 0#: If point > mReverseOperatorsTest.ConcretePointCount Then bar = steel.AcceptedSteelStress(point)
                TestWriteReverseBinaryVector handle, Array(point, side): TestWriteReverseBinaryVector handle, state
                TestWriteReverseBinaryVector handle, Array(tension, compression, bar): TestWriteReverseBinaryVector handle, stress
RECORD_BASELINE_TEST
            Next point
        Next side'''
baseline='''                For component = 1 To 19
                    Select Case component
                        Case 1 To 10: value = state(LBound(state) + component - 1)
                        Case 11: value = tension
                        Case 12: value = compression
                        Case 13: value = bar
                        Case Else: value = stress(component - 13)
                    End Select
BASELINE_ACTION_TEST
                Next component'''
for kind,action in [('INITIAL','                    baseline(record, component) = value'),
 ('CYCLE',None),('FINAL',None),('RESTORED','                    If value <> baseline(record, component) Then Err.Raise 5, , "Не восстановлены исходная история, энергия или напряжение."')]:
    anchor="    ' "+kind+'_HISTORY_TEST';assert entry.count(anchor)==1
    inserted=snapshot.replace('RECORD_BASELINE_TEST',baseline.replace('BASELINE_ACTION_TEST',action) if action else '')
    entry=entry.replace(anchor,inserted,1)
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*20+'\n\nPublic Function TestRunFourPartEventHistory()';assert code.count(anchor)==1
updates[name]=code.replace(anchor,prefix*21+'\n\n'+entry+'\n\n\nPublic Function TestRunFourPartEventHistory()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunBoundedIncomingSteps'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunFourPartEventHistory';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunBoundedIncomingSteps',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'Input/BoundedIncoming.bin':vector(context)+vector(source['fields'][-1])+vector(event['field'])+vector(ends)+vector(event['selected_event']),
    'SourceDetectedEventProof.json':(base/'Cold952DetectedLoadingBoundary_20261126/DetectedLoadingBoundaryProof.json').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunBoundedIncomingSteps.bas.txt':entry.replace('\n','\r\n').encode('utf-8-sig')})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingGuardedReturnForShortIncomingSteps',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact957=True,maximum_accepted_cycles=8,
    original_force_tolerances_and_half_shape_limit_preserved=True,existing_lambda_uncertainty_and_steel_guards_used=True,
    guide_and_all_rejected_attempts_share30_per_accepted_step=True,source_first_guide_cost2_counted=True,
    earlier_primary_and_predictor_effort_not_included=True,actual_history_accepted_only_after_existing_return_succeeds=True,
    original25_restoration_required=True,no_production_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    previous_goal_turn_classification='progress:955-957 genuine sequential history effect distinguished from geometric curvature and algebraic precision, independent force/steel/event/restoration checks archived595 files; d70adfe pushed successfully. Physical capacity still open.',
    next_action='958 prepared: compile/audit actual257 sources, run existing guarded return at bounded shorter incoming targets, charge guide and rejected attempts to30 per accepted step, update actual history only on full existing acceptance, independently verify and restore original25. Goal active; no new pause requested.')
raw=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for path in (statepath,here/'Continuation666_Resume20261123.json'):path.write_bytes(raw)
print(root,'EXISTING FULL/HALF/UNCERTAINTY/STEEL GUARDS;SHARED30 PER ACCEPTED STEP;RESTORE25')
