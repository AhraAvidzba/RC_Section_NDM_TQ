"""ДЛЯ ТЕСТОВ. Готовит отдельную книгу для одного полного и двух последовательных материальных шагов с точным восстановлением25."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json, numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold954RefinedEventApproach_20261126';root=base/'Cold955SequentialEventHistory_20261126'
proof=json.loads((base/'Cold952DetectedLoadingBoundary_20261126/DetectedLoadingBoundaryProof.json').read_bytes())
assert proof['independently_equilibrated'] and proof['effort'][0]==2
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'SequentialEventCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Input/DetectedEvent.bin').read_bytes());r.position=0
old_context,ends,_=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
context=[source['lambdas'][-1],proof['lambda_value'],source['options'][3],source['options'][4],source['options'][2],proof['effort'][0],old_context[1]]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
entry=(here/'TestRunSequentialEventHistory.bas.txt').read_text('utf-8-sig').rstrip()
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*17+'\n\nPublic Function TestRunRefinedEventApproach()';assert code.count(anchor)==1
updates[name]=code.replace(anchor,prefix*18+'\n\n'+entry+'\n\n\nPublic Function TestRunRefinedEventApproach()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunSequentialEventHistory'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunRefinedEventApproach';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunSequentialEventHistory',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'Input/SequentialEvent.bin':vector(context)+vector(np.array(proof['field']))+vector(ends)+vector(source['fields'][-1]),
    'SourceDetectedEventProof.json':(base/'Cold952DetectedLoadingBoundary_20261126/DetectedLoadingBoundaryProof.json').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunSequentialEventHistory.bas.txt':entry.replace('\n','\r\n').encode('utf-8-sig')})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='FullVersusTwoSequentialMaterialHistorySteps',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact954=True,half_solve_precision_factor=.0001,
    original_force_acceptance_preserved=True,total_shared_budget=30,known_full_event_source_cost=2,
    earlier_predictor_and_failed_primary_cost_not_included=True,temporary_half_acceptance_diagnostic_only=True,
    geometric_half_guard_not_changed_or_claimed_passed=True,original25_restored_twice_with_bitexact_native_guards=True,
    no_production_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_after_current_stage=False,pause_checkpoint_ready=False,pause_requested_after_logical_stage=False,
    goal_pause_tool_pending=False,goal_tool_pause_after_git_checkpoint_requested=False,
    no_new_experiments_added_after_pause_request=False,
    previous_goal_turn_classification='progress: completed951-954 evidence verified, own processes absent, pause checkpointa53a824 published; user explicitly resumed goal and revoked pause.',
    next_action='955 prepared: compile and audit all257 actual sources, run one-full versus two sequential material-history diagnostic, independently verify forces, steel, material event and exact original25 restoration. Goal active; no new pause without explicit request.')
raw=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for path in (statepath,here/'Continuation666_Resume20261123.json'):path.write_bytes(raw)
print(root,'SEQUENTIAL MATERIAL HISTORY DIAGNOSTIC;RESTORE25;RELEASE UNCHANGED')
