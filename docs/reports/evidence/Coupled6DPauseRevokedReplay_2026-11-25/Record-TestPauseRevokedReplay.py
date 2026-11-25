"""ДЛЯ ТЕСТОВ. Фиксирует отмену паузы и завершённую независимую проверку фактического холодного маршрута 832."""
from pathlib import Path
from datetime import datetime
from hashlib import sha256
from runpy import run_path
import json

repo=Path.cwd();here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold835GatedColdReplay_20261125')
proof=json.loads((root/'OperationalColdReplayProof.json').read_bytes())
assert proof['accepted_states']==49 and proof['states_above_arc_inner_tolerance']==4
assert proof['source_name']=='Cold832GatedLoadingCold_20261125'
assert not proof['physical_capacity_certified'] and not proof['whole_goal_completed']
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
out=repo/'docs/reports/evidence/Coupled6DPauseRevokedReplay_2026-11-25'
assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
files=['OperationalColdReplayProof.json','Preparation.json','CompilePoint.log','CompilePointClosure.json',
    'CompileClosure.json','PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json',
    'docs/regression/CodeUpdate.json','tests/modTestCoupledElasticState.bas']
files += [p.relative_to(root).as_posix() for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
for name in files:
    p=root/name
    if p.is_file():
        target=out/'835'/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
for name in (Path(__file__).name,'Check-TestOperationalColdReplay.py'):
    (out/name).write_bytes((here/name).read_bytes())
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    pause_requested_explicitly=False,pause_request_revoked=True,pause_after_current_stage=False,
    pause_checkpoint_ready=False,goal_pause_tool_pending=False,
    pause_revoked_user_message='Паузу не включай пока я сам об этом чётко не попрошу',
    active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
    stage835='Completed after explicit pause revocation:49 actual832 states independently rebuilt;4 early states above arc inner bound;256 actual compiled components audited,255 parent components preserved',
    pause_revoked_replay_evidence=out.relative_to(repo).as_posix(),
    pause_goal_tool_was_not_called=True,
    next_action='Goal ACTIVE; user explicitly revoked pause before any goal pause tool call.835 actual49-state cold stress replay/source audit VERIFIED, all native jobs terminal. Next implement generic CDP elastic principal-zero event in private candidate:833 math/834 exact86 input,831 gated proactive controller when integrating cold. Require phi<0 at both endpoints and solved face; generic earliest point/station/principal selection, no hardcoded3, preserve ASD behavior,30/96 and full/half/force/gauge/steel criteria. Verify source86 failure fix, continue branch, then actual cold and physical ONQ012,step/mesh/external and original complete goal.829 local linear sensitivity not physical ultimate; no release/main promotion; full acceptance pending. Checkpoint92948eb records historical pause preparation, subsequently revoked. GitHub push of92948eb failedHTTP408; retry latest checkpoint and verify actual remote ref before claiming sent.')
state['terminal_sessions_this_stage']=sorted(set(state.get('terminal_sessions_this_stage',[])+[30403,99354,56270]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(raw)
print('GOAL ACTIVE; PAUSE REVOKED;835 VERIFIED;ARCHIVED',len(pins),'FILES;NO PHYSICAL LIMIT')
