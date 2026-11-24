"""ДЛЯ ТЕСТОВ. Сохраняет точный максимум717 с незавершённым подтверждением и отдельным отрицательным контролем поздней ветви."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold717ShoulderAccuracy_20261124'
proof=json.loads((root/'ShoulderAccuracyProof.json').read_bytes())
out=repo/'docs/reports/evidence/Coupled6DShoulderAccuracy_2026-11-24';out.mkdir(parents=True,exist_ok=False)
paths=['Native/PostCrackTension.txt','Native/PostCrackTension.txt.ProcessClosure.json',
    'CompilePointClosure.json','PointTraceSourceComposition.json','Preparation.json',
    'ShoulderAccuracyProof.json','ShoulderAccuracyParsingFailure.json',
    'src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledNewtonSolver.cls','tests/modTestCoupledElasticState.bas']
paths += [str(p.relative_to(root)) for p in (root/'Native').iterdir() if p.is_file() and p.suffix in ('.csv','.bin')]
for name in paths:
    target=out/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((root/name).read_bytes())
for name in ('Stage-TestShoulderAccuracy.py','CandidateShoulderAccuracy.cls.txt','Check-TestShoulderAccuracy.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'ParsingCorrection.json').write_text(json.dumps(dict(
    superseded_file='ShoulderAccuracyParsingFailure.json',active_proof='ShoulderAccuracyProof.json',
    reason='Initial checker interpreted R diagnostics as N; corrected checker integrates sigma_zz and steel-minus-displaced-concrete stresses using point areas. Original0.5N threshold unchanged.',
    actual_remaining_reference_failure_count=1),ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    active_exec_session=66134,active_jobs=[dict(session=66134,root=str(base/'Cold718ControlledLoadBalance_20261124'),stage='native fixed-field lambda balance')],
    all_started_stages_terminal=False,working_numerical_sources_modified=False,
    previous_goal_turn_classification='progress: critical tightening reaches known uniaxial maximum to5.83e-7lambda,64fields independently verified; final precision/steel event still fail and one late homogeneous-reference miss retained; exact scalar balance718 now live',
    shoulder717=dict(proof='docs/reports/evidence/Coupled6DShoulderAccuracy_2026-11-24/ShoulderAccuracyProof.json',
        native=proof['native'],states=64,lower_lambda=proof['lower_lambda'],known_maximum_error=proof['lower_to_known_maximum'],
        reported_local_accuracy=proof['local_peak_reported_accuracy'],required_capacity_accuracy=.001,
        analytical_reference_misses=proof['analytical_force_misses'],steel_event_not_reached=True,physical_capacity_certified=False),
    next_action='Poll existing66134;718 reuses lambda-only balance before strain Newton while preserving material717,controller717,force tolerances,30iterations/96probes and all equilibrium checks. Independently verify entire curve,including late homogeneous-reference miss from717,and physical end/peak if reached. No production promotion before focused regressions and real coldall6/HardShear. Full original scope active.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[58105,18956]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',repo/'tmp/material-consistency/Continuation717_Active20261124.json',base/'Continuation717_Active20261124.json'):p.write_bytes(raw)
print('717 SAVED;64FIELDS,KNOWN MAXIMUM NEARLY EXACT,BUT FINAL CERTIFICATE AND LATE REFERENCE FAIL;718 LIVE')
