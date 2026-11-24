"""ДЛЯ ТЕСТОВ. Сохраняет незавершённый автоматический предел712 и текущую направленную проверку713."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold712PriorDilationCapacity_20261124'
proof=json.loads((root/'PriorDilationCapacityProof.json').read_bytes())
out=repo/'docs/reports/evidence/Coupled6DPriorDilationCapacity_2026-11-24';out.mkdir(parents=True,exist_ok=False)
paths=['Native/PostCrackTension.txt','Native/PostCrackTension.txt.ProcessClosure.json','CompilePointClosure.json',
    'PointTraceSourceComposition.json','Preparation.json','PriorDilationCapacityProof.json',
    'src/Solver/CCoupledDeformationContinuation.cls','tests/modTestCoupledElasticState.bas']
paths += [str(p.relative_to(root)) for p in (root/'Native').iterdir() if p.is_file() and p.suffix in ('.csv','.bin')]
for name in paths:
    target=out/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((root/name).read_bytes())
for name in ('Stage-TestPriorDilationCapacity.py','Check-TestPriorDilationCapacity.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,active_exec_session=28317,active_jobs=[dict(session=28317,root=str(base/'Cold713SmallLoadContinuation_20261124'),stage='native automatic small-load continuation')],
    all_started_stages_terminal=False,working_numerical_sources_modified=False,
    automatic712=dict(proof='docs/reports/evidence/Coupled6DPriorDilationCapacity_2026-11-24/PriorDilationCapacityProof.json',
        native=proof['native'],accepted_states=57,lower_lambda=proof['confirmed_lower_lambda'],probes=96,physical_capacity_certified=False),
    next_action='Poll existing session28317 only.713 accepts converged forward spatial steps with small lambda change and detects cumulative decline from accepted maximum; coarse local peak before verified design remains subject to original final precision. Preserve observer/checkpoints and all negative outcomes. Independently verify all fields and final event/peak if found, then focused regressions before production promotion; original full scope remains open.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[45282,77140,70831]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',repo/'tmp/material-consistency/Continuation713_Active20261124.json',base/'Continuation713_Active20261124.json'):p.write_bytes(raw)
print('712 NEGATIVE CAPACITY SAVED;57 FIELDS VERIFIED;713 LIVE; GOAL ACTIVE')
