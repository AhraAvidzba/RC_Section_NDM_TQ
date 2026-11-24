"""ДЛЯ ТЕСТОВ. Сохраняет опыты707/708 и точку продолжения без объявления численного отказа разрушением."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DExactCohesionCurve_2026-11-24';out.mkdir(parents=True,exist_ok=False)
for label,name in (('707','Cold707SectionDilationClosure_20261124'),('708','Cold708ExactCohesionClosure_20261124')):
    root=base/name
    paths=['Native/CurveMatrix.txt','Native/CurveMatrix.txt.ProcessClosure.json','Native/CurveMatrix.txt.Ownership.json',
       'CompilePointClosure.json','PointTraceSourceComposition.json','Preparation.json',
       'src/Materials/CCoupledConcreteResponse.cls','tests/modTestCoupledElasticState.bas']
    paths+=['SectionDilationTimeoutProof.json'] if label=='707' else ['ExactCohesionCurveProof.json']
    paths+=[str(p.relative_to(root)) for p in (root/'Matrix/Native').iterdir() if p.is_file()]
    for name in paths:
        source=root/name;target=out/label/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(source.read_bytes())
for name in ('Stage-TestSectionDilationClosure.py','CandidateDilationClosureGate.cls.txt','Stage-TestExactCohesionClosure.py',
 'TestDilationWork.cls.txt','TestDilationProgress.bas.txt','Check-TestExactCohesionCurve.py','Check-TestSectionDilationTimeout.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
proof=json.loads((base/'Cold708ExactCohesionClosure_20261124/ExactCohesionCurveProof.json').read_bytes())
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_revoked_user_message='На паузу не ставь пока я сам об этом не попрошу снова',
    active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
    working_numerical_sources_modified=False,physical_capacity_certified=False,
    previous_goal_turn_classification='progress: local full-step root and18 independently assembled section states through known uniaxial peak; next prescribed step still NumFail',
    section708=dict(proof='docs/reports/evidence/Coupled6DExactCohesionCurve_2026-11-24/708/ExactCohesionCurveProof.json',
        accepted_states=18,native=proof['native'],known_peak_lambda=proof['native_lambda_at_peak'],
        next_strain_numfail=.0025,automatic_capacity_confirmed=False),
    next_action='Diagnose post-yield prescribed step .002→.0025 on exact accepted18 material history; distinguish predictor/step size from local constitutive discontinuity. No tolerance loosening or physical failure inference. Then automatic peak trace and real coldall6/HardShear; original full scope remains open.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[38699,25261,55018]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',repo/'tmp/material-consistency/Continuation708_Active20261124.json',base/'Continuation708_Active20261124.json'):p.write_bytes(raw)
print('707/708 EVIDENCE SAVED; GOAL ACTIVE; PAUSE REVOKED; AUTOMATIC CAPACITY OPEN')
