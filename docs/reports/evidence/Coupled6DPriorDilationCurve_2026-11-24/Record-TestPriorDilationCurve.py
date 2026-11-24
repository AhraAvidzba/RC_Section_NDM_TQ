"""ДЛЯ ТЕСТОВ. Архивирует отрицательную локальную пробу709, прогноз710 и подтверждённую полную кривую711."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DPriorDilationCurve_2026-11-24';out.mkdir(parents=True,exist_ok=False)
for label,name,proof_name in (
 ('709','Cold709PostYieldLocalProbe_20261124','PostYieldLocalProbeProof.json'),
 ('710','Cold710PriorDilationSeed_20261124','PriorDilationSeedProof.json'),
 ('711','Cold711PriorDilationCurve_20261124','PriorDilationCurveProof.json')):
    root=base/name
    paths=['Native/CurveMatrix.txt','Native/CurveMatrix.txt.ProcessClosure.json','CompilePointClosure.json',
        'PointTraceSourceComposition.json','Preparation.json','src/Materials/CCoupledConcreteResponse.cls',
        'tests/modTestCoupledElasticState.bas',proof_name]
    paths += [str(p.relative_to(root)) for p in (root/'Matrix/Native').iterdir() if p.is_file()]
    for name in paths:
        target=out/label/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((root/name).read_bytes())
for name in ('Stage-TestPostYieldLocalProbe.py','Check-TestPostYieldLocalProbe.py','Stage-TestPriorDilationSeed.py',
    'Check-TestPriorDilationSeed.py','Stage-TestPriorDilationCurve.py','Check-TestPriorDilationCurve.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
proof=json.loads((base/'Cold711PriorDilationCurve_20261124/PriorDilationCurveProof.json').read_bytes());assert proof['full_original21_prescribed_curve_passed']
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    no_new_experiments_added_after_pause_request=False,subsequent_experiments_authorized_by_pause_revocation=True,
    active_exec_session=70831,active_jobs=[dict(session=70831,root=str(base/'Cold712PriorDilationCapacity_20261124'),stage='compile automatic capacity candidate')],
    all_started_stages_terminal=False,working_numerical_sources_modified=False,
    previous_goal_turn_classification='progress: full original21 realCDP+steel prescribed curve194/0 and all spatial fields independently verified; automatic capacity712 being compiled',
    section711=dict(proof='docs/reports/evidence/Coupled6DPriorDilationCurve_2026-11-24/711/PriorDilationCurveProof.json',
        accepted_states=21,native=proof['native'],full_original_curve_passed=True,automatic_capacity_confirmed=False),
    next_action='Finish existing712 compiler, verify actual sources, run automatic capacity from two elastic initial states with exact697 adaptive controller and711 material, default assembled tangent. Preserve96 probes/30 iterations/.001 accuracy; do not infer failure from NumFail. Then real coldall6/HardShear and full original scope.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[80666,50535,56016,24444,48427,72043]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',repo/'tmp/material-consistency/Continuation711_Active20261124.json',base/'Continuation711_Active20261124.json'):p.write_bytes(raw)
print('709-711 RAW EVIDENCE SAVED; ORIGINAL21 CURVE194/0; GOAL ACTIVE; AUTOMATIC CAPACITY PENDING')
