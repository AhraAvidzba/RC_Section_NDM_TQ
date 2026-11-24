"""ДЛЯ ТЕСТОВ. Сохраняет713/715/716 с отрицательными исходами, точными матрицами и фактическими остатками."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DPeakAccuracy_2026-11-24';out.mkdir(parents=True,exist_ok=False)
for label,name,proof in (
 ('713','Cold713SmallLoadContinuation_20261124','SmallLoadContinuationProof.json'),
 ('715','Cold715PeakAdjointProbe_20261124','PeakAdjointProbeProof.json'),
 ('716','Cold716ActualAdjointResidual_20261124','ActualAdjointResidualProof.json')):
    root=base/name
    paths=['Native/PostCrackTension.txt','Native/PostCrackTension.txt.ProcessClosure.json',
        'CompilePointClosure.json','PointTraceSourceComposition.json','Preparation.json',proof,
        'src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledNewtonSolver.cls','tests/modTestCoupledElasticState.bas']
    paths += [str(p.relative_to(root)) for p in (root/'Native').iterdir() if p.is_file() and p.suffix in ('.csv','.bin')]
    for name in paths:
        target=out/label/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((root/name).read_bytes())
failed=base/'Cold714PeakAdjointProbe_20261124'
for name in ('PreparationFailure.json','Preparation.json'):
    target=out/'714'/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes((failed/name).read_bytes())
for name in ('Stage-TestSmallLoadContinuation.py','CandidateSmallLoadContinuation.cls.txt','Check-TestSmallLoadContinuation.py',
    'Stage-TestPeakAdjointProbe.py','TestAdjointProbe.bas.txt','Check-TestPeakAdjointProbe.py',
    'Stage-TestActualAdjointResidual.py','Check-TestActualAdjointResidual.py',Path(__file__).name):
    (out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    active_exec_session=80580,active_jobs=[dict(session=80580,root=str(base/'Cold717ShoulderAccuracy_20261124'),stage='compile goal-directed critical accuracy')],
    all_started_stages_terminal=False,working_numerical_sources_modified=False,
    previous_goal_turn_classification='progress: full-rank adjoints and exact residuals falsify rank-failure hypothesis; excessive uniform uncertainty blocks peak identification; targeted tighter solves717 staged with original bound retained',
    peak_accuracy716=dict(proof='docs/reports/evidence/Coupled6DPeakAccuracy_2026-11-24/716/ActualAdjointResidualProof.json',
        raw_history715_exact=True,adjoints_full_rank=8,maximum_uniform_uncertainty=.32099988994524725,
        corresponding_actual_componentwise_linear_estimate=3.090377029485483e-6,
        original_formula_preserved=True,nonlinear_estimator_not_adopted=True,physical_capacity_certified=False),
    next_action='Finish compiler80580; finalize717 and run critical-point tightening with original adjoint bound,96probes,30iterations and .001capacity accuracy. All retries counted. Poll existing native handle when started; verify all fields/critical tolerances/peak or failure before promotion. Remaining original full scope unchanged.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[28317,23111,68966,51591,78149]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',repo/'tmp/material-consistency/Continuation716_Active20261124.json',base/'Continuation716_Active20261124.json'):p.write_bytes(raw)
print('713/715/716 EVIDENCE SAVED; UNIFORM ERROR BOUND DIAGNOSED;717 COMPILE LIVE; GOAL ACTIVE')
