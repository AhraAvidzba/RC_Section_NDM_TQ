"""ДЛЯ ТЕСТОВ. Сохраняет воспроизводимый скачок локального отклика и контроль корня при фиксированном дроблении."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json,subprocess

repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ');here=Path(__file__).resolve().parent
second=base/'Cold702PostCrackIntegration_20261124'
proof=json.loads((second/'PostCrackIntegrationProof.json').read_bytes())
assert proof['native701']['own_process_closed'] and proof['native702']['own_process_closed']
assert proof['accepted11_native_fields_and_calibration_bits_equal']
out=repo/'docs/reports/evidence/Coupled6DPostCrackIntegration_2026-11-24';out.mkdir(parents=True,exist_ok=False)
files={}
for label,root in (('701',base/'Cold701PostCrackLateralRoot_20261124'),('702',second)):
    for name in ('Native/CurveMatrix.txt','Native/CurveMatrix.txt.ProcessClosure.json','CompilePointClosure.json','PointTraceSourceComposition.json',
       'Matrix/Native/PostCrackPrescribed.Accepted11.bin'):
        target=out/(label+'_'+Path(name).name);target.write_bytes((root/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
for root,name in ((base/'Cold701PostCrackLateralRoot_20261124','LateralProbe.csv'),(second,'LateralFixedIntegration.csv')):
    target=out/name;target.write_bytes((root/'Matrix/Native'/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
for name in ('PostCrackPrescribed.Accepted.csv','PostCrackPrescribed.Accepted.Operators.csv'):
    target=out/name;target.write_bytes((second/'Matrix/Native'/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
for name in ('Stage-TestPostCrackLateral.py','Stage-TestPostCrackFixedIntegration.py','TestPostCrackLateralProbe.bas.txt',
   'TestPostCrackFixedIntegration.bas.txt','TestPostCrackFixedIntegration.cls.txt','Check-TestPostCrackIntegration.py',Path(__file__).name):
    target=out/name;target.write_bytes((here/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
(out/'Proof.json').write_bytes((second/'PostCrackIntegrationProof.json').read_bytes())
(out/'Files.json').write_text(json.dumps(files,indent=2,sort_keys=True)+'\n',encoding='utf-8')
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8',newline='\n')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
   active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
   working_numerical_sources_modified=False,working_commit=subprocess.check_output(['git','rev-parse','HEAD'],cwd=repo,text=True).strip(),
   previous_goal_turn_classification='progress: exact working kernel213/0+48SVD; material jump1.044934788MPa reproduced at same accepted11; local roots exist for2..64substeps',
   material_integration702=dict(proof='docs/reports/evidence/Coupled6DPostCrackIntegration_2026-11-24/Proof.json',
      material_methods_unchanged=True,automatic_substeps_both_sides_of_jump=1,
      observed_stress_jump_MPa=proof['results']['0']['nearest_sign_bracket']['stress_gap_MPa'],
      fixed_divisions_with_local_root=[int(k) for k,v in proof['results'].items() if v['roots']],
      original11_history_exact=True,local_root_not_section_capacity=True),
   next_action='Investigate local UpdatePoint branch selection/cohesion/dilation at failed epsilon_zz0.0008. One-step map jumps while2..64substeps recover same uniaxial root. Test a consistent local continuation/predictor or error-controlled integration; do not blindly force2substeps or relax global tolerances. Then full curve and actual cold capacity/all6 original scope.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[97029,37183,21859,94068]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',
   repo/'tmp/material-consistency/Continuation702_Active20261124.json',base/'Continuation702_Active20261124.json'):target.write_bytes(raw)
print('MATERIAL JUMP EVIDENCE SAVED; ALL PROCESSES TERMINAL; GOAL ACTIVE; NEXT LOCAL INTEGRATOR')
