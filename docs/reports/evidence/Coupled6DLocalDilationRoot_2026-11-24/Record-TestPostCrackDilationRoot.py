"""ДЛЯ ТЕСТОВ. Сохраняет положительный локальный корень полного шага и все предшествующие отрицательные результаты."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold706PostCrackDilationDense_20261124'
proof=json.loads((root/'PostCrackDilationRootProof.json').read_bytes())
assert proof['local_original_full_step_root_found'] and proof['original_full_step_acceptance_preserved']
out=repo/'docs/reports/evidence/Coupled6DLocalDilationRoot_2026-11-24';out.mkdir(parents=True,exist_ok=False)
files={}
for label,name in (('704','Cold704PostCrackLocalGuide_20261124'),('705','Cold705PostCrackDilationRoot_20261124'),('706',root.name)):
    folder=base/name
    for path in ('Native/CurveMatrix.txt','Native/CurveMatrix.txt.ProcessClosure.json','CompilePointClosure.json','PointTraceSourceComposition.json',
       'Matrix/Native/LateralFixedIntegration.csv','Matrix/Native/PostCrackPrescribed.Accepted11.bin'):
        target=out/(label+'_'+Path(path).name);target.write_bytes((folder/path).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
failed=base/'Cold703PostCrackLocalGuide_20261124'
for name in ('CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json'):
    target=out/('703_'+name);target.write_bytes((failed/name).read_bytes());files[target.name]=sha256(target.read_bytes()).hexdigest().upper()
for name in ('Stage-TestPostCrackLocalGuide.py','TestPostCrackLocalGuide.cls.txt','Check-TestPostCrackLocalGuide.py',
   'Stage-TestPostCrackDilationRoot.py','TestPostCrackDilationRoot.cls.txt','Stage-TestPostCrackDilationDense.py','Check-TestPostCrackDilationRoot.py',Path(__file__).name):
    target=out/name;target.write_bytes((here/name).read_bytes());files[name]=sha256(target.read_bytes()).hexdigest().upper()
(out/'Proof.json').write_bytes((root/'PostCrackDilationRootProof.json').read_bytes())
(out/'Files.json').write_text(json.dumps(files,indent=2,sort_keys=True)+'\n',encoding='utf-8')
(out/'.gitattributes').write_text('* -text whitespace=blank-at-eol,blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8',newline='\n')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
   active_exec_session=38699,active_jobs=[dict(session=38699,root='D:/CodexScratch/RC_Section_NDM_TQ/Cold707SectionDilationClosure_20261124',stage='native full prescribed CurveMatrix')],
   all_started_stages_terminal=False,working_numerical_sources_modified=False,
   previous_goal_turn_classification='progress: full original material step root recovered by scalar dilation closure; independent uniaxial control and exact accepted11 verified; global707 now live',
   local_dilation706=dict(proof='docs/reports/evidence/Coupled6DLocalDilationRoot_2026-11-24/Proof.json',original_full_step_root_found=True,
      original_acceptance_preserved=True,independent_uniaxial_reference=proof['independent_uniaxial_reference'],
      full_section_curve_not_yet_passed=True),
   next_action='Poll existing session38699; do not restart on observation timeout.707 enables general per-point closure for near-apex tensile return after nonhydrostatic accepted history; original valid response preserved on recovery failure. Check full curve192/0, then independent all-state equilibrium before capacity/coldall6.')
state['terminal_sessions_this_stage']=sorted(set(state['terminal_sessions_this_stage']+[76344,57786,40717,84774,99841,11153,55381,60225]))
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json',repo/'tmp/material-consistency/Continuation706_Active20261124.json',base/'Continuation706_Active20261124.json'):target.write_bytes(raw)
print('LOCAL FULL-STEP ROOT SAVED; GLOBAL707 LIVE; WORKING MATERIAL UNCHANGED; GOAL ACTIVE')
