"""ДЛЯ ТЕСТОВ. Сохраняет подтверждённые направления и отрицательные численные опыты, не закрывая открытый My."""
from pathlib import Path
from hashlib import sha256
import json
import shutil

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
evidence=repo/'docs/reports/evidence/Coupled6DASDPhysicalDirections_2026-11-27'
summary=json.loads((evidence/'Summary.json').read_bytes())
negative=base/'Cold1154ASDLinearTrustMy_20261127'
compiled=json.loads((negative/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert sha256((negative/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
closure=json.loads((negative/'Case03/Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert closure['ChildHasExited'] and closure['ChildExitCode']==0 and closure['Passed']
history=(negative/'Case03/Native/MainCapacityHistory.bin').read_bytes()
assert history==(base/'Cold1153ASDPhysicalDirections_20261127/Case03/Native/MainCapacityHistory.bin').read_bytes()
copied={}
for relative in ['src/Solver/CCoupledNewtonSolver.cls','Preparation.json','EarlySourceProof.json',
                 'CompilePoint.log','CompilePointClosure.json','docs/regression/CodeUpdate.json']:
    source=negative/relative;target=evidence/'Trial1154'/relative;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    copied[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
for source in (negative/'Case03/Native').iterdir():
    if not source.is_file() or source.suffix=='.csv':continue
    target=evidence/'Trial1154/Case03/Native'/source.name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    copied[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
extra_names=['IndependentMyForceCurve.json','IndependentMyForceCurveVerifiedSeed.json','IndependentMySteelEndpoint.json']
if (base/'Cold1153ASDPhysicalDirections_20261127/Case03/Native/IndependentMySteelEndpointRefined.json').exists():
    extra_names.append('IndependentMySteelEndpointRefined.json')
for name in extra_names:
    source=base/'Cold1153ASDPhysicalDirections_20261127/Case03/Native'/name;target=evidence/'AdditionalMy'/name
    target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    copied[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
for name in ['Stage-TestASDLinearTrustMy.py','Check-TestASDLinearTrustSource.py','Stage-TestASDLinearTrustDifferenceMy.py',
             'Probe-TestASDMyForceCurve.py','Probe-TestASDMySteelEndpoint.py','Record-TestASDPhysicalCheckpoint.py']:
    source=here/name;target=evidence/'AdditionalHelpers'/name;target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target)
    copied[target.relative_to(evidence).as_posix()]=sha256(source.read_bytes()).hexdigest().upper()
negative_summary=dict(root=negative.name,status='NumFail',physical_ultimate_confirmed=False,native_call_seconds=395.64453125,
                      maximum_lambda=6.575169710628729,final_saved_history_bit_exact_with_verified_original=True,
                      saved_history_sha256=sha256(history).hexdigest().upper(),not_adopted_in_production=True,
                      exact_private_book_retained_on_D=str(negative/'docs/regression/Diagnostic.xlsm'),
                      exact_private_book_sha256=compiled['WorkbookSHA256'],book_not_in_compact_trial_evidence=True,
                      no_solver_improvement_established=True,own_compiler_and_test_Excel_closed=True)
(evidence/'Trial1154/Summary.json').write_text(json.dumps(negative_summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
(evidence/'AdditionalFiles.json').write_text(json.dumps(copied,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
state.update(asd_physical_directions_20261127=summary,asd_linear_trust_negative_20261127=negative_summary,
             current_load_stage='N compression, negative Mx and pure T independently confirmed; My open; numerical derivative candidate active; no formatting work',
             active_jobs=[dict(root='Cold1155ASDLinearTrustDifferenceMy_20261127',case='Case03',exec_session=97093,
                               kind='Cold My; existing QR linear reserve with numerical difference 0.001; 600 s watchdog'),
                          dict(root='Cold1153ASDPhysicalDirections_20261127',case='Case03',exec_session=90911,
                               kind='Independent steel-endpoint refinement; no production change')],
             active_cold_exec_session=97093,all_started_stages_terminal=False,goal_status='active',entire_goal_completed=False,
             last_goal_turn_classification='progress',formatting_deferred_until_stable_calculation_methods_by_user_20261127=True)
state_path.write_text(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
plan_path=repo/'docs/Coupled6DLoadSequencePlan.md';plan=plan_path.read_text('utf-8-sig')
plan=plan.replace('Другие направления и скорость остаются открытыми','Другие направления и скорость остаются открытыми')
plan=plan.replace('Явный бетонный критерий и выбор ASD на Config подключены; другие направления и скорость остаются открытыми',
                  'Явный бетонный критерий и выбор ASD на Config подключены; отрицательный Mx подтверждён; My остаётся NumFail и дорабатывается, скорость открыта')
plan=plan.replace('Есть воспроизводимое арматурное событие около 3,662 кН·м и Fail при T=4 кН·м; общий бетонный предел не завершён',
                  'Арматурное событие около 3,662 кН·м повторено с явным бетонным критерием и reader Config; предел хомута подтверждён независимым верхним равновесием. Другие сечения, масштабы и скорость остаются открытыми')
paragraph='\n\n[Повтор N, отрицательного Mx, My и T](reports/Coupled6DASDPhysicalDirections_2026-11-27.md): подтверждены сжатие 735,94484 кН, продольная сталь при Mx≈−4,657714 кН·м и хомут при T≈3,661950 кН·м. Все 188 сохранённых состояний пересчитаны независимо. My остаётся NumFail до 6,575170 кН·м; внешним TRF найдены допустимые состояния выше этой точки. Расширение существующего линейного QR-резерва само по себе не устранило остановку и не принято; проверяется масштаб численной производной, без изменения материалов и силовых допусков.\n'
if '[Повтор N, отрицательного Mx, My и T]' not in plan:plan+=paragraph
plan_path.write_text(plan,encoding='utf-8')
print('PHYSICAL CHECKPOINT RECORDED; NEGATIVE QR EXPERIMENT NOT ADOPTED; MY STILL OPEN; GOAL ACTIVE')
