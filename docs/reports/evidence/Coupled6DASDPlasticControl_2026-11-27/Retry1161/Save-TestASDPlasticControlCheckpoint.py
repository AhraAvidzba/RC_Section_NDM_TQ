"""ДЛЯ ТЕСТОВ: фиксирует проверенные результаты и точный живой дескриптор повторного расчёта."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime, timezone
import json

repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ');here=Path(__file__).resolve().parent
resume_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';resume=json.loads(resume_path.read_text('utf-8-sig'))
for record in resume['terminal_ASD_trials_1156_1158_20261127']['records']:
    root=Path(record['root']);record['staged_input_book_sha256']=record.pop('book_sha256',record.get('staged_input_book_sha256'))
    record['compiled_book_path']=str(root/'docs/regression/Diagnostic.xlsm')
    record['compiled_book_sha256']=sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()

root=base/'Cold1159ASDPlasticSteelMy_20261127';summary=json.loads((root/'Case03/Native/PhysicalDirectionSummary.json').read_bytes())
assert summary['accepted_states']==111 and summary['all60_equations_and_six_resultants_verified']
resume['ASD_plastic_steel_control_1159']={
    'root':str(root),'compiled_book_sha256':sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper(),
    'independent_summary':summary,'projection_assertions_passed':432,'projection_assertions_failed':0,
    'maximum_projection_error':2.71050543121376e-20,'maximum_retained_lambda':6.5751697088477625,
    'maximum_epsilon0':0.003595949877634699,'maximum_steel_utilization':0.14383856173966098,
    'elapsed_seconds':313.64453125,'first50_parent_states_bitexact':True,
    'transient_diagnostic_maximum_6_662781_not_independently_verified':True,
    'physical_capacity_confirmed':False,'no_kernel_promotion':True,
    'archive':'docs/reports/evidence/Coupled6DASDPlasticControl_2026-11-27/Files.json',
}
old=base/'Cold1160ASDSteelLinearTrustMy_20261127'
resume['ASD_steel_QR_1160']={'root':str(old),'stage':'TEST_OBSERVER_IO_FAILURE','worker_closed':True,
    'not_a_numerical_or_physical_failure':True,'external_observer_snapshots':91,
    'missing_final_result':True,'numerics_repeated_in_1161_without_external_reader':True}
active=base/'Cold1161ASDSteelLinearRetry_20261127'
proof=json.loads((active/'EarlySourceProof.json').read_bytes())
assert proof['all262_other_parent_components_unchanged']
resume['ASD_steel_QR_retry_1161']={'root':str(active),'stage':'RUNNING','exec_session':5090,
    'compiled_book_sha256':proof['saved_book_sha256'],'only_test_observer_changed':True,
    'all_numerical_classes_exact_1160':True,'snapshot_copies_written_by_own_VBA_after_file_close':True,
    'no_external_snapshot_reader':True,'timeout_seconds':600,'full_original_physical_criteria_preserved':True}
resume['active_jobs']=[{'exec_session':5090,'kind':'single_cold_ASD_My_QR_retry','root':str(active),
    'authoritatively_running_when_launched':True,'do_not_restart_on_observation_timeout':True}]
resume['active_exec_session']=5090;resume['active_cold_exec_session']=5090;resume['active_cold_source_root']=str(active)
resume['current_load_stage']='ASD My: plastic-steel control reached epsilon0=0.00359595;111 states independently verified, maximum load still6.57517 and NumFail. QR trial1160 hit test observer I/O failure; same numerics now running1161 with internal snapshots.'
resume['next_action']='Poll exact live exec session5090 first; never restart from stale state or observation timeout. Confirm final process closure, independently replay final and highest retained accepted history with actual OpenSees material/all60 equations, then diagnose local-peak refinement cost and remaining NumFail. Archive1161 terminal evidence and keep full NMT/Q/NMTQ/CDP/performance/acceptance scope open.'
resume['current_goal_turn_classification']='progress: new steel-control projection proved,111 cold states independently replayed; physical My limit remains open; repaired diagnostic I/O and launched identical QR retry'
resume['updated_local']=datetime.now(timezone.utc).isoformat()
resume_path.write_text(json.dumps(resume,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8',newline='\n')
target=repo/'docs/reports/evidence/Coupled6DASDPlasticControl_2026-11-27/Retry1161'
target.mkdir(exist_ok=True)
for path in [here/'Stage-TestASDSteelLinearRetry.py',here/'Check-TestASDSteelLinearRetrySource.py',Path(__file__),
             active/'tests/CCoupledMaterialProgressTest.cls',active/'EarlySourceProof.json',active/'Preparation.json',
             active/'docs/regression/CodeUpdate.json',active/'Attempts/NoExcelCompile1/CompilePointFailure.json']:
    (target/path.name).write_bytes(path.read_bytes())
(target/'ActiveCheckpoint.json').write_text(json.dumps(resume['ASD_steel_QR_retry_1161'],sort_keys=True,indent=2)+'\n',encoding='utf-8',newline='\n')
print('CHECKPOINT SAVED;111 ACCEPTED STATES VERIFIED;EXACT LIVE SESSION5090 RETAINED;WHOLE GOAL OPEN')
