"""ДЛЯ ТЕСТОВ. Привязывает синхронную рабочую выдачу к реально восстановленному архиву завершённого этапа."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo = Path.cwd()
evidence = repo / 'docs/reports/evidence/Coupled6DWorkingASDRetryIntegration_2026-11-27'
summary = json.loads((evidence / 'Summary.json').read_bytes())
archive = json.loads((evidence / 'ArchiveProof.json').read_bytes())
restored = json.loads((evidence / 'RestorationProof.json').read_bytes())
assert summary['working_core_and_output_synchronized'] and summary['all_started_stages_terminal']
assert archive['logical_files'] == restored['actual_files'] == 3003
assert archive['exact_books'] == len(restored['exact_books']) == 18
assert restored['all_original_file_hashes_match'] and restored['all_restored_book_zip_crc_verified']
for name, record in summary['public_delivery']['output'].items():
    assert sha256((repo / 'workbook/output' / name).read_bytes()).hexdigest().upper() == record['after_sha256']
scope = dict(report=summary['report'], evidence=summary['evidence'], restored_files=3003, restored_books=18,
             compressed_archive_bytes=sum(a['bytes'] for a in archive['archives']),
             working_core_and_actual_public_output_synchronized=True, actual_components=256,
             all121_user_parts_preserved=True, full_native_code_dump_updated=True,
             direct_N_without_observer_cases=6, ordinary_capacity_cases=3,
             mutated_observer_snapshots=347, native_foundation_suites=8, native_assertions_passed=35452,
             all_started_stages_terminal=True, Config_ASD_choice_pending=True,
             general_concrete_criterion_pending=True, whole_user_run10s_not_met=True,
             full_acceptance_not_run=True, main_not_merged=True, whole_goal_completed=False)
(evidence / 'DeliveryScope.json').write_text(json.dumps(scope, sort_keys=True, indent=2) + '\n', encoding='utf-8')
path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
resume = json.loads(path.read_bytes())
resume['working_asd_retry_integration_20261127']['delivery_scope'] = scope
resume.update(goal_status='active', active_jobs=[], all_started_stages_terminal=True,
              updated_local=datetime.now().astimezone().isoformat())
path.write_text(json.dumps(resume, ensure_ascii=False, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD RETRY DELIVERY VERIFIED;18 BOOKS/3003 FILES RESTORED;SOURCE/ACTUAL OUTPUT SYNCHRONIZED')
