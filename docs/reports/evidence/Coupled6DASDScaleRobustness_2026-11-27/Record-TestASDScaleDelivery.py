"""ДЛЯ ТЕСТОВ. Привязывает завершённую контрольную точку к фактически восстановленному архиву и сохраняет следующий этап полной цели."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo = Path.cwd()
evidence = repo / 'docs/reports/evidence/Coupled6DASDScaleRobustness_2026-11-27'
summary = json.loads((evidence / 'Summary.json').read_bytes())
archive = json.loads((evidence / 'ArchiveProof.json').read_bytes())
restored = json.loads((evidence / 'RestorationProof.json').read_bytes())
assert not summary['latest_retry_four_scales_and_axial_gates_pending']
assert archive['logical_files'] == restored['actual_files'] == 3958
assert archive['exact_books'] == len(restored['exact_books']) == 26
assert restored['all_original_file_hashes_match'] and restored['all_restored_book_zip_crc_verified']
public_hash = sha256((repo / 'workbook/output/RC_Section_NDM.xlsm').read_bytes()).hexdigest().upper()
assert public_hash == '1BF23742F465C0551454166C6BE3DCB51E169219475EAF635AC55B92BB802D07'
delivery = dict(report='docs/reports/Coupled6DASDScaleRobustness_2026-11-27.md',
                evidence=evidence.relative_to(repo).as_posix(),
                restored_files=restored['actual_files'], restored_books=len(restored['exact_books']),
                compressed_archive_bytes=sum(p['bytes'] for p in archive['archives']),
                compiled_candidate='D:/CodexScratch/RC_Section_NDM_TQ/Cold1112ASDFreshCapacityRetryM4_20261127',
                same_book_four_scale_matrix='D:/CodexScratch/RC_Section_NDM_TQ/Cold1113ASDFreshRetryBendingScales_20261127',
                all_started_calculation_stages_terminal=True, working_core_promoted=False,
                public_workbook_sha256=public_hash, main_merged=False,
                full_acceptance_run=False, whole_goal_completed=False)
(evidence / 'DeliveryScope.json').write_text(json.dumps(delivery, sort_keys=True, indent=2) + '\n', encoding='utf-8')
path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
resume = json.loads(path.read_bytes())
resume['asd_scale_robustness_20261127']['delivery'] = delivery
resume.update(goal_status='active', active_jobs=[], all_started_stages_terminal=True,
              updated_local=datetime.now().astimezone().isoformat())
path.write_text(json.dumps(resume, ensure_ascii=False, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ASD SCALE DELIVERY RECORDED;3958 FILES/26 BOOKS RESTORED;NEXT CLEAN CORE/OUTPUT INTEGRATION')
