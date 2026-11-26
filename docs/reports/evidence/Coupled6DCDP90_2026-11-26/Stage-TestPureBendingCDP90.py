"""ДЛЯ ТЕСТОВ. По прямому запросу пользователя увеличивает отдельный CDP-контроль до 5400 с при неизменных книге и численных параметрах."""
from pathlib import Path
from hashlib import sha256
import shutil,json

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold1014PureBendingCDPLong_20261126'
root=base/'Cold1019PureBendingCDP90_20261126';assert not root.exists()
proof=json.loads((parent/'PureBendingCurvatureProof.json').read_bytes());assert proof['native']['timeout1800'] and proof['native']['own_process_closed']
compiled=json.loads((parent/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((parent/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
root.mkdir()
for directory in ('src','tests','tools','Input','workbook','docs/regression'):shutil.copytree(parent/directory,root/directory)
for directory in ('Native','Temp'):(root/directory).mkdir()
shutil.copyfile(parent/'Guard-IsolatedOffice.ps1',root/'Guard-IsolatedOffice.ps1')
runner=(parent/'Run-TestPureBendingCurvature.ps1').read_text('utf-8-sig');assert runner.count('-TimeoutSeconds 1800')==1
(root/'Run-TestPureBendingCurvature.ps1').write_bytes(runner.replace('-TimeoutSeconds 1800','-TimeoutSeconds 5400').replace('\n','\r\n').encode('utf-8-sig'))
(root/'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent),all257_components_bitexact=True,
    compiled_book_bitexact=True,changed_only_watchdog_seconds=[1800,5400]),indent=2)+'\n',encoding='utf-8')
shutil.copyfile(Path(__file__),root/Path(__file__).name)
prep=json.loads((parent/'Preparation.json').read_bytes());prep['inputs']={p.relative_to(root).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in root.rglob('*')
    if p.is_file() and p.relative_to(root).as_posix()!='docs/regression/Diagnostic.xlsm'}
prep['directed_scope'].update(case='PureBendingCDP90Minutes',watchdog_seconds=5400,explicit_user_authorized_90_minutes=True,
    reference_planar_limit_kNm=4.629057996544561,reference_ASD_steel_limit_kNm=4.658444129476362,
    numerical_budgets_tolerances_materials_and_book_unchanged=True,full_acceptance_not_run=True,
    ASD_development_priority_retained=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PREPARED1019; CDP5400S; SAME EXACT BOOK/MATERIALS/CONTROLS;USER AUTHORIZED90MIN')
