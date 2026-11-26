"""ДЛЯ ТЕСТОВ. Добавляет отсутствующий сторожевой harness после ошибки подготовки до запуска Excel-расчёта."""
from pathlib import Path
from hashlib import sha256
import json,shutil

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold1018SelectableMaterialCore_20261126';parent=base/'Cold1016PureBendingSteelLimitFixed_20261126'
assert not (root/'Native/RealColdReplay.txt').exists()
prep_path=root/'Preparation.json';before=prep_path.read_bytes();prep=json.loads(before)
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
(root/'PreparationBeforeHarnessFix.json').write_bytes(before)
added=[]
for source in (parent/'tools/build_workbook').rglob('*'):
    if not source.is_file():continue
    target=root/'tools/build_workbook'/source.relative_to(parent/'tools/build_workbook')
    if target.exists():continue
    target.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(source,target);added.append(target.relative_to(root).as_posix())
assert 'tools/build_workbook/Run-FocusedKaniWatchdog.ps1' in added
failure=dict(preparation_failed=True,reason='Missing private Run-FocusedKaniWatchdog.ps1 in fresh-build tool tree',
    Excel_calculation_not_started=True,compiled_book_unchanged=True,numerical_controls_unchanged=True,added_harness_files=added)
(root/'HarnessPreparationFailure.json').write_text(json.dumps(failure,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for name in added+['PreparationBeforeHarnessFix.json','HarnessPreparationFailure.json']:
    prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
prep_path.write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('MISSING HARNESS ADDED',len(added),'FILES;NO NUMERICAL RUN PRECEDED FIX;BOOK UNCHANGED')
