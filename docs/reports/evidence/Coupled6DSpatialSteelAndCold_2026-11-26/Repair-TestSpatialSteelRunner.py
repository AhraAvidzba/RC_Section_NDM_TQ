"""ДЛЯ ТЕСТОВ. Убирает унаследованный вызов H-диагностики из запуска отдельного теста стали, сохраняя неудачный отчёт."""
from pathlib import Path
from hashlib import sha256
import json
base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold881SpatialSteelLimits_20261126'
closure=json.loads((root/'Native/SpatialSteel.txt.ProcessClosure.json').read_bytes())
assert closure['ChildExitCode']==1 and closure['Closure']['FinalState']['State']=='Absent'
book=root/'docs/regression/Diagnostic.xlsm';digest=sha256(book.read_bytes()).hexdigest()
archive=root/'Native/FailedInheritedHRunner';archive.mkdir()
for file in (root/'Native').glob('SpatialSteel.txt*'):file.rename(archive/file.name)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
(archive/'Run-Audit02ModeTests.ps1').write_bytes((root/name).read_bytes())
data=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_bytes()
assert b'Unexpected verified-H setup macro.' not in data
(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
file=Path(__file__);(root/file.name).write_bytes(file.read_bytes());prep['inputs'][file.name]=sha256(file.read_bytes()).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert sha256(book.read_bytes()).hexdigest()==digest
print('ONLY INHERITED EXTRA DIAGNOSTIC RUNNER HOOK REMOVED;FAILED REPORT AND OWN CLOSURE PRESERVED;BOOK EXACT')
