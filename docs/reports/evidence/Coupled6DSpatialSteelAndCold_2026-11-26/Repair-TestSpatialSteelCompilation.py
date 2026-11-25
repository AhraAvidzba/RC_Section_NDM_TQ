"""ДЛЯ ТЕСТОВ. Сохраняет неудачную сборку и исправляет доступ читающих методов между экземплярами VBA-класса."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold881SpatialSteelLimits_20261126')
prep=json.loads((root/'Preparation.json').read_bytes())
closure=json.loads((root/'CompilePointClosure.json').read_bytes());assert closure['State']=='Absent'
parent=json.loads((root/'ParentCodeUpdate.json').read_bytes())
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==parent['WorkbookSHA256']
archive=root/'CompileFailurePrivateInstanceAccess';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json'):(root/name).rename(archive/name)
name='src/Materials/CCoupledReinforcedResponse.cls';code=(root/name).read_text('utf-8-sig')
for method in ('LocalSteelLimit','LocalSteelBoundaryWithinCorrection'):
 old='Private Function '+method+'(';new='Friend Function '+method+'('
 assert code.count(old)==1;code=code.replace(old,new,1)
(archive/'CCoupledReinforcedResponse.BeforeFriend.cls').write_bytes((root/name).read_bytes())
data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
for file in (here/'Stage-TestSpatialSteelLimits.py',Path(__file__)):
 (root/file.name).write_bytes(file.read_bytes());prep['inputs'][file.name]=sha256(file.read_bytes()).hexdigest().upper()
prep['directed_scope']['private_instance_access_compile_failure_preserved']=True
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FAILED UNSAVED BOOK EXACT PARENT;OWN EXCEL CLOSED;TWO READERS CHANGED TO FRIEND;FAILURE PRESERVED')
