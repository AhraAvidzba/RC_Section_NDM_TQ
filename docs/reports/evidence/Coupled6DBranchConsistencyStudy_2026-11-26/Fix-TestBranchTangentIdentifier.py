"""ДЛЯ ТЕСТОВ. Исправляет конфликт имени Scale с синтаксисом VBA после неуспешной компиляции; сохраняет все её доказательства."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold873BranchConsistentArcTangent20_20261126')
parent=json.loads((root/'ParentCodeUpdate.json').read_bytes());assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==parent['WorkbookSHA256']
closure=json.loads((root/'CompilePointClosure.json').read_bytes());assert closure['State']=='Absent'
archive=root/'FailedCompileReservedIdentifier';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json'):
    if (root/name).exists():(root/name).rename(archive/name)
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='orientation As Double, scale As Double, nextProbe() As Double';assert code.count(old)==1;code=code.replace(old,'orientation As Double, tangentFactor As Double, nextProbe() As Double',1)
old='        scale = Sgn(orientation) * spacing / length: nextProbe = accepted';assert code.count(old)==1;code=code.replace(old,'        tangentFactor = Sgn(orientation) * spacing / length: nextProbe = accepted',1)
old='accepted(i) + scale * tangent(i) / 1024#';assert code.count(old)==1;code=code.replace(old,'accepted(i) + tangentFactor * tangent(i) / 1024#',1)
prep=json.loads((root/'Preparation.json').read_bytes());data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes());prep['inputs'][Path(__file__).name]=sha256(Path(__file__).read_bytes()).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('RESERVED LOCAL IDENTIFIER RENAMED;FAILED COMPILE PRESERVED;PARENT BOOK UNCHANGED')
