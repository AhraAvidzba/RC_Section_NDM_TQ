"""ДЛЯ ТЕСТОВ. Отражает два пустых разделителя VBE перед вторым добавленным методом, не меняя исполняемый код."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold936RecordedYieldCoordinate_20261126')
name='tests/modTestCoupledElasticState.bas';p=root/name;raw=p.read_bytes()
before=root/'BeforeRecordedYieldHelperBlankAlignment.bas';assert not before.exists();before.write_bytes(raw)
code=raw.decode('utf-8-sig').replace('\r\n','\n')
anchor="' FOR TESTS: baseline entry point only; existing methods retained.\nPublic Sub TestWriteRecordedYieldStressTest("
assert code.count(anchor)==1
code=code.replace(anchor,anchor.replace('\nPublic','\n\n\nPublic'),1)
p.write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
prep=json.loads((root/'Preparation.json').read_bytes())
for file in (p,before):prep['inputs'][file.relative_to(root).as_posix()]=sha256(file.read_bytes()).hexdigest().upper()
for file in (here/'Stage-TestRecordedYieldCoordinate.py',Path(__file__)):
    (root/file.name).write_bytes(file.read_bytes());prep['inputs'][file.name]=sha256(file.read_bytes()).hexdigest().upper()
prep['directed_scope']['second_entry_VBE_blank_alignment_only']=True
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('VBE SECOND ENTRY TWO BLANKS ALIGNED;NO NUMERICAL CHANGE;REQUIRES FULL SOURCE AUDIT')
