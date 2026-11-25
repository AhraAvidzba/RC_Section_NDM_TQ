"""ДЛЯ ТЕСТОВ. Убирает одну лишнюю пустую строку после нового входа; книга и все остальные байты исходника неизменны."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold876VerifiedHCornerPairs20_20261126');name='tests/modTestCoupledElasticState.bas';old=(root/name).read_bytes();code=old.decode('utf-8-sig').replace('\r\n','\n')
anchor='End Function\n\n\n\nPublic Function TestReadSharedCurveCases()';assert code.count(anchor)==1
data=code.replace(anchor,'End Function\n\n\nPublic Function TestReadSharedCurveCases()',1).replace('\n','\r\n').encode('utf-8-sig')
assert old.replace(b'\r\n',b'').replace(b'\n',b'')==data.replace(b'\r\n',b'').replace(b'\n',b'')
book_hash=sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest();prep=json.loads((root/'Preparation.json').read_bytes())
for n,b in {name:data,'PreparedBeforeVerifiedHBlankAlignment.bas':old,Path(__file__).name:Path(__file__).read_bytes()}.items():(root/n).write_bytes(b);prep['inputs'][n]=sha256(b).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8');assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest()==book_hash
print('ONE EMPTY LINE REMOVED;ALL NON-NEWLINE SOURCE BYTES AND BOOK EXACT')
