"""ДЛЯ ТЕСТОВ. Учитывает единственную добавленную VBE пустую строку между новым методом и первым исходным методом."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold911NewtonRejectedCapture_20261126')
name='tests/modTestCoupledElasticState.bas';old=(root/name).read_bytes();code=old.decode('utf-8-sig').replace('\r\n','\n')
anchor='End Sub\n\nPublic Function TestRunArc32Diagnostics()';assert code.count(anchor)==1
code=code.replace(anchor,'End Sub\n\n\nPublic Function TestRunArc32Diagnostics()',1)
assert old.decode('utf-8-sig').replace('\r\n','\n').splitlines()==code.replace('End Sub\n\n\nPublic Function TestRunArc32Diagnostics()',anchor,1).splitlines()
book=root/'docs/regression/Diagnostic.xlsm';before=sha256(book.read_bytes()).hexdigest()
prep=json.loads((root/'Preparation.json').read_bytes())
for n,b in {name:code.replace('\n','\r\n').encode('utf-8-sig'),'BeforeCaptureBlankAlignment.bas':old,Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/n).write_bytes(b);prep['inputs'][n]=sha256(b).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert sha256(book.read_bytes()).hexdigest()==before
print('ONE BLANK LINE ONLY;BOOK EXACT')
