"""ДЛЯ ТЕСТОВ. Выравнивает только пустые строки двух новых входов с нативной вставкой VBE; книга и все лексемы неизменны."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold867FixedMetricKernel_20261126');name='tests/modTestCoupledElasticState.bas'
old=(root/name).read_bytes();code=old.decode('utf-8-sig').replace('\r\n','\n')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n\n\n'+prefix
assert code.count(anchor)==1
code=code.replace(anchor,'Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n'+prefix+'\n\n',1)
anchor=prefix+'Public Sub TestWriteFixedMetricGuide('
assert code.count(anchor)==1;code=code.replace(anchor,prefix+'\n\nPublic Sub TestWriteFixedMetricGuide(',1)
data=code.replace('\n','\r\n').encode('utf-8-sig')
assert old.replace(b'\r\n',b'').replace(b'\n',b'')==data.replace(b'\r\n',b'').replace(b'\n',b'')
book=root/'docs/regression/Diagnostic.xlsm';book_hash=sha256(book.read_bytes()).hexdigest().upper()
prep=json.loads((root/'Preparation.json').read_bytes())
for n,b in {name:data,'PreparedTestBeforeBlankAlignment.bas':old,Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/n).write_bytes(b);prep['inputs'][n]=sha256(b).hexdigest().upper()
prep['directed_scope']['two_new_entry_blank_lines_aligned_no_token_or_book_change']=True
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert sha256(book.read_bytes()).hexdigest().upper()==book_hash
print('SOURCE BLANK LINES ALIGNED;NO CODE OR BOOK CHANGE')
