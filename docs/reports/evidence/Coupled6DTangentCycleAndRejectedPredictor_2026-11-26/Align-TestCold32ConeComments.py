"""ДЛЯ ТЕСТОВ. Сопоставляет только перенос четырёх одинаковых служебных комментариев при добавлении нового входа VBE."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold907Cold32TangentConeIterations_20261126')
name='tests/modTestCoupledElasticState.bas';old=(root/name).read_bytes();code=old.decode('utf-8-sig').replace('\r\n','\n')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor='End Function\n'+prefix*5+'\n\nPublic Function TestReadFreshSpatialArcCases()'
assert code.count(anchor)==1
code=code.replace(anchor,'End Function\n\n\nPublic Function TestReadFreshSpatialArcCases()',1)
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n'+prefix
assert code.count(anchor)==1;code=code.replace(anchor,anchor+prefix*5,1)
def executable(text):
    """Сохраняет все неслужебные строки, включая пустые, строки и комментарии рабочих процедур."""
    return text.replace(prefix,'')
assert executable(old.decode('utf-8-sig').replace('\r\n','\n'))==executable(code)
book=root/'docs/regression/Diagnostic.xlsm';before=sha256(book.read_bytes()).hexdigest()
prep=json.loads((root/'Preparation.json').read_bytes())
for n,b in {name:code.replace('\n','\r\n').encode('utf-8-sig'),'PreparedCold32JacobianBeforeCommentAlignment.bas':old,Path(__file__).name:Path(__file__).read_bytes()}.items():
 (root/n).write_bytes(b);prep['inputs'][n]=sha256(b).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert sha256(book.read_bytes()).hexdigest()==before
print('FOUR IDENTICAL HEADER COMMENTS MOVED;ALL OTHER SOURCE LINES AND BOOK EXACT')
