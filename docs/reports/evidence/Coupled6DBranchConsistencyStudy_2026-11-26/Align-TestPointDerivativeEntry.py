"""ДЛЯ ТЕСТОВ. Совмещает расположение единственного комментария вставки с VBE; методы и книга неизменны."""
from pathlib import Path
from hashlib import sha256
import json
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold872PointDerivativeAudit20_20261126');name='tests/modTestCoupledElasticState.bas'
old=(root/name).read_bytes();text=old.decode('utf-8-sig').replace('\r\n','\n');prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor='End Function\n'+prefix+'\n\nPublic Sub TestRunFixedMetricArcWalk()';assert text.count(anchor)==1
text=text.replace(anchor,'End Function\n\n\nPublic Sub TestRunFixedMetricArcWalk()',1)
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert text.count(anchor)==1
text=text.replace(anchor,anchor+prefix,1)
data=text.replace('\n','\r\n').encode('utf-8-sig')
def TestNonFormatting(raw):
    """Сравнивает все оставшиеся байты после выделенного комментария вставки и переводов строк."""
    return raw.decode('utf-8-sig').replace('\r\n','\n').replace(prefix,'').replace('\n','')
assert TestNonFormatting(old)==TestNonFormatting(data)
book_hash=sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest()
prep=json.loads((root/'Preparation.json').read_bytes())
for n,b in {name:data,'PreparedBeforePointCommentAlignment.bas':old,Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/n).write_bytes(b);prep['inputs'][n]=sha256(b).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest()==book_hash
print('ONLY INSERTION COMMENT POSITION ALIGNED;METHODS AND BOOK EXACT')
