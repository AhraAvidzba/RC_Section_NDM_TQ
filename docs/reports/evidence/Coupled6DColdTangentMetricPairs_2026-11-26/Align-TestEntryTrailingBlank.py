"""ДЛЯ ТЕСТОВ. Удаляет единственную лишнюю пустую строку после нового входа без изменения лексем и фактической книги."""
from pathlib import Path
from hashlib import sha256
import json

root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold865Cold19TangentPairs_20261125');name='tests/modTestCoupledElasticState.bas'
old=(root/name).read_bytes();code=old.decode('utf-8-sig').replace('\r\n','\n')
anchor='    TestRunColdTangentPairs = "TOTAL_COLD_TANGENT_PAIRS_TEST: passed=1; failed=0; numerical outcomes recorded not asserted; capacity_certified=False"\nEnd Function\n\n\n'
assert code.count(anchor)==1
new=code.replace(anchor,anchor[:-1],1).replace('\n','\r\n').encode('utf-8-sig')
assert old.replace(b'\r\n',b'').replace(b'\n',b'')==new.replace(b'\r\n',b'').replace(b'\n',b'')
book=root/'docs/regression/Diagnostic.xlsm';book_hash=sha256(book.read_bytes()).hexdigest().upper()
prep=json.loads((root/'Preparation.json').read_bytes());values={name:new,'PreparedTestBeforeBlankCorrection.bas':old,Path(__file__).name:Path(__file__).read_bytes()}
for n,raw in values.items():(root/n).write_bytes(raw);prep['inputs'][n]=sha256(raw).hexdigest().upper()
prep['directed_scope']['one_extra_blank_after_new_entry_removed_without_code_change']=True
prep['directed_scope']['failed_token_alignment_was_not_applied']=True
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
assert sha256(book.read_bytes()).hexdigest().upper()==book_hash
print('ONLY ONE EMPTY SOURCE LINE REMOVED;ALL NON-NEWLINE BYTES AND COMPILED BOOK EXACT')
