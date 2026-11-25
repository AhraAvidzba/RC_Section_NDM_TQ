"""ДЛЯ ТЕСТОВ. Синхронизирует порядок нового тестового входа с фактической вставкой VBE без изменения книги и существующих методов."""
from pathlib import Path
from hashlib import sha256
import json

base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold864Cold19BranchTangents_20261125';parent=base/'Cold858BalancedInitialStressReplay_20261125'
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
marker="\n' FOR TESTS: baseline entry point only; existing methods retained.\nPublic Function TestReadColdBranchTangents()"
position=code.index(marker);baseline=code[:position];entry=code[position+1:]
assert baseline==(parent/name).read_text('utf-8-sig')
prefix,entry=entry.split('\n',1)
assert entry.startswith('Public Function TestReadColdBranchTangents()') and entry.rstrip().endswith('End Function')
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert baseline.count(anchor)==1
aligned=baseline.replace(anchor,anchor+prefix+'\n\n\n'+entry,1)
assert aligned.replace(prefix+'\n\n\n'+entry,'',1)==baseline
book=root/'docs/regression/Diagnostic.xlsm';book_hash=sha256(book.read_bytes()).hexdigest().upper()
prep=json.loads((root/'Preparation.json').read_bytes());values={name:aligned.replace('\n','\r\n').encode('utf-8-sig'),Path(__file__).name:Path(__file__).read_bytes()}
for n,raw in values.items():(root/n).write_bytes(raw);prep['inputs'][n]=sha256(raw).hexdigest().upper()
prep['directed_scope']['new_entry_inserted_before_existing_procedures_by_VBE']=True
prep['directed_scope']['existing_method_bodies_and_new_test_body_unchanged']=True
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
assert sha256(book.read_bytes()).hexdigest().upper()==book_hash
print('ONLY PREPARED TEST ENTRY ORDER ALIGNED;BASELINE AND NEW METHOD EXACT;COMPILED BOOK UNCHANGED')
