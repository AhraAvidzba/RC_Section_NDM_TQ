"""ДЛЯ ТЕСТОВ. Пересобирает собственную диагностическую книгу с русской шапкой внутри добавляемого метода."""
from pathlib import Path
from hashlib import sha256
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold911NewtonRejectedCapture_20261126';parent=base/'Cold909ResidualArcPredictorBaseline_20261126'
assert json.loads((root/'CompilePointClosure.json').read_bytes())['State']=='Absent'
archive=root/'InitialCaptureCompilation';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json','docs/regression/Diagnostic.xlsm','docs/regression/CodeUpdate.json','PointTraceSourceComposition.json'):
    (archive/Path(name).name).write_bytes((root/name).read_bytes())
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).unlink()
(root/'docs/regression/Diagnostic.xlsm').write_bytes((parent/'docs/regression/Diagnostic.xlsm').read_bytes())
name='tests/modTestCoupledElasticState.bas';old=(root/name).read_bytes()
entry=(here/'TestNewtonRejectedCapture.bas.txt').read_text('utf-8-sig')
head,entry=entry.split('Public Sub ',1);entry='Public Sub '+entry
anchor='    Static savedMatrix()';assert entry.count(anchor)==1
entry=entry.replace(anchor,''.join('    '+line+'\n' for line in head.rstrip().splitlines())+anchor,1)
code=(parent/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*7+'\n\nPublic Function TestRunArc32Diagnostics()'
assert code.count(anchor)==1
code=code.replace(anchor,prefix*8+'\n\n'+entry.rstrip()+'\n\nPublic Function TestRunArc32Diagnostics()',1)
prep=json.loads((root/'Preparation.json').read_bytes())
for n,b in {name:code.replace('\n','\r\n').encode('utf-8-sig'),'InitialPreparedCapture.bas':old,Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/n).write_bytes(b);prep['inputs'][n]=sha256(b).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('OWN CLOSED INITIAL COMPILE ARCHIVED;EXACT PARENT BOOK RESTORED;RUSSIAN TEST HEADER IN PROCEDURE')
