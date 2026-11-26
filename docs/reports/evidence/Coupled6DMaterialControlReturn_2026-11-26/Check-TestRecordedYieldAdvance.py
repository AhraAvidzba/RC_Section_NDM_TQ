"""ДЛЯ ТЕСТОВ. Применяет неизменённые независимые проверки936 к новым целям938 и сверяет неизменяемость книги."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold936RecordedYieldCoordinate_20261126';root=base/'Cold938RecordedYieldAdvance_20261126'
assert (parent/'docs/regression/Diagnostic.xlsm').read_bytes()==(root/'docs/regression/Diagnostic.xlsm').read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];inputs=[]
for folder in (parent,root):
    r=R((folder/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
    inputs.append([r.VectorTest() for _ in range(5)]);assert r.position==len(r.raw)
for before,after in zip(*inputs):
    assert before.shape==after.shape
for index in (1,2,3,4):assert (inputs[0][index]==inputs[1][index]).all()
for index in range(10):
    if index not in (2,9):assert inputs[0][0][index]==inputs[1][0][index]
assert inputs[1][0][2]==64*inputs[0][0][9] and inputs[1][0][9]==512*inputs[0][0][9]
script=here/'Check-TestRecordedYieldCoordinate.py';code=script.read_text('utf-8-sig')
old="root=base/'Cold936RecordedYieldCoordinate_20261126'";assert code.count(old)==1
code=code.replace(old,"root=base/'Cold938RecordedYieldAdvance_20261126'",1)
scope={'__file__':str(script),'__name__':'recorded_yield_advance_test'}
exec(compile(code,str(script),'exec'),scope)
proof=json.loads((root/'RecordedYieldCoordinateProof.json').read_bytes())
proof.update(immutable_compiled936_reused=True,targets_only_changed_with_verified_input_fields=True,
    independent_validator936_sha256=sha256(script.read_bytes()).hexdigest().upper())
(root/'RecordedYieldCoordinateProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
