"""ДЛЯ ТЕСТОВ. Сохраняет неудачную компиляцию и заменяет конфликтующие имена локальных переменных читающего теста."""
from pathlib import Path
from hashlib import sha256
import json,re

repo=Path.cwd();here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold907Cold32TangentConeIterations_20261126')
parent=root.parent/'Cold895FreshSpatialArc_20261126'
closure=json.loads((root/'CompilePointClosure.json').read_bytes());assert closure['State']=='Absent' and closure['PID']==19952
book=root/'docs/regression/Diagnostic.xlsm';assert book.read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
archive=root/'CompileFailureLocalIdentifiers';assert not archive.exists();archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json'):
    source=root/name;target=archive/name
    assert source.resolve().is_relative_to(root.resolve()) and target.resolve().is_relative_to(root.resolve())
    source.rename(target)
rel='tests/modTestCoupledElasticState.bas';old=(root/rel).read_bytes();(archive/'PreparedFailedTestModule.bas').write_bytes(old)
entry=here/'TestCold32TangentConeIterations.bas.txt';text=entry.read_text('utf-8-sig')
(archive/entry.name).write_bytes(entry.read_bytes())
def TestRename(text):
    """Меняет только имена двух локальных переменных, оставляя формулы и ветви теста теми же."""
    return re.sub(r'\bright\b','rightHandSide',re.sub(r'\bscale\b','tangentScale',text))
entry.write_text(TestRename(text),encoding='utf-8')
code=old.decode('utf-8-sig').replace('\r\n','\n')
match,=re.findall(r'(?ms)^Public Function TestReadCold32TangentConeIterations\(.*?^End Function',code)
new=code.replace(match,TestRename(match),1)
assert TestRename(match)!=match
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in {rel:new.replace('\n','\r\n').encode('utf-8-sig'),Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert book.read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print('FAILED COMPILE PRESERVED;OWN EXCEL ABSENT;PARENT BOOK EXACT;TWO LOCAL IDENTIFIERS RENAMED')
