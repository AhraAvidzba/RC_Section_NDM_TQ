"""ДЛЯ ТЕСТОВ. Сверяет три изменённых компонента частного раннего продолжения и 253 сохранённых компонента книги."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from runpy import run_path
import json,re

here=Path(__file__).resolve().parent;repo=Path.cwd();root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1026ASDEarlyDeformation_20261126')
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
changed={'CCoupledLoadContinuation':'src/Solver/CCoupledLoadContinuation.cls','CCoupledSectionAnalysis':'src/Solver/CCoupledSectionAnalysis.cls',
    'modTestSelectableMaterialCore':'tests/modTestSelectableMaterialCore.bas'}
assert compiled['Compiled'] and not compiled['ProductionChanged'] and set(compiled['Modules'])==set(changed)
assert json.loads((root/'CompilePointClosure.json').read_text('utf-8-sig'))['State']=='Absent'
prep=json.loads((root/'Preparation.json').read_bytes())
for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,name
parent=Path(json.loads((root/'CompiledSourceReuse.json').read_bytes())['parent'])
before=(parent/'docs/regression/Diagnostic.xlsm').read_bytes();after=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(before).hexdigest().upper()==json.loads((root/'ParentCodeUpdate.json').read_bytes())['WorkbookSHA256']
assert sha256(after).hexdigest().upper()==compiled['WorkbookSHA256']
helper,body,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
with ZipFile(BytesIO(before)) as z:old_bin=z.read('xl/vbaProject.bin')
with ZipFile(BytesIO(after)) as z:new_bin=z.read('xl/vbaProject.bin')
_,old=reader(helper['TestContainer'](old_bin));_,new=reader(helper['TestContainer'](new_bin));assert len(old)==256 and set(old)==set(new)

def TestIdentifiers(text):
    """Сохраняет литералы и комментарии; допускает только регистр имён VBE."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',lambda m:m.group() if m.group().startswith(('"',"'")) else m.group().lower(),text)

for name in old:
    if name not in changed:assert TestIdentifiers(old[name])==TestIdentifiers(new[name]),name
normal=run_path(str(here/'Canonical-TestCoreVBA.py'))['TestCanonical'];normalization={}
for name,path in changed.items():
    expected,_=helper['TestMVTCompilerBody'](body,(root/path).read_text('utf-8-sig'),name)
    before=TestIdentifiers(expected);after=TestIdentifiers(body(new[name],name)[0])
    if before!=after:
        before,record=normal(before);after,_=normal(after);normalization[name]=record
    assert before==after,name
proof=dict(actual_components=256,all253_other_components_unchanged_except_identifier_case=True,
    three_changed_components_exact=True,saved_book_sha256=compiled['WorkbookSHA256'],own_compiler_closed=True,
    exact_numeric_value_and_comment_codepage_normalization=normalization,runtime_string_literals_unchanged=True,
    physical_capacity_not_assumed=True,observer_only_copies_fields_and_immutable_operators=True)
(root/'EarlySourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('EARLY ASD SOURCE VERIFIED;3 EXACT CHANGED COMPONENTS;253 OTHERS UNCHANGED')
