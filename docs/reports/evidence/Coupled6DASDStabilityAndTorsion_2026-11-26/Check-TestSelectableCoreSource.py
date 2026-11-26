"""ДЛЯ ТЕСТОВ. Сверяет единственный новый тестовый модуль и сохранность всех 255 компонентов проверенной свежей книги."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from runpy import run_path
import json,re

repo=Path.cwd();here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1018SelectableMaterialCore_20261126')
prep=json.loads((root/'Preparation.json').read_bytes());compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged'] and compiled['Modules']==['modTestSelectableMaterialCore']
assert json.loads((root/'CompilePointClosure.json').read_text('utf-8-sig'))['State']=='Absent'
for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,name
parent=Path(json.loads((root/'CompiledSourceReuse.json').read_bytes())['parent'])
before=(parent/'workbook/output/SelectableLaws.xlsm').read_bytes();after=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(before).hexdigest().upper()==json.loads((root/'ParentCodeUpdate.json').read_bytes())['WorkbookSHA256']
assert sha256(after).hexdigest().upper()==compiled['WorkbookSHA256']
helper,body,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
with ZipFile(BytesIO(before)) as archive:old_bin=archive.read('xl/vbaProject.bin')
with ZipFile(BytesIO(after)) as archive:new_bin=archive.read('xl/vbaProject.bin')
_,old=reader(helper['TestContainer'](old_bin));_,new=reader(helper['TestContainer'](new_bin))
assert len(old)==255 and set(new)==set(old)|{'modTestSelectableMaterialCore'}

def TestIdentifiers(text):
    """Сохраняет строки и комментарии, разрешая только регистр исполняемых имён VBE."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',lambda m:m.group() if m.group().startswith(('"',"'")) else m.group().lower(),text)

for name in old:assert TestIdentifiers(old[name])==TestIdentifiers(new[name]),name
name='modTestSelectableMaterialCore';expected,_=helper['TestMVTCompilerBody'](body,(root/'tests'/str(name+'.bas')).read_text('utf-8-sig'),name)
actual=body(new[name],name)[0]
assert TestIdentifiers(expected)==TestIdentifiers(actual),name
proof=dict(actual_components=256,all255_parent_components_unchanged_except_identifier_case=True,
    new_test_module_exact=True,own_compiler_closed=True,saved_book_sha256=compiled['WorkbookSHA256'],
    calculations_not_yet_verified=True,native_dropdown_not_yet_verified=True,whole_goal_completed=False)
(root/'SelectableCoreSourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SELECTABLE CORE SOURCE VERIFIED;256 COMPONENTS;255 EXACT PARENT COMPONENTS EXCEPT IDENTIFIER CASE')
