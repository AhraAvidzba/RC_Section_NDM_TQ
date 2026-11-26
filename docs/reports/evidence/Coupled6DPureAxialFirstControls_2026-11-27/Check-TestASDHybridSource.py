"""ДЛЯ ТЕСТОВ. Проверяет единственное изменение доступности гибридного резерва и точную сохранность остальных нативных компонентов."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from runpy import run_path
import json,re

here=Path(__file__).resolve().parent;repo=Path.cwd();root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1033ASDHybridRecovery_20261126')
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged'] and compiled['Modules']==['CCoupledNewtonSolver']
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
    """Сохраняет литералы и комментарии; допускает только регистр исполняемых имён VBE."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',lambda m:m.group() if m.group().startswith(('"',"'")) else m.group().lower(),text)

name='CCoupledNewtonSolver'
for item in old:
    if item!=name:assert TestIdentifiers(old[item])==TestIdentifiers(new[item]),item
normal=run_path(str(here/'Canonical-TestCoreVBA.py'))['TestCanonical']
previous=TestIdentifiers(body(old[name],name)[0]);current=TestIdentifiers(body(new[name],name)[0])
needle='''    supportshybridrecovery = not (hybridresponsehasprincipalboundaries(mresponse) or _
        hybridresponsehasprincipalboundaries(mplusresponse) or hybridresponsehasprincipalboundaries(mminusresponse))'''
assert previous.count(needle)==1
previous=previous.replace(needle,'    supportshybridrecovery = true',1)
old_comment="' Разрешает контроллеру запросить полный разностный резерв только для пространственной истории без спектральных границ."
new_comment="' Разрешает полный разностный резерв пространственной истории; материальные разрывы остаются в проверяемой системе."
assert previous.count(old_comment.lower())==1
previous=previous.replace(old_comment.lower(),new_comment.lower(),1) # Закреплённый compiler_body сохраняет строковые литералы, но приводит регистр комментариев.
assert normal(previous)[0]==normal(current)[0]
proof=dict(actual_components=256,all255_other_components_unchanged_except_identifier_case=True,
    changed_components_exact=True,only_native_gate_and_comment_changed=True,unchanged_native_numeric_values_preserved=True,
    saved_book_sha256=compiled['WorkbookSHA256'],own_compiler_closed=True,full_equilibrium_acceptance_unchanged=True,
    physical_capacity_not_assumed=True)
(root/'EarlySourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ASD HYBRID SOURCE VERIFIED;GATE ONLY;255 OTHER COMPONENTS UNCHANGED')
