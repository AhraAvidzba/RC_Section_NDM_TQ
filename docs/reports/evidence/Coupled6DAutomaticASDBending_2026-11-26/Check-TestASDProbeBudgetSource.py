"""ДЛЯ ТЕСТОВ. Проверяет единственное изменение частного контроллера и исправляет ошибочную трактовку признака сохранности основной книги в обвязке."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from runpy import run_path
import json,re

here=Path(__file__).resolve().parent;repo=Path.cwd()
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1022ASDBendingProbeBudget_20261126')
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged'] and compiled['Modules']==['CCoupledDeformationContinuation']
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
_,old=reader(helper['TestContainer'](old_bin));_,new=reader(helper['TestContainer'](new_bin))
assert len(old)==256 and set(old)==set(new)

def TestIdentifiers(text):
    """Сохраняет строки и комментарии; допускает только регистр исполняемых имён VBE."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',lambda m:m.group() if m.group().startswith(('"',"'")) else m.group().lower(),text)

name='CCoupledDeformationContinuation'
for item in old:
    if item!=name:assert TestIdentifiers(old[item])==TestIdentifiers(new[item]),item
expected,_=helper['TestMVTCompilerBody'](body,(root/'src/Solver'/str(name+'.cls')).read_text('utf-8-sig'),name)
assert TestIdentifiers(expected)==TestIdentifiers(body(new[name],name)[0]),name
needle='private const probe_limit as long = 96'
previous=TestIdentifiers(body(old[name],name)[0]);current=TestIdentifiers(body(new[name],name)[0])
assert previous.count(needle)==1
assert previous.replace(needle,'private const probe_limit as long = 192')==current
# ProductionChanged относится к защищённой основной книге, а не к типу импортируемого модуля.
runner=root/'Run-TestASDBendingMainCapacity.ps1';text=runner.read_text('utf-8-sig')
assert '!$compiled.ProductionChanged' in text and not (root/'Native/RealColdReplay.txt').exists()
(root/'PreparationBeforeHarnessFix.json').write_bytes((root/'Preparation.json').read_bytes())
(root/'RunnerBeforeHarnessFix.ps1.txt').write_bytes(runner.read_bytes())
text=text.replace('!$compiled.ProductionChanged','$compiled.ProductionChanged',1)
runner.write_bytes(text.replace('\n','\r\n').encode('utf-8-sig'))
prep['inputs'][runner.name]=sha256(runner.read_bytes()).hexdigest().upper()
prep['harness_correction']='ProductionChanged=False means published workbook unchanged; required alongside exact imported controller name.'
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
proof=dict(actual_components=256,all255_other_components_unchanged_except_identifier_case=True,
    only_controller_change='PROBE_LIMIT96to192',saved_book_sha256=compiled['WorkbookSHA256'],own_compiler_closed=True,
    published_workbook_unchanged=True,numerical_calculation_not_yet_checked=True)
(root/'ProbeBudgetSourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ASD BUDGET SOURCE VERIFIED; 96to192 ONLY; RUNNER RELEASE-PROTECTION GUARD CORRECTED BEFORE EXECUTION')
