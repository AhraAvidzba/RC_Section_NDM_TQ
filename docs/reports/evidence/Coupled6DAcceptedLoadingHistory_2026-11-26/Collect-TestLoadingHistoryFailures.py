"""ДЛЯ ТЕСТОВ. Сохраняет отрицательные981–983 как ошибки подготовки, не подменяя ими численную несходимость модели."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import ast,ctypes,json,re

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold984AcceptedLoadingHistory_20261126';out=root/'PriorPreparationFailures';assert not out.exists();out.mkdir()
class TestMemory(ctypes.Structure):
    """ДЛЯ ТЕСТОВ. Ограничивает независимое чтение исходников при работающей нативной проверке."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory)) and memory.free_physical>6500*1024**2 and memory.free_page>6*1024**3
source=(here/'Record-TestSequentialEventHistory.py').read_text('utf-8-sig')
function=next(node for node in ast.parse(source).body if isinstance(node,ast.FunctionDef) and node.name=='TestCellContents')
env={'ET':ET,'ns':{'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}}
exec(compile(ast.Module(body=[function],type_ignores=[]),str(here/'Record-TestSequentialEventHistory.py'),'exec'),env)
reference=(base/'Cold954RefinedEventApproach_20261126/docs/regression/Diagnostic.xlsm').read_bytes()
helper,_,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
proof={}
for n in (981,982,983):
    old=base/('Cold%dAcceptedLoadingHistory_20261126'%n);target=out/str(n);target.mkdir()
    book=(old/'docs/regression/Diagnostic.xlsm').read_bytes()
    with ZipFile(BytesIO(reference)) as before,ZipFile(BytesIO(book)) as after:
        assert set(before.namelist())==set(after.namelist())
        sheets=[name for name in after.namelist() if re.fullmatch(r'xl/worksheets/sheet\d+\.xml',name)];assert len(sheets)==5
        for name in sheets:assert env['TestCellContents'](before.read(name))==env['TestCellContents'](after.read(name))
        for name in after.namelist():
            if name=='xl/sharedStrings.xml' or name.startswith('xl/externalLinks/') or name=='xl/connections.xml':assert before.read(name)==after.read(name)
        vba=after.read('xl/vbaProject.bin')
    (target/'ActualVBA.bin').write_bytes(vba)
    _,modules=reader(helper['TestContainer'](vba));assert len(modules)==257
    (target/'ActualTestModule.bas').write_text(modules['modTestCoupledElasticState'],encoding='utf-8')
    names=['Preparation.json','CompilePoint.log','CompilePointFailure.json','CompilePointClosure.json','docs/regression/CodeUpdate.json',
        'ParentCodeUpdate.json','ParentCompileClosure.json','Compile-TestExternalPoint.ps1','Stage-TestAcceptedLoadingHistory.py',
        'TestRunAcceptedLoadingHistory.bas.txt','tests/modTestCoupledElasticState.bas','PointTraceSourceComposition.json']
    names += [p.relative_to(old).as_posix() for p in (old/'Native').rglob('*') if p.is_file()]
    for name in names:
        path=old/name
        if path.is_file():
            destination=target/name;destination.parent.mkdir(parents=True,exist_ok=True);destination.write_bytes(path.read_bytes())
    closure=json.loads((old/'CompilePointClosure.json').read_bytes());assert closure['State']=='Absent'
    item=dict(all5_sheet_values_formulas_types_addresses_exact_public954=True,shared_strings_and_external_links_unchanged=True,
        actual_book_sha256=sha256(book).hexdigest().upper(),actual_VBA_sha256=sha256(vba).hexdigest().upper(),
        compile_own_process_closed=True,numerical_core_change_claimed=False,physical_failure_claimed=False)
    if n==981:
        assert (old/'CompilePointFailure.json').exists() and not (old/'Native/RealColdReplay.txt').exists()
        assert book==(base/'Cold980LoadingBudgetPair_20261126/docs/regression/Diagnostic.xlsm').read_bytes()
        item.update(outcome='compile rejected: private helper not included in Public-only entry injection',new_code_not_saved=True,native_numerical_test_not_started=True)
    elif n==982:
        assert json.loads((old/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))['Compiled']
        assert not (old/'Native/RealColdReplay.txt').exists()
        item.update(outcome='source audit rejected: entry insertion order and comments differ from prepared source',native_numerical_test_not_started=True)
    else:
        assert (old/'Native/MaterialMidpointPair.bin').read_bytes()==(base/'Cold980LoadingBudgetPair_20261126/Native/MaterialMidpointPair.bin').read_bytes()
        assert (old/'Native/LoadingHistory.bin').stat().st_size==16 and (old/'Native/HistoryStateBefore.1.bin').stat().st_size==512
        closure=json.loads((old/'Native/RealColdReplay.txt.ProcessClosure.json').read_text('utf-8-sig'));assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode']==1
        item.update(outcome='test snapshot failed before first point write: unallocated CopyAcceptedStress destination',first_loading_pair_bitexact980=True,
            no_state_acceptance=True,native_own_process_closed=True)
    proof[str(n)]=item
(out/'Summary.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PRESERVED NEGATIVE PREPARATION981-983;NO PHYSICAL FAILURE CLAIM;ALL DATA EXACT PUBLIC954')
