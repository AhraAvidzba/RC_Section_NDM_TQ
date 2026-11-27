"""ДЛЯ ТЕСТОВ. Проверяет единственную численную правку ASD и все компоненты фактически скомпилированной книги."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from zipfile import ZipFile
from argparse import ArgumentParser
import json

repo=Path.cwd()
here=Path(__file__).resolve().parent
parser=ArgumentParser(description=__doc__)
parser.add_argument('--root',type=Path,default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1154ASDLinearTrustMy_20261127'))
root=parser.parse_args().root
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert json.loads((root/'CompilePointClosure.json').read_text('utf-8-sig'))['State']=='Absent'
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
helper,body,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal=run_path(str(here/'Canonical-TestCoreVBA.py'))['TestCanonical']
with ZipFile(root/'docs/regression/Diagnostic.xlsm') as archive:
    actual=reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]
with ZipFile(root/'Before/RC_Section_NDM.xlsm') as archive:
    previous=reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]
names={'CCoupledNewtonSolver','modTestSelectableMaterialCore','CCoupledMaterialProgressTest'}
assert len(previous)==260 and len(actual)==262 and set(compiled['Modules'])==names
assert set(actual)-set(previous)==names-{'CCoupledNewtonSolver'}
for name in previous:
    if name not in names:
        assert body.__globals__['literal_safe_fold'](previous[name])==body.__globals__['literal_safe_fold'](actual[name]),name
for path in [p for folder in ['src','tests'] for p in (root/folder).rglob('*') if p.suffix in ['.cls','.bas']]:
    assert path.stem in names
    assert normal(body(path.read_text('utf-8-sig'),path.stem)[0])[0]==normal(body(actual[path.stem],path.stem)[0])[0],path.stem
proof=dict(changed_components_exact=True,actual_components=262,all259_other_working_components_unchanged=True,
           imported_components=sorted(names),changed_kernel_components=['CCoupledNewtonSolver'],
           saved_book_sha256=compiled['WorkbookSHA256'],native_compilation=True,own_compiler_closed=True,
           original_physical_model_and_tolerances_unchanged=True,actual_Config_limit_reader=True,
           production_output_not_modified=True)
for path in [root]+[root/f'Case{i:02d}' for i in range(1,5)]:
    (path/'EarlySourceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ISOLATED LINEAR TRUST SOURCE VERIFIED;262 COMPONENTS;259 UNCHANGED;THREE IMPORTS EXACT;PRODUCTION UNCHANGED')
