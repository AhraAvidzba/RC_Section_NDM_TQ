"""ДЛЯ ТЕСТОВ. Готовит проверку прямого ASD с восстановлением точного сочетания из существующей деформационной ветви."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json, shutil

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1061ASDPureNDirectBaseline_20261127'
root = base / 'Cold1062ASDPureNDirectRecovery_20261127'
assert not root.exists()
checked = json.loads((parent / 'ASDDirectPureNExecutionProof.json').read_bytes())
assert checked['native']['own_process_closed'] and checked['all56plus4_equations_and_six_resultants_verified']
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
for directory in ('src', 'tests', 'tools', 'Input', 'workbook'):
    shutil.copytree(parent / directory, root / directory)
for directory in ('Native', 'Temp', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for index in range(1, 7):
    case = root / f'Case{index:02d}'
    shutil.copytree(parent / f'Case{index:02d}/Input', case / 'Input')
    for directory in ('Native', 'docs/regression'):
        (case / directory).mkdir(parents=True, exist_ok=True)
patch = run_path(str(here / 'Patch-TestASDDirectRecovery.py'))
files = {'docs/regression/Diagnostic.xlsm': book,
         'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
         'CompiledSourceReuse.json': (json.dumps(dict(parent=str(parent), variant='direct-recovery'), indent=2) + '\n').encode()}
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDDirectPureN.ps1'):
    files[name] = (parent / name).read_bytes()
compiler = (parent / 'Compile-TestASDDirectPureN.ps1').read_text('utf-8-sig')
before = "@('CCoupledSectionAnalysis','modTestSelectableMaterialCore')"
assert compiler.count(before) == 1
files['Compile-TestASDDirectRecovery.ps1'] = compiler.replace(before, "@('CCoupledSectionAnalysis','CCoupledLoadContinuation')", 1).replace('\n', '\r\n').encode('utf-8-sig')
for name, function in [('CCoupledSectionAnalysis', 'TestPatchDirectRecoveryAnalysis'),
                        ('CCoupledLoadContinuation', 'TestPatchDirectRecoveryContinuation')]:
    path = f'src/Solver/{name}.cls'
    text = patch[function]((parent / path).read_text('utf-8-sig'))
    assert max(map(len, text.splitlines())) < 1024
    files[path] = text.replace('\n', '\r\n').encode('utf-8-sig')
for name in ('Stage-TestASDDirectRecovery.py', 'Patch-TestASDDirectRecovery.py', 'ASDDirectRecoveryMethod.bas.txt'):
    files[name] = (here / name).read_bytes()
for name, data in files.items():
    (root / name).write_bytes(data)
prep = dict(parent=str(parent), directed_scope=dict(case='ASDDirectPureNRecovery',
    actual_import_this_pass=['CCoupledSectionAnalysis', 'CCoupledLoadContinuation'],
    existing_deformation_solver_reused=True, exact_load_solved_from_same_left_history=True,
    original_tolerances_and_materials_unchanged=True, no_CDP_run=True, physical_success_not_assumed=True),
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ASD DIRECT RECOVERY PREPARED;SAME SIX CASES,MATERIALS,TOLERANCES;NEW EXACT-LOAD RECOVERY;NO CDP')
