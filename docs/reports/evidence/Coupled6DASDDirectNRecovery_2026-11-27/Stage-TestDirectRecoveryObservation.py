"""ДЛЯ ТЕСТОВ. Готовит читающий контроль неудачной точной пробы прямого ASD с сохранением всей деформационной ветви."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json, shutil

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1062ASDPureNDirectRecovery_20261127'
root = base / 'Cold1063ASDPureNDirectObservation_20261127'
assert not root.exists()
proof = json.loads((parent / 'ASDDirectPureNExecutionProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['all56plus4_equations_and_six_resultants_verified']
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
for name in ('Guard-IsolatedOffice.ps1', 'Run-TestASDDirectPureN.ps1'):
    shutil.copyfile(parent / name, root / name)
(root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'ParentCodeUpdate.json')
(root / 'CompiledSourceReuse.json').write_text(json.dumps(dict(parent=str(parent), variant='observe-direct'), indent=2) + '\n', encoding='utf-8')
code = (parent / 'src/Solver/CCoupledLoadContinuation.cls').read_text('utf-8-sig')
code = run_path(str(here / 'Patch-TestDirectRecoveryObservation.py'))['TestPatchDirectRecoveryObservation'](code)
(root / 'src/Solver/CCoupledLoadContinuation.cls').write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
compiler = (parent / 'Compile-TestASDDirectRecovery.ps1').read_text('utf-8-sig')
before = "@('CCoupledSectionAnalysis','CCoupledLoadContinuation')"
assert compiler.count(before) == 1
compiler = compiler.replace(before, "'CCoupledLoadContinuation'", 1)
(root / 'Compile-TestDirectRecoveryObservation.ps1').write_bytes(compiler.replace('\n', '\r\n').encode('utf-8-sig'))
for name in ('Stage-TestDirectRecoveryObservation.py', 'Patch-TestDirectRecoveryObservation.py'):
    shutil.copyfile(here / name, root / name)
prep = dict(parent=str(parent), directed_scope=dict(case='ASDDirectPureNObservation',
    actual_import_this_pass=['CCoupledLoadContinuation'], only_reading_observations_added=True,
    numerical_algorithm_unchanged=True, physical_acceptance_not_assumed=True, no_CDP_run=True),
    inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
            for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix() != 'docs/regression/Diagnostic.xlsm'})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('DIRECT ASD OBSERVATION PREPARED;NUMERICAL EXPRESSIONS UNCHANGED')
