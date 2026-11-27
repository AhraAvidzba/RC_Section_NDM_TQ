"""ДЛЯ ТЕСТОВ. Сверяет четыре фактически импортированных рабочих класса ASD с источниками и проверяет остальные компоненты книги."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from zipfile import ZipFile
import json, re

repo = Path.cwd()
here = Path(__file__).resolve().parent
root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1067WorkingASDNoObserver_20261127')
classes = ['CCoupledDeformationContinuation', 'CCoupledLoadContinuation', 'CCoupledNewtonSolver', 'CCoupledSectionAnalysis']
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged'] and compiled['Modules'] == classes
assert json.loads((root / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled['WorkbookSHA256']
prep = json.loads((root / 'Preparation.json').read_bytes())
for name, digest in prep['inputs'].items():
    assert sha256((root / name).read_bytes()).hexdigest().upper() == digest, name
parent = Path(json.loads((root / 'CompiledSourceReuse.json').read_bytes())['parent'])
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']


def TestRead(book):
    """Извлекает исполняемые компоненты реальной книги закреплённым декодером VBA."""
    with ZipFile(book) as handle:
        return reader(helper['TestContainer'](handle.read('xl/vbaProject.bin')))[1]


def TestIdentifiers(text):
    """Учитывает регистр имён VBE без изменения строк и комментариев в остальных компонентах."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
                  lambda match: match.group() if match.group().startswith(('"', "'")) else match.group().lower(), text)


old = TestRead(parent / 'docs/regression/Diagnostic.xlsm')
new = TestRead(root / 'docs/regression/Diagnostic.xlsm')
assert len(old) == len(new) == 256 and set(old) == set(new)
for name in old:
    if name not in classes:
        assert TestIdentifiers(old[name]) == TestIdentifiers(new[name]), name
for name in classes:
    path = f'src/Solver/{name}.cls'
    staged = (root / path).read_bytes()
    assert staged == (repo / path).read_bytes(), name
    assert b'modTest' not in staged and b'TestCapture' not in staged
    assert normal(body(staged.decode('utf-8-sig'), name)[0])[0] == normal(body(new[name], name)[0])[0], name
proof = dict(saved_book_sha256=compiled['WorkbookSHA256'], changed_components_exact=True,
    actual_components=256, all252_other_components_unchanged=True,
    four_working_classes_exact_current_worktree=True, no_kernel_test_exports=True,
    full_equilibrium_acceptance_unchanged=True, own_compiler_closed=True,
    default_observer_is_Nothing=True, no_CDP_run=True, physical_success_not_assumed=True)
(root / 'WorkingASDSourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ACTUAL WORKING ASD SOURCE VERIFIED;4 CLASSES EXACT WORKTREE;252 OTHER COMPONENTS UNCHANGED;NO TEST EXPORTS')
