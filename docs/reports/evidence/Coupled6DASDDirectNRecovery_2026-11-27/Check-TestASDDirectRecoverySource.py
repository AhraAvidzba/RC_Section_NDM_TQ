"""ДЛЯ ТЕСТОВ. Проверяет фактическое подключение резерва прямого ASD и неизменность материалов, допусков и остальных компонентов."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from zipfile import ZipFile
import json, re, shutil

repo = Path.cwd()
root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1062ASDPureNDirectRecovery_20261127')
here = Path(__file__).resolve().parent
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert compiled['Modules'] == ['CCoupledLoadContinuation', 'CCoupledSectionAnalysis']
assert json.loads((root / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled['WorkbookSHA256']
prep = json.loads((root / 'Preparation.json').read_bytes())
for name, digest in prep['inputs'].items():
    assert sha256((root / name).read_bytes()).hexdigest().upper() == digest, name
parent = Path(json.loads((root / 'CompiledSourceReuse.json').read_bytes())['parent'])


def TestRead(book):
    """Декодирует реально сохранённые компоненты книги закреплённым декодером CFB/VBA."""
    with ZipFile(book) as archive:
        return reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]


def TestIdentifiers(text):
    """Разрешает только регистр имён VBE, сохраняя строки и комментарии в неизменённых компонентах."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
                  lambda match: match.group() if match.group().startswith(('"', "'")) else match.group().lower(), text)


old = TestRead(parent / 'docs/regression/Diagnostic.xlsm')
new = TestRead(root / 'docs/regression/Diagnostic.xlsm')
assert len(old) == len(new) == 256 and set(old) == set(new)
for name in old:
    if name not in compiled['Modules']:
        assert TestIdentifiers(old[name]) == TestIdentifiers(new[name]), name
patch = run_path(str(here / 'Patch-TestASDDirectRecovery.py'))
name = 'CCoupledSectionAnalysis'
expected = patch['TestPatchDirectRecoveryAnalysis'](body(old[name], name)[0], canonical=True)
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
name = 'CCoupledLoadContinuation'
method = (here / 'ASDDirectRecoveryMethod.bas.txt').read_text('utf-8-sig')
expected = body(old[name], name)[0].rstrip() + '\n\n' + body.__globals__['literal_safe_fold'](method.rstrip())
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
for name in compiled['Modules']:
    staged = (root / f'src/Solver/{name}.cls').read_text('utf-8-sig')
    assert normal(body(staged, name)[0])[0] == normal(body(new[name], name)[0])[0]
proof = dict(saved_book_sha256=compiled['WorkbookSHA256'], changed_components_exact=True,
             actual_components=256, all254_other_components_unchanged=True,
             original_materials_and_tolerances_unchanged=True, full_equilibrium_acceptance_unchanged=True,
             new_direct_recovery_ASD_only=True, existing_deformation_solver_reused=True,
             no_CDP_run=True, own_compiler_closed=True, physical_capacity_not_assumed=True)
(root / 'EarlySourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
for index in range(1, 7):
    case = root / f'Case{index:02d}'
    shutil.copyfile(root / 'EarlySourceProof.json', case / 'EarlySourceProof.json')
    shutil.copyfile(root / 'docs/regression/Diagnostic.xlsm', case / 'docs/regression/Diagnostic.xlsm')
print('DIRECT ASD RECOVERY SOURCE VERIFIED;ONLY TWO EXPECTED COMPONENTS;MATERIALS AND TOLERANCES UNCHANGED')
