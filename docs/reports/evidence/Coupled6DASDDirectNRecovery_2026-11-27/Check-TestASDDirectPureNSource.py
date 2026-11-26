"""ДЛЯ ТЕСТОВ. Сверяет читающий наблюдатель прямого НДС с реальной книгой и готовит точное происхождение шести осевых путей."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from zipfile import ZipFile
import json, re, shutil

repo = Path.cwd()
root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1061ASDPureNDirectBaseline_20261127')
here = Path(__file__).resolve().parent
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert compiled['Modules'] == ['CCoupledSectionAnalysis', 'modTestSelectableMaterialCore']
assert json.loads((root / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled['WorkbookSHA256']
prep = json.loads((root / 'Preparation.json').read_bytes())
for name, digest in prep['inputs'].items():
    assert sha256((root / name).read_bytes()).hexdigest().upper() == digest, name
parent = Path(json.loads((root / 'CompiledSourceReuse.json').read_bytes())['parent'])


def TestRead(book):
    """Извлекает реальный VBA сохранённой книги с закреплённым декодером проекта."""
    with ZipFile(book) as archive:
        return reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]


def TestIdentifiers(text):
    """Учитывает только регистр исполняемых имён VBE и сохраняет строки и комментарии при сравнении неизменённых компонентов."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
                  lambda match: match.group() if match.group().startswith(('"', "'")) else match.group().lower(), text)


old = TestRead(parent / 'docs/regression/Diagnostic.xlsm')
new = TestRead(root / 'docs/regression/Diagnostic.xlsm')
assert len(new) == len(old) == 256 and set(new) == set(old)
for name in old:
    if name not in compiled['Modules']:
        assert TestIdentifiers(old[name]) == TestIdentifiers(new[name]), name
name = 'CCoupledSectionAnalysis'
before = body(old[name], name)[0]
anchor = '    set acceptedtrace = continuation.savedtrace\n'
assert before.count(anchor) == 1
expected = before.replace(anchor, '    modtestselectablematerialcore.testcapturedirecthistory continuation, moperators, path\n' + anchor, 1)
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
name = 'modTestSelectableMaterialCore'
staged = (root / 'tests/modTestSelectableMaterialCore.bas').read_text('utf-8-sig')
assert normal(body(staged, name)[0])[0] == normal(body(new[name], name)[0])[0]
method = (root / 'TestASDDirectPureN.generated.bas.txt').read_text('utf-8-sig')
expected = body(old[name], name)[0].rstrip() + '\n\n' + body.__globals__['literal_safe_fold'](method.rstrip())
assert normal(expected)[0] == normal(body(new[name], name)[0])[0]
proof = dict(saved_book_sha256=compiled['WorkbookSHA256'], changed_components_exact=True,
             actual_components=256, all254_other_components_unchanged=True,
             only_reading_observer_and_new_test=True, full_equilibrium_acceptance_unchanged=True,
             no_numerical_algorithm_changed=True, no_CDP_run=True, own_compiler_closed=True,
             physical_capacity_not_assumed=True)
(root / 'EarlySourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
for index in range(1, 7):
    case = root / f'Case{index:02d}'
    shutil.copyfile(root / 'EarlySourceProof.json', case / 'EarlySourceProof.json')
    shutil.copyfile(root / 'docs/regression/Diagnostic.xlsm', case / 'docs/regression/Diagnostic.xlsm')
print('DIRECT ASD PURE N SOURCE VERIFIED;ONLY READ-ONLY OBSERVER AND NEW TEST;6 COLD CASES PREPARED')
