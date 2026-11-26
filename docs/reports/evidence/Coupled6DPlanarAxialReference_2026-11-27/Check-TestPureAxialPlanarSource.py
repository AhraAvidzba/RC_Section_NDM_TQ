"""ДЛЯ ТЕСТОВ. Проверяет фактический импорт только осевого теста и неизменность всех остальных компонентов и плоского ядра."""
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
from runpy import run_path
import json, re

repo = Path.cwd()
root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1059PureAxialPlanarTension_20261127')
here = Path(__file__).resolve().parent
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert compiled['Modules'] == ['modTestCoupledElasticState']
assert json.loads((root / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == compiled['WorkbookSHA256']
prep = json.loads((root / 'Preparation.json').read_bytes())
for name, digest in prep['inputs'].items():
    assert sha256((root / name).read_bytes()).hexdigest().upper() == digest, name
parent = Path(json.loads((root / 'CompiledSourceReuse.json').read_bytes())['parent'])


def TestRead(book):
    """Читает исполняемые компоненты из реальной сохранённой книги, а не из предполагаемой копии исходников."""
    with ZipFile(book) as archive:
        return reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]


old = TestRead(parent / 'docs/regression/Diagnostic.xlsm')
new = TestRead(root / 'docs/regression/Diagnostic.xlsm')
assert set(old) == set(new)
module = 'modTestCoupledElasticState'


def TestIdentifiers(text):
    """Допускает регистр имён VBE во всех компонентах, включая пустые модули листов, не меняя строк и комментариев."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
                  lambda match: match.group() if match.group().startswith(('"', "'")) else match.group().lower(), text)


for name in old:
    if name != module:
        assert TestIdentifiers(old[name]) == TestIdentifiers(new[name]), name
method = (root / 'TestPureAxialPlanarReference.bas.txt').read_text('utf-8-sig').rstrip()
method = body.__globals__['literal_safe_fold'](method)
before = body(old[module], module)[0]
anchor = 'public function testrunpurebendingplanarreference('
assert before.count(anchor) == 1
expected = before.replace(anchor, method + '\n\n\n' + anchor, 1)  # Шаблон метода заканчивается переводом строки, к которому сборщик добавляет ещё два.
assert normal(expected)[0] == normal(body(new[module], module)[0])[0]
current = json.loads((root / 'CurrentPlanarKernelProof.json').read_bytes())
assert current['all9_actual_planar_classes_exact_current_output']
for name in current['classes']:
    assert normal(body(old[name], name)[0])[0] == normal(body(new[name], name)[0])[0]
proof = dict(saved_book_sha256=compiled['WorkbookSHA256'], only_new_axial_test_method=True,
             actual_components=len(new), all_other_components_unchanged=True,
             actual_planar_kernel_matches_current_output=True, compiler_process_closed=True,
             physical_capacity_not_assumed=True)
(root / 'PlanarAxialSourceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('ACTUAL PLANAR PURE N TEST VERIFIED', len(new), 'COMPONENTS;ONLY NEW TEST METHOD;CURRENT NUMERICAL KERNEL UNCHANGED')
