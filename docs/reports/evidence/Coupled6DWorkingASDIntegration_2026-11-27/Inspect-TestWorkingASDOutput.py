"""ДЛЯ ТЕСТОВ. Проверяет фактическую VBA-сборку, повторное открытие, полный дамп и сохранность пакета пользовательской книги."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from zipfile import ZipFile
from xml.etree import ElementTree as ET
from argparse import ArgumentParser
import json, re, difflib

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, default=base / 'Cold1073WorkingASDReopenCompile_20261127')
root = parser.parse_args().root
book = root / 'docs/regression/Diagnostic.xlsm'
compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
prep = json.loads((root / 'Preparation.json').read_bytes())
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert set(compiled['Modules']) == set(prep['imports']) and len(compiled['Modules']) == 59
assert json.loads((root / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
reopened = json.loads((root / 'docs/regression/OutputReopenProof.json').read_text('utf-8-sig'))
assert reopened['SavedReopened'] and reopened['ReadOnly'] and reopened['Compiled'] and reopened['Components'] == 255
assert sha256(book.read_bytes()).hexdigest().upper() == compiled['WorkbookSHA256']
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']


def TestRead(path):
    """Читает исходный текст каждого реально сохранённого компонента VBA."""
    with ZipFile(path) as archive:
        return reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]


def TestCode(text):
    """Исключает скрытые атрибуты, не являющиеся строками CodeModule; числовые значения не округляет."""
    return re.sub(r'(?m)^Attribute [^\r\n]*(?:\r?\n|$)', '', text).replace('\r\n', '\n').strip('\r\n')


before = TestRead(root / 'Before/RC_Section_NDM.xlsm')
actual = TestRead(book)
reference = TestRead(base / 'Cold1067WorkingASDNoObserver_20261127/docs/regression/Diagnostic.xlsm')
assert len(before) == 226 and len(actual) == 255 and len(reference) == 256
assert set(actual) == set(reference) - {'modTestSelectableMaterialCore'}
source_checks, differences = [], []
for name in sorted(actual):
    expected = reference[name]
    if TestCode(actual[name]) != TestCode(expected):
        try:
            match = normal(body(actual[name], name)[0])[0] == normal(body(expected, name)[0])[0]
        except (IndexError, AssertionError):
            match = False
        if not match:
            differences.append(dict(name=name, diff=list(difflib.unified_diff(TestCode(expected).splitlines(), TestCode(actual[name]).splitlines(), lineterm=''))[:28]))
    if name in prep['imports']:
        source = [p for folder in ('src', 'tests') for p in (root / folder).rglob(name + '.*') if p.suffix in ('.bas', '.cls')]
        assert len(source) == 1
        assert source[0].read_bytes() == (repo / source[0].relative_to(root)).read_bytes()
        match = normal(body(source[0].read_text('utf-8-sig'), name)[0])[0] == normal(body(actual[name], name)[0])[0]
        source_checks.append(dict(name=name, exact_current_source=match))

dump = (root / 'docs/regression/VBA_All_Code.txt').read_text('utf-8-sig').splitlines()
dump_components, index, dump_differences = {}, 0, []
while index < len(dump):
    if not dump[index].startswith('COMPONENT: '):
        index += 1
        continue
    name = dump[index][11:]
    assert name not in dump_components
    count = int(dump[index + 2][7:])
    text = '\n'.join(dump[index + 4:index + 4 + count])
    dump_components[name] = text
    if TestCode(actual[name]) != text.strip('\r\n'):
        dump_differences.append(dict(name=name, diff=list(difflib.unified_diff(TestCode(actual[name]).splitlines(), text.splitlines(), lineterm=''))[:20]))
    index += 4 + count
assert set(dump_components) == set(actual)

with ZipFile(root / 'Before/RC_Section_NDM.xlsm') as old_zip, ZipFile(book) as new_zip:
    package_removed = sorted(set(old_zip.namelist()) - set(new_zip.namelist()))
    package_added = sorted(set(new_zip.namelist()) - set(old_zip.namelist()))
    package_differences = []
    for name in old_zip.namelist():
        if name not in new_zip.namelist():
            continue
        old, new = old_zip.read(name), new_zip.read(name)
        if old != new and name != 'xl/vbaProject.bin':
            item = dict(name=name, old_bytes=len(old), new_bytes=len(new))
            if name.endswith(('.xml', '.rels')):
                a, b = ET.canonicalize(old), ET.canonicalize(new)
                item['canonical_equal'] = a == b
                if a != b:
                    item['diff'] = list(difflib.unified_diff(a.replace('><', '>\n<').splitlines(), b.replace('><', '>\n<').splitlines(), lineterm=''))[:35]
            package_differences.append(item)
proof = dict(source_public_book_sha256=prep['source_public_book_sha256'], saved_candidate_sha256=compiled['WorkbookSHA256'],
    reopened=reopened, actual_components=len(actual), source_checks=source_checks,
    numerical_reference_differences=differences, dump_differences=dump_differences,
    package_differences=package_differences, package_removed=package_removed, package_added=package_added,
    full_acceptance_not_run=True, publication_not_authorized_by_this_inspection=True)
(root / 'OutputInspection.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print(json.dumps(dict(components=len(actual), source_mismatch=[x['name'] for x in source_checks if not x['exact_current_source']],
    numerical_reference_differences=differences, dump_differences=dump_differences,
    package_removed=package_removed, package_added=package_added,
    package_differences=[dict(name=x['name'], canonical_equal=x.get('canonical_equal')) for x in package_differences]), sort_keys=True))
