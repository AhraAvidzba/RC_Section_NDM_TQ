"""ДЛЯ ТЕСТОВ. Показывает разницу фактического VBA выдаваемой и проверенной частной книги, не меняя книги или исходники."""
from pathlib import Path
from runpy import run_path
from zipfile import ZipFile
from hashlib import sha256
import json, re, difflib

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
helper, body, reader = run_path(str(here / 'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
normal = run_path(str(here / 'Canonical-TestCoreVBA.py'))['TestCanonical']


def TestRead(path):
    """Читает исходники из реального бинарного проекта VBA закреплённым декодером."""
    with ZipFile(path) as archive:
        return reader(helper['TestContainer'](archive.read('xl/vbaProject.bin')))[1]


def TestIdentifiers(text):
    """Допускает только регистр идентификаторов VBE, сохраняя строки и комментарии."""
    return re.sub(r'"(?:[^"\n]|"")*"|\'[^\n]*|[A-Za-z_][\w]*',
        lambda m: m.group() if m.group().startswith(('"', "'")) else m.group().lower(), text)


public_path = repo / 'workbook/output/RC_Section_NDM.xlsm'
private_path = base / 'Cold1067WorkingASDNoObserver_20261127/docs/regression/Diagnostic.xlsm'
public, private = TestRead(public_path), TestRead(private_path)
changed = [name for name in public.keys() & private.keys() if TestIdentifiers(public[name]) != TestIdentifiers(private[name])]
sources = [p for folder in ('src', 'tests') for p in (repo / folder).rglob('*') if p.suffix in ('.cls', '.bas')]
byname = {p.stem: p for p in sources}
assert len(sources) == len(byname)
records = []
for name in sorted(set(changed) | (private.keys() - public.keys())):
    path = byname.get(name)
    match = False
    if path:
        match = normal(body(path.read_text('utf-8-sig'), name)[0])[0] == normal(body(private[name], name)[0])[0]
    records.append(dict(name=name, source=path.relative_to(repo).as_posix() if path else None,
        private_exact_current_source=match, added=name not in public))
    if path and not match:
        expected, actual = normal(body(path.read_text('utf-8-sig'), name)[0])[0], normal(body(private[name], name)[0])[0]
        print('PRIVATE_VS_SOURCE', name, '\n'.join(list(difflib.unified_diff(expected.splitlines(), actual.splitlines(), lineterm=''))[:45]))
all_records = []
for name, path in sorted(byname.items()):
    source_text = path.read_text('utf-8-sig')
    if name == 'modSectionRecipeApiGuide':
        continue
    if 'Option Explicit' not in source_text:
        assert name == 'modSectionRecipeApiGuide'
        continue
    try:
        expected = normal(body(source_text, name)[0])[0]
    except IndexError:
        print('SOURCE_BODY_UNAVAILABLE', name, path)
        raise
    matches = {}
    for label, data in [('public', public), ('private', private)]:
        matches[label] = name in data and expected == normal(body(data[name], name)[0])[0]
    all_records.append(dict(name=name, source=path.relative_to(repo).as_posix(), **matches))
proof = dict(public_sha256=sha256(public_path.read_bytes()).hexdigest().upper(),
    public_count=len(public), private_count=len(private), public_only=sorted(public.keys() - private.keys()),
    private_only=sorted(private.keys() - public.keys()), changed_or_added=records, all_current_sources=all_records)
(base / 'CleanASDWorkingPromotion_20261127/PublicDependencyInspection.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('COMPONENT_COUNTS', len(public), len(private), len(sources))
print('CURRENT_SOURCES_NOT_MATCHING_PRIVATE', [item['name'] for item in all_records if not item['private']])
print('CURRENT_SOURCES_NOT_MATCHING_PUBLIC', [item['name'] for item in all_records if not item['public']])
