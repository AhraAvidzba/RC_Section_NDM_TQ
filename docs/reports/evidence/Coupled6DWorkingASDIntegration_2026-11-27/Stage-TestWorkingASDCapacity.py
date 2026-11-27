"""ДЛЯ ТЕСТОВ. Повторяет два ранее проверенных осевых предела на точных рабочих классах без наблюдателя."""
from pathlib import Path
from hashlib import sha256
import json, shutil

repo = Path.cwd()
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1067WorkingASDNoObserver_20261127'
proof = json.loads((parent / 'WorkingASDNoObserverProof.json').read_bytes())
assert proof['all6_no_observer_results_bitexact_verified_reference']
book = (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256(book).hexdigest().upper() == compiled['WorkbookSHA256']
for key, name, reference in [
    ('Compression', 'Cold1069WorkingASDCapacityCompression_20261127', 'Cold1056ASDPureNFullCompressionBranch_20261127'),
    ('Tension', 'Cold1070WorkingASDCapacityTension_20261127', 'Cold1049ASDPureNSeedHighLoad_20261127')
]:
    root, original = base / name, base / reference
    assert not root.exists()
    for directory in ('src', 'tests', 'tools', 'workbook'):
        shutil.copytree(parent / directory, root / directory)
    shutil.copytree(original / 'Input', root / 'Input')
    for directory in ('Native', 'Temp', 'docs/regression'):
        (root / directory).mkdir(parents=True, exist_ok=True)
    for file in ('Guard-IsolatedOffice.ps1', 'Run-TestASDSpatialMainCapacity.ps1'):
        shutil.copyfile(original / file, root / file)
    (root / 'docs/regression/Diagnostic.xlsm').write_bytes(book)
    shutil.copyfile(parent / 'docs/regression/CodeUpdate.json', root / 'docs/regression/CodeUpdate.json')
    shutil.copyfile(parent / 'WorkingASDSourceProof.json', root / 'WorkingASDSourceProof.json')
    (root / 'WorkingCapacityReference.json').write_text(json.dumps(dict(reference=str(original), parent=str(parent),
        case=key, same_compiled_working_classes=True, no_observer=True, same_input=True), indent=2) + '\n', encoding='utf-8')
    shutil.copyfile(Path(__file__), root / Path(__file__).name)
    prep = dict(parent=str(parent), reference=str(original), scope='Directed working ASD capacity regression; ' + key,
        inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                for p in root.rglob('*') if p.is_file() and p.relative_to(root).as_posix()
                not in ('docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json')})
    (root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
    print('WORKING ASD CAPACITY PREPARED', key, root)
