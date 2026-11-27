"""ДЛЯ ТЕСТОВ. Сохраняет все исходные части пользовательской книги и заменяет только реально скомпилированный VBA-проект."""
from pathlib import Path
from hashlib import sha256
from zipfile import ZipFile
import json, shutil

repo = Path.cwd()
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1073WorkingASDReopenCompile_20261127'
root = base / 'Cold1074WorkingASDPreservedOutput_20261127'
assert not root.exists()
compiled = json.loads((parent / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
reopened = json.loads((parent / 'docs/regression/OutputReopenProof.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and reopened['Compiled'] and reopened['SavedReopened']
assert json.loads((parent / 'CompilePointClosure.json').read_text('utf-8-sig'))['State'] == 'Absent'
source = repo / 'workbook/output/RC_Section_NDM.xlsm'
assert source.read_bytes() == (parent / 'Before/RC_Section_NDM.xlsm').read_bytes()
for directory in ('Native', 'Temp', 'docs/regression', 'Before', 'workbook/output'):
    (root / directory).mkdir(parents=True, exist_ok=True)
shutil.copytree(parent / 'tools', root / 'tools')
shutil.copyfile(parent / 'Guard-IsolatedOffice.ps1', root / 'Guard-IsolatedOffice.ps1')
shutil.copytree(parent / 'Before', root / 'Before', dirs_exist_ok=True)
shutil.copyfile(source, root / 'workbook/output/RC_Section_NDM.xlsm')
candidate = root / 'docs/regression/Diagnostic.xlsm'
with ZipFile(parent / 'docs/regression/Diagnostic.xlsm') as compiled_zip:
    vba = compiled_zip.read('xl/vbaProject.bin')
with ZipFile(source) as original, ZipFile(candidate, 'x') as target:
    target.comment = original.comment
    for item in original.infolist():
        target.writestr(item, vba if item.filename == 'xl/vbaProject.bin' else original.read(item.filename))
with ZipFile(source) as original, ZipFile(candidate) as target:
    assert original.namelist() == target.namelist() and original.comment == target.comment
    pins = {name: sha256(original.read(name)).hexdigest().upper() for name in original.namelist() if name != 'xl/vbaProject.bin'}
    for name, digest in pins.items():
        assert sha256(target.read(name)).hexdigest().upper() == digest, name
shutil.copyfile(Path(__file__), root / Path(__file__).name)
proof = dict(parent=str(parent), source_public_book_sha256=sha256(source.read_bytes()).hexdigest().upper(),
    saved_candidate_sha256=sha256(candidate.read_bytes()).hexdigest().upper(),
    actual_compiled_vba_sha256=sha256(vba).hexdigest().upper(), non_vba_parts=pins,
    all_non_vba_payloads_bitexact=True, exact_member_order_and_zip_comment=True,
    native_reopen_and_tests_pending=True, public_output_unchanged=True)
(root / 'PreservedPackageProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PRESERVED VBA-ONLY CANDIDATE CREATED;NON-VBA PARTS EXACT', len(pins), ';NATIVE REOPEN PENDING')
