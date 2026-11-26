"""ДЛЯ ТЕСТОВ. Закрепляет реальные осевые расчёты плоской НДМ, диаграммы, исходные книги и проверку всех байтов архива."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json, tarfile, shutil, ctypes

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {1059: base / 'Cold1059PureAxialPlanarTension_20261127',
         1060: base / 'Cold1060PureAxialPlanarCompression_20261127'}
out = repo / 'docs/reports/evidence/Coupled6DPlanarAxialReference_2026-11-27'
assert not out.exists()
memory = ctypes.create_string_buffer(64)
ctypes.c_ulong.from_buffer(memory).value = 64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory, 16).value > 6500 * 1024**2
assert ctypes.c_ulonglong.from_buffer(memory, 32).value > 6 * 1024**3
raw, proofs, books = {}, {}, {}
for case, root in roots.items():
    proof = json.loads((root / 'PureAxialPlanarReferenceProof.json').read_bytes())
    assert proof['native']['own_process_closed'] and proof['planar_reference_physical_limit_verified']
    assert not proof['artificial_diagram_extension'] and proof['replacement_concrete_subtracted_at_bars']
    prep = json.loads((root / 'Preparation.json').read_bytes())
    files = set(prep['inputs']) | {'Preparation.json', 'docs/regression/Diagnostic.xlsm', 'docs/regression/CodeUpdate.json'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file())
    files.update(p.relative_to(root).as_posix() for p in (root / 'Native').rglob('*') if p.is_file())
    for name in sorted(files):
        data = (root / name).read_bytes()
        if name in prep['inputs']:
            assert sha256(data).hexdigest().upper() == prep['inputs'][name], (case, name)
        raw[f'{case}/{name}'] = data
    digest = sha256(raw[f'{case}/docs/regression/Diagnostic.xlsm']).hexdigest().upper()
    assert digest == json.loads((root / 'PlanarAxialSourceProof.json').read_bytes())['saved_book_sha256']
    books[str(case)] = dict(member=f'{case}/docs/regression/Diagnostic.xlsm', sha256=digest)
    proofs[str(case)] = proof
assert books['1059']['sha256'] == books['1060']['sha256']
out.mkdir(parents=True)
(out / '.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
index = {name: sha256(data).hexdigest().upper() for name, data in raw.items()}
canonical = {}
archive = out / 'RawEvidence.tar.xz'
with tarfile.open(archive, 'w:xz', preset=3) as handle:
    for name, data in raw.items():
        info = tarfile.TarInfo(name)
        digest = index[name]
        if digest in canonical:
            info.type = tarfile.LNKTYPE
            info.linkname = canonical[digest]
            handle.addfile(info)
        else:
            canonical[digest] = name
            info.size = len(data)
            handle.addfile(info, BytesIO(data))
    data = (json.dumps(index, sort_keys=True, indent=2) + '\n').encode('utf-8')
    info = tarfile.TarInfo('RawFiles.json')
    info.size = len(data)
    handle.addfile(info, BytesIO(data))
assert archive.stat().st_size < 100 * 1024**2
seen, saved = set(), {}
with tarfile.open(archive, 'r|xz') as handle:
    for info in handle:
        if info.name == 'RawFiles.json':
            assert json.load(handle.extractfile(info)) == index
            continue
        if info.islnk():
            data = saved[info.linkname]
        else:
            data = handle.extractfile(info).read()
            saved[info.name] = data
        assert sha256(data).hexdigest().upper() == index[info.name]
        seen.add(info.name)
assert seen == set(raw)
helpers = set('''Record-TestPureAxialPlanarEvidence.py Stage-TestPureAxialPlanarReference.py
Stage-TestPureAxialPlanarCompression.py Check-TestPureAxialPlanarReference.py
Check-TestPureAxialPlanarSource.py TestPureBendingPlanarReference.bas.txt
Compile-TestExternalPoint.ps1 Check-TestPersistentHEvidence.py
Inspect-TestResolutionVBA.py Canonical-TestCoreVBA.py Read-TestColdRawCapture.py'''.split())
catalog = {p.name: p for p in here.iterdir() if p.is_file() and p.suffix in ('.py', '.ps1', '.txt')}
while True:
    expanded = set(helpers)
    for name in helpers:
        text = catalog[name].read_text('utf-8-sig')
        expanded.update(candidate for candidate in catalog if candidate in text)
    if expanded == helpers:
        break
    helpers = expanded
for name in sorted(helpers):
    shutil.copyfile(here / name, out / name)
summary = dict(created_local=datetime.now().astimezone().isoformat(), logical_files=len(raw),
               unique_payloads=len(canonical), archive_bytes=archive.stat().st_size,
               all_raw_files_read_back_and_hashed=True, actual_saved_books=books, proofs=proofs,
               actual9_planar_classes_exact_current_output=True, no_numerical_class_changed=True,
               same_compiled_book_for_both_N_signs=True,
               source_audit_test_corrections=[
                   'Source copies contain only previously imported modules; the numerical kernel was therefore compared in actual saved books.',
                   'VBE displays two finite-number guard literals with fewer decimal digits; exact comparison used the same actual literals from the current output workbook, without rounding scientific inputs.',
                   'Unchanged empty worksheet modules require complete-source comparison rather than the Option Explicit body parser.',
                   'Expected insertion retained the exact third newline from the method template.'
               ],
               working_solver_not_changed=True, CDP_N_and_direct_state_checks_pending=True,
               full_acceptance_not_run=True, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, sort_keys=True, indent=2) + '\n', encoding='utf-8')
pins = {p.name: sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PLANAR AXIAL EVIDENCE VERIFIED', len(raw), 'FILES', archive.stat().st_size, 'BYTES;BOTH PHYSICAL LIMITS;WORKING BOOK UNCHANGED')
