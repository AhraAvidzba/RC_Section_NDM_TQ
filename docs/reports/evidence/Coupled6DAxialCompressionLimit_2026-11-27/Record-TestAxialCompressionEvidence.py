"""ДЛЯ ТЕСТОВ. Сохраняет точные книги, отрицательные контроли и независимые осевые ветви с проверкой каждого байтового объекта архива."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from io import BytesIO
from datetime import datetime
import json, tarfile, shutil, ctypes

repo = Path.cwd()
here = Path(__file__).resolve().parent
context = run_path(str(here / 'Check-TestAxialCompressionEvidence.py'))
roots = context['roots']
native_cases = context['native_cases']
proof = context['proof']
out = repo / 'docs/reports/evidence/Coupled6DAxialCompressionLimit_2026-11-27'
assert not out.exists()
memory = ctypes.create_string_buffer(64)
ctypes.c_ulong.from_buffer(memory).value = 64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory, 16).value > 6500 * 1024**2
assert ctypes.c_ulonglong.from_buffer(memory, 32).value > 6 * 1024**3
raw, books, executions = {}, {}, {}
for case, root in roots.items():
    prep = json.loads((root / 'Preparation.json').read_bytes())
    files = set(prep['inputs']) | {'Preparation.json', 'EarlySourceProof.json', 'docs/regression/Diagnostic.xlsm'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file())
    files.update(p.relative_to(root).as_posix() for p in (root / 'Native').rglob('*') if p.is_file())
    if case in native_cases:
        files.update(('MainCapacityExecutionProof.json', 'docs/regression/CodeUpdate.json'))
        executions[str(case)] = {k: v for k, v in context['observed'][case].items() if k not in ('trace', 'diagnostics')}
    for child_name in ('FullCompressionBranchTestReplay', 'IndependentCompressionPeakTestReplay',
                       'IndependentCompressionFineTestReplay'):
        child = root / child_name
        if child.exists():
            files.update(p.relative_to(root).as_posix() for p in child.rglob('*') if p.is_file())
    for name in sorted(files):
        data = (root / name).read_bytes()
        if name in prep['inputs']:
            assert sha256(data).hexdigest().upper() == prep['inputs'][name], (case, name)
        raw[f'{case}/{name}'] = data
    book_name = f'{case}/docs/regression/Diagnostic.xlsm'
    digest = sha256(raw[book_name]).hexdigest().upper()
    assert digest == json.loads(raw[f'{case}/EarlySourceProof.json'])['saved_book_sha256']
    books[str(case)] = dict(member=book_name, sha256=digest,
                           actual_native_execution=case in native_cases,
                           exact_saved_book_included=True)
out.mkdir(parents=True)
(out / '.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
index = {name: sha256(data).hexdigest().upper() for name, data in raw.items()}
archive = out / 'RawEvidence.tar.xz'
canonical = {}
print('COMPRESSION ARCHIVE WRITING', len(raw), 'LOGICAL FILES', flush=True)
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
            assert info.isfile()
            data = handle.extractfile(info).read()
            saved[info.name] = data
        assert sha256(data).hexdigest().upper() == index[info.name], info.name
        seen.add(info.name)
assert seen == set(raw)
helpers = set('''Record-TestAxialCompressionEvidence.py Check-TestAxialCompressionEvidence.py
Stage-TestASDAxialController.py Check-TestASDAxialControllerSource.py
Patch-TestPeakPrecision.py Patch-TestFullCompressionBranch.py
Stage-TestIndependentCompressionPeak.py Stage-TestIndependentCompressionFine.py
Solve-TestCompressionPeakOpenSees.py Solve-TestPureAxialOpenSees.py
Stage-TestCompressionIndependentReplay.py Stage-TestAxialCompressionReplays.py
Stage-TestAxialCompressionOverload.py Replay-TestMainCapacityEquilibrium.py
Check-TestASDBendingMainCapacity.py'''.split())
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
for case in (1054, 1057):
    driver = json.loads(raw[f'{case}/Native/IndependentCompressionPeakTRF.json'])
    assert sha256((out / 'Solve-TestCompressionPeakOpenSees.py').read_bytes()).hexdigest().upper() == driver['source_helper_sha256']
    assert sha256((out / 'Solve-TestPureAxialOpenSees.py').read_bytes()).hexdigest().upper() == driver['common_driver_sha256']
summary = dict(created_local=datetime.now().astimezone().isoformat(), logical_files=len(raw),
               unique_payloads=len(canonical), archive_bytes=archive.stat().st_size,
               all_raw_files_read_back_and_hashed=True, books=books,
               all_eight_saved_books_included=True, six_native_executions=executions,
               compression_limit_proof=proof,
               test_harness_corrections=[
                   'First observer audit passed an already header-stripped module into a header parser; corrected to use the same literal-safe body fold.',
                   'First compression certificate read published PeakAccuracy for a rejected peak; that output is zero by contract. The rejected estimate is checked directly in the observed completion CSV.'
               ],
               output_book_sha256=context['TestDigest'](repo / 'workbook/output/RC_Section_NDM.xlsm'),
               output_VBA_dump_sha256=context['TestDigest'](repo / 'workbook/output/VBA_All_Code.txt'),
               working_solver_not_changed=True, pure_N_stage_complete=False,
               full_acceptance_not_run=True, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, sort_keys=True, indent=2) + '\n', encoding='utf-8')
pins = {p.name: sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('COMPRESSION LIMIT EVIDENCE SAVED', len(raw), 'RAW FILES', archive.stat().st_size,
      'BYTES;ALL READ BACK;6 NATIVE RUNS;2 INDEPENDENT CURVES;NO GLOBAL PASS', flush=True)
