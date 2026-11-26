"""ДЛЯ ТЕСТОВ. Сохраняет исходные книги, неудачи чистого N и независимое продолжение к пределу стали без потери байтов."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json, tarfile, shutil, ctypes

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {1039: base / 'Cold1039ASDPureNTension_20261127',
         1040: base / 'Cold1040ASDPureNCompression_20261127',
         1041: base / 'Cold1041ASDPureNHybrid_20261127'}
out = repo / 'docs/reports/evidence/Coupled6DPureAxialFirstControls_2026-11-27'
assert not out.exists()
memory = ctypes.create_string_buffer(64)
ctypes.c_ulong.from_buffer(memory).value = 64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory, 16).value > 6500 * 1024**2
assert ctypes.c_ulonglong.from_buffer(memory, 32).value > 6 * 1024**3
raw, execution, external, books = {}, {}, {}, {}
for case, root in roots.items():
    prep = json.loads((root / 'Preparation.json').read_bytes())
    observed = json.loads((root / 'MainCapacityExecutionProof.json').read_bytes())
    replay = json.loads((root / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    assert observed['native']['own_process_closed'] and observed['source_unchanged']
    assert replay['all56plus4_equations_and_six_resultants_verified']
    compiled = json.loads((root / 'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    source = json.loads((root / 'EarlySourceProof.json').read_bytes())
    assert source['saved_book_sha256'] == compiled['WorkbookSHA256']
    files = set(prep['inputs']) | {'Preparation.json', 'MainCapacityExecutionProof.json',
                                 'EarlySourceProof.json', 'docs/regression/CodeUpdate.json',
                                 'docs/regression/Diagnostic.xlsm'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file()
                 and p.suffix in ('.json', '.log', '.ndjson', '.ps1', '.py', '.txt'))
    files.update(p.relative_to(root).as_posix() for p in (root / 'Native').rglob('*') if p.is_file())
    if case == 1039:
        child = root / 'IndependentAxialTRFTestReplay'
        files.update(p.relative_to(root).as_posix() for p in child.rglob('*') if p.is_file())
        independent = json.loads((child / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
        assert independent['accepted_states'] == 43 and independent['reference_tolerance_violated_states'] == 0
    for name in sorted(files):
        data = (root / name).read_bytes()
        if name in prep['inputs']:
            assert sha256(data).hexdigest().upper() == prep['inputs'][name], (case, name)
        raw[f'{case}/{name}'] = data
    book_name = f'{case}/docs/regression/Diagnostic.xlsm'
    assert sha256(raw[book_name]).hexdigest().upper() == compiled['WorkbookSHA256']
    books[str(case)] = dict(member=book_name, sha256=compiled['WorkbookSHA256'], exact_saved_book_included=True)
    execution[str(case)] = {k: v for k, v in observed.items() if k not in ('trace', 'diagnostics')}
    external[str(case)] = {k: v for k, v in replay.items() if k not in ('records', 'observed_OpenSees_concrete_stresses')}
out.mkdir(parents=True)
(out / '.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
index = {name: sha256(data).hexdigest().upper() for name, data in raw.items()}
archive_file = out / 'RawEvidence.tar.xz'
canonical = {}
with tarfile.open(archive_file, 'w:xz', preset=3) as handle:
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
assert archive_file.stat().st_size < 100 * 1024**2
seen, canonical_bytes = {}, {}
with tarfile.open(archive_file, 'r|xz') as handle:
    for info in handle:
        if info.name == 'RawFiles.json':
            assert json.load(handle.extractfile(info)) == index
            continue
        if info.islnk():
            data = canonical_bytes[info.linkname]
        else:
            assert info.isfile()
            data = handle.extractfile(info).read()
            canonical_bytes[info.name] = data
        assert sha256(data).hexdigest().upper() == index[info.name], info.name
        seen[info.name] = True
assert set(seen) == set(raw)
helpers = set('''Record-TestASDPureAxialEvidence.py Stage-TestASDPureAxialCapacity.py
Solve-TestPureAxialOpenSees.py Stage-TestAxialIndependentReplay.py Stage-TestASDAxialHybrid.py
Check-TestASDAxialHybridSource.py Stage-TestASDHybridRecovery.py Check-TestASDHybridSource.py
Check-TestASDBendingMainCapacity.py Replay-TestMainCapacityEquilibrium.py'''.split())
while True:
    expanded = set(helpers)
    for name in helpers:
        text = (here / name).read_text('utf-8-sig')
        expanded.update(p.name for p in here.iterdir() if p.is_file() and p.suffix in ('.py', '.ps1', '.txt') and p.name in text)
    if expanded == helpers:
        break
    helpers = expanded
for name in sorted(helpers):
    shutil.copyfile(here / name, out / name)
driver = json.loads(raw['1039/Native/IndependentAxialTRF.json'])
assert sha256((out / 'Solve-TestPureAxialOpenSees.py').read_bytes()).hexdigest().upper() == driver['source_helper_sha256']
summary = dict(created_local=datetime.now().astimezone().isoformat(), logical_files=len(raw),
               unique_payloads=len(canonical), archive_bytes=archive_file.stat().st_size,
               all_raw_files_read_back_and_hashed=True, all_three_actual_saved_books_included=True,
               books=books, native_execution=execution, native_equilibrium_replay=external,
               independent_axial_continuation=dict(inherited_states=22, new_states=21,
                   complete_path_replayed_states=43, final_N_N=driver['records'][-1]['force_N'],
                   final_steel_utilization=driver['records'][-1]['steel_utilization'],
                   elapsed_seconds=driver['elapsed_seconds'], original_tolerances_unchanged=True,
                   actual_OpenSees_material=True, independent_SciPy_global_driver=True,
                   independent_OpenSees_structural_model=False),
               next_stage='Pure N; budget192 test1042 is not part of this closed packet',
               working_solver_not_changed=True, physical_capacity_certified=False,
               full_acceptance_not_run=True, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, sort_keys=True, indent=2) + '\n', encoding='utf-8')
pins = {p.name: sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PURE N FIRST CONTROLS SAVED', len(raw), 'RAW FILES;', archive_file.stat().st_size, 'BYTES;43 EXTERNAL STATES;NO GLOBAL PASS')
