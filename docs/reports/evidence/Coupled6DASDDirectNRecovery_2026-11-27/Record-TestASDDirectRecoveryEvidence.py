"""ДЛЯ ТЕСТОВ. Сохраняет реальные книги, отрицательные контроли и холодный повтор восстановления прямого ASD с чтением каждого объекта архива."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json, tarfile, shutil, ctypes

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
names = {1061: 'Cold1061ASDPureNDirectBaseline_20261127',
         1062: 'Cold1062ASDPureNDirectRecovery_20261127',
         1063: 'Cold1063ASDPureNDirectObservation_20261127',
         1064: 'Cold1064ASDPureNDirectWarmTrial_20261127',
         1065: 'Cold1065ASDPureNDirectEarlyPrefix_20261127',
         1066: 'Cold1066ASDPureNDirectColdRepeat_20261127'}
roots = {case: base / name for case, name in names.items()}
out = repo / 'docs/reports/evidence/Coupled6DASDDirectNRecovery_2026-11-27'
assert not out.exists()
memory = ctypes.create_string_buffer(64)
ctypes.c_ulong.from_buffer(memory).value = 64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory, 16).value > 6500 * 1024**2
assert ctypes.c_ulonglong.from_buffer(memory, 32).value > 6 * 1024**3
raw, pool, books, proofs = {}, {}, {}, {}
for case, root in roots.items():
    proof = json.loads((root / 'ASDDirectPureNExecutionProof.json').read_bytes())
    assert proof['native'] == dict(passed=6, failed=0, macros=1, timeout600=False, own_process_closed=True)
    assert proof['all56plus4_equations_and_six_resultants_verified']
    prep = json.loads((root / 'Preparation.json').read_bytes())
    source = json.loads((root / 'EarlySourceProof.json').read_bytes())
    assert source['changed_components_exact'] and source['full_equilibrium_acceptance_unchanged']
    files = set(prep['inputs']) | {'Preparation.json', 'EarlySourceProof.json',
        'ASDDirectPureNExecutionProof.json', 'docs/regression/CodeUpdate.json', 'docs/regression/Diagnostic.xlsm'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file())
    files.update(p.relative_to(root).as_posix() for p in (root / 'Native').rglob('*') if p.is_file())
    for index in range(1, 7):
        child = root / f'Case{index:02d}'
        files.update(p.relative_to(root).as_posix() for p in child.rglob('*') if p.is_file())
        replay = json.loads((child / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
        assert replay['all56plus4_equations_and_six_resultants_verified'] and replay['reference_tolerance_violated_states'] == 0
        assert replay['input_history_sha256'] == sha256((child / 'Native/MainCapacityHistory.bin').read_bytes()).hexdigest().upper()
        if case >= 1063 and index in (2, 3, 4, 6):
            full = json.loads((child / 'FullRecoveryTestReplay/Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
            assert full['all56plus4_equations_and_six_resultants_verified'] and full['reference_tolerance_violated_states'] == 0
    for name in sorted(files):
        data = (root / name).read_bytes()
        digest = sha256(data).hexdigest().upper()
        if name in prep['inputs']:
            assert digest == prep['inputs'][name], (case, name)
        if digest in pool:
            assert pool[digest] == data
        else:
            pool[digest] = data
        raw[f'{case}/{name}'] = pool[digest]
    member = f'{case}/docs/regression/Diagnostic.xlsm'
    digest = sha256(raw[member]).hexdigest().upper()
    assert digest == source['saved_book_sha256']
    books[str(case)] = dict(member=member, sha256=digest, exact_saved_book_included=True)
    proofs[str(case)] = proof
expected = {1061: ['OK', 'NumFail', 'NumFail', 'NumFail', 'OK', 'NumFail'],
            1062: ['OK', 'NumFail', 'NumFail', 'FAIL', 'OK', 'FAIL'],
            1063: ['OK', 'NumFail', 'NumFail', 'FAIL', 'OK', 'FAIL'],
            1064: ['OK', 'OK', 'NumFail', 'FAIL', 'OK', 'FAIL'],
            1065: ['OK', 'OK', 'OK', 'FAIL', 'OK', 'FAIL'],
            1066: ['OK', 'OK', 'OK', 'FAIL', 'OK', 'FAIL']}
for case in roots:
    assert [record['status'] for record in proofs[str(case)]['cases']] == expected[case]
for index in range(1, 7):
    name = f'Case{index:02d}/Native/MainCapacityHistory.bin'
    assert raw['1062/' + name] == raw['1063/' + name]
    assert raw['1065/' + name] == raw['1066/' + name]
for index in (1, 5):
    name = f'Case{index:02d}/Native/MainCapacityHistory.bin'
    assert len({sha256(raw[f'{case}/' + name]).digest() for case in roots}) == 1
cold = json.loads(raw['1066/ColdDirectRepeatProof.json'])
assert cold['all6_cold_histories_bitexact'] and cold['all6_expected_direct_states_verified']
out.mkdir(parents=True)
(out / '.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
archives = {}
for number, group in enumerate(((1061, 1062, 1063), (1064, 1065, 1066)), 1):
    selected = {name: data for name, data in raw.items() if int(name.split('/', 1)[0]) in group}
    index = {name: sha256(data).hexdigest().upper() for name, data in selected.items()}
    filename = f'RawEvidence{number:02d}.tar.xz'
    archive = out / filename
    canonical = {}
    print('WRITING DIRECT ASD ARCHIVE', number, len(selected), 'LOGICAL FILES', flush=True)
    with tarfile.open(archive, 'w:xz', preset=3) as handle:
        for name, data in selected.items():
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
    assert seen == set(selected)
    archives[filename] = dict(cases=list(group), bytes=archive.stat().st_size,
        sha256=sha256(archive.read_bytes()).hexdigest().upper(), logical_files=len(selected),
        all_objects_read_back_and_hashed=True)
helpers = set('''Record-TestASDDirectRecoveryEvidence.py TestASDDirectPureN.bas.txt
ASDDirectRecoveryMethod.bas.txt Stage-TestASDDirectPureN.py Check-TestASDDirectPureNSource.py
Stage-TestASDDirectRecovery.py Patch-TestASDDirectRecovery.py Check-TestASDDirectRecoverySource.py
Stage-TestDirectRecoveryObservation.py Patch-TestDirectRecoveryObservation.py Check-TestDirectRecoveryObservationSource.py
Stage-TestDirectWarmTrial.py Patch-TestDirectWarmTrial.py Check-TestDirectWarmTrialSource.py
Stage-TestDirectRecoveryPrefix.py Patch-TestDirectRecoveryPrefix.py Check-TestDirectRecoveryPrefixSource.py
Stage-TestDirectRecoveryColdRepeat.py Stage-TestDirectFullRecoveryReplay.py
Check-TestASDDirectPureNEvidence.py Replay-TestMainCapacityEquilibrium.py'''.split())
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
    unique_payloads=len(pool), archives=archives, actual_saved_books=books, native_jobs=6,
    native_macro_invocations=6, independently_started_cases=36, execution_proofs=proofs,
    all_objects_read_back_and_hashed=True, all6_cold_histories_bitexact=True,
    observation_all6_histories_bitexact=True, successful_elastic_cases_unchanged=True,
    original_materials_and_equilibrium_tolerances_unchanged=True,
    last_revision_all6_expected_direct_statuses_verified=True,
    scope='Pure N ASD,one section and chosen materials;working output not updated',
    independent_OpenSees_material_replay=True, independent_OpenSees_structural_analysis=False,
    public_solver_not_updated=True, public_workbook_not_updated=True,
    full_user_run_10second_target_not_proved=True, full_acceptance_not_run=True,
    source_harness_corrections=[
        'Explicit zero-length vectors belong only to absent returned NDS,not to the nonempty history-vector contract.',
        'Compiler imports dependencies before their consumer; source audit checks the actual expected module order.',
        'Method insertion comparison preserves the exact blank line added by the prepared source template.'
    ], whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, sort_keys=True, indent=2) + '\n', encoding='utf-8')
pins = {p.name: sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('DIRECT ASD RECOVERY EVIDENCE SAVED', len(raw), 'FILES;2 ARCHIVES;6 EXACT BOOKS;36 CASES;NO GLOBAL PASS', flush=True)
