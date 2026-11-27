"""ДЛЯ ТЕСТОВ. Сохраняет рабочий перенос ASD, нативные книги, отрицательные попытки и проверку сохранности без подмены полной приёмки."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import ctypes, json, tarfile, shutil

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
names = {
    1067: 'Cold1067WorkingASDNoObserver_20261127',
    1068: 'Cold1068WorkingASDObserverMutation_20261127',
    1069: 'Cold1069WorkingASDCapacityCompression_20261127',
    1070: 'Cold1070WorkingASDCapacityTension_20261127',
    1071: 'Cold1071WorkingASDOutput_20261127',
    1072: 'Cold1072WorkingASDDependenciesOutput_20261127',
    1073: 'Cold1073WorkingASDReopenCompile_20261127',
    1074: 'Cold1074WorkingASDPreservedOutput_20261127'}
roots = {key: base / value for key, value in names.items()}
out = repo / 'docs/reports/evidence/Coupled6DWorkingASDIntegration_2026-11-27'
assert not out.exists()
memory = ctypes.create_string_buffer(64)
ctypes.c_ulong.from_buffer(memory).value = 64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory, 16).value > 6500 * 1024**2
assert ctypes.c_ulonglong.from_buffer(memory, 32).value > 6 * 1024**3
checks = {
    1067: 'WorkingASDNoObserverProof.json', 1068: 'WorkingASDObserverMutationProof.json',
    1069: 'WorkingASDCapacityProof.json', 1070: 'WorkingASDCapacityProof.json',
    1074: 'WorkingASDPreservedOutputProof.json'}
proofs = {str(key): json.loads((roots[key] / file).read_bytes()) for key, file in checks.items()}
assert proofs['1067']['all6_no_observer_results_bitexact_verified_reference']
assert proofs['1068']['all6_mutating_observer_results_bitexact_verified_reference']
assert all(proofs[str(key)]['all_numerical_capacity_outputs_bitexact_reference'] for key in (1069, 1070))
assert proofs['1074']['all121_non_vba_payloads_bitexact']
publication = json.loads((roots[1074] / 'OutputPublication.json').read_bytes())
for name, digest in publication['after'].items():
    assert sha256((repo / 'workbook/output' / name).read_bytes()).hexdigest().upper() == digest
raw, pool = {}, {}
for key, root in roots.items():
    prep_path = root / 'Preparation.json'
    if prep_path.exists():
        prep = json.loads(prep_path.read_bytes())
        for name, digest in prep['inputs'].items():
            assert sha256((root / name).read_bytes()).hexdigest().upper() == digest, (key, name)
    for file in sorted(root.rglob('*')):
        if not file.is_file():
            continue
        name = file.relative_to(root).as_posix()
        if name.startswith('Temp/'):
            continue
        data = file.read_bytes()
        digest = sha256(data).hexdigest().upper()
        if digest not in pool:
            pool[digest] = data
        else:
            assert pool[digest] == data
        raw[f'{key}/{name}'] = pool[digest]
promotion = base / 'CleanASDWorkingPromotion_20261127'
for file in sorted(promotion.rglob('*')):
    if file.is_file():
        raw['Promotion/' + file.relative_to(promotion).as_posix()] = file.read_bytes()
out.mkdir(parents=True)
(out / '.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
archives = {}
for number, group in enumerate(((1067, 1068, 1069, 1070), (1071, 1072, 1073, 1074)), 1):
    selected = {name: data for name, data in raw.items() if name.split('/', 1)[0] in {str(x) for x in group} or (number == 1 and name.startswith('Promotion/'))}
    pins = {name: sha256(data).hexdigest().upper() for name, data in selected.items()}
    path = out / f'RawEvidence{number:02d}.tar.xz'
    seen_payloads = {}
    print('WRITING WORKING ASD INTEGRATION ARCHIVE', number, len(selected), flush=True)
    with tarfile.open(path, 'w:xz', preset=3) as archive:
        for name, data in selected.items():
            info = tarfile.TarInfo(name)
            digest = pins[name]
            if digest in seen_payloads:
                info.type, info.linkname = tarfile.LNKTYPE, seen_payloads[digest]
                archive.addfile(info)
            else:
                seen_payloads[digest] = name
                info.size = len(data)
                archive.addfile(info, BytesIO(data))
        data = (json.dumps(pins, sort_keys=True, indent=2) + '\n').encode('utf-8')
        info = tarfile.TarInfo('RawFiles.json'); info.size = len(data)
        archive.addfile(info, BytesIO(data))
    assert path.stat().st_size < 100 * 1024**2
    seen, payloads = set(), {}
    with tarfile.open(path, 'r|xz') as archive:
        for info in archive:
            if info.name == 'RawFiles.json':
                assert json.load(archive.extractfile(info)) == pins
                continue
            if info.islnk():
                data = payloads[info.linkname]
            else:
                assert info.isfile()
                data = archive.extractfile(info).read(); payloads[info.name] = data
            assert sha256(data).hexdigest().upper() == pins[info.name]
            seen.add(info.name)
    assert seen == set(selected)
    archives[path.name] = dict(bytes=path.stat().st_size, sha256=sha256(path.read_bytes()).hexdigest().upper(),
        logical_files=len(selected), all_objects_read_back_and_hashed=True)
helpers = set('''Record-TestWorkingASDIntegration.py Promote-TestASDNumericalCore.py
Stage-TestWorkingASDNoObserver.py Check-TestWorkingASDSource.py Check-TestWorkingASDNoObserver.py
Stage-TestWorkingASDObserverMutation.py Check-TestWorkingASDObserverSource.py
Stage-TestWorkingASDCapacity.py Check-TestWorkingASDCapacity.py Inspect-TestWorkingASDCapacity.py
Stage-TestWorkingASDOutput.py Inspect-TestPublicASDDependencies.py Stage-TestWorkingASDDependenciesOutput.py
Inspect-TestWorkingASDOutput.py Stage-TestWorkingASDPreservedOutput.py Test-ReopenPreservedASDOutput.ps1
Run-TestWorkingASDFoundations.ps1 Check-TestWorkingASDPreservedOutput.py Publish-TestWorkingASDOutput.py'''.split())
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
    shutil.copyfile(catalog[name], out / name)
summary = dict(created_local=datetime.now().astimezone().isoformat(), logical_files=len(raw), unique_payloads=len(pool),
    archives=archives, checks=proofs, publication=publication, failed_four_class_only_compile_retained=True,
    failed_compile_state_gate_retained=True, COM_save_package_changes_retained_and_rejected=True,
    actual_full_output_and_dump_updated=True, all121_non_vba_payloads_bitexact=True,
    no_full_acceptance_run=True, default_concrete_law_not_changed=True, Config_ASD_selector_UI_pending=True,
    native_direct_cases=12, native_capacity_cases=2, native_foundation_suites=7,
    no_new_global_OpenSees_equilibrium_solve_claimed=True, ten_seconds_whole_run_not_met=True,
    main_not_merged=True, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, sort_keys=True, indent=2) + '\n', encoding='utf-8')
files = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out / 'Files.json').write_text(json.dumps(files, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('WORKING ASD INTEGRATION ARCHIVED', json.dumps(dict(logical_files=len(raw), archives=archives, helpers=len(helpers)), sort_keys=True))
