"""ДЛЯ ТЕСТОВ. Сохраняет971–972 с точной ссылкой на уже опубликованную книгу968 без повторного хранения её байтов."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from runpy import run_path
from datetime import datetime
import json,tarfile,re,ctypes

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={971:base/'Cold971MaterialPrefixCoordinate_20261126',972:base/'Cold972ActiveUnloadingCoordinate_20261126',973:base/'Cold973OutgoingMaterialPair_20261126'}
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs'] and state['goal_status']=='active'
class TestMemory(ctypes.Structure):
    """Не запускает упаковку при недостаточном свободном физическом и общем выделяемом объёме."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in
        ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
prior=repo/'docs/reports/evidence/Coupled6DPrefixNewtonDirection_2026-11-26';prior_index=json.loads((prior/'Files.json').read_bytes())
assert sha256((prior/'RawEvidence.tar.xz').read_bytes()).hexdigest().upper()==prior_index['RawEvidence.tar.xz']
prior_summary=json.loads((prior/'Summary.json').read_bytes());book_reference=prior_summary['book_reconstruction']['968']
parent=base/'Cold968PrefixNewtonCapture_20261126';book=(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(book).hexdigest().upper()==book_reference['actual_saved_book_sha256']
raw={};proofs={};native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for n,root in roots.items():
    proof=json.loads((root/'MaterialPrefixCoordinateProof.json').read_bytes());actual,_=native(root,'RealColdReplay')
    assert actual==proof['native'] and actual['failed']==0 and actual['own_process_closed']
    assert proof['all257_actual_components_and_saved_book_bitexact968'] and not proof['physical_capacity_certified']
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==book
    hashes={}
    for folder in ('src','tests'):
        for p in (parent/folder).rglob('*'):
            if p.is_file():
                name=p.relative_to(parent).as_posix();data=(root/name).read_bytes();assert p.read_bytes()==data
                hashes[name]=sha256(data).hexdigest().upper()
    assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
    reference=dict(actual_saved_book_sha256=sha256(book).hexdigest().upper(),all_declared_src_and_tests_hashes=hashes,
        actual257_components_proven_by_book_bitexact_prior968=True,book_reconstruction=book_reference,
        book_reconstruction_evidence=prior.relative_to(repo).as_posix(),
        book_reconstruction_archive_sha256=prior_index['RawEvidence.tar.xz'],
        original_sheet_data_and_source_history_exact_public968=True)
    raw[f'{n}/VerifiedSourceBookReuse.json']=(json.dumps(reference,sort_keys=True,indent=2)+'\n').encode()
    files=[p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.py','.ps1','.txt','.bin')]
    for folder in ('Native','Input'):
        files.extend(p.relative_to(root).as_posix() for p in (root/folder).rglob('*') if p.is_file())
    files+=['docs/regression/CodeUpdate.json','tools/build_workbook/Run-Audit02ModeTests.ps1',
        'tools/build_workbook/Run-FocusedKaniWatchdog.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    proofs[str(n)]=proof
out=repo/'docs/reports/evidence/Coupled6DMaterialCoordinateCases_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};aliases=0
archive_file=out/'RawEvidence.tar.xz'
with tarfile.open(archive_file,'w:xz',preset=6) as archive:
    for name,data in raw.items():
        item=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:item.type=tarfile.LNKTYPE;item.linkname=canonical[digest];archive.addfile(item);aliases+=1
        else:canonical[digest]=name;item.size=len(data);archive.addfile(item,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();item=tarfile.TarInfo('RawFiles.json');item.size=len(data);archive.addfile(item,BytesIO(data))
checked={};actual_pins=None
with tarfile.open(archive_file,'r|xz') as archive:
    for item in archive:
        if item.islnk():digest=checked[item.linkname]
        else:
            data=archive.extractfile(item).read();digest=sha256(data).hexdigest().upper()
            if item.name=='RawFiles.json':actual_pins=json.loads(data);continue
        assert digest==pins[item.name];checked[item.name]=digest
assert checked==pins==actual_pins
helpers=set('''Stage-TestMaterialPrefixCoordinate.py Stage-TestActiveUnloadingCoordinate.py Check-TestMaterialPrefixCoordinate.py
Stage-TestOutgoingMaterialPair.py
Record-TestMaterialCoordinateCases.py Check-TestRecordedYieldCoordinate.py Set-TestActiveStage.py Set-TestTerminalStage.py
Rebuild-TestBookSegments.py'''.split());todo=list(helpers)
while todo:
    path=here/todo.pop();assert path.is_file()
    for name in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',path.read_text('utf-8-sig')):
        if name not in helpers and (here/name).is_file():helpers.add(name);todo.append(name)
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
for name,digest in prior_summary['release_hashes_unchanged'].items():assert sha256((repo/name).read_bytes()).hexdigest().upper()==digest
summary=dict(proofs=proofs,verified_book_and_source_reuse_reference=book_reference,
    prior_book_evidence=prior.relative_to(repo).as_posix(),prior_archive_sha256=prior_index['RawEvidence.tar.xz'],
    raw_files=len(raw),raw_archive_bytes=archive_file.stat().st_size,exact_duplicates=aliases,all_raw_hashes_read_back=True,
    all_actual_books_and257_components_bitexact_public968=True,release_hashes_unchanged=prior_summary['release_hashes_unchanged'],
    no_new_compilation=True,no_material_history_acceptance=True,physical_capacity_certified=False,cold_gain_claimed=False,
    full_acceptance_not_run=True,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_material_coordinate_evidence=out.relative_to(repo).as_posix(),latest_material_coordinate_cases=proofs)
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode()
for target in (statepath,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED MATERIAL CASES',len(raw),'RAW;',archive_file.stat().st_size,'BYTES;EXACT PUBLIC968 SOURCE/BOOK REUSE;GOAL ACTIVE')
