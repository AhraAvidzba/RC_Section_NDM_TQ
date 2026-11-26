"""ДЛЯ ТЕСТОВ. Сохраняет все положительные и отрицательные ASD-контроли, повторно используя проверенные данные родительского архива по SHA-256."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from runpy import run_path
from datetime import datetime
import json,tarfile,shutil,ast,ctypes
from openpyxl import load_workbook

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent_folder=repo/'docs/reports/evidence/Coupled6DAutomaticASDBending_2026-11-26';parent_archive=parent_folder/'RawEvidence.tar.xz'
parent_pins=json.loads((parent_folder/'Files.json').read_bytes());assert sha256(parent_archive.read_bytes()).hexdigest().upper()==parent_pins['RawEvidence.tar.xz']
with tarfile.open(parent_archive,'r:xz') as archive:parent_raw=json.load(archive.extractfile('RawFiles.json'))
external_by_hash={}
for member,digest in parent_raw.items():external_by_hash.setdefault(digest,member)
parent_summary=json.loads((parent_folder/'Summary.json').read_bytes())
parent_book=(base/'Cold1026ASDEarlyDeformation_20261126/docs/regression/Diagnostic.xlsm').read_bytes()
assert sha256(parent_book).hexdigest().upper()==parent_summary['reconstruction']['1026']['actual_book_sha256']
memory=ctypes.create_string_buffer(64);ctypes.c_ulong.from_buffer(memory).value=64;assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory,16).value>6500*1024**2 and ctypes.c_ulonglong.from_buffer(memory,32).value>6*1024**3
names={1027:'Cold1027ASDCapacityScaling_20261126',1028:'Cold1028ASDCapacityDoubleInput_20261126',1029:'Cold1029ASDTwoPointStart_20261126',
1030:'Cold1030ASDScale1M_20261126',1031:'Cold1031ASDScale4M_20261126',1032:'Cold1032ASDScale10M_20261126',
1033:'Cold1033ASDHybridRecovery_20261126',1034:'Cold1034ASDTolerance1E5_20261126',1035:'Cold1035ASDPureTorsion_20261126',
1036:'Cold1036ASDTorsionOverload_20261126',1037:'Cold1037ASDNMTCapacity_20261126'}
out=repo/'docs/reports/evidence/Coupled6DASDStabilityAndTorsion_2026-11-26';assert not out.exists()
segments=run_path(str(here/'Rebuild-TestBookSegments.py'))
reader_source=(here/'Record-TestSelectableMaterialEvidence.py').read_text('utf-8-sig')
method=next(n for n in ast.parse(reader_source).body if isinstance(n,ast.FunctionDef) and n.name=='TestWorkbookCells')
scope=dict(BytesIO=BytesIO,load_workbook=load_workbook);exec(compile(ast.get_source_segment(reader_source,method),'<TestWorkbookCells>','exec'),scope)
parent_cells=scope['TestWorkbookCells'](parent_book)
logical={};payloads={};books={};proofs={};replays={}

def TestAdd(name,data):
    """Создаёт ссылку на неизменённые родительские байты либо один новый полезный блок с точным хешем."""
    digest=sha256(data).hexdigest().upper()
    if digest in external_by_hash:
        logical[name]=dict(sha256=digest,parent_member=external_by_hash[digest])
    else:
        payloads.setdefault(digest,data);logical[name]=dict(sha256=digest,payload='Payloads/'+digest)

for n,folder in names.items():
    root=base/folder;prep=json.loads((root/'Preparation.json').read_bytes())
    for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,(n,name)
    execution=json.loads((root/'MainCapacityExecutionProof.json').read_bytes())
    assert execution['native']['own_process_closed'] and not execution['native']['timeout600'] and execution['source_unchanged']
    replay=json.loads((root/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes());assert replay['all56plus4_equations_and_six_resultants_verified']
    manifest=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'));book=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
    assert sha256(book).hexdigest().upper()==manifest['WorkbookSHA256']
    assert scope['TestWorkbookCells'](book)==parent_cells,(n,'workbook cell change')
    patch,blobs=segments['TestMakeSegments'](parent_book,book)
    assert segments['TestRebuildSegments'](parent_book,patch,blobs)==book
    for digest,data in blobs.items():TestAdd('BookSegmentBlobs/'+digest,data)
    member=f'{n}/ActualSavedBookSegments.json';TestAdd(member,(json.dumps(patch,sort_keys=True,indent=2)+'\n').encode())
    books[str(n)]=dict(manifest_member=member,actual_saved_book_sha256=manifest['WorkbookSHA256'],blobs_prefix='BookSegmentBlobs/')
    files=set(prep['inputs'])|{'Preparation.json','docs/regression/CodeUpdate.json','MainCapacityExecutionProof.json'}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.log','.ndjson','.ps1','.py','.txt'))
    files.update(p.relative_to(root).as_posix() for p in (root/'Native').rglob('*') if p.is_file())
    for name in sorted(files):TestAdd(f'{n}/{name}',(root/name).read_bytes())
    proofs[str(n)]={k:v for k,v in execution.items() if k not in ('trace','diagnostics')}
    replays[str(n)]={k:v for k,v in replay.items() if k not in ('records','observed_OpenSees_concrete_stresses')}
out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
parent_record=dict(archive=parent_archive.relative_to(repo).as_posix(),archive_sha256=parent_pins['RawEvidence.tar.xz'],
    raw_index_member='RawFiles.json',base_book_sha256=sha256(parent_book).hexdigest().upper(),
    base_book_manifest_member='1026/ActualSavedBookSegments.json',base_book_reconstruction=parent_summary['reconstruction']['1026'])
archivefile=out/'RawOverlay.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=3) as archive:
    for digest,data in payloads.items():
        info=tarfile.TarInfo('Payloads/'+digest);info.size=len(data);archive.addfile(info,BytesIO(data))
    for name,obj in (('LogicalFiles.json',logical),('ParentArchive.json',parent_record)):
        data=(json.dumps(obj,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo(name);info.size=len(data);archive.addfile(info,BytesIO(data))
assert archivefile.stat().st_size<100*1024**2
read_payloads={}
with tarfile.open(archivefile,'r|xz') as archive:
    for info in archive:
        data=archive.extractfile(info).read()
        if info.name.startswith('Payloads/'):
            digest=info.name.split('/')[-1];assert sha256(data).hexdigest().upper()==digest;read_payloads[digest]=data
        elif info.name=='LogicalFiles.json':assert json.loads(data)==logical
        elif info.name=='ParentArchive.json':assert json.loads(data)==parent_record
        else:raise AssertionError(info.name)
for name,record in logical.items():
    digest=record['sha256']
    if 'parent_member' in record:assert parent_raw[record['parent_member']]==digest
    else:assert sha256(read_payloads[digest]).hexdigest().upper()==digest
# Для повторной сборки книг нужные блоки родителя считываются одним проходом, без повторного распаковывания архива на каждую ссылку.
needed=set()
for record in books.values():
    patch=json.loads(read_payloads[logical[record['manifest_member']]['sha256']])
    for item in patch['segments']:
        if 'blob' in item and item['blob'] not in read_payloads:needed.add(item['blob'])
external_bytes={}
if needed:
    with tarfile.open(parent_archive,'r|xz') as archive:
        for info in archive:
            if info.name not in parent_raw or info.islnk():continue
            digest=parent_raw[info.name]
            if digest in needed:
                data=archive.extractfile(info).read();assert sha256(data).hexdigest().upper()==digest;external_bytes[digest]=data
    assert set(external_bytes)==needed
for record in books.values():
    patch=json.loads(read_payloads[logical[record['manifest_member']]['sha256']]);blobs={}
    for item in patch['segments']:
        if 'blob' in item:blobs[item['blob']]=read_payloads.get(item['blob'],external_bytes.get(item['blob']))
    rebuilt=segments['TestRebuildSegments'](parent_book,patch,blobs)
    assert sha256(rebuilt).hexdigest().upper()==record['actual_saved_book_sha256']
helpers=set('''Record-TestASDStabilityOverlay.py Rebuild-TestBookSegments.py Record-TestSelectableMaterialEvidence.py
Check-TestASDBendingMainCapacity.py Replay-TestMainCapacityEquilibrium.py
Stage-TestASDCapacityScaling.py Check-TestASDCapacityScalingSource.py Stage-TestASDCapacityDoubleInput.py Stage-TestASDTwoPointStart.py Check-TestASDTwoPointSource.py
Stage-TestASDCapacityScaleBundle.py Stage-TestASDHybridRecovery.py Check-TestASDHybridSource.py Stage-TestASDToleranceStudy.py Check-TestASDToleranceSource.py
Stage-TestASDPureTorsionCapacity.py Check-TestASDPureTorsionSource.py Stage-TestASDTorsionOverload.py Stage-TestASDNMTCapacity.py Check-TestASDNMTSource.py'''.split())
while True:
    expanded=set(helpers)
    for name in helpers:
        path=here/name;assert path.is_file(),name;text=path.read_text('utf-8-sig')
        expanded.update(p.name for p in here.iterdir() if p.is_file() and p.suffix in ('.py','.ps1','.txt') and p.name in text)
    if expanded==helpers:break
    helpers=expanded
for name in sorted(helpers):shutil.copyfile(here/name,out/name)
summary=dict(created_local=datetime.now().astimezone().isoformat(),logical_files=len(logical),unique_new_payloads=len(payloads),
    parent_references=sum('parent_member' in r for r in logical.values()),overlay_bytes=archivefile.stat().st_size,
    all_new_payloads_read_back_and_hashed=True,all_parent_references_match_pinned_parent_raw_index=True,
    all11_actual_books_rebuilt_bitexact_from_read_back_payloads=True,parent_archive=parent_record,book_reconstruction=books,
    native_execution=proofs,independent_equilibrium=replays,all_populated_workbook_cells_equal_published_fixture=True,
    production_solver_tolerances_unchanged=True,private_controller_not_promoted=True,full_global_concrete_capacity_unresolved=True,
    full_acceptance_not_run=True,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
pins={p.name:sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ASD STABILITY OVERLAY VERIFIED',len(logical),'LOGICAL FILES;',summary['parent_references'],'PARENT REFERENCES;',archivefile.stat().st_size,'BYTES;11 EXACT BOOKS;NO GLOBAL PASS')
