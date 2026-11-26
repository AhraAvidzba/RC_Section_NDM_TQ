"""ДЛЯ ТЕСТОВ. Сохраняет отрицательные контроли и первый автоматически найденный арматурный предел ASD с независимым равновесием."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from runpy import run_path
from datetime import datetime
import json,tarfile,re,shutil,ctypes

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={1020:'Cold1020ASDBendingMainCapacity_20261126',1021:'Cold1021ASDBendingMainCapacityRetry_20261126',
    1022:'Cold1022ASDBendingProbeBudget_20261126',1023:'Cold1023ASDGeometricContinuation_20261126',
    1024:'Cold1024ASDGeometricDecline_20261126',1025:'Cold1025ASDGeometricDivision_20261126',1026:'Cold1026ASDEarlyDeformation_20261126'}
out=repo/'docs/reports/evidence/Coupled6DAutomaticASDBending_2026-11-26';assert not out.exists()
foundation=repo/'docs/reports/evidence/Coupled6DSelectableMaterials_2026-11-26'
foundation_pins=json.loads((foundation/'Files.json').read_bytes());parent_archive=foundation/'RawEvidence.tar.xz'
assert sha256(parent_archive.read_bytes()).hexdigest().upper()==foundation_pins[parent_archive.name]
parent_member='1018/docs/regression/Diagnostic.xlsm'
with tarfile.open(parent_archive,'r:xz') as archive:parent=archive.extractfile(parent_member).read()
assert sha256(parent).hexdigest().upper()=='AD8C0B4E38E25230EA036377FD06BABCB415025197125D1AE38DFF1821ECCAF0'
memory=ctypes.create_string_buffer(64);ctypes.c_ulong.from_buffer(memory).value=64
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(memory)
assert ctypes.c_ulonglong.from_buffer(memory,16).value>6500*1024**2 and ctypes.c_ulonglong.from_buffer(memory,32).value>6*1024**3
segment=run_path(str(here/'Rebuild-TestBookSegments.py'));cell_reader=run_path(str(here/'Record-TestSelectableMaterialEvidence.py')) if False else None
# Код проверки ячеек извлекается без повторного запуска архиватора и без записи в книги.
import ast
source=(here/'Record-TestSelectableMaterialEvidence.py').read_text('utf-8-sig');tree=ast.parse(source)
method=next(node for node in tree.body if isinstance(node,ast.FunctionDef) and node.name=='TestWorkbookCells')
from openpyxl import load_workbook
scope=dict(BytesIO=BytesIO,load_workbook=load_workbook);exec(compile(ast.get_source_segment(source,method),'<TestWorkbookCells>','exec'),scope)
parent_cells=scope['TestWorkbookCells'](parent)
raw={};reconstruction={};proofs={}
for n,folder in names.items():
    root=base/folder;prep=json.loads((root/'Preparation.json').read_bytes())
    for name,digest in prep['inputs'].items():assert sha256((root/name).read_bytes()).hexdigest().upper()==digest,(n,name)
    manifest='docs/regression/CodeUpdate.json' if n!=1020 else 'ParentCodeUpdate.json'
    compiled=json.loads((root/manifest).read_text('utf-8-sig'));book=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
    assert sha256(book).hexdigest().upper()==compiled['WorkbookSHA256']
    assert scope['TestWorkbookCells'](book)==parent_cells,(n,'changed workbook cells')
    if n==1020:
        proof=json.loads((root/'CompilerRPCFailureProof.json').read_bytes());assert proof['own_compiler_closed'] and proof['numerical_calculation_not_started']
    else:
        proof=json.loads((root/'MainCapacityExecutionProof.json').read_bytes());assert proof['native']['own_process_closed'] and proof['source_unchanged']
        if n==1026:assert proof['has_steel_limit']==1 and proof['has_ultimate_point']==0
        else:assert proof['has_steel_limit']==0 and proof['has_ultimate_point']==0
    proofs[str(n)]=proof
    patch,blobs=segment['TestMakeSegments'](parent,book)
    assert segment['TestRebuildSegments'](parent,patch,blobs)==book
    for digest,data in blobs.items():raw['BookSegmentBlobs/'+digest]=data
    member=f'{n}/ActualSavedBookSegments.json';raw[member]=(json.dumps(patch,sort_keys=True,indent=2)+'\n').encode()
    reconstruction[str(n)]=dict(actual_book_sha256=patch['target_sha256'],manifest_member=member,all_saved_bytes_rebuilt_exact=True,
        parent_archive=parent_archive.relative_to(repo).as_posix(),parent_archive_sha256=foundation_pins[parent_archive.name],
        parent_book_member=parent_member,blobs_prefix='BookSegmentBlobs/')
    files=set(prep['inputs'])|{'Preparation.json',manifest}
    files.update(p.relative_to(root).as_posix() for p in root.iterdir() if p.is_file() and p.suffix in ('.json','.log','.ndjson','.ps1','.py','.txt'))
    files.update(p.relative_to(root).as_posix() for p in (root/'Native').rglob('*') if p.is_file())
    for name in sorted(files):raw[f'{n}/{name}']=(root/name).read_bytes()
external=json.loads((base/names[1026]/'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
assert external['all56plus4_equations_and_six_resultants_verified'] and external['accepted_states']==53
sourceproof=json.loads((base/names[1026]/'EarlySourceProof.json').read_bytes());assert sourceproof['three_changed_components_exact']
out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};aliases=0
archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=3) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);aliases+=1
        else:canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo('RawFiles.json');info.size=len(data);archive.addfile(info,BytesIO(data))
assert archivefile.stat().st_size<100*1024**2
with tarfile.open(archivefile,'r:xz') as archive:
    assert json.load(archive.extractfile('RawFiles.json'))==pins
    for name,digest in pins.items():assert sha256(archive.extractfile(name).read()).hexdigest().upper()==digest,name
    for n,record in reconstruction.items():
        patch=json.load(archive.extractfile(record['manifest_member']))
        blobs={item['blob']:archive.extractfile('BookSegmentBlobs/'+item['blob']).read() for item in patch['segments'] if 'blob' in item}
        assert sha256(segment['TestRebuildSegments'](parent,patch,blobs)).hexdigest().upper()==record['actual_book_sha256']
helpers=set('''Record-TestAutomaticASDBending.py Record-TestSelectableMaterialEvidence.py Rebuild-TestBookSegments.py
Stage-TestASDBendingMainCapacity.py Stage-TestASDBendingMainCapacityRetry.py Check-TestASDBendingMainSource.py Check-TestASDBendingMainCapacity.py
Stage-TestASDBendingProbeBudget.py Check-TestASDProbeBudgetSource.py Stage-TestASDGeometricContinuation.py Check-TestASDGeometricSource.py
Stage-TestASDGeometricDecline.py Check-TestASDGeometricDeclineSource.py Stage-TestASDGeometricDivision.py Check-TestASDGeometricDivisionSource.py
Stage-TestASDEarlyDeformation.py Check-TestASDEarlySource.py Replay-TestMainCapacityEquilibrium.py TestCaptureMainCapacity.bas.txt TestASDBendingMainCapacity.bas.txt'''.split())
while True:
    expanded=set(helpers)
    for name in helpers:
        path=here/name;assert path.is_file(),name;text=path.read_text('utf-8-sig')
        expanded.update(p.name for p in here.iterdir() if p.is_file() and p.suffix in ('.py','.ps1','.txt') and p.name in text)
    if expanded==helpers:break
    helpers=expanded
for name in sorted(helpers):shutil.copyfile(here/name,out/name)
summary=dict(created_local=datetime.now().astimezone().isoformat(),raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive_bytes=archivefile.stat().st_size,
    exact_duplicate_hardlinks=aliases,all_raw_hashes_read_back=True,all7_actual_books_rebuilt_from_archive_bitexact=True,
    reconstruction=reconstruction,sourceproof=sourceproof,execution_proofs=proofs,
    external_equilibrium={k:v for k,v in external.items() if k not in ('records','observed_OpenSees_concrete_stresses')},
    final_state=external['records'][-1],maximum_physical_residual_MPa=max(s['physical_residual_MPa'] for s in external['records']),
    public_solver_controller_not_changed=True,public_release_book_unchanged=True,
    physical_global_capacity_not_certified=True,concrete_limit_unresolved=True,other_loads_not_yet_tested=True,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.name:sha256(p.read_bytes()).hexdigest().upper() for p in out.iterdir() if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('AUTOMATIC ASD BENDING ARCHIVED',len(raw),'RAW FILES;',archivefile.stat().st_size,'BYTES;53 STATES VERIFIED;STEEL EVENT ONLY')
