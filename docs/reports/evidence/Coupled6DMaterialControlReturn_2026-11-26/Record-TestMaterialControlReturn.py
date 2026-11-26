"""ДЛЯ ТЕСТОВ. Сохраняет точные материальные корни и возврат исходной цели с отрицательными исходами и проверенными нативными книгами."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,re,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={935:'Cold935Point11Matrices_20261126',936:'Cold936RecordedYieldCoordinate_20261126',
    938:'Cold938RecordedYieldAdvance_20261126',940:'Cold940RecordedControlBracket_20261126',941:'Cold941OriginalControlReturn_20261126',
    942:'Cold942OriginalControlFarReturn_20261126'}
roots={n:base/name for n,name in names.items()}
proofnames={935:['SurfaceMatricesProof.json','SurfaceDirectionSignsProof.json'],
    936:['RecordedYieldCoordinateProof.json'],938:['RecordedYieldCoordinateProof.json'],
    940:['RecordedYieldCoordinateProof.json'],941:['OriginalControlReturnProof.json'],942:['OriginalControlReturnProof.json']}
proofs={str(n):{key:json.loads((roots[n]/key).read_bytes()) for key in keys} for n,keys in proofnames.items()}
statefile=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statefile.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None
for n,items in proofs.items():
    for key,p in items.items():
        assert not p.get('physical_capacity_certified',False)
        if 'native' in p:assert p['native']['own_process_closed'] and p['native']['failed']==0
for n in (936,938,940):
    p=proofs[str(n)]['RecordedYieldCoordinateProof.json']
    for c in p['cases']:
        assert c['steps']<=30
        if c['native_solved']:assert c['independently_equilibrated']
for n in (941,942):
    for c in proofs[str(n)]['OriginalControlReturnProof.json']['cases']:
        assert c['steps']<=30
        if c['native_solved']:assert c['independently_solves_original_control']
class TestMemory(ctypes.Structure):
    """Проверяет общий запас памяти перед сжатием доказательств."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in
        ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DMaterialControlReturn_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
    'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json','CompiledSourceReuse.json',
    'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1','Run.ndjson',
    'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
    'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','SourceMaterialRoot.bin','OriginalControlBracket.json',
    'VerifiedStartingMaterialRoot.json','BeforeRecordedYieldHelperBlankAlignment.bas']
raw={}
for n,root in roots.items():
    files=fixed+proofnames[n]
    files += [p.relative_to(root).as_posix() for folder in ('src','tests','Native','Input','RecordedYieldCoordinateCompilationAncestor',
        'OriginalControlReturnCompilationAncestor') for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    book=root/'docs/regression/Diagnostic.xlsm';manifest=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert sha256(book.read_bytes()).hexdigest().upper()==manifest['WorkbookSHA256']
    with ZipFile(book) as z:raw[f'{n}/ActualCompiledVBA.bin']=z.read('xl/vbaProject.bin')
    if n in (936,941):raw[f'{n}/ActualCompiledWorkbook.xlsm']=book.read_bytes()
pins={key:sha256(value).hexdigest().upper() for key,value in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as bundle:
    for key,value in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(key);info.size=len(value);bundle.addfile(info,BytesIO(value))
with tarfile.open(archive,'r:xz') as bundle:
    for key,digest in pins.items():assert sha256(bundle.extractfile(key).read()).hexdigest().upper()==digest,key
helpers=set('''Stage-TestPoint11Matrices.py Check-TestPoint11Matrices.py Check-TestPoint11DirectionSigns.py
Stage-TestRecordedYieldCoordinate.py Align-TestRecordedYieldHelperBlank.py Check-TestRecordedYieldCoordinate.py
TestSolveRecordedYieldCoordinate.bas.txt Stage-TestRecordedYieldAdvance.py Check-TestRecordedYieldAdvance.py
Stage-TestRecordedYieldWarm.py Read-TestRecordedYieldCoordinate.py Stage-TestRecordedControlBracket.py
Check-TestRecordedControlBracket.py TestSolveOriginalControlReturn.bas.txt Stage-TestOriginalControlReturn.py
Check-TestOriginalControlReturn.py Stage-TestOriginalControlFarReturn.py Check-TestOriginalControlFarReturn.py
Finalize-TestExisting257Project.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())
helpers.add(Path(__file__).name);todo=list(helpers)
while todo:
    p=here/todo.pop();assert p.is_file()
    for dep in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',p.read_text('utf-8-sig')):
        if dep not in helpers and (here/dep).is_file():helpers.add(dep);todo.append(dep)
for name in sorted(helpers):(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive=archive.name,
    raw_archive_bytes=archive.stat().st_size,all_raw_hashes_read_back=True,exact_compiled936_and941_books_included=True,
    archived_helpers=len(helpers),release_hashes_unchanged=release,full_acceptance_not_run=True,
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False,
    no_working_release_main_promotion=True,all_started_stages_terminal=True,
    prepared939_not_launched=True,local_hook_check='docs/reports/figures/MaterialState_2026-11-26/HookTrajectoryProof.json')
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',all_started_stages_terminal=True,
    latest_material_control_return_evidence=out.relative_to(repo).as_posix(),
    latest_material_control_return_summary={k:v for k,v in summary.items() if k!='proofs'},
    latest_original_control941=proofs['941']['OriginalControlReturnProof.json'],
    latest_original_control942=proofs['942']['OriginalControlReturnProof.json'],
    previous_goal_turn_classification='progress:935 reproduced current matrices and reversal;936 found constrained roots;938 larger positive material target crossed original scalar goal;940 verified intermediate roots;941 tested return to all original equations from verified material root. Negative outcomes retained; no capacity or cold gain claim.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for target in (statefile,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED935/936/938/940/941/942',len(raw),'RAW;',archive.stat().st_size,'BYTES;',len(helpers),'HELPERS;GOAL ACTIVE')
