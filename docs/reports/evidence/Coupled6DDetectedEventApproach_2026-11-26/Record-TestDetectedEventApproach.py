"""ДЛЯ ТЕСТОВ. Сохраняет поиск события951–954, фактические книги, отрицательную локальность и проверку точности половинной пробы."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import json,re,tarfile,ctypes
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={951:'Cold951Point18EventRoots_20261126',952:'Cold952DetectedLoadingBoundary_20261126',
    953:'Cold953DetectedEventApproach_20261126',954:'Cold954RefinedEventApproach_20261126'}
roots={n:base/name for n,name in names.items()}
keys={951:['Point18EventRootsProof.json','RecordedYieldCoordinateProof.json'],952:['DetectedLoadingBoundaryProof.json'],
    953:['DetectedEventApproachProof.json'],954:['DetectedEventApproachProof.json','RefinedEventApproachProof.json']}
proofs={str(n):{key:json.loads((roots[n]/key).read_bytes()) for key in group} for n,group in keys.items()}
statefile=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statefile.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None and state['goal_status']=='active'
for items in proofs.values():
    for p in items.values():assert p['native']['own_process_closed'] and p['native']['failed']==0 and not p['physical_capacity_certified']
assert not any(c['independently_equilibrated'] for c in proofs['951']['Point18EventRootsProof.json']['cases'][:2])
assert proofs['952']['DetectedLoadingBoundaryProof.json']['independently_equilibrated']
assert proofs['952']['DetectedLoadingBoundaryProof.json']['effort'][0]==2
for n in (953,954):
    p=proofs[str(n)]['DetectedEventApproachProof.json'];assert p['native_half_solved'] and not p['independent_half_verified']
    assert p['combined_steps']<=30
class TestMemory(ctypes.Structure):
    """Проверяет RAM и общий запас памяти до упаковки завершённых доказательств."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
m=TestMemory();m.length=ctypes.sizeof(m);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
assert m.free_physical>3*1024**3 and m.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DDetectedEventApproach_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','docs/regression/CodeUpdate.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
    'CompilePoint.log','CompilePointClosure.json','CompileClosure.json','Run.ndjson','ParentCodeUpdate.json','ParentCompileClosure.json',
    'Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1','Compile-TestExternalPoint.ps1',
    'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
    'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','SourceOriginalRoot.bin','SourceWorkingGuidePairs.bin',
    'SourceDetectedEventProof.json','SourceAcceptedYieldCrossingProof.json']
raw={}
for n,root in roots.items():
    files=fixed+keys[n]
    files += [p.relative_to(root).as_posix() for folder in ('Native','Input','src','tests','DetectedEventCompilationAncestor','EventApproachCompilationAncestor','RefinedEventApproachCompilationAncestor')
        for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    book=(root/'docs/regression/Diagnostic.xlsm').read_bytes();manifest=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert sha256(book).hexdigest().upper()==manifest['WorkbookSHA256']
    with ZipFile(BytesIO(book)) as z:raw[f'{n}/ActualCompiledVBA.bin']=z.read('xl/vbaProject.bin')
    if n==954:raw[f'{n}/ActualCompiledWorkbook.xlsm']=book
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};canonical={};aliases=0;archivefile=out/'RawEvidence.tar.xz'
with tarfile.open(archivefile,'w:xz',preset=6) as archive:
    for name,data in raw.items():
        info=tarfile.TarInfo(name);digest=pins[name]
        if digest in canonical:
            info.type=tarfile.LNKTYPE;info.linkname=canonical[digest];archive.addfile(info);aliases+=1
        else:
            canonical[digest]=name;info.size=len(data);archive.addfile(info,BytesIO(data))
    data=(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode();info=tarfile.TarInfo('RawFiles.json');info.size=len(data);archive.addfile(info,BytesIO(data))
checked={};readpins=None
with tarfile.open(archivefile,'r|xz') as archive:
    for info in archive:
        if info.islnk():assert info.linkname in checked;digest=checked[info.linkname]
        else:
            data=archive.extractfile(info).read();digest=sha256(data).hexdigest().upper()
            if info.name=='RawFiles.json':readpins=json.loads(data);continue
        assert digest==pins[info.name];checked[info.name]=digest
assert checked==pins==readpins
helpers=set('''Read-TestOriginalControlReturn.py Stage-TestPoint18EventRoots.py Check-TestPoint18EventRoots.py
TestRunDetectedLoadingBoundary.bas.txt Stage-TestDetectedLoadingBoundary.py Check-TestDetectedLoadingBoundary.py
TestRunDetectedEventApproach.bas.txt Stage-TestDetectedEventApproach.py Check-TestDetectedEventApproach.py
Stage-TestRefinedEventApproach.py Check-TestRefinedEventApproach.py Finalize-TestExisting257Project.py
Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())
helpers.add(Path(__file__).name);todo=list(helpers)
while todo:
    p=here/todo.pop();assert p.is_file()
    for dep in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',p.read_text('utf-8-sig')):
        if dep not in helpers and (here/dep).is_file():helpers.add(dep);todo.append(dep)
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive_bytes=archivefile.stat().st_size,
    exact_duplicate_hardlinks=aliases,all_logical_raw_hashes_read_back=True,actual_compiled954_book_included=True,
    actual_compiled_projects951_to954_included=True,original_force_acceptance_half_policy_and_budget_unchanged=True,
    diagnostic_half_solve_precision_factor954=.0001,numerical_classes949_unchanged=True,
    generic_event_equilibrium_found_without_hardcoded_point=True,event_locality_not_accepted_yet=True,
    release_hashes_unchanged=release,all_numerical_stages_terminal=True,full_acceptance_not_run=True,
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_detected_event_approach_evidence=out.relative_to(repo).as_posix(),latest_detected_event_approach_summary={k:v for k,v in summary.items() if k!='proofs'},
    latest_detected_event952=proofs['952']['DetectedLoadingBoundaryProof.json'],latest_event_approach953=proofs['953']['DetectedEventApproachProof.json'],
    latest_refined_half954=proofs['954']['RefinedEventApproachProof.json'],
    previous_goal_turn_classification='progress:951 exact half reproduced but far-seed eventF0 failed.952 existing generic detector selected event18 and solved it in2steps with independent force/yield verification.953 same-tolerance event half failed geometric criterion.954 tighter half residual~8.45e-12 leaves3.574% geometric defect, so ordinary half solve tolerance is not main cause. No thresholds relaxed, history accepted, capacity or cold gain claimed.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for target in (statefile,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED951-954',len(raw),'RAW;',archivefile.stat().st_size,'BYTES;',aliases,'DUPLICATES;GOAL ACTIVE')
