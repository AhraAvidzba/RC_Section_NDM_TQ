"""ДЛЯ ТЕСТОВ. Архивирует рабочий возврат947–950, сохранённые ограничения и независимый разбор полной/половинной пары."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import json,re,tarfile,ctypes
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={947:'Cold947WorkingSecantReturn_20261126',948:'Cold948QuarterWorkingReturn_20261126',
    949:'Cold949WorkingGuidePair_20261126',950:'Cold950GuidePairBaseScan_20261126'}
roots={n:base/name for n,name in names.items()}
keys={947:['WorkingSecantReturnProof.json'],948:['WorkingSecantReturnProof.json'],
    949:['WorkingSecantReturnProof.json','WorkingGuidePairProof.json'],950:['GuidePairBaseScanProof.json']}
proofs={str(n):{key:json.loads((roots[n]/key).read_bytes()) for key in group} for n,group in keys.items()}
statefile=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statefile.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None and state['goal_status']=='active'
for items in proofs.values():
    for proof in items.values():
        assert proof['native']['own_process_closed'] and proof['native']['failed']==0 and not proof['physical_capacity_certified']
for n in (947,948):
    p=proofs[str(n)]['WorkingSecantReturnProof.json'];assert not p['cases'][0]['native_solved']
    assert p['cases'][1]['synthetic_exhausted_budget'] and p['cases'][1]['total_steps_including_guide']==30
assert (roots[949]/'Native/WorkingGuideReturn.bin').read_bytes()==(roots[947]/'Native/WorkingGuideReturn.bin').read_bytes()
pair=proofs['949']['WorkingGuidePairProof.json']['pairs'][0]
assert not pair['branch_differences'] and pair['field_error']>.005
class TestMemory(ctypes.Structure):
    """Проверяет RAM и общий запас выделяемой памяти до упаковки закрытых доказательств."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
m=TestMemory();m.length=ctypes.sizeof(m);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(m))
assert m.free_physical>3*1024**3 and m.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DWorkingGuideCurvature_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','docs/regression/CodeUpdate.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
    'CompilePoint.log','CompilePointClosure.json','CompileClosure.json','Run.ndjson','ParentCodeUpdate.json','ParentCompileClosure.json',
    'Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1','Compile-TestExternalPoint.ps1',
    'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
    'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','VerifiedGuide938.json','SourceMaterialRoot.bin','SourceWorkingGuidePairs.bin','FullControlTargetReference.json']
raw={}
for n,root in roots.items():
    files=fixed+keys[n]
    files += [p.relative_to(root).as_posix() for folder in ('Native','Input','src','tests','WorkingSecantCompilationAncestor','WorkingGuidePairCompilationAncestor')
        for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    book=(root/'docs/regression/Diagnostic.xlsm').read_bytes();manifest=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert sha256(book).hexdigest().upper()==manifest['WorkbookSHA256']
    with ZipFile(BytesIO(book)) as z:raw[f'{n}/ActualCompiledVBA.bin']=z.read('xl/vbaProject.bin')
    if n in (947,949):raw[f'{n}/ActualCompiledWorkbook.xlsm']=book
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
helpers=set('''Stage-TestWorkingSecantReturn.py TestMaterialGuideOriginalReturn.cls.txt TestRunWorkingGuideReturn.bas.txt
Check-TestWorkingSecantReturn.py Stage-TestQuarterWorkingReturn.py Check-TestQuarterWorkingReturn.py
Stage-TestWorkingGuidePair.py TestRunWorkingGuidePair.bas.txt Check-TestWorkingGuidePair.py
Stage-TestGuidePairBaseScan.py Check-TestGuidePairBaseScan.py Finalize-TestExisting257Project.py
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
    exact_duplicate_hardlinks=aliases,all_logical_raw_hashes_read_back=True,actual_compiled947_and949_books_included=True,
    guide17_steps_counted_in30=True,half_field_limit_preserved=.005,earlier_failed_primary_effort_not_included=True,
    actual_full_and_half_states_verified=True,source_base_reread_matches_saved_stress=True,
    generic_secant_is_private_working_variant_not_release=True,release_hashes_unchanged=release,
    all_numerical_stages_terminal=True,full_acceptance_not_run=True,physical_capacity_certified=False,
    cold_gain_claimed=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
index={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(index,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_working_guide_curvature_evidence=out.relative_to(repo).as_posix(),latest_working_guide_curvature_summary={k:v for k,v in summary.items() if k!='proofs'},
    latest_actual_half_pair949=proofs['949']['WorkingGuidePairProof.json'],latest_base_reread950=proofs['950']['GuidePairBaseScanProof.json'],
    previous_goal_turn_classification='progress:947 generic secant integrated into private working return with17guide steps charged and all half/uncertainty/direction guards retained.948 smaller target still rejected.949 exact passive result and full/half equilibrium/metric independently verified, no active-branch change.950 base material reread exact. No physical limit or new cold gain claimed.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for target in (statefile,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED947-950',len(raw),'RAW;',archivefile.stat().st_size,'BYTES;',aliases,'DUPLICATES;GOAL ACTIVE')
