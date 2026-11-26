"""ДЛЯ ТЕСТОВ. Архивирует непрерывность поверхности, фактические касательные и проверку найденного диапазона спуска без подмены физического предела."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,re,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={930:'Cold930FineLinePlusSurface_20261126',931:'Cold931SurfaceMatrices_20261126',
 932:'Cold932SurfaceCorrectionScan_20261126',933:'Cold933ProspectiveExtraHalf_20261126',934:'Cold934ConsistentSpatialLine_20261126'}
roots={n:base/name for n,name in names.items()}
proof_names={930:['FineLinePlusSurfaceProof.json'],931:['SurfaceMatricesProof.json','SurfaceDirectionSignsProof.json'],
 932:['SurfaceCorrectionScanProof.json'],933:['ProspectiveExtraHalfProof.json'],934:['ConsistentSpatialLineProof.json']}
proofs={str(n):{name:json.loads((roots[n]/name).read_bytes()) for name in files} for n,files in proof_names.items()}
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None
for n,files in proofs.items():
    for name,proof in files.items():
        if 'native' in proof:assert proof['native']['own_process_closed'] and proof['native']['failed']==0
        assert not proof.get('physical_capacity_certified',False)
assert proofs['930']['FineLinePlusSurfaceProof.json']['actual_capture_probes_matched']==17
assert proofs['931']['SurfaceMatricesProof.json']['base_residual_max_error_MPa']==0
assert proofs['932']['SurfaceCorrectionScanProof.json']['improving_endpoints']>0
for n,key in ((933,'ProspectiveExtraHalfProof.json'),(934,'ConsistentSpatialLineProof.json')):
    for case in proofs[str(n)][key]['cases']:
        assert case['effort'][0]<=30
        if case['solved']:assert case['independently_equilibrated']

class TestMemory(ctypes.Structure):
    """Проверяет RAM и общий запас commit перед выделением буфера сжатия."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in
        ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DSurfaceDescent_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json','CompiledSourceReuse.json',
 'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1',
 'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Run-FocusedKaniWatchdog.ps1',
 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
raw={}
for n,root in roots.items():
    files=fixed+proof_names[n]
    files += [p.relative_to(root).as_posix() for folder in ('src','tests','Native','Input')
        for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{n}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
    if n in (931,934):raw[f'{n}/ActualCompiledWorkbook.xlsm']=(root/'docs/regression/Diagnostic.xlsm').read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=set('''Stage-TestFineLinePlusSurface.py Check-TestFineLinePlusSurface.py
TestReadSurfaceMatrices.bas.txt Stage-TestSurfaceMatrices.py Check-TestSurfaceMatrices.py Check-TestSurfaceDirectionSigns.py
Stage-TestSurfaceCorrectionScan.py Check-TestSurfaceCorrectionScan.py Stage-TestProspectiveExtraHalf.py
Check-TestProspectiveExtraHalf.py Stage-TestConsistentSpatialLine.py Check-TestConsistentSpatialLine.py
Finalize-TestExisting257Project.py Set-TestActiveStage.py Set-TestTerminalStage.py'''.split())
helpers.add(Path(__file__).name);todo=list(helpers)
while todo:
    name=todo.pop();p=here/name;assert p.is_file()
    for dep in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',p.read_text('utf-8-sig')):
        if dep not in helpers and (here/dep).is_file():helpers.add(dep);todo.append(dep)
for name in sorted(helpers):(out/name).write_bytes((here/name).read_bytes())
release={name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary=dict(proofs=proofs,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive=archive.name,
    raw_archive_bytes=archive.stat().st_size,all_raw_hashes_read_back=True,exact_compiled931_and934_books_included=True,
    free_physical_before=memory.free_physical,free_commit_before=memory.free_page,archived_helpers=len(helpers),
    release_hashes_unchanged=release,full_acceptance_not_run=True,physical_capacity_certified=False,
    whole_goal_completed=False,no_working_release_main_promotion=True,all_started_stages_terminal=True)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',all_started_stages_terminal=True,
    latest_surface_descent_evidence=out.relative_to(repo).as_posix(),latest_surface_descent_summary={k:v for k,v in summary.items() if k!='proofs'},
    latest_shared_line934=proofs['934']['ConsistentSpatialLineProof.json'],
    previous_goal_turn_classification='progress:930 actual plus-station surface continuous, all17 recorded probes independently reconstructed.931 exact current R and refined matrix reproduced; incoming/outgoing Newton directions independently checked.932 actual nonlinear material reveals missed descent window1/256.933 one extra helper half accepts the previously missed branch step but no root;934 shared line range outcome preserved explicitly. No physical capacity, cold gain or release promotion claimed.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for target in (statepath,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED930-934',len(raw),'RAW;',archive.stat().st_size,'BYTES;',len(helpers),'HELPERS;GOAL ACTIVE')
