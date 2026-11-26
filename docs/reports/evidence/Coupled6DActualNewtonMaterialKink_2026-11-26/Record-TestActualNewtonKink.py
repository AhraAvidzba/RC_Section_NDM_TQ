"""ДЛЯ ТЕСТОВ. Сохраняет реальные отклонённые системы Newton, отрицательную проверку демпфирования и локальную смену поверхности."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={911:base/'Cold911NewtonRejectedCapture_20261126',912:base/'Cold912ExtendedSpatialDamping_20261126',913:base/'Cold913RejectedNewtonStressScan_20261126'}
names={911:'NewtonRejectedCaptureProof.json',912:'ExtendedSpatialDampingProof.json',913:'RejectedNewtonStressScanProof.json'}
proofs={n:json.loads((r/names[n]).read_bytes()) for n,r in roots.items()}
diagnosis=json.loads((roots[913]/'ActualNewtonKinkDiagnosisProof.json').read_bytes())
assert diagnosis['extended_damping_not_selected_as_a_solution'] and diagnosis['additional_arc_cases_solved']==0
assert all(p['native']['own_process_closed'] for p in proofs.values())
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes());assert not state['active_jobs']
class TestMemory(ctypes.Structure):
    """Проверяет общий запас RAM/commit до выделения словаря XZ."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DActualNewtonMaterialKink_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
raw={}
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json','CompiledSourceReuse.json',
 'Input/Reverse1.bin','Input/SourceColdRawHistory.bin','Input/ArcScanFields.bin','Input/ArcScanCases.json','Input/ActualRejectedNewton.bin',
 'NewtonRejectedCaptureProof.json','ExtendedSpatialDampingProof.json','RejectedNewtonStressScanProof.json','ActualNewtonKinkDiagnosisProof.json',
 'SourceNewtonRejectedCaptureProof.json','BeforeCaptureBlankAlignment.bas','InitialPreparedCapture.bas',
 'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1',
 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','tools/build_workbook/Run-Audit02ModeTests.ps1']
for number,root in roots.items():
    files=fixed+[p.relative_to(root).as_posix() for folder in ('src','tests','Native','InitialCaptureCompilation') for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
with ZipFile(roots[911]/'InitialCaptureCompilation/Diagnostic.xlsm') as book:raw['911/InitialCaptureCompilation/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
baseline=base/'Cold909ResidualArcPredictorBaseline_20261126'
for name in ('ResidualArcPredictorBaselineProof.json','Native/Arc32Diagnostics.bin'):
    raw['Referenced909/'+name]=(baseline/name).read_bytes()
for path in (baseline/'Native').glob('Arc32.Case*'):
    if path.is_file():raw['Referenced909/Native/'+path.name]=path.read_bytes()
parent=base/'Cold900Cold32OutgoingJacobian_20261126'
for path in (parent/'src/Materials').glob('*'):
    if path.is_file():assert path.read_bytes()==(roots[911]/path.relative_to(parent)).read_bytes(),path.name
assert (parent/'src/Solver/CCoupledElasticState.cls').read_bytes()==(roots[911]/'src/Solver/CCoupledElasticState.cls').read_bytes()
for name in ('PointTraceSourceComposition.json','Cold32OutgoingJacobianProof.json'):
    raw['Referenced900/'+name]=(parent/name).read_bytes()
raw['Referenced911Source/CCoupledNewtonSolver.cls']=(baseline/'src/Solver/CCoupledNewtonSolver.cls').read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['TestNewtonRejectedCapture.bas.txt','Stage-TestNewtonRejectedCapture.py','Repair-TestNewtonCaptureHeader.py',
 'Align-TestNewtonCaptureBlank.py','Check-TestNewtonRejectedCapture.py','Stage-TestExtendedSpatialDamping.py','Check-TestExtendedSpatialDamping.py',
 'Stage-TestRejectedNewtonStressScan.py','Check-TestRejectedNewtonStressScan.py','Check-TestNewtonKinkDiagnosis.py',
 'Check-TestCold32OutgoingStressScan.py','Check-TestArc56Diagnostics.py','Check-TestPersistentHEvidence.py',
 'Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',
 'Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Finalize-TestOperationalArc.py','Finalize-TestStockYield.py',
 'Audit-TestPrivateCoupledBook.py','Inspect-TestResolutionVBA.py','Canonical-TestCoreVBA.py','Compile-TestExternalPoint.ps1',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(proofs=proofs,diagnosis=diagnosis,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
 raw_archive=archive.name,raw_archive_bytes=archive.stat().st_size,all_raw_hashes_read_back=True,
 free_physical_before=memory.free_physical,free_commit_before=memory.free_page,
 all_started_stages_terminal=True,initial_compile_header_position_repaired=True,
 additional_floor_halvings_not_selected=True,full_acceptance_not_run=True,no_new_cold_limit_claim=True,
 no_working_release_main_promotion=True,physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
 latest_actual_newton_kink_evidence=out.relative_to(repo).as_posix(),latest_actual_newton_kink_diagnosis=diagnosis,
 all_started_stages_terminal=True,
 previous_goal_turn_revalidation='Previous user-facing acknowledgement only deferred pure-bending comparison; no numerical progress claimed. Revalidated clean worktree and terminal jobs; continued actual Newton capture instead of repeating a plan.',
 previous_goal_turn_classification='progress:911 actual rejected augmented matrix/RHS/correction/valid probes captured with exact accepted32 history; all43 baseline native files exact; actual256 sources verified.912 five more halvings yield0.44% local norm reduction then rejection repeats; all40 root/operator exports exact baseline, zero gained/lost cases; not selected.913 immutable compiled900 local stress scan independently rebuilds actual failed residual within2.43e-13MPa and captured probes6.06e-13; point5 plus-station trial-yield crosses between1/256 and1/128; near direction agrees approximately withA, finite kink is localized. No physical ultimate or production promotion. All native processes closed.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (statepath,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED',len(raw),'RAW FILES',archive.stat().st_size,'BYTES;ACTUAL MATERIAL KINK;NO CAPACITY CLAIM;GOAL ACTIVE')
