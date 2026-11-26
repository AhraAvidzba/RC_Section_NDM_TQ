"""ДЛЯ ТЕСТОВ. Сохраняет воспроизводимый холодный контекст922–929, локальные проверки и отрицательные варианты без объявления несущей способности."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,re,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={922:'Cold922ColdNewtonCapture_20261126',923:'Cold923RecordedNewtonCold96_20261126',
 924:'Cold924RecordedHistoryStressBridge_20261126',925:'Cold925RecordedTailSystem_20261126',
 926:'Cold926RecordedTailStressScan_20261126',927:'Cold927BoundedBranchRecovery_20261126',
 928:'Cold928RecordedTailBoundaryScan_20261126',929:'Cold929RecordedTailFineLine_20261126'}
roots={n:base/name for n,name in names.items()}
proof_names={923:['RecordedNewtonCold96Proof.json','RecordedCheckpointsIdentityProof.json','ColdRejectedContextsProof.json'],
 924:['VerifiedRecordedStressBridgeProof.json'],925:['RecordedTailSystemProof.json'],
 926:['RecordedTailStressScanProof.json','RecordedTailKinksProof.json'],927:['BoundedBranchRecoveryProof.json'],
 928:['RecordedTailStressScanProof.json','RecordedTailBoundaryScanProof.json'],929:['RecordedTailFineLineProof.json']}
proofs={str(n):{name:json.loads((roots[n]/name).read_bytes()) for name in files} for n,files in proof_names.items()}
statepath=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(statepath.read_bytes())
assert not state['active_jobs'] and state['active_exec_session'] is None
cold=proofs['923']['RecordedNewtonCold96Proof.json'];contexts=proofs['923']['ColdRejectedContextsProof.json']
assert cold['entire_raw_trajectory_and_intermediate_frames_bit_identical920'] and cold['all257_saved_sources_verified']
assert cold['cold']['native']['own_process_closed'] and cold['cold']['capacity']=='NumFail'
assert proofs['923']['RecordedCheckpointsIdentityProof.json']['raw_checkpoints_entire_bytes_exact920']
assert len(contexts['records'])==55
for n,key in ((925,'RecordedTailSystemProof.json'),(927,'BoundedBranchRecoveryProof.json'),(929,'RecordedTailFineLineProof.json')):
    p=proofs[str(n)][key];assert p['native']['own_process_closed']
    assert p['base_residual_max_error_MPa']==0 and p['refined_matrix_max_relative_row_error']<2e-12
    assert all(c['effort'][0]<=30 for c in p['cases'])
for n,key in ((926,'RecordedTailStressScanProof.json'),(928,'RecordedTailBoundaryScanProof.json')):
    assert proofs[str(n)][key]['native']['own_process_closed']
release={'workbook/output/RC_Section_NDM.xlsm':'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450',
 'workbook/output/VBA_All_Code.txt':'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'}
for name,digest in release.items():assert sha256((repo/name).read_bytes()).hexdigest().upper()==digest

class TestMemory(ctypes.Structure):
    """Проверяет общий запас физической памяти и commit перед упаковкой завершённых данных."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in
        ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DRecordedNewtonTail_2026-11-26'
assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json',
 'PointTraceSourceComposition.json','ParentCodeUpdate.json','ParentCompileClosure.json','CompiledSourceReuse.json',
 'docs/regression/CodeUpdate.json','Input/Reverse1.bin','Input/SourceColdRawHistory.bin','Input/FrozenNewton.bin',
 'Input/ActualRejectedNewton.bin','Input/ActualRefinedNewton.bin','Input/ArcScanFields.bin','Input/ArcScanCases.json',
 'Input/StockColdOptions.bin','SourceRecordedTailSystemProof.json','SourceRejectedContextsProof.json',
 'SourceRecordedTailKinksProof.json','SourceColdProof.json','PriorDirectedRegressionProof.json',
 'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1','Guard-IsolatedOffice.ps1']
raw={}
for n,root in roots.items():
    files=fixed+proof_names.get(n,[])
    files += [p.relative_to(root).as_posix() for folder in ('src','tests','Native','Input','tools/build_workbook')
        for p in (root/folder).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    book=root/'docs/regression/Diagnostic.xlsm'
    if book.is_file():
        with ZipFile(book) as archive:raw[f'{n}/ActualCompiledVBA.bin']=archive.read('xl/vbaProject.bin')
        if n in (925,929):raw[f'{n}/ActualCompiledWorkbook.xlsm']=book.read_bytes()
saved=(roots[923]/'SavedWorkbook.xlsm').read_bytes()
assert sha256(saved).hexdigest().upper()==cold['cold']['saved_book_sha256']
raw['923/ActualSavedWorkbook.xlsm']=saved
with ZipFile(BytesIO(saved)) as book:raw['923/ActualSavedVBA.bin']=book.read('xl/vbaProject.bin')
for name in ('VerifiedColdStressReplayProof.json','Native/EnvelopeEnergy.Accepted.csv','Native/EnvelopeEnergy.Accepted.Operators.csv','Input/Reverse1.bin'):
    raw['Referenced921/'+name]=(base/'Cold921AffineRecoveryColdStress_20261126'/name).read_bytes()
for name in ('Native/Stock.Checkpoints.csv','Native/Stock.RawHistory.bin','AffineRecoveryCold96Proof.json'):
    raw['Referenced920/'+name]=(base/'Cold920AffineRecoveryCold96_20261126'/name).read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()}
archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name

helpers=set('''Stage-TestColdNewtonCapture.py Finalize-TestColdNewtonCapture.py Stage-TestRecordedNewtonCold96.py
Check-TestRecordedNewtonCold96.py Check-TestColdRejectedContexts.py Inspect-TestLiveColdRejectedContexts.py
Check-TestRecordedHistoryStressBridge.py TestReadRecordedTailSystem.bas.txt Stage-TestRecordedTailSystem.py
Check-TestRecordedTailSystem.py Finalize-TestExisting257Project.py Stage-TestRecordedTailStressScan.py
Check-TestRecordedTailStressScan.py Check-TestRecordedTailKinks.py Stage-TestBoundedBranchRecovery.py
Check-TestBoundedBranchRecovery.py Stage-TestRecordedTailBoundaryScan.py Check-TestRecordedTailBoundaryScan.py
Stage-TestRecordedTailFineLine.py Check-TestRecordedTailFineLine.py Check-TestCold32OutgoingStressScan.py
Read-TestColdRawCapture.py Check-TestActualReverseCapture.py Check-TestHardHalfSteps.py Validate-TestHardPrepared.py
Check-TestPersistentHEvidence.py Inspect-TestResolutionVBA.py Canonical-TestCoreVBA.py Compile-TestExternalPoint.ps1'''.split())
helpers.add(Path(__file__).name);todo=list(helpers)
while todo:
    name=todo.pop();p=here/name;assert p.is_file(),name
    for dependency in re.findall(r'[A-Za-z][A-Za-z0-9_-]*\.(?:py|ps1|bas\.txt|cls\.txt)',p.read_text('utf-8-sig')):
        if dependency not in helpers and (here/dependency).is_file():helpers.add(dependency);todo.append(dependency)
for name in sorted(helpers):(out/name).write_bytes((here/name).read_bytes())
summary=dict(proofs=proofs,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive=archive.name,
    raw_archive_bytes=archive.stat().st_size,all_raw_hashes_read_back=True,
    saved_reopened257_component_workbook_in_archive=True,actual_compiled925_and929_workbooks_in_archive=True,
    reference921_tensors_included=True,release_hashes_unchanged=release,
    archived_helpers=len(helpers),free_physical_before=memory.free_physical,free_commit_before=memory.free_page,
    central_surface_continuity_only_not_all_station_continuity=True,
    captured_final_pair54_55_has_exact_accepted25_prefix=True,
    all_intermediate_removed_prefixes_not_independently_replayed=True,
    final_diagnostic_norm_not_minimum_over_all_attempts=True,
    full_acceptance_not_run=True,physical_capacity_certified=False,whole_goal_completed=False,
    no_working_release_main_promotion=True,all_started_stages_terminal=True)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',all_started_stages_terminal=True,
    latest_recorded_newton_tail_evidence=out.relative_to(repo).as_posix(),
    latest_recorded_newton_tail_summary={k:v for k,v in summary.items() if k!='proofs'},
    latest_exact_system925=proofs['925']['RecordedTailSystemProof.json'],
    latest_bounded_branch927=proofs['927']['BoundedBranchRecoveryProof.json'],
    latest_fine_line929=proofs['929']['RecordedTailFineLineProof.json'],
    previous_goal_turn_revalidation='Previous acknowledgement only deferred pure-bending CDP/SP comparison; no numerical progress claimed. Revalidated branch, terminal handles, staged926 and exact925 evidence before continuing.',
    previous_goal_turn_classification='progress:923 cold trace and entire checkpoints bit-exact920,55 real failed Newton contexts captured;924 all25 final states reassembled from exact921 tensors.925 R exact,A1.37e-16, both reference policies fail30.926 localized point5 finite active-set crossing, near direction accurate2.38e-6 relative.927 bounded3 uncommitted corrector gains no root; not promoted.928 central surface stress difference scales from1.34e-5 to1.34e-11MPa as width decreases1e6; no jump at that surface.929 directed fine line outcome preserved separately. Exact books/proofs archived; physical ultimate and full SOW remain open.')
data=(json.dumps(state,ensure_ascii=False,sort_keys=True,indent=2)+'\n').encode('utf-8')
for target in (statepath,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED922-929',len(raw),'RAW FILES;',archive.stat().st_size,'BYTES;',len(helpers),'HELPERS;GOAL ACTIVE')
