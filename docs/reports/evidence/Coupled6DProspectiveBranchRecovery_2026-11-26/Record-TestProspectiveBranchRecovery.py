"""ДЛЯ ТЕСТОВ. Архивирует завершённые локальные проверки активной ветви, сохраняя незавершённый холодный процесс отдельным этапом."""
from pathlib import Path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={914:base/'Cold914ProspectiveBranchMatrix_20261126',915:base/'Cold915ProspectiveBranchRecovery_20261126',
 916:base/'Cold916ProspectiveAffineRoots_20261126',917:base/'Cold917ProspectiveRecoveryRegressions_20261126',
 918:base/'Cold918AffineBranchRecovery_20261126',919:base/'Cold919AffineRecoveryRegressions_20261126'}
names={914:'ProspectiveBranchMatrixProof.json',915:'ProspectiveBranchRecoveryProof.json',916:'ProspectiveAffineRootsProof.json',
 917:'DirectedRegressionProof.json',918:'AffineBranchRecoveryProof.json',919:'DirectedRegressionProof.json'}
proofs={n:json.loads((root/names[n]).read_bytes()) for n,root in roots.items()}
for n in (914,915,916,918):assert proofs[n]['native']['own_process_closed']
for n in (917,919):assert proofs[n]==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
assert proofs[914]['base_residual_maximum_difference_from_actual_capture_MPa']==0
assert proofs[916]['independently_equilibrated_affine_predictions']==5
for n in (915,918):
    assert not proofs[n]['lost_original_cases'] and proofs[n]['actual_prospective_diagnostic_records']>0
    assert next(x for x in proofs[n]['cases'] if x['case']==7)['full_root_found']
class TestMemory(ctypes.Structure):
    """Оставляет резерв для Codex и параллельного холодного Excel перед выделением словаря XZ."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>6*1024**3 and memory.free_page>6*1024**3
out=repo/'docs/reports/evidence/Coupled6DProspectiveBranchRecovery_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
raw={};fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json','CompiledSourceReuse.json',
 'Input/Reverse1.bin','Input/SourceColdRawHistory.bin','Input/ArcScanFields.bin','Input/ArcScanCases.json','Input/FrozenNewton.bin',
 'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestOperationalArcRegressions.ps1','Guard-IsolatedOffice.ps1',
 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1','tools/build_workbook/Run-Audit02ModeTests.ps1']
for n,root in roots.items():
    files=fixed+[names[n]]+[p.relative_to(root).as_posix() for folder in ('src','tests','Native') for p in (root/folder).rglob('*') if p.is_file()]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{n}/{name}']=(root/name).read_bytes()
    if n in (914,915,916,918):
        with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{n}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
for n in (915,918):
    lines=(roots[n]/'Native/Arc32Diagnostics.txt').read_text('cp1251').splitlines();case=None;selected=[]
    for line in lines:
        if line.startswith('CASE;'):case=int(line.split(';')[1])
        if case==7:selected.append(line)
    raw[f'{n}/Case7NativeSummary.txt']=('\n'.join(selected)+'\n').encode('utf-8')
reference=base/'Cold911NewtonRejectedCapture_20261126'
for name in ('NewtonRejectedCaptureProof.json','Native/NewtonRejected.00001.bin','Native/NewtonRejected.00002.bin'):
    raw['Referenced911/'+name]=(reference/name).read_bytes()
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['TestProspectiveBranchMatrix.bas.txt','Stage-TestProspectiveBranchMatrix.py','Check-TestProspectiveBranchMatrix.py',
 'CandidateProspectiveBranchStep.cls.txt','Stage-TestProspectiveBranchRecovery.py','Check-TestProspectiveBranchRecovery.py',
 'Stage-TestProspectiveAffineRoots.py','Check-TestProspectiveAffineRoots.py','Stage-TestAffineBranchRecovery.py','Check-TestAffineBranchRecovery.py',
 'Stage-TestPreparedRegressionRun.py','Check-TestArc56Diagnostics.py','Check-TestCold32OutgoingStressScan.py',
 'Check-TestPersistentHEvidence.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Check-TestHardHalfSteps.py',
 'Validate-TestHardPrepared.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Finalize-TestOperationalArc.py',
 'Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py','Inspect-TestResolutionVBA.py','Canonical-TestCoreVBA.py',
 'Compile-TestExternalPoint.ps1',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(proofs=proofs,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),raw_archive=archive.name,
 raw_archive_bytes=archive.stat().st_size,all_raw_hashes_read_back=True,free_physical_before=memory.free_physical,
 free_commit_before=memory.free_page,all_archived_native_stages_terminal=True,
 current_cold920_explicitly_not_included_as_completed=True,no_clean_timing_claim=True,
 no_working_release_main_promotion=True,physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),latest_prospective_recovery_evidence=out.relative_to(repo).as_posix(),
 latest_prospective_recovery_summary=dict(actual_failed_current_reproduced_bit_exact=True,
    root_case7_full_newly_available=True,all5_original_arc_acceptances_retained=True,
    modified_recovery_regression917=proofs[917],affine_recovery_regression919=proofs[919],
    five_affine_predictions_independently_equilibrated=True,physical_capacity_certified=False),
 previous_goal_turn_classification='progress: actual failed Newton capture911, finite surface kink point5 confirmed913 and extended damping912 not selected. Current continuation914 reproduces exact failed residual and refined matrix,915/918 solve previously failedCase7Full within original control/material budget, all5 original Arc cases retained,917/919 all164PASS,916 five affine roots independently rebuilt. Cold920 is separate ongoing stage; no completion or physical-capacity claim.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED COMPLETED914-919',len(raw),'RAW',archive.stat().st_size,'BYTES;COLD920 NOT CLAIMED COMPLETE;GOAL ACTIVE')
