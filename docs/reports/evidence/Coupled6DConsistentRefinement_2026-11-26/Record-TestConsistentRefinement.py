"""ДЛЯ ТЕСТОВ. Сохраняет проверку реальных производных и парный контроль одного изменения Newton, не выдавая неактивный путь за улучшение."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile,numpy as np

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={899:base/'Cold899Cold32DirectionalJacobian_20261126',900:base/'Cold900Cold32OutgoingJacobian_20261126',
       901:base/'Cold901ConsistentControlRefinement_20261126',902:base/'Cold902ConsistentRefinementBaseline_20261126'}
proofs={n:json.loads((root/('Cold32DirectionalJacobianProof.json' if n==899 else 'Cold32OutgoingJacobianProof.json' if n==900 else 'ConsistentRefinementProof.json')).read_bytes()) for n,root in roots.items()}
assert all(p['native']['own_process_closed'] for p in proofs.values())
regroot=base/'Cold903ConsistentRefinementRegressions_20261126';coldroot=base/'Cold904ConsistentRefinementCold96_20261126'
reg,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](regroot,'OperationalArcRegressions')
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
assert proofs[900]['outgoing_working_predictor']['relative_difference_from_independent_linear_solve']<1e-8
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
def TestCases(root):
    """Читает исходные Double-поля без округления для парного сравнения кандидата и контрольной книги."""
    r=R((root/'Native/Cold32RefinementSteps.bin').read_bytes());meta=r.VectorTest();projection=r.VectorTest();cases=[]
    for _ in range(int(meta[0])):
        info=r.VectorTest();effort=r.VectorTest();field=r.VectorTest() if info[2] else None;actual=r.VectorTest() if info[2] else None
        cases.append((info,effort,field,actual))
    assert r.position==len(r.raw)
    return cases
left,right=TestCases(roots[901]),TestCases(roots[902]);assert len(left)==len(right)==3
for a,b in zip(left,right):
    np.testing.assert_array_equal(a[0][:4],b[0][:4]);np.testing.assert_array_equal(a[1][:2],b[1][:2])
    np.testing.assert_array_equal(a[2],b[2]);np.testing.assert_array_equal(a[3],b[3])
assert all(c['solved'] and not c['refined_tangent_used'] for n in (901,902) for c in proofs[n]['independent_cold32_probes'])
comparison=dict(all_three_lambda_fields_and_resultants_bit_identical=True,iterations_and_evaluations_identical=True,
    refinement_not_activated_on_any_local_probe=True,no_robustness_or_speed_improvement_proven=True,
    physical_capacity_certified=False)
(roots[901]/'BaselineComparison.json').write_text(json.dumps(comparison,sort_keys=True,indent=2)+'\n',encoding='utf-8')
class TestMemory(ctypes.Structure):
    """Проверяет резерв RAM и commit перед сжатием четырёх фактических VBA-проектов."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory)
assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DConsistentRefinement_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
raw={};fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json',
 'PointTraceSourceComposition.json','docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'Cold32DirectionalJacobianProof.json','Cold32OutgoingJacobianProof.json','ConsistentRefinementProof.json','BaselineComparison.json',
 'Input/Reverse1.bin','SourceStress32Proof.json','Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1',
 'Guard-IsolatedOffice.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1','tools/build_workbook/Run-Audit02ModeTests.ps1']
for number,root in roots.items():
    names=fixed+[p.relative_to(root).as_posix() for directory in ('src','tests','Native') for p in (root/directory).rglob('*') if p.is_file()]
    for name in sorted(set(names)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
raw['Referenced895/CCoupledNewtonSolver.cls']=(base/'Cold895FreshSpatialArc_20261126/src/Solver/CCoupledNewtonSolver.cls').read_bytes()
raw['Referenced898/VerifiedColdStressReplayProof.json']=(base/'Cold898FreshArcColdStress_20261126/VerifiedColdStressReplayProof.json').read_bytes()
for name in fixed+[p.relative_to(regroot).as_posix() for p in (regroot/'Native').rglob('*') if p.is_file()]:
    if (regroot/name).is_file():raw['903/'+name]=(regroot/name).read_bytes()
for name in ('Preparation.json','PriorDirectedRegressionProof.json','Run-TestHMixedStock.ps1',
             'Native/HBackendStock.txt.Ownership.json','Native/HBackendStock.txt.memory.json'):
    raw['904/StaticLaunchOnly/'+name]=(coldroot/name).read_bytes()
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestCold32DirectionalJacobian.py','Check-TestCold32DirectionalJacobian.py','Align-TestCold32DirectionalComments.py',
 'Stage-TestCold32OutgoingJacobian.py','Check-TestCold32OutgoingJacobian.py','Align-TestCold32OutgoingComments.py',
 'Stage-TestConsistentRefinement.py','TestCold32RefinementSteps.bas.txt','Check-TestConsistentRefinement.py',
 'Align-TestConsistentRefinementComments.py','Align-TestConsistentBaselineComments.py','Stage-TestPreparedRegressionRun.py',
 'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py',
 'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py','Stage-TestEnergyNumerics.py',Path(__file__).name]
helpers+=['Stage-TestConsistentRefinementCold96.py','Check-TestConsistentRefinementCold96.py',
 'Stage-TestFreshArcCold96.py','Check-TestFreshArcCold96.py','Stage-TestStrictGrowthCold96.py',
 'Check-TestStrictGrowthCold96.py','Check-TestCriticalHCold96.py','Check-TestFixedMetricCold.py',
 'Check-TestSpatialProjectionCompatibility.py']
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(proofs=proofs,paired_comparison=comparison,directed_regressions903=reg,
 cold904_only_static_launch_not_outcome=True,raw_archive=archive.name,raw_files=len(raw),
 raw_bytes=sum(map(len,raw.values())),all_raw_hashes_read_back=True,xz_preset=9,free_physical_before=memory.free_physical,
 free_commit_before=memory.free_page,whole_goal_completed=False,no_release_or_main_promotion=True,
 test_wrapper_first_failure900='Assertion before data read: expected four indentation spaces in code anchor, actual one; anchor corrected, numerical data/book/tolerances unchanged')
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
 latest_consistent_refinement_evidence=out.relative_to(repo).as_posix(),latest_consistent_refinement_local_comparison=comparison,
 latest_cold32_outgoing_predictor=proofs[900]['outgoing_working_predictor'],
 latest_consistent_refinement_regressions903=reg,
 previous_goal_turn_classification='progress:897 completed cold/directOK/capacityNumFail32states max1.1042833676987807 in820.4443359375s;898all32stresses verified;899actual directional derivative differs by side at accepted boundary;900operative1/1024 predictor independently verified rel1.174e-12.901one existing Newton refinement gate corrected,902original baseline; three strict roots bitidentical/4,4,1iterations/new refinement inactive. No robustness or capacity claim. Later directed regressions/cold activity preserved by active_jobs.',
 next_action=state.get('next_action','Verify directed regression903 then cold904 before claiming effect of consistentControl refinement.'))
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED',len(raw),'RAW FILES',archive.stat().st_size,'BYTES;LOCAL A/B EXACT;NEW REFINEMENT NOT ACTIVATED;GOAL ACTIVE')
