"""ДЛЯ ТЕСТОВ. Архивирует завершённый холодный повтор, точный мост напряжений, двухцикл касательных и отклонённый прогноз без уменьшения объёма цели."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from zipfile import ZipFile
from datetime import datetime
import ctypes,json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={904:base/'Cold904ConsistentRefinementCold96_20261126',905:base/'Cold905ConsistentRefinementColdStress_20261126',
       906:base/'Cold906Cold32OutgoingStressScan_20261126',907:base/'Cold907Cold32TangentConeIterations_20261126',
       908:base/'Cold908ResidualArcPredictor_20261126',909:base/'Cold909ResidualArcPredictorBaseline_20261126',
       910:base/'Cold910ResidualArcPredictorRegressions_20261126'}
names={904:'ConsistentRefinementCold96Proof.json',905:'VerifiedColdStressBridgeProof.json',906:'Cold32OutgoingStressScanProof.json',
       907:'Cold32TangentConeIterationsProof.json',908:'ResidualArcPredictorProof.json',909:'ResidualArcPredictorBaselineProof.json',910:'DirectedRegressionProof.json'}
proofs={n:json.loads((root/names[n]).read_bytes()) for n,root in roots.items()}
assert proofs[904]['accepted_states']==32 and proofs[904]['cold']['capacity']=='NumFail' and proofs[904]['cold']['all_saved_sources_exact']==256
assert proofs[905]['entire_raw_trajectory_bit_identical897_and898'] and proofs[905]['no_new_native_replay_needed_for_exact_trajectory']
assert proofs[906]['zero_increment_stress_drift_MPa']==0 and proofs[907]['all12_available']
loss=json.loads((roots[908]/'ArcPredictorRegressionProof.json').read_bytes())
assert not loss['candidate_accepted'] and loss['baseline_solved']==5 and loss['candidate_solved']==4 and [x['case'] for x in loss['lost']]==[1] and not loss['gained']
assert proofs[910]==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
for n in (906,907,908,909):assert proofs[n]['native']['own_process_closed']
class TestMemory(ctypes.Structure):
    """Проверяет RAM и запас commit перед сжатием фактических VBA-проектов большим словарём."""
    _fields_=[('length',ctypes.c_ulong),('load',ctypes.c_ulong)]+[(n,ctypes.c_ulonglong) for n in ('physical','free_physical','page','free_page','virtual','free_virtual','extended')]
memory=TestMemory();memory.length=ctypes.sizeof(memory);assert ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(memory))
assert memory.free_physical>3*1024**3 and memory.free_page>3*1024**3
out=repo/'docs/reports/evidence/Coupled6DTangentCycleAndRejectedPredictor_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompilePointFailure.json','CompileClosure.json',
 'PointTraceSourceComposition.json','docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'Input/Reverse1.bin','Input/SourceColdRawHistory.bin','Input/ArcScanFields.bin','Input/ArcScanCases.json','PriorDirectedRegressionProof.json',
 'ReplaySourceScope.json','SourceStress32Proof.json','SourceOutgoingJacobianProof.json',
 'ConsistentRefinementCold96Proof.json','VerifiedColdStressBridgeProof.json','Cold32OutgoingStressScanProof.json',
 'Cold32TangentConeIterationsProof.json','Cold32CycleLinearAlgebraProof.json','ResidualArcPredictorProof.json',
 'ResidualArcPredictorBaselineProof.json','ArcPredictorRegressionProof.json','DirectedRegressionProof.json',
 'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1','Run-TestOperationalArcRegressions.ps1',
 'Guard-IsolatedOffice.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1','tools/build_workbook/Run-Audit02ModeTests.ps1']
raw={}
for number,root in roots.items():
    files=fixed+[p.relative_to(root).as_posix() for directory in ('src','tests','Native','CompileFailureLocalIdentifiers') for p in (root/directory).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
    if number in (906,907,908,909):
        with ZipFile(root/'docs/regression/Diagnostic.xlsm') as book:raw[f'{number}/ActualCompiledVBA.bin']=book.read('xl/vbaProject.bin')
with ZipFile(roots[904]/'SavedWorkbook.xlsm') as book:raw['904/ActualSavedVBA.bin']=book.read('xl/vbaProject.bin')
reference=base/'Cold898FreshArcColdStress_20261126'
for name in ['Input/Reverse1.bin','VerifiedColdStressReplayProof.json','Native/EnvelopeEnergy.Accepted.csv','Native/EnvelopeEnergy.Accepted.Operators.csv']+[p.relative_to(reference).as_posix() for p in (reference/'src').rglob('*') if p.is_file()]:
    raw['Referenced898/'+name]=(reference/name).read_bytes()
raw['Referenced897/Stock.RawHistory.bin']=(base/'Cold897FreshArcCold96_20261126/Native/Stock.RawHistory.bin').read_bytes()
raw['Referenced901/CCoupledArcContinuation.cls']=(base/'Cold901ConsistentControlRefinement_20261126/src/Solver/CCoupledArcContinuation.cls').read_bytes()
pins={n:sha256(b).hexdigest().upper() for n,b in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=9) as bundle:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,sort_keys=True,indent=2)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);bundle.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as bundle:
    for name,digest in pins.items():assert sha256(bundle.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Check-TestConsistentRefinementCold96.py','Stage-TestConsistentRefinementCold96.py',
 'Check-TestCold904ExactMaterialBridge.py','Check-TestFreshArcCold96.py','Stage-TestFreshArcCold96.py',
 'Check-TestStrictGrowthCold96.py','Stage-TestStrictGrowthCold96.py','Check-TestCriticalHCold96.py','Check-TestFixedMetricCold.py',
 'Check-TestSpatialProjectionCompatibility.py','Check-TestCriticalHColdStress.py',
 'Stage-TestCold32OutgoingStressScan.py','Check-TestCold32OutgoingStressScan.py',
 'Stage-TestCold32TangentConeIterations.py','TestCold32TangentConeIterations.bas.txt','Repair-TestCold32ConeCompilation.py',
 'Align-TestCold32ConeComments.py','Check-TestCold32TangentConeIterations.py','Check-TestCold32CycleLinearAlgebra.py',
 'Stage-TestCold32DirectionalJacobian.py','Align-TestCold32DirectionalComments.py',
 'CandidateArcPredictorDefects.cls.txt','Stage-TestResidualArcPredictor.py','Check-TestResidualArcPredictor.py',
 'Align-TestResidualPredictorComments.py','Stage-TestResidualArcPredictorBaseline.py','Check-TestResidualArcPredictorBaseline.py',
 'Align-TestResidualPredictorBaselineComments.py','Check-TestArcPredictorRegression.py','Stage-TestArc56Diagnostics.py','Check-TestArc56Diagnostics.py',
 'Stage-TestPreparedRegressionRun.py','Check-TestSharedCurveHControl.py','Check-TestPersistentHEvidence.py',
 'Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',
 'Stage-TestEnergyNumerics.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(proofs=proofs,cycle_linear_algebra=json.loads((roots[907]/'Cold32CycleLinearAlgebraProof.json').read_bytes()),
 predictor_rejection=loss,raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
 all_raw_hashes_read_back=True,xz_preset=9,free_physical_before=memory.free_physical,free_commit_before=memory.free_page,
 all_started_native_stages_terminal=True,all164_not_sufficient_for_actual_arc_regression=True,
 full_acceptance_not_run=True,no_new_cold_for_rejected908=True,no_release_main_or_working_source_promotion=True,
 physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes());assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
 latest_tangent_cycle_evidence=out.relative_to(repo).as_posix(),latest_completed_cold904=proofs[904],
 latest_exact_material_bridge905=proofs[905],latest_outgoing_local_stress906=proofs[906],
 latest_tangent_two_cycle907=proofs[907],latest_predictor_candidate_rejected908=loss,
 latest_residual_predictor_regressions910=proofs[910],all_started_stages_terminal=True,
 previous_goal_turn_classification='progress: cold904 completed814.57421875s/directOK/capacityNumFail/max1.1042833676987807; all32 raw fields/fractions/lambdas exact897; actual refined marker observed.905 exact material bridge rebuilt all6/56+4/all3steel without duplicate native replay.906zero drift/all129 local returns/independent yield7.25e-14.907all12 tangent reads verified, persistent two-cycle and geometric row53 error gain1021.83, no physical limit.908actual residual predictor active7cases and4 verified pairs; original9095pairs, candidate loses original event case1; strict regression gateFAIL candidate rejected despite910all164PASS. No promotion. All native stages closed.',
 next_action='Capture the actual rejected Newton iterate, full correction/Jacobian/control and exact accepted material history at a real failing transition; accepted-state32 probes are not a substitute. Independently localize directional derivative and substep/active-branch changes; inspect numerical material error budget against gradient amplification. Do not repeat known tolerance tightening blindly (Progress653/earlier1e-14 tests). Do not promote908 or run cold for it; it lost a valid event. Preserve full SOW, ONQ010/012 and deferred pure bending CDP/SP comparison; goal remains active.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('ARCHIVED',len(raw),'RAW FILES',archive.stat().st_size,'BYTES;908 REJECTED DESPITE164PASS;GOAL ACTIVE;ALL NATIVE TERMINAL')
