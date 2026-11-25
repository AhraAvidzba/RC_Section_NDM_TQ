"""ДЛЯ ТЕСТОВ. Закрепляет фактическое изменение общего численного ядра, направленные проверки и честный результат холодного прогона."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={867:base/'Cold867FixedMetricKernel_20261126',868:base/'Cold868FixedMetricCold_20261126',869:base/'Cold869FixedMetricNextScan_20261126'}
warm=json.loads((roots[867]/'FixedMetricKernelProof.json').read_bytes());cold=json.loads((roots[868]/'FixedMetricColdProof.json').read_bytes())
scan=json.loads((roots[869]/'ColdAcceptedStressContinuityProof.json').read_bytes())
tangent_residual=json.loads((roots[869]/'FailedTangentResidualProof.json').read_bytes())
assert warm['new_accepted_equilibria']==1 and warm['total_calls']==36 and warm['all_new_H_strictly_growing']
assert warm['native']['failed']==0 and warm['native']['own_process_closed']
assert cold['regressions']==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
assert cold['cold'].get('own_process_closed',cold['cold'].get('native',{}).get('own_process_closed'))
assert scan['native']['failed']==0 and scan['native']['own_process_closed']
assert scan['postcommit_stress_drift_within1e_minus11_MPa']
out=repo/'docs/reports/evidence/Coupled6DFixedMetricKernel_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
raw={}
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json','docs/regression/CodeUpdate.json',
 'Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1','Run-TestOperationalArcRegressions.ps1','Run.ndjson',
 'Stage-TestFixedMetricKernel.py','Stage-TestFixedMetricCold.py','Finalize-TestFixedMetricPreparation.py','Align-TestFixedMetricEntries.py',
 'PreparedTestBeforeBlankAlignment.bas','FixedMetricKernelProof.json','FixedMetricColdProof.json','ColdAcceptedStressContinuityProof.json','FailedTangentResidualProof.json',
 'Stage-TestFixedMetricNextScan.py','SourceVerifiedWarm20Proof.json',
 'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
for number,root in roots.items():
    inputs=['Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','ArcScanFields.bin']
    names=fixed+['Input/'+name for name in inputs]+[p.relative_to(root).as_posix() for directory in ('src','tests','Native') for p in (root/directory).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(names)):
        p=root/name
        if p.is_file():raw[f'{number}/{name}']=p.read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()}
archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestFixedMetricKernel.py','Finalize-TestFixedMetricPreparation.py','Align-TestFixedMetricEntries.py','Stage-TestFixedMetricCold.py',
 'Check-TestFixedMetricKernel.py','Check-TestFixedMetricCold.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py',
 'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',
 'Run-TestOperationalArcRegressions.ps1','Stage-TestFixedMetricNextScan.py','Check-TestFixedMetricNextScan.py','Check-TestColdAcceptedStressContinuity.py','Check-TestFixedMetricTangentResidual.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(warm=warm,cold=cold,next_state20_readonly_scan=scan,failed_tangent_residual=tangent_residual,raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
    actual_general_numerical_kernel_changed=True,four_imports_in_actual256_verified=True,other252_components_verified=True,
    fixed_reference_scale_and_refined_outgoing_tangent_in_common_classes=True,all4_pair_gates_use_constant_joint_metric=True,
    original_materials_and_force_tolerances=True,original96_global_call_budget_and30_primary_phase_iterations=True,
    attempts_per_revision6_includes_ratio1_over16=True,one_verified_local_transition_not_ultimate=True,
    next20_state_six_cutbacks_no_equilibrium=True,no_release_or_main_promotion=True,
    physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_fixed_metric_kernel_evidence=out.relative_to(repo).as_posix(),latest_fixed_metric_kernel_proof=warm,latest_fixed_metric_cold_proof=cold,latest_fixed_metric_next20_stress_scan=scan,latest_fixed_metric_failed_tangent_residual=tangent_residual,
    previous_goal_turn_classification='progress:generalNewton/Arc/Deformationkernel867actuallychangedandcompiled;fixedalphaandrefinedoutgoingtangent,4jointpairgates,StepErroradaptation,6cutbackswithin96/30guards;actual256auditPASS4imports252othercomponents;164nativechecksPASS;warm19automatic1verifiednewstate20lambda1.1042633271726336,36calls,fullandhalfstrictforce/gauge/steel/H/metricPASS;next20sixcutbacksfailed equilibrium. Cold868outcomearchivedhonestly;no ultimate or release promotion.')
raw_state=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(raw_state)
print('GENERAL NUMERICAL KERNEL ARCHIVED',len(raw),'RAW FILES',archive.stat().st_size,'BYTES;ONE VERIFIED TRANSITION;GOAL ACTIVE')
