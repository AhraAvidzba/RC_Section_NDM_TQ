"""ДЛЯ ТЕСТОВ. Сохраняет общую метрику и доказанную критическую H-пару, отрицательную регрессию и её исправление, не объявляя итог ещё работающего холодного прогона."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile
repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={874:base/'Cold874SharedCurveHControl20_20261126',875:base/'Cold875SharedCurveRegressions_20261126',876:base/'Cold876VerifiedHCornerPairs20_20261126',877:base/'Cold877VerifiedHRegressions_20261126',878:base/'Cold878CriticalHRecoveryPairs_20261126',879:base/'Cold879CriticalHRegressions_20261126'}
shared=json.loads((roots[874]/'SharedCurveHControlProof.json').read_bytes());paired=json.loads((roots[876]/'VerifiedHCornerPairsProof.json').read_bytes());critical=json.loads((roots[878]/'CriticalHRecoveryPairsProof.json').read_bytes())
assert shared['all14_native_analytical_outcomes_verified'] and paired['verified_working_H_pairs']==critical['verified_working_H_pairs']==1
assert critical['verified_case1_full_half_stresses_and_operators_bit_identical876']
native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];reg={str(n):native(roots[n],'OperationalArcRegressions')[0] for n in (875,877,879)}
assert reg['875']==reg['879']==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
assert reg['877']==dict(passed=23,failed=1,macros=1,timeout600=False,own_process_closed=True)
out=repo/'docs/reports/evidence/Coupled6DCriticalHRecovery_2026-11-26';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
fixed=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json',
 'ParentCodeUpdate.json','ParentCompileClosure.json','docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run-TestOperationalArcRegressions.ps1','Run.ndjson',
 'SharedCurveHControlProof.json','VerifiedHCornerPairsProof.json','CriticalHRecoveryPairsProof.json','PriorVerifiedHCornerPairsProof.json','NegativeMandatoryAllHRegression.json',
 'Stage-TestSharedCurveHControl.py','Finalize-TestSharedCurvePreparation.py','Stage-TestSharedCurveRegressions.py','Stage-TestVerifiedHCornerPairs.py','Align-TestVerifiedHEntry.py','PreparedBeforeVerifiedHBlankAlignment.bas',
 'Stage-TestVerifiedHRegressions.py','Stage-TestCriticalHRecoveryPairs.py','Stage-TestPreparedRegressionRun.py','SourceArcMetricCounterexample.json','SourceSharedCurveHControlProof.json',
 'tools/build_workbook/Run-Audit02ModeTests.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
inputs=['Reverse1.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json','ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','SharedCurveCases.bin']
raw={}
for number,root in roots.items():
    names=fixed+['Input/'+name for name in inputs]+[p.relative_to(root).as_posix() for directory in ('src','tests','Native') for p in (root/directory).rglob('*') if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(names)):
        if (root/name).is_file():raw[f'{number}/{name}']=(root/name).read_bytes()
# Холодный запуск880 ещё работает: сохраняются только неизменные исходники запуска и личность/ресурсы, без промежуточного расчётного результата.
cold=base/'Cold880CriticalHCold96_20261126'
for name in ('Preparation.json','CompiledSourceReuse.json','docs/regression/CodeUpdate.json','Run-TestHMixedStock.ps1','Guard-IsolatedOffice.ps1','Stage-TestCriticalHCold96.py',
 'SourceCriticalHRecoveryPairsProof.json','PriorDirectedRegressionProof.json','Native/HBackendStock.txt.Ownership.json','Native/HBackendStock.txt.memory.json'):
    if (cold/name).is_file():raw['880/'+name]=(cold/name).read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    for name,data in dict(raw,**{'RawFiles.json':(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()}).items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=['Stage-TestSharedCurveHControl.py','Finalize-TestSharedCurvePreparation.py','Check-TestSharedCurveHControl.py','Stage-TestSharedCurveRegressions.py',
 'Stage-TestVerifiedHCornerPairs.py','Align-TestVerifiedHEntry.py','Check-TestVerifiedHCornerPairs.py','Stage-TestVerifiedHRegressions.py',
 'Stage-TestCriticalHRecoveryPairs.py','Check-TestCriticalHRecoveryPairs.py','Stage-TestPreparedRegressionRun.py','Stage-TestCriticalHCold96.py','Check-TestCriticalHCold96.py',
 'Check-TestFixedMetricCold.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py',
 'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(shared_curve_control=shared,verified_H_pair=paired,critical_H_recovery=critical,regressions=reg,
    failed_unconditional_H_pair_regression_preserved_and_fixed=True,known0point9peak_target_and96budget_not_weakened=True,
    raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
    cold880=dict(status='running_no_final_result_in_this_checkpoint',exec_session=9088,source_book='878',global_probe_limit=96,primary_phase_iterations=30,wall_watchdog_seconds=1200,
        no_warm_fields_passed=True,no_clean_speed_claim=True),
    generic_runtime_three_station_steel_limit_coverage_incomplete=['CCoupledNewtonSolver.PrincipalSteelWithinLimits','CCoupledNewtonSolver.TryEnvelopeStep half steel guard','CCoupledDeformationContinuation.BeyondSteelLimit'],
    current_benchmark_full_and_half_roots_all3stations_checked=True,physical_capacity_certified=False,whole_goal_completed=False,no_release_or_main_promotion=True,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,sort_keys=True,indent=2)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()};(out/'Files.json').write_text(json.dumps(top,sort_keys=True,indent=2)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
assert any(job['session']==9088 for job in state['active_jobs'])
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,pause_request_revoked=True,
    latest_critical_H_recovery_evidence=out.relative_to(repo).as_posix(),latest_shared_curve_control_proof=shared,latest_verified_H_pair_proof=paired,
    latest_critical_H_recovery_proof=critical,latest_critical_H_regressions=reg,
    generic_runtime_three_station_steel_limit_audit_open=summary['generic_runtime_three_station_steel_limit_coverage_incomplete'],
    active_cold_source_root='Cold880CriticalHCold96_20261126',active_cold_exec_session=9088,
    previous_goal_turn_classification='progress:sharedCurveLambdaBound in generalNewton/Arc twohalf-gates,14nativeanalyticPASS;874strictfullHroot3itersfoundattrue20butincominghalf27fails.876workingHhalfseedmidpointfull gives strictverifiedfull+half3+2iters,field.003724998joint.003922040,oldlambdaerrorboundwouldreject;all6/56+4/H/all3stationsteel verified. MandatoryeveryH876causedknown0.9peakregression87723PASS1FAIL budget, fixedscope878criticalrecovery/turnonly, preservingordinaryHbehavior/knownpeak/96/30;8782rootsandstressesbitidentical876;879all164PASS. Actualcold880launchedsession9088onexact878,wall1200only,noresultyet. Generic3stationruntime steelhelpergap found/currenttoyall3independentlyverified;physicalultimate/wholegoalopen.')
data=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(data)
print('CRITICAL H RECOVERY ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;164REGRESSIONPASS;COLD880SESSION9088RUNNING;GOALACTIVE')
