"""ДЛЯ ТЕСТОВ. Сохраняет отрицательный бюджетный опыт и проверку маршрутов управления с исходными Double и контрольными суммами."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from io import BytesIO
from datetime import datetime
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={854:'Cold854ExtendedColdBudget192_20261125',855:'Cold855BalancedInitialControl_20261125',856:'Cold856ExtendedColdStressReplay_20261125',857:'Cold857SpatialDeclineRecovery_20261125'}
roots={n:base/name for n,name in names.items()}
proofs={854:'ExtendedColdBudgetProof.json',855:'BalancedInitialColdProof.json',856:'VerifiedColdStressReplayProof.json',857:'SpatialDeclineColdProof.json'}
results={n:json.loads((roots[n]/name).read_bytes()) for n,name in proofs.items()}
assert results[854]['accepted_states']==64 and results[854]['negative96_exact_prefix_states']==49
assert results[854]['maximum_accepted_lambda']==1.1058869810286156
assert results[856]['source_states']==64 and results[856]['all_working_sources_exact_source']
assert results[855]['accepted_states']==7
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](roots[857],'OperationalArcRegressions')
assert native==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
for n in (854,855):assert results[n]['cold']['native']['own_process_closed'] and not results[n]['cold'].get('workflow_result_missing',False)
assert results[857]['cold']['workflow_result_missing'] and results[857]['cold']['timeout_seconds']==600
assert results[857]['cold']['own_process_closed'] and results[857]['accepted_states']==19
out=repo/'docs/reports/evidence/Coupled6DColdBudgetAndControls_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1',
 'Run-TestHMixedStock.ps1','Run.ndjson','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
 'src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledElasticState.cls',
 'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
 'src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas']
stages={854:'Stage-TestExtendedColdBudget.py',855:'Stage-TestBalancedInitialControl.py',856:'Stage-TestVerifiedColdStressReplay.py',857:'Stage-TestSpatialDeclineRecovery.py'}
extra={854:[],855:['BalancedInitialControlProof.json'],856:['SourceColdProof.json'],857:['SpatialDeclineRoutingProof.json','NegativeInitialControlColdProof.json']}
inputs=['Reverse1.bin','SourceColdRawHistory.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json',
 'ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin',
 'InitialControl.Fields.bin','InitialControl.Cases.json']
raw={}
for n,r in roots.items():
    files=common+[proofs[n],stages[n]]+extra[n]+['Input/'+name for name in inputs]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
pins={n:sha256(data).hexdigest().upper() for n,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    assert len(z.getnames())==len(pins)+1
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=list(stages.values())+['Check-TestExtendedColdBudget.py','Check-TestBalancedInitialControl.py',
 'Check-TestVerifiedColdStressReplay.py','Check-TestSpatialDeclineRouting.py','Stage-TestEnergyNumerics.py',
 'Stage-TestColdRawReplay.py','Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py',
 'Read-TestColdRawCapture.py','Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py',
 'Inspect-TestResolutionVBA.py','Run-TestOperationalArcRegressions.ps1',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
summary=dict(results={str(n):v for n,v in results.items()},targeted_regressions=native,raw_archive=archive.name,
 raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),budget_doubling_not_physical_progress=True,
 spatial_decline_cold_timeout600_not_NumFail_or_failure_certificate=True,
 CDP_versus_plane_SP_pure_bending_explicitly_deferred_by_user=True,full_acceptance_not_run=True,
 no_working_or_release_promotion=True,physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
 pause_request_revoked=True,pause_after_current_stage=False,all_started_stages_terminal=not state.get('active_jobs',[]),
 latest_cold_budget_and_controls_evidence=out.relative_to(repo).as_posix(),
 latest_spatial_decline_cold_source=str(roots[857]),latest_spatial_decline_cold_summary=results[857],
 extended192_did_not_improve_load_maximum=True,verified_extended_cold_stress_source=str(roots[856]),
 deferred_CDP_versus_SP_pure_bending_check=dict(requested='2026-11-25',after='primary coupled convergence and confirmed ultimate problems',status='not run',plan='docs/ValidationPlan.md'),
 next_action='Continue original goal;DO NOT PAUSE. Inspect archived854–857 outcomes before changing budget/control again.192budget64states SAMEmax1.1058869810286156,no meaningful physical progress;all64forces+56/4equations+3stationsteelindepverified856.855initialbalancedselector7readonlycasesPASS butcoldstoppedat7 on unverifiedcumulativedecline;857 only3TraceroutinggatesallowexistingReplayUnverifiedDeclineforspatialgeneralizedtwist,allothermathunchanged and164targetedregressionspassed.857coldTIMEOUT600,19acceptedstates,max1.1042833676987807,NO finalResults/savedbook and NO returnedNumFail;exactown10332closedbywatchdog. Latestdefaultauthoritativecold853remains49/max1.1058869810286156/NumFail. Do NOT promote857orblindlyincreaseitsruntimebudget;inspectcoldcheckpointcost (12shoulderprobes/localpeakverification,slowfailedNewtonsteps).858frozen7stresscompilationmaystillberunning;checkactivejobs thenfinishactual256audit/native/exportverifier.855/857distinctstressesnotyetindependentlyverified. NO physicalcapacity certificate,fullacceptance or working/releasepromotion. PurebendingCDP-vs-SP explicitlyDEFERREDuntilprimaryproblemsresolved. Preserve main/release/userdata.')
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('COLD BUDGET/CONTROLS ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;NO LIMIT CERTIFICATE;GOAL ACTIVE')
