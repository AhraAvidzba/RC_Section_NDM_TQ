"""ДЛЯ ТЕСТОВ. Сохраняет проверку холодных напряжений, отрицательный дуговой опыт и отсутствие дрейфа принятого CDP-поля."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from io import BytesIO
from datetime import datetime
import json,tarfile

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={858:'Cold858BalancedInitialStressReplay_20261125',859:'Cold859SpatialDeclineStressReplay_20261125',860:'Cold860BalancedImmutableStressReplay_20261125',861:'Cold861Cold19AutomaticArcRestart_20261125',862:'Cold862ColdAcceptedStressContinuity_20261125'}
roots={n:base/name for n,name in names.items()}
proofs={859:'VerifiedColdStressReplayProof.json',860:'VerifiedColdStressReplayProof.json',861:'AutomaticLoadingWalkProof.json',862:'ColdAcceptedStressContinuityProof.json'}
result={n:json.loads((roots[n]/name).read_bytes()) for n,name in proofs.items()}
assert result[859]['source_states']==19 and result[860]['source_states']==7
assert result[861]['new_accepted_equilibria']==0 and result[861]['total_calls']==12
assert result[861]['independently_rebuilt_half_root_count']==0 and not result[861]['all_half_roots_independently_rebuilt']
assert result[862]['maximum_postcommit_concrete_stress_difference_MPa']==result[862]['maximum_equation_change_MPa']==0
native={};helper=run_path(str(here/'Check-TestPersistentHEvidence.py'))
for n in (859,860,861,862):
    native[n],_=helper['TestNative'](roots[n],'RealColdReplay')
    assert native[n]==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
audit=json.loads((roots[858]/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
out=repo/'docs/reports/evidence/Coupled6DColdStressContinuity_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
common=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
 'CompiledSourceReuse.json','CompiledSourceReuseProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
 'docs/regression/CodeUpdate.json','Compile-TestExternalPoint.ps1','Guard-IsolatedOffice.ps1','Run-TestRealColdReplay.ps1','Run.ndjson',
 'src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
 'src/Solver/CCoupledElasticState.cls','src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
 'src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas','tools/build_workbook/Run-Audit02ModeTests.ps1']
inputs=['Reverse1.bin','SourceColdRawHistory.bin','EnvelopeOptions.bin','NativeTolerance.bin','NativeSolver.json',
 'ArcCallAllowance.bin','OperationalArcRatio.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin',
 'InitialControl.Fields.bin','InitialControl.Cases.json','ArcScanFields.bin']
extra=['SourceColdProof.json','SourceColdStressProof.json','SourceColdOutcome.json']
stages=['Stage-TestVerifiedColdStressReplay.py','Stage-TestImmutableColdStressReplay.py','Stage-TestColdRestartAutomaticWalk.py','Stage-TestColdAcceptedStressContinuity.py']
raw={}
for n,r in roots.items():
    files=common+extra+stages+([proofs[n]] if n in proofs else [])+['Input/'+name for name in inputs]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():raw[f'{n}/{name}']=p.read_bytes()
pins={name:sha256(data).hexdigest().upper() for name,data in raw.items()};archive=out/'RawEvidence.tar.xz'
with tarfile.open(archive,'w:xz',preset=6) as z:
    content=dict(raw);content['RawFiles.json']=(json.dumps(pins,indent=2,sort_keys=True)+'\n').encode()
    for name,data in content.items():
        info=tarfile.TarInfo(name);info.size=len(data);z.addfile(info,BytesIO(data))
with tarfile.open(archive,'r:xz') as z:
    assert len(z.getnames())==len(pins)+1
    for name,digest in pins.items():assert sha256(z.extractfile(name).read()).hexdigest().upper()==digest,name
helpers=stages+['Check-TestVerifiedColdStressReplay.py','Check-TestAutomaticLoadingWalk.py','Check-TestColdAcceptedStressContinuity.py',
 'Stage-TestColdYieldCoordinateCapture.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py',
 'Check-TestPersistentHEvidence.py','Check-TestActualReverseCapture.py','Read-TestColdRawCapture.py',
 'Check-TestHardHalfSteps.py','Validate-TestHardPrepared.py','TestUnacceptedMechanics.py',Path(__file__).name]
for name in sorted(set(helpers)):(out/name).write_bytes((here/name).read_bytes())
summary=dict(results={str(n):v for n,v in result.items()},native={str(n):v for n,v in native.items()},
 compilation858_actual256_one_test_import=True,compilation858_not_used_as_NDS_or_capacity_result=True,
 actual_compiled_book_immutable_reuse859_through862=True,zero_new_equilibria861=True,
 no_postcommit_stress_or_weak_equation_drift862=True,raw_archive=archive.name,raw_files=len(raw),raw_bytes=sum(map(len,raw.values())),
 previous855_and857_independent_stress_replay_now_complete=True,full_cold_workflow857_still_timeout600=True,
 more_detailed_yield_diagnostic863_not_completed_by_this_archive=True,no_working_or_release_promotion=True,
 CDP_versus_SP_pure_bending_deferred=True,physical_capacity_certified=False,whole_goal_completed=False,goal_active=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
top={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(top,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
 pause_request_revoked=True,pause_after_current_stage=False,latest_cold_stress_continuity_evidence=out.relative_to(repo).as_posix(),
 latest_cold19_arc_restart_summary=dict(new_equilibria=0,calls=12,ratios=[1,.5,.25,.125],physical_capacity_certified=False),
 latest_cold19_postcommit_drift=dict(stress_MPa=0,weak_equation_MPa=0,last_equilibrium_residual_MPa=result[862]['after_commit_zero_increment_equilibrium']['residual_MPa']),
 previous855_and857_independent_stresses_now_verified=True,
 previous_goal_turn_classification='progress:verified 64cold stresses and negative192budget outcome;testedbalancedinitialmeasureandthree routinggates;164regressionsPASS butcold857timeout600at19;archived359files andpushed4e43. Currentturn:859/860bothdistincthistoriesindepstressverified;858actual256audited;861all4arcattemptsfailedwithoutacceptance;862excludedpostcommitstressdriftandresidualfloorat19,found10sampledyieldcrossings. ExistingprivateVBAprocedurescanbecalledfromisolatedPSafterpurefrozenexportwithoutrecompilingbook;863detailedscalarcapturelaunchedtoidentifycause. Notmerestatecountprogress;NO confirmedphysicalcapacity.')
encoded=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (path,here/'Continuation666_Resume20261123.json'):target.write_bytes(encoded)
print('COLD STRESSES/CONTINUITY/NEGATIVE ARC ARCHIVED',len(raw),'FILES',archive.stat().st_size,'BYTES;NO LIMIT CERTIFICATE;GOAL ACTIVE')
