"""ДЛЯ ТЕСТОВ. Архивирует проверку численной точности, локализацию перехода и рабочее восстановление от прогноза802–812."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={802:'Cold802NestedMaterialAccuracy_20261125',803:'Cold803NestedAccuracyEndpoint_20261125',
 804:'Cold804HistoryAccuracyEndpoint_20261125',805:'Cold805EndpointLineWitness_20261125',
 806:'Cold806EndpointPointWitness_20261125',807:'Cold807PredictorLoadingEvent_20261125',
 808:'Cold808PredictorEventCold_20261125',809:'Cold809PredictorEventFourCalls_20261125',
 810:'Cold810PredictorEventFiveCalls_20261125',811:'Cold811PredictorEventColdReplay_20261125',
 812:'Cold812PredictorEventOriginalState_20261125'}
roots={k:base/v for k,v in names.items()}
selected={802:['OperationalColdProof.json'],803:['OperationalColdReplayProof.json','ColdEndpointArcProof.json'],
 804:['ColdEndpointArcProof.json'],805:['EndpointLineWitnessProof.json'],806:['EndpointPointWitnessProof.json'],
 807:['ColdEndpointArcProof.json'],808:['OperationalColdProof.json'],809:['ColdEndpointArcProof.json'],
 810:['ColdEndpointArcProof.json'],811:['OperationalColdReplayProof.json'],812:['ColdEndpointArcProof.json']}
proofs={str(k):{n:json.loads((roots[k]/n).read_bytes()) for n in ns} for k,ns in selected.items()}
assert proofs['802']['OperationalColdProof.json']['accepted_states']==45
assert proofs['803']['OperationalColdReplayProof.json']['source_cold_exact']
assert not proofs['803']['ColdEndpointArcProof.json']['returned_verified_candidate']
assert not proofs['804']['ColdEndpointArcProof.json']['returned_verified_candidate']
assert proofs['805']['EndpointLineWitnessProof.json']['original804_arc_packet_bitidentical']
assert proofs['806']['EndpointPointWitnessProof.json']['predictor_brackets_real_loading_transition']
assert proofs['807']['ColdEndpointArcProof.json']['returned_verified_candidate']
assert not proofs['809']['ColdEndpointArcProof.json']['returned_verified_candidate']
assert proofs['810']['ColdEndpointArcProof.json']['returned_verified_candidate']
assert proofs['809']['ColdEndpointArcProof.json']['arc_call_allowance']==4
assert proofs['810']['ColdEndpointArcProof.json']['arc_call_allowance']==5
assert proofs['812']['ColdEndpointArcProof.json']['returned_verified_candidate']
assert not proofs['812']['ColdEndpointArcProof.json']['material_state_tolerance_tightened_only']
assert proofs['811']['OperationalColdReplayProof.json']['source_cold_exact']
assert proofs['811']['OperationalColdReplayProof.json']['accepted_states']==proofs['808']['OperationalColdProof.json']['accepted_states']
assert not any(p['physical_capacity_certified'] for d in proofs.values() for p in d.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for k in roots:
 jobs=['HBackendStock','PriorBranchMaterials','OperationalArcRegressions'] if k==802 else ['RealColdReplay']
 if k==804:jobs.append('PriorBranchMaterials')
 if k==807:jobs.append('OperationalArcRegressions')
 if k==808:jobs=['HBackendStock']
 for job in jobs:
  n,_=check(roots[k],job);assert n['failed']==0 and n['own_process_closed'] and not n['timeout600']
  if job=='OperationalArcRegressions':assert n['passed']==164 and n['macros']==5
  if job=='PriorBranchMaterials':assert n['passed']==87 and n['macros']==3
  native[f'{k}/{job}']=n
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
assert not state['active_jobs']
out=repo/'docs/reports/evidence/Coupled6DPredictorLoadingEvent_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for k,r in roots.items():
 files=selected[k]+['Preparation.json','CompilePoint.log','CompilePointClosure.json',
  'PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json',
  'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
  'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
  'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
  'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1','Run-TestPriorBranchMaterials.ps1',
  'Guard-IsolatedOffice.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
 files+=['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin','OperationalArcRatio.bin',
  'NativeTolerance.bin','NativeSolver.json','EnvelopeOptions.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin')]
 files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
 for relative in files:
  p=r/relative
  if p.is_file():
   dest=out/str(k)/relative;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
helpers=['Stage-TestColdNestedAccuracy.py','Stage-TestColdEndpointArc.py','Check-TestColdEndpointArc.py',
 'Stage-TestEndpointHistoryAccuracy.py','Stage-TestEndpointLineWitness.py','Check-TestEndpointLineWitness.py',
 'Stage-TestEndpointPointWitness.py','Check-TestEndpointPointWitness.py','Stage-TestPredictorLoadingEvent.py',
 'Stage-TestPredictorEventControls.py','Stage-TestPredictorEventOriginalState.py','Stage-TestAdaptiveColdReplay.py','Check-TestOperationalCold.py',
 'Check-TestOperationalColdReplay.py','Finalize-TestOperationalArc.py','Run-TestOperationalArcRegressions.ps1',
 'Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Compile-TestExternalPoint.ps1','Finalize-TestStockYield.py',
 'Audit-TestPrivateCoupledBook.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py',
 'Check-TestPersistentHEvidence.py','Validate-TestHardPrepared.py','Check-TestHardHalfSteps.py',Path(__file__).name]
for n in helpers:(out/n).write_bytes((here/n).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
(out/'Summary.json').write_text(json.dumps(dict(proofs=proofs,native=native,release_hashes=release,
 exact_endpoint25_plus1_event_recovery_confirmed=True,all_started_stages_terminal=True,
 physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
cold=proofs['808']['OperationalColdProof.json']
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
 predictor_loading_event_evidence='docs/reports/evidence/Coupled6DPredictorLoadingEvent_2026-11-25',
 cold808_final_summary={k:cold[k] for k in ('accepted_states','maximum_accepted_lambda','last_accepted_lambda','arc_calls','arc_recovered_directions')},
 previous_goal_turn_classification='progress: paired nested precision80287materials164regressions,45cold equilibria803; state precision80487materials still fails endpoint.805 exact Newton witness linear residual3.47e-23 but actual row53 reverses,25/+1 dominates.806 independent yield/branch shows actual loading transition25/+1 crossed by failed correction, original predictors bracket it but both root failures prevent any event search.807 permits event solve from predictor bracket only; restores verified event after5calls, no history acceptance/criteria change.8094call negative and8105call bitexact positive pass.807reg164 pass.808 actualcold still NumFail43/max1.1058869810286156/464.796875s;811 independent43equilibria verified.812 isolates same CArc fix with original802 state tolerance: also verified5calls. No higher lambda/physical capacity/speedup; production unchanged.',
 next_action='Use812 as isolated current candidate (same802 material state tolerance, only verified CArc fix), not unsuccessful804 state tightening. Cold808 still NumFail43/max1.1058869810286156 at95probes, no ultimate or performance improvement. Reproduce actual remaining cold808 endpoint and verify outgoing material branch/control after a resolved event; determine whether incoming linear arc hyperplane can track the outgoing direction before further cold runs. Finite local alpha alone does not exclude a nonsmooth corner, but a corner/local crack is not confirmed full RC failure; require actual post-transition equilibria and physical criteria. Keep30/96, force/gauges/field0.005/lambda checks. Need confirmed capacity, cold reproducibility, step/mesh/external comparisons, all LoadPaths, original workbook/UI/CAD deliverables and scoped full acceptance. Goal active.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (state_path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('802-812 ARCHIVED',len(pins),'PINS; EVENT FROM PREDICTOR VERIFIED; COLD808',cold['accepted_states'],cold['maximum_accepted_lambda'],'GOAL ACTIVE')
