"""ДЛЯ ТЕСТОВ. Сохраняет завершённый исходящий участок813–815 и точную точку возобновления перед явно запрошенной паузой."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={813:base/'Cold813ActualRemainingEndpoint_20261125',814:base/'Cold814ActiveLoadingCoordinate_20261125',815:base/'Cold815ActiveLoadingNeighborhood_20261125'}
selected={813:'ColdEndpointArcProof.json',814:'ActiveLoadingCoordinateProof.json',815:'ActiveLoadingCoordinateProof.json'}
proofs={str(k):json.loads((roots[k]/n).read_bytes()) for k,n in selected.items()}
assert proofs['813']['source_states']==43 and not proofs['813']['returned_verified_candidate']
assert not proofs['814']['pair']['verified'] and not proofs['814']['root_witnesses']
assert proofs['815']['pair']['verified'] and set(proofs['815']['root_witnesses'])=={'Full','Half'}
assert proofs['815']['pair']['both_lambdas_below_source']
assert proofs['815']['pair']['mechanical_direction_cosine']>0 and proofs['815']['pair']['incoming_joint_direction_cosine']<0
assert proofs['815']['half_effort'][0]+proofs['815']['full_effort'][0]==15
assert not any(p['physical_capacity_certified'] for p in proofs.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for k,r in roots.items():
 n,_=check(r,'RealColdReplay');assert n==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
 native[str(k)]=n
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
assert state['pause_requested_explicitly'] and not state['pause_request_revoked'] and not state['active_jobs']
out=repo/'docs/reports/evidence/Coupled6DOutgoingLoadingPause_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for k,r in roots.items():
 files=[selected[k],'Preparation.json','CompilePoint.log','CompilePointClosure.json','PointTraceSourceComposition.json',
  'CompiledSourceReuse.json','CompiledSourceReuseProof.json','docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
  'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
  'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
  'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
 files+=['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin','OperationalArcRatio.bin',
  'NativeTolerance.bin','NativeSolver.json','EnvelopeOptions.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin')]
 files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
 for relative in files:
  p=r/relative
  if p.is_file():
   dest=out/str(k)/relative;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
helpers=['Stage-TestColdEndpointArc.py','Check-TestColdEndpointArc.py','Stage-TestPredictorLoadingEvent.py',
 'Stage-TestActiveLoadingCoordinate.py','Stage-TestActiveLoadingNeighborhood.py','Check-TestActiveLoadingCoordinate.py',
 'Finalize-TestOperationalArc.py','Stage-TestEnergyNumerics.py','Stage-TestColdRawReplay.py','Compile-TestExternalPoint.ps1',
 'Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py',
 'Check-TestPersistentHEvidence.py','Validate-TestHardPrepared.py','Check-TestHardHalfSteps.py',Path(__file__).name]
for n in helpers:(out/n).write_bytes((here/n).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
(out/'Summary.json').write_text(json.dumps(dict(proofs=proofs,native=native,release_hashes=release,
 independently_verified_outgoing_equilibrium_pair=True,automatic_outgoing_coordinate_not_implemented=True,
 all_started_stages_terminal=True,user_requested_pause=True,goal_pause_tool_pending=True,
 physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='pause_requested',
 pause_checkpoint_ready=True,goal_pause_tool_pending=True,active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
 outgoing_loading_pause_evidence='docs/reports/evidence/Coupled6DOutgoingLoadingPause_2026-11-25',
 outgoing815_summary=dict(source_lambda=proofs['815']['meta'][2],half_lambda=proofs['815']['meta'][10],full_lambda=proofs['815']['meta'][11],
  pair=proofs['815']['pair'],shared_primary=15,material_history_not_accepted=True),
 previous_goal_turn_classification='progress:813 reproduces actual remaining43-state cold808 endpoint with both arc roots failing and no new resolved event.814 mapped outgoing coordinate candidates rejected by seed-relative locality before force recheck; no equilibria claimed.815 uses original full-predictor neighborhood fixed before solves, half1.5/full3, and existing CDP coordinate25/+1 with unchanged strict force/gauges/30; independently verifies both roots, field half error3.915e-5, lambda half error1.321e-12<8.424e-10, mechanical cosine+0.4017 versus incoming joint-0.6999, total15primary. Exact source history retained. No automatic continuation/physical capacity implemented. User explicitly requested pause at logical stage completion; no new research/solver changes after request, only finishing815 and archiving.',
 next_action='AFTER COMMIT AND SUCCESSFUL PUSH: call update_goal paused, report returned status and STOP. Resume only after explicit user request. On resume, retain full original objective.813-815 evidence establishes outgoing equilibria on cold808 exact endpoint: source1.1058869810286156, half1.1058869198622188, full1.105886858697143; surface25/+1 positive targets. Old incoming arc has joint cosine-0.6999, mechanical forward+0.4017. Need automatic outgoing-coordinate continuation with recent resolved station selected from solver/controller state, never hardcode25; use neighborhood from original predictor before solving, strict full/half field0.005 and lambda errors, force/gauges,30/96 and exact restoration; verify actual material limits, steel/hoops, LoadPaths.815 is TEST ENTRY ONLY, no physical peak/capacity claim; lambda uncertainty, wider postbranch, cold repeat, mesh/step/external controls, production integration/workbook, performance and full scoped acceptance remain incomplete. Prefer original state tolerance as812 unless new evidence justifies804 tightening. Do not restart cold808 or mark failure merely from NumFail.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (state_path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('813-815 ARCHIVED',len(pins),'PINS; OUTGOING PAIR VERIFIED; PAUSE REQUEST READY; GOAL TOOL STILL PENDING')
