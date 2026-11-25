"""ДЛЯ ТЕСТОВ. Сохраняет холодный маршрут, протяжённую исходящую ветвь и независимую проверку холодных напряжений."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots={824:base/'Cold824AutomaticLoadingCold_20261125',825:base/'Cold825ExtendedLoadingWalk_20261125',826:base/'Cold826AutomaticColdReplay_20261125'}
selected={824:'OperationalColdProof.json',825:'AutomaticLoadingWalkProof.json',826:'OperationalColdReplayProof.json'}
proofs={str(k):json.loads((roots[k]/v).read_bytes()) for k,v in selected.items()}
cold=proofs['824'];walk=proofs['825'];replay=proofs['826']
assert cold['accepted_states']==51 and cold['exact_prefix_with_warm_input']==43 and cold['arc_recovered_directions']==3
assert cold['cold']['capacity']=='NumFail' and cold['cold']['direct']=='OK' and not cold['cold']['limit']
assert cold['cold']['all_saved_sources_exact']==256 and cold['cold']['results_checksum']
assert cold['cold']['seconds']==301.1875
assert walk['new_accepted_equilibria']==26 and walk['total_calls']==95 and len(walk['all_equilibria']['states'])==69
assert walk['all_half_roots_independently_rebuilt'] and walk['all_steel_strain_limits_verified_at_all_three_stations']
assert replay['accepted_states']==51 and replay['states_above_arc_inner_tolerance']==4
assert all(not p['physical_capacity_certified'] for p in proofs.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={}
for k,r in roots.items():
    value,_=check(r,'HBackendStock' if k==824 else 'RealColdReplay')
    assert value['passed']==(28 if k==824 else 1) and value['failed']==0 and value['own_process_closed'] and not value['timeout600']
    native[str(k)]=value
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
assert not state['active_jobs'] and state['goal_status']=='active' and not state['pause_requested_explicitly']
out=repo/'docs/reports/evidence/Coupled6DLoadingColdBranch_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
stage={824:'Stage-TestAutomaticLoadingCold.py',825:'Stage-TestExtendedLoadingWalk.py',826:'Stage-TestAdaptiveColdReplay.py'}
for k,r in roots.items():
    files=[selected[k],stage[k],'Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json',
        'PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json','docs/regression/CodeUpdate.json',
        'ParentCodeUpdate.json','ParentCompileClosure.json','Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1',
        'Run-TestHMixedStock.ps1','Guard-IsolatedOffice.ps1','Run.ndjson',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
        'tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
    files += ['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin','OperationalArcRatio.bin',
        'NativeTolerance.bin','NativeSolver.json','EnvelopeOptions.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin')]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p=r/name
        if p.is_file():
            target=out/str(k)/name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
for n in ('Stage-TestAutomaticLoadingCold.py','Stage-TestExtendedLoadingWalk.py','Stage-TestAutomaticLoadingWalk.py',
    'Stage-TestAdaptiveColdReplay.py','Stage-TestColdRawReplay.py','Stage-TestEnergyNumerics.py','Check-TestOperationalCold.py',
    'Check-TestAutomaticLoadingWalk.py','Check-TestOperationalColdReplay.py','Check-TestPersistentHEvidence.py',
    'Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Validate-TestHardPrepared.py','Check-TestHardHalfSteps.py',
    'Finalize-TestOperationalArc.py','Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py','Compile-TestExternalPoint.ps1',Path(__file__).name):
    (out/n).write_bytes((here/n).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
(out/'Summary.json').write_text(json.dumps(dict(proofs=proofs,native=native,release_hashes=release,
    cold_progressed_beyond_source43_with_same_maximum_lambda=True,cold_stopped_at96_global_probes=True,
    warm_branch26_not_full_cold_result=True,all_started_stages_terminal=True,
    full_acceptance_completed=False,physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,active_exec_session=None,active_jobs=[],all_started_stages_terminal=True,
    automatic_loading_cold_branch_evidence=out.relative_to(repo).as_posix(),
    latest_verified_cold_source=str(roots[824]),latest_verified_extended_branch=str(roots[825]),
    automatic_loading_cold_summary=dict(states51=True,max_lambda=cold['maximum_accepted_lambda'],last_lambda=cold['last_accepted_lambda'],
        seconds=cold['cold']['seconds'],capacity='NumFail',cause='96 global probes',actual_saved_reopened_components=256),
    automatic_loading_extended_summary=dict(new_equilibria=26,total_calls=95,last_lambda=walk['all_equilibria']['states'][-1]['lambda'],
        maximum_three_station_steel_utilization=walk['maximum_three_station_steel_utilization'],physical_capacity_certified=False),
    previous_goal_turn_classification='progress: automatic revision-bound recent CDP departure and working arc integrated/verified.824 cold actual workflow passes28 checks, directOK, capacityNumFail at96 global probes,51 accepted states versus43; same maximum1.1058869810286156, final1.1058853590245683,301.1875s instrumented, actual saved/reopened256 sources exact.825 unchanged working kernel follows26 new equilibria95 calls to1.09961588357691 from exact43 prefix, all accepted and half roots and all three-station steel limits independently verified.826 independently rebuilds all51 actual cold equilibria; only4 early states above arc inner tolerance. No physical ultimate, uncertainty, mesh/external or full acceptance claimed.',
    next_action='Continue full original goal. Checkpoint automatic CDP departure is verified; production/main/release unchanged. Source824 is authoritative cold NumFail from96 global probes, not lack of equilibrium; warm825 continues26 equilibria to1.09961588357691 with same working kernel. Inspect retaining verified signed-arc continuation after initial admission and earlier use at unique recent loading boundaries, to avoid repeated tiny fixed-projection geometric probes consuming the96 budget. Preserve all existing steel/hoop/physical/peak verification and step/force/gauge criteria; do not merely relabel decline as destruction. Need lambda sensitivity/uncertainty and peak/step/mesh/external controls, then physical admissibility ONQ012 and complete original production/workbook/LoadPaths/UI/CAD/performance/full scoped acceptance. Internal state5e-13 retained because ordinary5e-10 variant822 fails outgoing coordinate; parameters unchanged. New full long suites require scoped consent; no active native jobs.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (state_path,here/'Continuation666_Resume20261123.json'):target.write_bytes(raw)
print('824-826 ARCHIVED',len(pins),'PINS;51 COLD STATES;26 NEW WARM EQUILIBRIA;NO PHYSICAL LIMIT;GOAL ACTIVE')
