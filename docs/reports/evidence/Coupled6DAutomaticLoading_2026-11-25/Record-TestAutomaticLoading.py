"""ДЛЯ ТЕСТОВ. Сохраняет проверенную автоматическую исходящую ветвь, отрицательные контроли и точку продолжения полной цели."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={816:'Cold816RecentLoadingBoundary_20261125',817:'Cold817RecentLoadingBoundary_20261125',
    818:'Cold818AutomaticLoadingCoordinate_20261125',819:'Cold819AutomaticLoadingArc_20261125',
    820:'Cold820LoadingArcBudget4_20261125',821:'Cold821PriorityLoadingArc_20261125',
    822:'Cold822LoadingOriginalState_20261125',823:'Cold823AutomaticLoadingWalk_20261125'}
roots={k:base/v for k,v in names.items()};selected={818:'AutomaticLoadingCoordinateProof.json',819:'AutomaticLoadingArcProof.json',
    820:'AutomaticLoadingArcProof.json',821:'AutomaticLoadingArcProof.json',822:'AutomaticLoadingArcProof.json',823:'AutomaticLoadingWalkProof.json'}
stage_helpers={816:'Stage-TestRecentLoadingBoundary.py',817:'Stage-TestRecentLoadingBoundary.py',818:'Stage-TestAutomaticLoadingCoordinate.py',
    819:'Stage-TestAutomaticLoadingArc.py',820:'Stage-TestLoadingArcBudget.py',821:'Stage-TestPriorityLoadingArc.py',
    822:'Stage-TestLoadingOriginalState.py',823:'Stage-TestAutomaticLoadingWalk.py'}
proofs={str(k):json.loads((roots[k]/v).read_bytes()) for k,v in selected.items()}
assert proofs['818']['pair']['verified'] and proofs['818']['all_numerical_root_fields_efforts_and_original_stresses_bitexact815']
assert proofs['819']['pair']['verified'] and proofs['819']['info'][18]==5
assert not proofs['820']['pair'] and proofs['820']['info'][18]==2
assert proofs['821']['pair']['verified'] and proofs['821']['info'][18]==3 and proofs['821']['info'][19]==15
assert proofs['822']['numerical_coordinate_failure'] and not proofs['822']['pair'] and proofs['822']['info'][18]==1
assert proofs['823']['new_accepted_equilibria']==8 and proofs['823']['total_calls']==24
assert proofs['823']['all_steel_strain_limits_verified_at_all_three_stations']
assert all(not value['physical_capacity_certified'] for value in proofs.values())
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'];native={};regressions={}
for k,r in roots.items():
    value,_=check(r,'RealColdReplay');assert value==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    native[str(k)]=value
    for label,expected in [('PriorBranchMaterials',87),('OperationalArcRegressions',164)]:
        if (r/'Native'/(label+'.txt')).exists():
            value,_=check(r,label);assert value['passed']==expected and value['failed']==0 and value['own_process_closed'] and not value['timeout600']
            regressions[str(k)+':'+label]=value
state_path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(state_path.read_bytes())
assert not state['active_jobs'] and not state['pause_requested_explicitly'] and state['goal_status']=='active'
out=repo/'docs/reports/evidence/Coupled6DAutomaticLoading_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for k,r in roots.items():
    files=['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json','PointTraceSourceComposition.json',
        'CompiledSourceReuse.json','CompiledSourceReuseProof.json','docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
        'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Guard-IsolatedOffice.ps1','Run.ndjson',
        'tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
    if k in selected:files.append(selected[k])
    files += ['Input/'+n for n in ('Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin','OperationalArcRatio.bin',
        'NativeTolerance.bin','NativeSolver.json','EnvelopeOptions.bin','HPathTestOptions.bin','ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin')]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    files.append(stage_helpers[k])
    for relative in sorted(set(files)):
        p=r/relative
        if p.is_file():
            dest=out/str(k)/relative;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
helpers=['Stage-TestRecentLoadingBoundary.py','Stage-TestAutomaticLoadingCoordinate.py','Check-TestAutomaticLoadingCoordinate.py',
    'Stage-TestAutomaticLoadingArc.py','Check-TestAutomaticLoadingArc.py','Stage-TestLoadingArcBudget.py','Stage-TestPriorityLoadingArc.py',
    'Stage-TestLoadingOriginalState.py','Stage-TestAutomaticLoadingWalk.py','Check-TestAutomaticLoadingWalk.py',
    'Check-TestActiveLoadingCoordinate.py','Run-TestPriorBranchMaterials.ps1','Run-TestOperationalArcRegressions.ps1',
    'Finalize-TestOperationalArc.py','Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py','Stage-TestEnergyNumerics.py',
    'Check-TestPersistentHEvidence.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py','Validate-TestHardPrepared.py',
    'Check-TestHardHalfSteps.py','Compile-TestExternalPoint.ps1',Path(__file__).name]
for n in helpers:(out/n).write_bytes((here/n).read_bytes())
release={n:sha256((repo/n).read_bytes()).hexdigest().upper() for n in ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm']=='E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt']=='DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
(out/'Summary.json').write_text(json.dumps(dict(proofs=proofs,native=native,regressions=regressions,
    release_hashes=release,automatic_recent_station_and_working_arc_implemented_in_private_candidate=True,
    source815_test_entry_replaced_by_working_coordinate=True,all_started_stages_terminal=True,
    instrumented_inactive_exports816_817_produced_no_boundary_evidence=True,
    ordinary_internal_state_tolerance822_negative_preserved=True,selected_state_tolerance=5e-13,
    full_acceptance_completed=False,physical_capacity_certified=False,whole_goal_completed=False),ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',pause_requested_explicitly=False,
    pause_request_revoked=True,pause_checkpoint_ready=False,goal_pause_tool_pending=False,active_exec_session=None,
    active_jobs=[],all_started_stages_terminal=True,automatic_loading_evidence=out.relative_to(repo).as_posix(),
    automatic_loading_candidate=str(roots[823]),automatic_loading_summary=dict(new_equilibria=8,total_calls=24,
        last_lambda=proofs['823']['all_equilibria']['states'][-1]['lambda'],physical_capacity_certified=False),
    previous_goal_turn_classification='progress: resumed explicitly; automatic unique recent loading-boundary detection before commit and replay, revision-bound coordinate, working CArc outgoing full/half verification and exact restore.818 reproduces815 pair bitexact without manual station.819 succeeds5calls.820 insufficient4 rejects.821 prioritizes unique active departure, same pair3calls15primary, zero/unload/reset guards.822 original internal state tolerance rejects first root11primary, retained as negative; selected tighter internal iterations unchanged physical parameters.823 accepts8 outgoing equilibria24calls, all accepted/half weak and six-force equations and all three-station steel limits independently verified. Selected source87material+164continuation regressions passed. Physical ultimate, uncertainty, cold consistency and entire original scope remain open.',
    next_action='Continue full original goal. Use exact verified private candidate823/821 with CDP state tolerance5e-13, no material parameter change. Need cold-from-zero actual workflow and wider outgoing branch/uncertainty plus physical admissibility ONQ012, mesh/step/external controls. Current8 outgoing equilibria do not prove ultimate and are not production integration. Original-state822 failed strict coordinate; do not silently claim it works. Keep30phase/96global, field0.005/lambda/force/gauge/steel/hoop and LoadPath contracts; NumFail is not physical failure. Production source/workbook and main unchanged; full ON/OFF/UI/Focused/AutoCAD and performance not accepted.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (state_path,here/'Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('AUTOMATIC LOADING ARCHIVED',len(pins),'FILES;8 NEW EQUILIBRIA;NO PHYSICAL CAPACITY;GOAL ACTIVE')
