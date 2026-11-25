"""ДЛЯ ТЕСТОВ. Сохраняет проверенные этапы 827–834 и незапущенную проверку 835 перед явной паузой."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
from runpy import run_path
import json

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {n: next(base.glob(f'Cold{n}*')) for n in range(827, 836)}
selected = {827:'RetainedArcContextProof.json',828:'OperationalColdProof.json',
    829:'OutgoingArcUncertaintyProof.json',830:'OperationalColdReplayProof.json',
    831:'RetainedArcContextProof.json',832:'OperationalColdProof.json',
    833:'AutomaticLoadingWalkProof.json',834:'UnresolvedArcScanProof.json'}
proofs = {str(n):json.loads((roots[n]/name).read_bytes()) for n,name in selected.items()}
assert proofs['828']['accepted_states'] == 44
assert proofs['830']['accepted_states'] == 44
assert proofs['831']['total_context_cases'] == 15
assert proofs['832']['accepted_states'] == 49
assert proofs['832']['exact_prefix_with_warm_input'] == 43
assert proofs['832']['cold']['capacity'] == 'NumFail'
assert not proofs['832']['cold']['limit']
assert proofs['833']['new_accepted_equilibria'] == 17
assert len(proofs['833']['all_equilibria']['states']) == 86
assert proofs['833']['total_calls'] == 93
assert proofs['834']['concrete_point3_elastic_largest_principal_crossing_at_all_three_stations']
assert all(not p['physical_capacity_certified'] for p in proofs.values())
native = {}
check = run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for n in selected:
    value,_ = check(roots[n], 'HBackendStock' if n in (828,832) else 'RealColdReplay')
    assert value['failed'] == 0 and value['own_process_closed'] and not value['timeout600']
    native[str(n)] = value
for n in (827,831):
    value,_ = check(roots[n], 'OperationalArcRegressions')
    assert value['passed'] == 164 and value['failed'] == 0 and value['own_process_closed']
    native[f'{n}_regressions'] = value
closure835 = json.loads((roots[835]/'CompilePointClosure.json').read_text(encoding='utf-8-sig'))
assert closure835['State'] == 'Absent'
assert not any((roots[835]/'Native').iterdir()), '835 must be recorded as compiled only, without native run'
out = repo/'docs/reports/evidence/Coupled6DSpectralPause_2026-11-25'
assert not out.exists()
out.mkdir(parents=True)
(out/'.gitattributes').write_text('* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
stage = {827:'Stage-TestRetainedArcControl.py',828:'Stage-TestRetainedArcCold.py',
    829:'Stage-TestOutgoingArcUncertainty.py',830:'Stage-TestAdaptiveColdReplay.py',
    831:'Stage-TestGatedLoadingArc.py',832:'Stage-TestGatedLoadingCold.py',
    833:'Stage-TestContinuedOutgoingWalk.py',834:'Stage-TestUnresolvedArcScan.py',835:'Stage-TestAdaptiveColdReplay.py'}
common = ['Preparation.json','CompilePoint.log','CompilePointClosure.json','CompileClosure.json',
    'PointTraceSourceComposition.json','CompiledSourceReuse.json','CompiledSourceReuseProof.json',
    'docs/regression/CodeUpdate.json','ParentCodeUpdate.json','ParentCompileClosure.json',
    'Compile-TestExternalPoint.ps1','Run-TestRealColdReplay.ps1','Run-TestHMixedStock.ps1',
    'Run-TestOperationalArcRegressions.ps1','Guard-IsolatedOffice.ps1','Run.ndjson',
    'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls',
    'src/Materials/CCoupledConcreteReturn.cls','tests/modTestCoupledElasticState.bas',
    'tools/build_workbook/Update-IsolatedOptimizerCode.ps1']
inputs = ['Reverse1.bin','SourceColdRawHistory.bin','ArcCallAllowance.bin','OperationalArcRatio.bin',
    'NativeTolerance.bin','NativeSolver.json','EnvelopeOptions.bin','HPathTestOptions.bin',
    'ColdHPrefix.bin','ColdHRequest.bin','MaterialProbeRatio.bin']
for n,r in roots.items():
    files = common + [stage[n]] + ( [selected[n]] if n in selected else [] )
    files += ['Input/'+name for name in inputs]
    # Дополнительные входы относятся только к диагностике этой точки, без унаследованных каталогов.
    files += [p.relative_to(r).as_posix() for p in (r/'Input').glob('*')
        if p.is_file() and ('ArcScan' in p.name or 'Uncertainty' in p.name)]
    files += [p.relative_to(r).as_posix() for p in (r/'Native').iterdir()
        if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    for name in sorted(set(files)):
        p = r/name
        if p.is_file():
            target = out/str(n)/name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(p.read_bytes())
helpers = list(stage.values()) + ['Check-TestRetainedArcControl.py','Check-TestOperationalCold.py',
    'Check-TestOutgoingArcUncertainty.py','Check-TestOperationalColdReplay.py',
    'Check-TestAutomaticLoadingWalk.py','Check-TestUnresolvedArcScan.py',
    'Check-TestPersistentHEvidence.py','Read-TestColdRawCapture.py','Check-TestActualReverseCapture.py',
    'Validate-TestHardPrepared.py','Check-TestHardHalfSteps.py','Stage-TestColdRawReplay.py',
    'Stage-TestEnergyNumerics.py','Stage-TestAutomaticLoadingWalk.py',
    'Finalize-TestOperationalArc.py','Finalize-TestStockYield.py','Audit-TestPrivateCoupledBook.py',
    'Inspect-TestResolutionVBA.py',Path(__file__).name]
for name in sorted(set(helpers)):
    p = here/name
    assert p.is_file(), name
    (out/name).write_bytes(p.read_bytes())
release = {name:sha256((repo/name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm','workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary = dict(proofs=proofs,native=native,release_hashes=release,
    stage835_compile_closed=closure835,stage835_native_not_started=True,
    stage832_independent_cold_stress_replay_pending=True,
    spectral_event_fix_not_implemented_or_verified=True,
    all_started_stages_terminal=True,physical_capacity_certified=False,
    full_acceptance_completed=False,whole_goal_completed=False,
    user_requested_pause=True)
(out/'Summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins = {p.relative_to(out).as_posix():sha256(p.read_bytes()).hexdigest().upper()
    for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
state_path = repo/'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(state_path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='paused',
    pause_requested_explicitly=True,pause_request_revoked=False,pause_after_current_stage=True,
    pause_request='До технологического завершения текущего этапа поставим задачу на паузу.',
    pause_checkpoint_ready=True,goal_pause_tool_pending=False,
    no_new_experiments_added_after_pause_request=True,
    active_jobs=[],active_exec_session=None,all_started_stages_terminal=True,
    spectral_pause_evidence=out.relative_to(repo).as_posix(),
    stage835='Compilation complete, own Excel26792 absent; native replay NOT started before pause',
    latest_verified_cold_source=str(roots[832]),latest_verified_extended_branch=str(roots[833]),
    physical_capacity_certified=False,full_acceptance_completed=False,
    previous_goal_turn_classification='progress: gated/proactive controller cold49, maximum unchanged1.1058869810286156; continuation69->86 to0.5832175966249188 with independent equilibrium and three-station steel checks; source86 spectral elastic principal sign event independently localized, no new yield crossing; physical ultimate unconfirmed',
    next_action='Resume only on explicit user request. Preserve full original goal. First finalize835 compiled-only source audit and run pending pure stress replay of832 actual49-state cold trajectory, verify source exact; no replay PASS claimed yet. Then implement generic CDP elastic principal-zero transition using existing projected material-event solver in private candidate based on833 math and831 gated/proactive controller. Require negative yield surface at both endpoints and event, preserve ASD flag/semantics, all full/half field/lambda/force/gauge/steel/30iteration/96probe contracts. Reproduce source86 frozen failure, test event correction and continue branch, then actual cold. Do not hardcode point3/station. CDP AcceptedPointDamage exposes energy coordinates, not damage factors. Retained-arc ungated828 slower negative variant preserved; local lambda uncertainty829 excludes nonlinear remainder and is not ultimate proof. No production/main/release promotion or full acceptance. Pending physical admissibility ONQ012, mesh/step/external and full original LoadPaths/hoop/UI/CAD/performance/workbook checks remain.')
state['terminal_sessions_this_stage'] = sorted(set(state.get('terminal_sessions_this_stage',[])+[37351]))
raw = (json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for target in (state_path,here/'Continuation666_Resume20261123.json'):
    target.write_bytes(raw)
print('PAUSE CHECKPOINT:',len(pins),'exact files;',sum(p.stat().st_size for p in out.rglob('*') if p.is_file()),'bytes;835 compiled only, native replay pending')
