"""ДЛЯ ТЕСТОВ. Архивирует автоматические события779–784, отрицательные свидетели и независимые проверки, не объявляя несущую способность."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
from runpy import run_path
import json

repo, here = Path.cwd(), Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {key: base / name for key, name in {
    '779': 'Cold779AutomaticYieldArcWalk_20261125', '780': 'Cold780RejectedYieldEvent_20261125',
    '781': 'Cold781GaugeProjectionWitness_20261125', '782': 'Cold782ProjectedEventRefinement_20261125',
    '783': 'Cold783SmallEventCorrection_20261125', '784': 'Cold784AutomaticEventRefinement_20261125'}.items()}
proofs = {key: json.loads((roots[key] / name).read_bytes()) for key, name in {
    '779': 'ConfirmedArcWalkProof.json', '781': 'GaugeProjectionWitnessProof.json',
    '782': 'ProjectedEventRefinementProof.json', '783': 'ArcRootDerivativeProof.json', '784': 'ConfirmedArcWalkProof.json'}.items()}
assert proofs['779']['accepted_states'] == 43 and proofs['779']['repeated_yield_event_recovery']
assert proofs['781']['geometric_projection_does_not_solve_force_equilibrium']
assert proofs['782']['targeted_regressions']['passed'] == 105
assert proofs['784']['resolution_uses_actual_inner_tolerance'] and proofs['784']['event_third_difference_refinement']
native_check = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative']
for root in roots.values():
    native, _ = native_check(root, 'RealColdReplay')
    assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for key in ('782', '784'):
    regression, _ = native_check(roots[key], 'WorkingGuideRegressions')
    assert regression == dict(passed=105, failed=0, macros=3, timeout600=False, own_process_closed=True)
state_path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(state_path.read_bytes()); assert not state['active_jobs']
out = repo / 'docs/reports/evidence/Coupled6DAutomaticYieldRefinement_2026-11-25'
assert not out.exists(); out.mkdir(parents=True)
(out / '.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
for key, root in roots.items():
    for relative in ('ConfirmedArcWalkProof.json', 'ArcYieldSplitProof.json', 'GaugeProjectionWitnessProof.json',
        'ProjectedEventRefinementProof.json', 'ArcRootDerivativeProof.json', 'GaugeProjectionPreparation.json',
        'Preparation.json', 'CompilePoint.log', 'CompilePointClosure.json', 'PointTraceSourceComposition.json',
        'CompiledSourceReuse.json', 'docs/regression/CodeUpdate.json', 'ParentCodeUpdate.json', 'ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls',
        'tests/modTestCoupledElasticState.bas', 'TestArcYieldBoundary.cls.txt',
        'Run-TestRealColdReplay.ps1', 'Run-TestWorkingGuideRegressions.ps1', 'Guard-IsolatedOffice.ps1',
        'Compile-TestExternalPoint.ps1', 'Run.ndjson', 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'):
        path = root / relative
        if path.is_file():
            target = out / key / relative; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(path.read_bytes())
    for directory in ('Input', 'Native'):
        for path in (root / directory).iterdir():
            if path.is_file() and path.suffix.lower() not in ('.xlsm', '.xlsx'):
                target = out / key / directory / path.name; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(path.read_bytes())
helpers = ('Stage-TestAutomaticYieldArcWalk.py', 'Stage-TestRejectedYieldEvent.py', 'Stage-TestGaugeProjectionWitness.py',
    'Stage-TestProjectedEventRefinement.py', 'Stage-TestSmallEventCorrection.py', 'Stage-TestAutomaticEventRefinement.py',
    'Check-TestProjectedEventRefinement.py', 'Check-TestGaugeProjectionWitness.py', 'Check-TestArcRootDerivative.py',
    'Check-TestConfirmedArcWalk.py', 'TestAutoArcEventProof.py', 'Check-TestArcYieldSplit.py', 'TestArcYieldBoundary.cls.txt',
    'Check-TestPersistentHEvidence.py', 'Read-TestColdRawCapture.py', 'Check-TestActualReverseCapture.py',
    'Validate-TestHardPrepared.py', 'Check-TestHardHalfSteps.py', 'Stage-TestColdRawReplay.py', 'Stage-TestEnergyNumerics.py',
    'Compile-TestExternalPoint.ps1', 'Stage-TestRealPostCrack.py', 'Finalize-TestStockYield.py', Path(__file__).name)
for name in helpers:
    (out / name).write_bytes((here / name).read_bytes())
release = {name: sha256((repo / name).read_bytes()).hexdigest().upper() for name in ('workbook/output/RC_Section_NDM.xlsm', 'workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary = dict(proofs=proofs, release_hashes=release, private_event_method_changed=True,
    production_sources_and_release_unchanged=True, physical_capacity_certified=False, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
pins = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, indent=2, sort_keys=True) + '\n', encoding='utf-8')
state.update(updated_local=datetime.now().astimezone().isoformat(), goal_status='active',
    automatic_yield_refinement_evidence='docs/reports/evidence/Coupled6DAutomaticYieldRefinement_2026-11-25',
    automatic_yield_refinement784_summary={k: proofs['784'][k] for k in ('new_accepted_states','accepted_states','maximum_lambda','last_lambda','nonlinear_calls','total_primary_steps','stop')},
    previous_goal_turn_classification='progress: independently verified782 full/half event, all44 weak/all6 equilibria,105 directed regressions;783 actual492 finite differences show field correction below0.0036% of increment. Integrated third-difference refinement in bounded automatic784 with resolution tied to actual inner solve; archived779-784 including negative780/781. No physical capacity certified.',
    next_action='Inspect784 stop and unresolved events. Integrate only verified automatic continuation into working controller and test cold from zero using exact current six-force input, with shared budgets and half-step/raw-equilibrium checks. Preserve detection of physical failure distinct from NumFail; do not promote warm-history advance as cold capacity, do not chase lambda increments alone. Then step/mesh/external controls and remaining full scope, actual workbook and performance acceptance. Full long suites need scoped consent.')
raw = (json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode('utf-8')
for target in (state_path, repo / 'tmp/material-consistency/Continuation666_Resume20261123.json'): target.write_bytes(raw)
print('779-784 ARCHIVED', len(pins), 'PINS; NO PHYSICAL CAPACITY; GOAL ACTIVE')
