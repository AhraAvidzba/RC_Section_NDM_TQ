"""ДЛЯ ТЕСТОВ. Сохраняет подтверждения, отрицательные результаты и исходники знаковой дуги без выпуска неподтверждённой несущей способности."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
from runpy import run_path
import json
import re

repo, here = Path.cwd(), Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
names = {
    '760': 'Cold760SignedArcPair_20261125', '761': 'Cold761SignedArcHybrid_20261125',
    '762': 'Cold762SignedArcRatio0125_20261125', '763': 'Cold763SignedArcRatio00625_20261125',
    '764': 'Cold764SignedArcTightPair_20261125', '765': 'Cold765AcceptedMapContinuity_20261125',
    '766': 'Cold766SignedArcPrefix32_20261125', '767': 'Cold767SignedArcPrefix32Quarter_20261125',
    '768': 'Cold768SignedArcPrefix32Small_20261125', '769': 'Cold769SignedArcPrefix32Strict_20261125',
    '770': 'Cold770ConfirmedArcWalk_20261125', '771': 'Cold771ArcMaterialAccuracy_20261125',
    '772': 'Cold772ArcConsistentAccuracy_20261125', '773': 'Cold773HalfFirstArcWalk_20261125'}
roots = {key: base / name for key, name in names.items()}
native_check = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative']
regression, _ = native_check(roots['761'], 'WorkingGuideRegressions')
assert regression == dict(passed=105, failed=0, macros=3, timeout600=False, own_process_closed=True)
proofs = {}
for key, root in roots.items():
    if key == '771':
        closure = json.loads((root / 'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
        text = (root / 'Native/RealColdReplay.txt').read_text('utf-8-sig')
        assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode'] == 1
        assert 'RUNTIME ERROR' in text
        rows = run_path(str(here / 'Validate-TestHardPrepared.py'))['TestRows'](root / 'Native/EnvelopeEnergy.Accepted.csv')
        residuals = [list(map(float, row[1:])) for row in rows if row[0] == 'R']
        assert len(residuals) == 5 and residuals[-1][1] > residuals[-1][2]
        proof = dict(native_failed_before_arc_walk=True, own_process_closed=True,
            first_failed_frozen_state=int(residuals[-1][0]), actual_residual_MPa=residuals[-1][1],
            required_bound_MPa=residuals[-1][2], old_scalar_return_tolerance_looser_than_new_outer=True,
            failure_preserved_not_accepted_as_physical_failure=True, physical_capacity_certified=False)
        proof_name = 'FailedFrozenReplayProof.json'
        (root / proof_name).write_text(json.dumps(proof, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    else:
        native, _ = native_check(root, 'RealColdReplay')
        assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
        proof_name = 'AcceptedMapContinuityProof.json' if key == '765' else ('ConfirmedArcWalkProof.json' if key in ('770', '772', '773') else 'SignedArcPairProof.json')
        proof = json.loads((root / proof_name).read_bytes())
    if key in ('771', '772'):
        material, _ = native_check(root, 'PriorBranchMaterials')
        assert material == dict(passed=87, failed=0, macros=3, timeout600=False, own_process_closed=True)
        proof['material_regressions'] = material
    proofs[key] = dict(proof_file=proof_name, result=proof)
assert proofs['769']['result']['locally_confirmed_by_half_step']
assert proofs['770']['result']['new_accepted_states'] == 2
assert proofs['772']['result']['new_accepted_states'] == 2
out = repo / 'docs/reports/evidence/Coupled6DSignedArcAccuracy_2026-11-25'
assert not out.exists()
out.mkdir(parents=True)
(out / '.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
for key, root in roots.items():
    paths = [root / proofs[key]['proof_file']]
    for name in ('Preparation.json', 'CompilePoint.log', 'CompilePointClosure.json', 'PointTraceSourceComposition.json',
        'CompiledSourceReuse.json', 'docs/regression/CodeUpdate.json', 'ParentCodeUpdate.json', 'ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls',
        'tests/modTestCoupledElasticState.bas', 'Input/Reverse1.bin', 'Input/Source752History.bin',
        'Input/SignedArcOptions.bin', 'Input/MapContinuityOptions.bin', 'Run.ndjson',
        'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'):
        if (root / name).is_file():
            paths.append(root / name)
    paths.extend(p for p in (root / 'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx', '.xlsm'))
    for source in paths:
        target = out / key / source.relative_to(root)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(source.read_bytes())
    if key == '771':
        text = (root / 'Native/RealColdReplay.txt').read_text('utf-8-sig')
        path, = re.findall(r'(?m)^([A-Z]:[^\r\n]+_execution_report\.txt)\r?$', text)
        path = Path(path).resolve()
        assert path.is_relative_to(root.resolve())
        (out / key / 'FailedFrozenExecutionReport.txt').write_bytes(path.read_bytes())
helpers = ('Stage-TestSignedArcPair.py', 'Stage-TestSignedArcHybrid.py', 'Stage-TestSignedArcRatio.py',
    'Stage-TestSignedArcTightPair.py', 'Stage-TestSignedArcPrefix.py', 'Check-TestSignedArcPair.py',
    'Stage-TestAcceptedMapContinuity.py', 'Check-TestAcceptedMapContinuity.py',
    'Check-TestWeakResultantClosure.py', 'Stage-TestConfirmedArcWalk.py', 'TestConfirmedArcWalk.bas.txt',
    'Check-TestConfirmedArcWalk.py', 'Stage-TestArcMaterialAccuracy.py', 'Stage-TestArcConsistentAccuracy.py',
    'Stage-TestHalfFirstArcWalk.py', 'Check-TestPersistentHEvidence.py', 'Validate-TestHardPrepared.py',
    'Check-TestHardHalfSteps.py', 'Read-TestColdRawCapture.py', 'Check-TestActualReverseCapture.py',
    'Stage-TestEnergyNumerics.py', 'Stage-TestColdRawReplay.py', 'Compile-TestExternalPoint.ps1',
    'Run-TestWorkingGuideRegressions.ps1', 'Run-TestPriorBranchMaterials.ps1', Path(__file__).name)
for name in helpers:
    (out / name).write_bytes((here / name).read_bytes())
(out / 'WeakResultantClosure.json').write_bytes((base / 'WeakResultantClosure_20261125.json').read_bytes())
release = {name: sha256((repo / name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm', 'workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary = dict(cases=proofs, regression761=regression, release_hashes=release,
    production_kernel_changed=False, physical_capacity_certified=False, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
pins = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, indent=2, sort_keys=True) + '\n', encoding='utf-8')
path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(path.read_bytes())
assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(), goal_status='active',
    signed_arc_accuracy_evidence='docs/reports/evidence/Coupled6DSignedArcAccuracy_2026-11-25',
    signed_arc_accuracy_summary={key: {k: v for k, v in item['result'].items() if k in
        ('locally_confirmed_by_half_step', 'tensor_relative_half_error', 'new_accepted_states', 'maximum_lambda',
         'last_lambda', 'nonlinear_calls', 'total_primary_steps', 'stop', 'first_failed_frozen_state', 'actual_residual_MPa')} for key, item in proofs.items()},
    previous_goal_turn_classification='progress: strict signed-arc pair769 independently confirmed; automatic770 continuation verified2 new states;771 isolated outer precision broke frozen5;772 consistent nested precision restored replay and87 regressions but no further advancement;773 half-first predictor evaluated. No physical capacity certified.',
    next_action='773 half-first predictor still accepts only2 states. Diagnose remaining field correction and tangent conditioning at roots of770/773; force residual alone may not resolve displacement accuracy near turning. Distinguish this from material branch transition using actual full/half root derivatives, rather than more blind step shrinking. Keep numerical stop distinct from physical failure; preserve full original six-force workbook scope, unchanged production kernel, main and release.')
raw = (json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode('utf-8')
for target in (path, repo / 'tmp/material-consistency/Continuation666_Resume20261123.json'):
    target.write_bytes(raw)
print('SIGNED ARC EVIDENCE ARCHIVED', len(pins), 'PINS; PHYSICAL CAPACITY NOT CERTIFIED; FULL GOAL ACTIVE')
