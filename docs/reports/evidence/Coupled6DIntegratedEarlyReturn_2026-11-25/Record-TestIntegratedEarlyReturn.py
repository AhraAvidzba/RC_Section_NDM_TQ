"""ДЛЯ ТЕСТОВ. Сохраняет холодный ранний H со встроенным возвратом и проверенные поля, без объявления неподтверждённого предела."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
from runpy import run_path
import json
import re

repo, here = Path.cwd(), Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {'754': base / 'Cold754IntegratedEarlyEnvelope_20261125',
    '755': base / 'Cold755IntegratedEarlyFrozenReplay_20261125',
    '756': base / 'Cold756GuideHalfConfirmation_20261125',
    '757': base / 'Cold757HalfGuardedMaterialReturn_20261125',
    '758': base / 'Cold758PositiveHalfGuard_20261125',
    '759': base / 'Cold759HalfGuardLoadAllocationFix_20261125'}
cold = json.loads((roots['754'] / 'IntegratedEarlyColdProof.json').read_bytes())
replay = json.loads((roots['755'] / 'IntegratedEarlyReplayProof.json').read_bytes())
half = json.loads((roots['756'] / 'GuideHalfConfirmationProof.json').read_bytes())
corrected = json.loads((roots['757'] / 'HalfGuardedMaterialReturnProof.json').read_bytes())
positive = json.loads((roots['759'] / 'PositiveHalfGuardProof.json').read_bytes())
guard_regression, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](roots['757'], 'WorkingGuideRegressions')
assert guard_regression == dict(passed=105, failed=0, macros=3, timeout600=False, own_process_closed=True)
final_regression, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](roots['759'], 'WorkingGuideRegressions')
assert final_regression == guard_regression
failed_positive_text = (roots['758'] / 'Native/RealColdReplay.txt').read_text('utf-8-sig')
failed_positive_closure = json.loads((roots['758'] / 'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert 'RUNTIME ERROR' in failed_positive_text and 'Subscript out of range' in failed_positive_text
assert failed_positive_closure['Passed'] and failed_positive_closure['ChildHasExited'] and failed_positive_closure['ChildExitCode'] == 1
assert half['half_solved'] and not half['locally_confirmed_by_half_step']
assert corrected['nonlocal756_candidate_rejected'] and positive['guarded_half_check_returned_true']
assert cold['cold']['native']['own_process_closed'] and replay['native']['own_process_closed']
assert replay['states'] == cold['accepted_states'] and replay['all_adjacent_spatial_increments_forward']
assert cold['regression']['passed'] == 105 and cold['regression']['failed'] == 0
out = repo / 'docs/reports/evidence/Coupled6DIntegratedEarlyReturn_2026-11-25'
assert not out.exists()
out.mkdir(parents=True)
(out / '.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
for key, root in roots.items():
    paths = []
    proof_names = {'754': ['IntegratedEarlyColdProof.json'],
        '755': ['IntegratedEarlyReplayProof.json', 'ActualColdMaterialMeasures.json'],
        '756': ['GuideHalfConfirmationProof.json'], '757': ['HalfGuardedMaterialReturnProof.json'],
        '758': [], '759': ['PositiveHalfGuardProof.json']}
    for name in proof_names[key]:
        assert (root / name).is_file()
        paths.append(root / name)
    for name in ('Preparation.json', 'CompilePoint.log', 'CompilePointClosure.json', 'PointTraceSourceComposition.json',
        'CompiledSourceReuse.json', 'docs/regression/CodeUpdate.json',
        'src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls', 'tests/modTestCoupledElasticState.bas',
        'Input/Reverse1.bin', 'Input/Source754History.bin', 'Input/ExpectedOriginalControl.bin',
        'ParentCodeUpdate.json', 'ParentCompileClosure.json', 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1', 'Run.ndjson'):
        if (root / name).is_file():
            paths.append(root / name)
    paths.extend(p for p in (root / 'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx', '.xlsm'))
    for source in paths:
        target = out / key / source.relative_to(root)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(source.read_bytes())
    if key == '758':
        execution_path, = re.findall(r'(?m)^([A-Z]:[^\r\n]+_execution_report\.txt)\r?$', failed_positive_text)
        execution_file = Path(execution_path).resolve()
        assert execution_file.is_relative_to(root.resolve())
        (out / key / 'FailedPositiveExecutionReport.txt').write_bytes(execution_file.read_bytes())
for name in ('Stage-TestIntegratedEarlyEnvelope.py', 'Check-TestIntegratedEarlyCold.py',
    'Stage-TestIntegratedEarlyReplay.py', 'Check-TestIntegratedEarlyReplay.py', 'Read-TestColdRawCapture.py',
    'Check-TestPersistentHEvidence.py', 'Check-TestActualReverseCapture.py', 'Validate-TestHardPrepared.py',
    'Analyze-TestRealColdMaterials.py', 'Stage-TestColdRawReplay.py', Path(__file__).name):
    (out / name).write_bytes((here / name).read_bytes())
for name in ('Stage-TestGuideHalfConfirmation.py', 'Check-TestGuideHalfConfirmation.py',
    'Stage-TestHalfGuardedMaterialReturn.py', 'Check-TestHalfGuardedMaterialReturn.py',
    'Stage-TestPositiveHalfGuard.py', 'Check-TestPositiveHalfGuard.py', 'Stage-TestHalfGuardLoadAllocationFix.py'):
    (out / name).write_bytes((here / name).read_bytes())
release = {name: sha256((repo / name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm', 'workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary = dict(actual_cold754=cold, independent_replay755=replay, failed_local_confirmation756=half,
    half_guard_fix757=corrected, guard_regression757=guard_regression,
    failed_positive758_subscript_error_preserved=True, failed_positive758_process_closed=failed_positive_closure['Passed'],
    positive_half_guard759=positive, final_regression759=final_regression, release_hashes=release,
    physical_capacity_certified=False, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
pins = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, indent=2, sort_keys=True) + '\n', encoding='utf-8')
path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(), goal_status='active',
    integrated_early_cold754=cold, integrated_early_replay755=replay, guide_half_confirmation756=half,
    half_guard_correction757=corrected, half_guard_regression757=guard_regression, positive_half_guard759=positive,
    half_guard_final_regression759=final_regression,
    previous_goal_turn_classification='progress:754 automatic return adds27 equilibrium states but756 halfstep disproves local-step consistency (tensorerror0.98711);757 adds shared-budget halfguard, rejection20/30/history preserved;758 positive test reveals unallocated full-restoration load array, fixed759 and positive plus105 final regressions verified. Physicalcapacity and fullgoal remain open.',
    next_action='Use corrected private759 halfguard (working757 plus six-component array allocation; testwrapper from758), reject unverified guide752/754 promotion. Seek connected continuation with a different parameter near turning states; verify full/half before adoption, not just positive projection and equilibrium. Positive guard restores fullstate, but cold search after correction and automatic physical limit remain unproved. Preserve full task and main/release.')
raw = (json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode('utf-8')
for target in (path, repo / 'tmp/material-consistency/Continuation666_Resume20261123.json'):
    target.write_bytes(raw)
print('754755 ARCHIVED', len(pins), 'PINS; PHYSICAL CAPACITY NOT CERTIFIED; FULL GOAL ACTIVE')
