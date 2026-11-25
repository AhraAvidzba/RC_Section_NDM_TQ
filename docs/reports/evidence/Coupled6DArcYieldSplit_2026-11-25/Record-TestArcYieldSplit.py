"""ДЛЯ ТЕСТОВ. Сохраняет трассу материальной ветви и проверенное разделение дуги по событию без объявления предела."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
from runpy import run_path
import re
import json

repo, here = Path.cwd(), Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {'776': base / 'Cold776ArcPointRoute_20261125', '777': base / 'Cold777ArcYieldSplit_20261125', '778': base / 'Cold778ArcYieldContextFix_20261125'}
route = json.loads((roots['776'] / 'ArcPointRouteProof.json').read_bytes())
event = json.loads((roots['778'] / 'ArcYieldSplitProof.json').read_bytes())
walk = json.loads((roots['778'] / 'ConfirmedArcWalkProof.json').read_bytes())
assert route['all775_residual_and_stress_path_bytes_exact'] and route['all_three_station_accepted_histories_unchanged']
assert all(r['substeps'] <= 1 and r['operational_counts'][2:4] == [0, 0] and r['operational_interface_stress_max_difference_MPa'] == 0 for r in route['records'])
assert event['event_verified'] and event['total_primary_steps'] == 3
assert walk['new_accepted_states'] == 9 and walk['accepted_states'] == 43
native_check = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative']
for key in ('776', '778'):
    native, _ = native_check(roots[key], 'RealColdReplay')
    assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
failure_text = (roots['777'] / 'Native/RealColdReplay.txt').read_text('utf-8-sig')
failure_closure = json.loads((roots['777'] / 'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert 'RUNTIME ERROR' in failure_text and failure_closure['Passed'] and failure_closure['ChildExitCode'] == 1
execution_path, = re.findall(r'(?m)^([A-Z]:[^\r\n]+_execution_report\.txt)\r?$', failure_text)
execution = Path(execution_path).resolve(); assert execution.is_relative_to(roots['777'].resolve())
execution_raw = execution.read_bytes()
execution_encoding = 'utf-16' if execution_raw.startswith((b'\xff\xfe', b'\xfe\xff')) else 'utf-8-sig'
assert 'Subscript out of range' in execution_raw.decode(execution_encoding)
out = repo / 'docs/reports/evidence/Coupled6DArcYieldSplit_2026-11-25'
assert not out.exists(); out.mkdir(parents=True)
(out / '.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
for key, root in roots.items():
    for relative in ('ArcPointRouteProof.json', 'ArcYieldSplitProof.json', 'ConfirmedArcWalkProof.json',
        'Preparation.json', 'CompilePoint.log', 'CompilePointClosure.json', 'PointTraceSourceComposition.json',
        'CompiledSourceReuse.json', 'docs/regression/CodeUpdate.json', 'ParentCodeUpdate.json', 'ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls',
        'tests/modTestCoupledElasticState.bas', 'Input/Reverse1.bin', 'Input/ArcRoots.bin', 'Input/ArcPathOptions.bin',
        'Input/SignedArcOptions.bin', 'TestArcYieldBoundary.cls.txt', 'TestArcYieldSplit.bas.txt',
        'TestPointUpdateRoute.cls.txt', 'Run.ndjson', 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'):
        path = root / relative
        if path.is_file():
            target = out / key / relative; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(path.read_bytes())
    for path in (root / 'Native').iterdir():
        if path.is_file() and path.suffix.lower() not in ('.xlsm', '.xlsx'):
            target = out / key / 'Native' / path.name; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(path.read_bytes())
(out / '777/FailedTestContextExecutionReport.txt').write_bytes(execution_raw)
for name in ('Stage-TestArcPointRoute.py', 'TestPointUpdateRoute.cls.txt', 'Check-TestArcPointRoute.py',
    'Stage-TestArcYieldSplit.py', 'Stage-TestArcYieldContextFix.py', 'TestArcYieldBoundary.cls.txt',
    'TestArcYieldSplit.bas.txt', 'Check-TestArcYieldSplit.py', 'Check-TestConfirmedArcWalk.py',
    'Check-TestPersistentHEvidence.py', 'Read-TestColdRawCapture.py', 'Check-TestActualReverseCapture.py',
    'Validate-TestHardPrepared.py', 'Check-TestHardHalfSteps.py', 'Stage-TestColdRawReplay.py',
    'Stage-TestEnergyNumerics.py', 'Compile-TestExternalPoint.ps1', Path(__file__).name):
    (out / name).write_bytes((here / name).read_bytes())
release = {name: sha256((repo / name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm', 'workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary = dict(route776=route, event778=event, post_event_walk778=walk,
    failed777_test_context_source_and_report_preserved=True, failed777_own_process_closed=failure_closure['Passed'],
    release_hashes=release, working_methods_changed=False, physical_capacity_certified=False, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
pins = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, indent=2, sort_keys=True) + '\n', encoding='utf-8')
path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(path.read_bytes()); assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(), goal_status='active',
    arc_yield_split_evidence='docs/reports/evidence/Coupled6DArcYieldSplit_2026-11-25',
    arc_yield_split778_event=event,
    arc_yield_split778_walk_summary={k: walk[k] for k in ('new_accepted_states','accepted_states','maximum_lambda','last_lambda','nonlinear_calls','total_primary_steps','stop')},
    previous_goal_turn_classification='progress:776 excludes cache/apex/substeps/closure atpoint15; side+1 ordinary update changes from elastic0 iterations to plastic7-12.777 test context failed due unallocated projection, fixed778. Existing projected solver automatically selects15/+1, event and half verified in3 shared primary iterations; then9 locally confirmed arc states, all43 accepted equilibria independently verified. Physicalcapacity not certified.',
    next_action='Integrate repeated automatic loading-yield event splits into bounded private arc walk after full/half consistency rejection. Use778 verified generic event finder and actual failed full/half bracket, not point15 hardcode or future fields. Preserve96 global probe budget and shared30 iterations per declared step; record separate event attempts and their exact source histories/parameters so every event and post-event state can be independently checked. Test from current33 source first, then cold from zero and affected regressions before any production promotion. Do not confuse old maximum~1.109 with a certified capacity; full original six-force scope remains active.')
raw = (json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode('utf-8')
for target in (path, repo / 'tmp/material-consistency/Continuation666_Resume20261123.json'):
    target.write_bytes(raw)
print('776-778 ARCHIVED', len(pins), 'PINS; EVENT+9 STEPS VERIFIED; NO PHYSICAL LIMIT CLAIM; GOAL ACTIVE')
