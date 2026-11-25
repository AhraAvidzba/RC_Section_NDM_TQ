"""ДЛЯ ТЕСТОВ. Закрепляет фактические производные корней и локализованную материальную нелинейность без изменения рабочего решателя."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
from runpy import run_path
import json

repo, here = Path.cwd(), Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {'774': base / 'Cold774ArcRootDerivative_20261125', '775': base / 'Cold775ArcRootPath_20261125'}
proofs = {}
for key, root in roots.items():
    native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
    assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
    name = 'ArcRootDerivativeProof.json' if key == '774' else 'ArcRootPathProof.json'
    proofs[key] = json.loads((root / name).read_bytes())
assert proofs['774']['all492_raw_difference_samples_verified']
assert proofs['775']['zero_stress_and_residual_repeat_bitidentical']
out = repo / 'docs/reports/evidence/Coupled6DArcRootWitness_2026-11-25'
assert not out.exists()
out.mkdir(parents=True)
(out / '.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
for key, root in roots.items():
    for relative in ('ArcRootDerivativeProof.json', 'ArcRootPathProof.json', 'Preparation.json', 'CompilePoint.log',
        'CompilePointClosure.json', 'PointTraceSourceComposition.json', 'docs/regression/CodeUpdate.json',
        'CompiledSourceReuse.json', 'ParentCodeUpdate.json', 'ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls',
        'tests/modTestCoupledElasticState.bas', 'Input/Reverse1.bin', 'Input/ArcRoots.bin',
        'Input/SourceArcWalk001.bin', 'Input/SourceArcWalk002.bin', 'Input/ArcPathOptions.bin',
        'Run.ndjson', 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1'):
        path = root / relative
        if path.is_file():
            target = out / key / relative; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(path.read_bytes())
    for path in (root / 'Native').iterdir():
        if path.is_file() and path.suffix.lower() not in ('.xlsx', '.xlsm'):
            target = out / key / 'Native' / path.name; target.parent.mkdir(parents=True, exist_ok=True); target.write_bytes(path.read_bytes())
for name in ('Stage-TestArcRootDerivative.py', 'TestArcRootDerivative.bas.txt', 'Check-TestArcRootDerivative.py',
    'Stage-TestArcRootPath.py', 'TestArcRootPath.bas.txt', 'Check-TestArcRootPath.py',
    'Check-TestPersistentHEvidence.py', 'Read-TestColdRawCapture.py', 'Check-TestActualReverseCapture.py',
    'Validate-TestHardPrepared.py', 'Check-TestHardHalfSteps.py', 'Stage-TestEnergyNumerics.py',
    'Stage-TestColdRawReplay.py', 'Compile-TestExternalPoint.ps1', Path(__file__).name):
    (out / name).write_bytes((here / name).read_bytes())
release = {name: sha256((repo / name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm', 'workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
summary = dict(cases=proofs, release_hashes=release, production_methods_changed=False,
    physical_capacity_certified=False, whole_goal_completed=False)
(out / 'Summary.json').write_text(json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
pins = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, indent=2, sort_keys=True) + '\n', encoding='utf-8')
state_file = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(state_file.read_bytes()); assert not state['active_jobs']
state.update(updated_local=datetime.now().astimezone().isoformat(), goal_status='active',
    arc_root_witness_evidence='docs/reports/evidence/Coupled6DArcRootWitness_2026-11-25',
    previous_goal_turn_classification='progress:774 full61-dimensional derivatives/492 readonly samples independently verified. Remaining root corrections are1e-8 of increment, refuting insufficient Newton convergence for these roots.775 all43x3 net tensors and residual curve independently verified; point15 at(72,0) dominates nonlinear row53 near half-step transition. No physical failure or capacity certified.',
    next_action='Trace actual operational CDP update at concrete point15 on each z-station along775 eta path, including substep counts, ordinary-return/apex/conditional-dilatancy-closure route and yield/dilatancy residuals, without accepting material or changing physical laws. Distinguish a genuine smooth/kinked material path from a numerical root-switch discontinuity before any production fix. Preserve full six-force workbook scope; do not substitute this witness for automatic capacity validation.')
raw = (json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode('utf-8')
for target in (state_file, repo / 'tmp/material-consistency/Continuation666_Resume20261123.json'):
    target.write_bytes(raw)
print('ARC ROOT WITNESS ARCHIVED', len(pins), 'SHA PINS; NO PHYSICAL LIMIT CLAIM; GOAL ACTIVE')
