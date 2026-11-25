"""ДЛЯ ТЕСТОВ. Архивирует встроенный возврат, холодный запуск и независимое равновесие без объявления физического предела."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo = Path.cwd()
here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {'751': base / 'Cold751IntegratedMaterialGuide_20261125',
    '752': base / 'Cold752IntegratedGuideDirected_20261125',
    '753': base / 'Cold753IntegratedColdFrozenReplay_20261125'}
directed = json.loads((roots['752'] / 'IntegratedGuideDirectedProof.json').read_bytes())
cold = json.loads((roots['752'] / 'IntegratedMaterialColdProof.json').read_bytes())
replay = json.loads((roots['753'] / 'IntegratedColdReplayProof.json').read_bytes())
assert directed['independent_NDS_and_original_control_verified'] and directed['total_primary_steps'] == 17
assert cold['cold']['native']['own_process_closed'] and replay['states'] == cold['accepted_states']
assert cold['regression']['passed'] == 105 and cold['regression']['failed'] == 0
out = repo / 'docs/reports/evidence/Coupled6DIntegratedMaterialReturn_2026-11-25'
assert not out.exists()
out.mkdir(parents=True)
(out / '.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n', encoding='utf-8')
for key, root in roots.items():
    paths = []
    for name in ('Preparation.json', 'CompilePoint.log', 'CompilePointClosure.json', 'PointTraceSourceComposition.json',
        'CompiledSourceReuse.json', 'docs/regression/CodeUpdate.json', 'IntegratedGuideDirectedProof.json',
        'IntegratedMaterialColdProof.json', 'IntegratedColdReplayProof.json', 'ActualColdMaterialMeasures.json',
        'CompileDiagnosticRepair.json', 'tools/build_workbook/Update-IsolatedOptimizerCode.ps1',
        'src/Solver/CCoupledNewtonSolver.cls', 'src/Solver/CCoupledDeformationContinuation.cls',
        'src/Materials/CCoupledConcreteResponse.cls', 'tests/modTestCoupledElasticState.bas',
        'Input/Reverse1.bin', 'Input/Source745History.bin', 'Input/Source752History.bin',
        'Input/ExpectedOriginalControl.bin', 'Run.ndjson', 'ParentCodeUpdate.json', 'ParentCompileClosure.json'):
        if (root / name).is_file():
            paths.append(root / name)
    if (root / 'Native').is_dir():
        paths.extend(p for p in (root / 'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsx', '.xlsm'))
    if (root / 'CompileAttempt1Failure').is_dir():
        paths.extend(p for p in (root / 'CompileAttempt1Failure').rglob('*') if p.is_file())
    for source in paths:
        target = out / key / source.relative_to(root)
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(source.read_bytes())
for name in ('Stage-TestIntegratedMaterialGuide.py', 'Stage-TestIntegratedGuideDirected.py',
    'Check-TestIntegratedGuideDirected.py', 'Check-TestIntegratedMaterialCold.py',
    'Stage-TestIntegratedColdReplay.py', 'Check-TestIntegratedColdReplay.py', 'Repair-TestFrozenCompilerDiagnostic.py',
    'Analyze-TestRealColdMaterials.py',
    'Read-TestColdRawCapture.py', 'Check-TestPersistentHEvidence.py', 'Check-TestActualReverseCapture.py',
    'Validate-TestHardPrepared.py', 'Stage-TestColdRawReplay.py', Path(__file__).name):
    (out / name).write_bytes((here / name).read_bytes())
release = {name: sha256((repo / name).read_bytes()).hexdigest().upper() for name in
    ('workbook/output/RC_Section_NDM.xlsm', 'workbook/output/VBA_All_Code.txt')}
assert release['workbook/output/RC_Section_NDM.xlsm'] == 'E47ACB7083D92310F9915FD5EB59DA6A7B6C52F352D5EE274A2DCDC6BA567450'
assert release['workbook/output/VBA_All_Code.txt'] == 'DB8023F9384B0059E1D1C51D59FE0E62F9B8CD94AC4D9B329DD0EE810AF676CC'
(out / 'Summary.json').write_text(json.dumps(dict(directed_return=directed, actual_cold=cold,
    independent_replay=replay, release_hashes=release, physical_capacity_certified=False, whole_goal_completed=False),
    ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
pins = {p.relative_to(out).as_posix(): sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out / 'Files.json').write_text(json.dumps(pins, indent=2, sort_keys=True) + '\n', encoding='utf-8')
path = repo / 'docs/reports/Coupled6DResumeState_2026-11-24.json'
state = json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(), goal_status='active',
    integrated_material_return752=directed, integrated_material_cold752=cold, integrated_cold_replay753=replay,
    previous_goal_turn_classification='progress: integrated original-target return on unchanged material history independently verified, actual255 audit and105 focused regressions; actual cold all6 and independent accepted-state reintegration recorded. Physical capacity still unproven; entire goal active.',
    next_action='Use actual cold752 terminal reason and frozen753 material measures to choose next continuation correction. Do not claim auxiliary reverse roots or NumFail as physical collapse; retain full goal, no main merge or production promotion from directed evidence.')
raw = (json.dumps(state, ensure_ascii=False, indent=2, sort_keys=True) + '\n').encode('utf-8')
for p in (path, repo / 'tmp/material-consistency/Continuation666_Resume20261123.json'):
    p.write_bytes(raw)
print('751752753 ARCHIVED', len(pins), 'PINS; INTEGRATED RETURN VERIFIED; FULL PHYSICAL CAPACITY GOAL OPEN')
