"""ДЛЯ ТЕСТОВ. Сохраняет достигнутый предел стали и отрицательный автоматический поиск без подмены области проверки."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DPrescribedSteelEnd_2026-11-24';assert not out.exists();out.mkdir(parents=True)
cases={'723':'Cold723PriorBranchCapacity_20261124','724':'Cold724PriorBranchSteelEnd_20261124',
    '725':'Cold725SteelEndDirectoryRetry_20261124','726':'Cold726PostYieldProfile_20261124'}
for number,name in cases.items():
    root=base/name
    names=['Preparation.json','CompilePointClosure.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
        'docs/regression/CodeUpdate.json','Run.ndjson','PriorBranchCapacityTimeoutProof.json',
        'MaterialRegressionProof.json','HardShearDirectProof.json','HardShearDirectComparisonProof.json',
        'PriorBranchSteelEndProof.json','DirectProbeProof.json','PreparationFailure724.json']
    paths=[root/n for n in names if (root/n).is_file()]
    for folder in ('Native','Matrix/Native'):
        if (root/folder).is_dir():paths += [p for p in (root/folder).iterdir() if p.is_file()]
    if number in ('724','726'):paths += [root/'tests/modTestCoupledElasticState.bas']
    if number=='726':paths += [root/'src/Solver/CCoupledNewtonSolver.cls',root/'Input/PostCrackReplay.bin',root/'Input/PostCrackReplayOptions.bin']
    if number=='724':paths += [root/'Stage-TestPriorBranchSteelEnd.py']
    for p in paths:
        target=out/number/p.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
baseline=base/'Cold719ScopedPeakRefinement_20261124'
for p in (baseline/'Native').glob('PriorBranchHardShearDirect*'):
    if p.is_file():
        target=out/'719Baseline/Native'/p.name;target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
helpers=['Stage-TestPriorBranchCapacity.py','Stage-TestPriorBranchSteelEnd.py','Stage-TestRetrySteelEnd.py',
    'Check-TestPriorBranchSteelEnd.py','Stage-TestPostYieldProfile.py','Check-TestPostCrackDirectProbe.py',
    'Run-TestPriorBranchMaterials.ps1','Run-TestPriorBranchHardShear.ps1','Check-TestPostCrackFields.py',
    'Check-TestShoulderAccuracy.py','Check-TestPersistentHEvidence.py','Validate-TestHardPrepared.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно, включая пробелы пустых полей отчёта.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
capacity=json.loads((base/cases['723']/'PriorBranchCapacityTimeoutProof.json').read_bytes())
steel=json.loads((base/cases['725']/'PriorBranchSteelEndProof.json').read_bytes())
direct=json.loads((base/cases['726']/'DirectProbeProof.json').read_bytes())
speed=json.loads((base/cases['723']/'HardShearDirectComparisonProof.json').read_bytes())
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    prescribed_steel_end=dict(proof=str(out.relative_to(repo)).replace('\\','/'),
        steel725_native=steel['native'],steel725_limit_reached=.02,all25_states_independently_verified=True,
        all_original21_history_bits_exact=True,maximum_force_error_N=max(p['error_N'] for p in steel['curve']),
        automatic723=capacity,next_direct726_solved=True,next_direct726_error_N=direct['last_analytical']['error_N'],
        six_force_HardShear_Direct_comparison=speed,working_material_unchanged=True,physical_capacity_certified=False),
    previous_goal_turn_classification='progress: prescribed curve reaches existing steel limit .02 with230/0 and25 independent mechanics checks; cold723 still timesout600 and has32 analytical misses; exact53prefix and next ordinary step independently verified in726; actual six-force HardShear DirectOK19/0 takes17.25s vs10.43s prior material. Critical shoulder accuracy now isolated rather than repeating cold search.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('723-726 SAVED',len(pins),'FILES; STEEL LIMIT PRESCRIBED ONLY; AUTOMATIC CAPACITY AND PERFORMANCE STILL OPEN')
