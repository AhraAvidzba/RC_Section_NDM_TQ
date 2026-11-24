"""ДЛЯ ТЕСТОВ. Сохраняет отрицательный холодный контроль, точные повторы и независимые матрицы без заявления о пределе."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
names={'734':'Cold734CapacityScaledGrowth_20261124','735':'Cold735FrozenGrowthReplay_20261125',
    '736':'Cold736FrozenTailMatrix_20261125','737':'Cold737TailMatrixPathRetry_20261125'}
roots={key:base/name for key,name in names.items()}
growth=json.loads((roots['734']/'CapacityScaledGrowthProof.json').read_bytes())
regression=json.loads((roots['734']/'GrowthRegressionProof.json').read_bytes())
replay=json.loads((roots['735']/'FrozenCapacityReplayProof.json').read_bytes())
matrices=json.loads((roots['737']/'TailMatrixProof.json').read_bytes())
assert growth['native']['timeout600'] and not growth['numerical_search_completed']
assert regression['passed']==105 and regression['failed']==0
assert replay['source_states']==70 and replay['all_fields_independent_weak_gauge_and_six_force_checks']
assert matrices['native']['passed']==1 and matrices['native']['failed']==0
growth.update(independent_all_fields_pending=False,independent_all70_fields_verified_via735=True,
    analytical_force_misses=len(replay['analytical_force_misses']),
    maximum_analytical_force_error_N=max(p['error_N'] for p in replay['analytical_forces']))
(roots['734']/'CapacityScaledGrowthProof.json').write_text(json.dumps(growth,indent=2,sort_keys=True)+'\n',encoding='utf-8')
out=repo/'docs/reports/evidence/Coupled6DPostPeakAdaptation_2026-11-25';assert not out.exists();out.mkdir(parents=True)
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
for number,root in roots.items():
    files=[]
    for name in ['Preparation.json','CompilePoint.log','CompilePointFailure.json','CompilePointClosure.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
        'docs/regression/CodeUpdate.json','Run.ndjson','CapacityScaledGrowthProof.json','GrowthRegressionProof.json',
        'FrozenCapacityReplayProof.json','TailMatrixProof.json','ParentCodeUpdate.json','ParentCompileClosure.json',
        'src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
        'src/Solver/CCoupledHybridFactorization.cls','src/Materials/CCoupledConcreteResponse.cls',
        'tests/modTestCoupledElasticState.bas','OriginalNativeFailure736.txt','OriginalNativeFailure736Closure.json',
        'SavedCompilationRecovery.json','CodeUpdateBeforeRecovery.json','tools/build_workbook/Update-IsolatedOptimizerCode.ps1']:
        if (root/name).is_file():files.append(root/name)
    files += [p for p in (root/'Native').iterdir() if p.is_file() and p.suffix.lower() not in ('.xlsm','.xlsx')]
    files += [p for p in root.glob('Stage-Test*.py') if p.name in ('Stage-TestCapacityScaledGrowth.py',
        'Stage-TestFrozenCapacityReplay.py','Stage-TestFrozenTailMatrix.py','Stage-TestTailMatrixPathRetry.py')]
    for name in ('Input/PostCrackReplay.bin','Input/PostCrackReplayOptions.bin','Input/PostCrackSourceRawHistory.bin'):
        if (root/name).is_file():files.append(root/name)
    for source in files:
        dest=out/number/source.relative_to(root);dest.parent.mkdir(parents=True,exist_ok=True);assert not dest.exists();dest.write_bytes(source.read_bytes())
helpers=['Stage-TestCapacityScaledGrowth.py','Check-TestCapacityScaledGrowth.py','Stage-TestFrozenCapacityReplay.py',
    'Check-TestFrozenCapacityReplay.py','Stage-TestFrozenTailMatrix.py','Stage-TestTailMatrixPathRetry.py',
    'Check-TestFrozenTailMatrix.py','Check-TestPersistentHEvidence.py','Read-TestColdRawCapture.py',
    'Check-TestPostCrackFields.py','Validate-TestHardPrepared.py','Check-TestShoulderAccuracy.py',
    'Recover-TestSavedCompilation.ps1','Recover-TestSavedCompilation.OriginalNoBOM.ps1',
    'CandidateTestMinimumNormRankStep.cls.txt','Check-TestMinimumNormDerivation.py','MinimumNormDerivationProof.json',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
summary=dict(cold734=growth,regressions734=regression,replay735_native=replay['native'],
    all70_independent_mechanics_verified=True,analytical_misses=len(replay['analytical_force_misses']),
    tail_strain_field=replay['independent_fields']['point_fields'][-9:],matrices737=matrices['matrices'],
    working_code_promoted=False,physical_capacity_certified=False,whole_goal_completed=False)
(out/'Summary.json').write_text(json.dumps(summary,indent=2,sort_keys=True)+'\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    postpeak_adaptation734=growth,frozen735_replay=dict(native=replay['native'],all70_independent_fields_verified=True,
        analytical_misses=len(replay['analytical_force_misses'])),tail737_matrices=matrices,
    previous_goal_turn_classification='progress: tested isolated capacity-scaled load-growth gate: no improved automatic result, timeout600 before outcome.105/0 targeted regressions; exact735 replay verifies all70 states and preserves36analytical misses. New native controlled-matrix diagnostics736 failed only unset export path, corrected737 without core changes; independent SVD/QR now determines next action. No physicalcapacity andfullscope staysactive.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for path in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):path.write_bytes(raw)
print('734-737 ARCHIVE WRITTEN:',len(pins),'PINNED FILES; NO CAPACITY CLAIM')
