"""ДЛЯ ТЕСТОВ. Сохраняет четыре сравнения точности/стоимости и незавершённый диагностический экспорт727."""
from pathlib import Path
from hashlib import sha256
from datetime import datetime
import json

repo=Path.cwd();here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
out=repo/'docs/reports/evidence/Coupled6DShoulderCost_2026-11-24';assert not out.exists();out.mkdir(parents=True)
cases={'727':'Cold727CriticalShoulderProfile_20261124','728':'Cold728FastShoulderWitness_20261124',
    '729':'Cold729TighterLocalShoulder_20261124','730':'Cold730ProvidedJacobianShoulder_20261124',
    '731':'Cold731OuterToleranceShoulder_20261124'}
proofs={}
for number,name in cases.items():
    root=base/name
    names=['Preparation.json','CompilePointClosure.json','CompiledSourceReuse.json','PointTraceSourceComposition.json',
        'docs/regression/CodeUpdate.json','Run.ndjson','ShoulderWitnessProof.json',
        'Input/PostCrackReplay.bin','Input/PostCrackReplayOptions.bin',
        'src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
        'src/Materials/CCoupledConcreteResponse.cls','tests/modTestCoupledElasticState.bas']
    paths=[root/n for n in names if (root/n).is_file()]
    paths += [p for p in (root/'Native').iterdir() if p.is_file()]
    for p in paths:
        target=out/number/p.relative_to(root);target.parent.mkdir(parents=True,exist_ok=True);target.write_bytes(p.read_bytes())
    if number!='727':
        proof=json.loads((root/'ShoulderWitnessProof.json').read_bytes())
        assert proof['quality_confirmed'] and proof['independent_final_field']['states']==1
        proofs[number]=proof
baseline=base/cases['728']
for number in ('729','730','731'):
    root=base/cases[number]
    for name in ('Input/PostCrackReplay.bin','Input/PostCrackReplayOptions.bin'):
        assert (root/name).read_bytes()==(baseline/name).read_bytes(),(number,name)
comparison=dict(same_original52_binary_prefix_and_target=True,original_quality_target=.001/16,
    cases={k:dict(quality_confirmed=v['quality_confirmed'],elapsed_before_export_seconds=v['elapsed_before_export_seconds'],
        final_uncertainty=v['final_uncertainty'],analytical_error_N=v['analytical']['error_N'],
        analytical_within_original_halfNewton=v['analytical_within_original_halfNewton']) for k,v in proofs.items()},
    no_full_capacity_result_in_these_cases=True,working_code_promoted=False,physical_capacity_certified=False,
    prototype731_requires_fallback_for_rejected_early_candidate=True,entire_goal_completed=False)
(out/'Comparison.json').write_text(json.dumps(comparison,indent=2,sort_keys=True)+'\n',encoding='utf-8')
helpers=['Stage-TestCriticalShoulderProfile.py','Stage-TestFastShoulderWitness.py','Check-TestShoulderWitness.py',
    'Stage-TestTighterLocalShoulder.py','Stage-TestProvidedJacobianShoulder.py','Stage-TestOuterToleranceShoulder.py',
    'Stage-TestProvidedJacobianCapacity.py','Check-TestPostCrackFields.py','Check-TestShoulderAccuracy.py',
    'Validate-TestHardPrepared.py','Check-TestPersistentHEvidence.py',Path(__file__).name]
for name in helpers:(out/name).write_bytes((here/name).read_bytes())
(out/'.gitattributes').write_text('# Сырые свидетельства сохраняются побитно.\n* -text whitespace=-blank-at-eol,-blank-at-eof,space-before-tab,cr-at-eol\n',encoding='utf-8')
pins={str(p.relative_to(out)).replace('\\','/'):sha256(p.read_bytes()).hexdigest().upper() for p in out.rglob('*') if p.is_file()}
(out/'Files.json').write_text(json.dumps(pins,indent=2,sort_keys=True)+'\n',encoding='utf-8')
path=repo/'docs/reports/Coupled6DResumeState_2026-11-24.json';state=json.loads(path.read_bytes())
state.update(updated_local=datetime.now().astimezone().isoformat(),goal_status='active',
    critical_shoulder_cost=dict(evidence=str(out.relative_to(repo)).replace('\\','/'),**comparison),
    previous_goal_turn_classification='progress: exact same critical shoulder independently verified in four controlled variants; originalquality U target met inall, cost146->168(local tighter)->87(providedJacobian)->37.88s(inner stop aligned); .5N analytical check passes onlylocal tighter. Same globalquality/caps preserved; no productionpromotion. Actualcold732 from original2elastic fields now tests whether numerical improvement completes automatic capacity; full6 goal unchanged.')
raw=(json.dumps(state,ensure_ascii=False,indent=2,sort_keys=True)+'\n').encode('utf-8')
for p in (path,repo/'tmp/material-consistency/Continuation666_Resume20261123.json'):p.write_bytes(raw)
print('727-731 SAVED',len(pins),'FILES; CONTROLLED COST COMPARISON; NEGATIVE ANALYTICAL CHECKS RETAINED; NO CAPACITY CERTIFICATE')
