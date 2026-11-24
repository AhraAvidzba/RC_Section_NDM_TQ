"""ДЛЯ ТЕСТОВ. Фиксирует численный исход734 отдельно от ещё не выполненной независимой проверки полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold734CapacityScaledGrowth_20261124';parent=base/'Cold732ProvidedJacobianCapacity_20261124'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
assert native['own_process_closed']
if native['timeout600']:
    frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Native/Stock.RawHistory.bin')
    assert raw['incomplete_tail_bytes']==0;frame=frames[-1]
    assert not (root/'Native/AutomaticCapacityOutcome.txt').exists()
    proof=dict(native=native,source_raw_history=raw,accepted_states=len(frame['fields']),
        last_control_strain=float(frame['fields'][-1,0]),largest_accepted_lambda=float(max(frame['lambdas'])),
        numerical_search_completed=False,numerical_stop_reason_not_established=True,
        native_capacity_result_missing=True,independent_all_fields_pending=True,
        physical_capacity_certified=False,entire_goal_completed=False)
    (root/'CapacityScaledGrowthProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('734 TIMEOUT600 DURING SEARCH; NO NUMERICAL OUTCOME; ACCEPTED',proof['accepted_states'],'LAST_EPS',proof['last_control_strain'])
    raise SystemExit(0)
outcome=(root/'Native/AutomaticCapacityOutcome.txt').read_text('cp1251')
flags,=re.findall(r'AUTOMATIC_RESULT: hasPeak=(True|False); peak=([^;]+); accuracy=([^;]+); designVerified=(True|False); steelEvent=(\d+)',outcome)
passed,failed=map(int,re.findall(r'PREEXPORT_TESTS: passed=(\d+); failed=(\d+)',outcome)[0])
assert (passed,failed)==(native['passed'],native['failed'])
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0;frame=frames[-1]
initial=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Native/PostCrackPrescribed.Accepted2.bin')
for key in ('fields','fractions','lambdas'):assert np.array_equal(frame[key][:2],initial[key]),key
name='src/Solver/CCoupledDeformationContinuation.cls'
old=(parent/name).read_text('utf-8-sig');new=(root/name).read_text('utf-8-sig')
anchor='    If Abs(trial.Lambda - current.Lambda - predictedLoad) > _\n            GeomMax(Abs(predictedLoad) * POST_PEAK_SECANT_TOLERANCE, LoadResolution(current.Lambda)) Then Exit Function'
replacement='''    Dim growthResolutionTest As Double: growthResolutionTest = LoadResolution(current.Lambda)
    If mFollowingLocalPeak Then growthResolutionTest = GeomMax(growthResolutionTest, mLambdaTolerance / 16#) ' Численный масштаб прогноза шага; не допуск приёмки равновесия или максимума.
    If Abs(trial.Lambda - current.Lambda - predictedLoad) > _
            GeomMax(Abs(predictedLoad) * POST_PEAK_SECANT_TOLERANCE, growthResolutionTest) Then Exit Function'''
assert new.replace(replacement,anchor,1)==old
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Materials/CCoupledConcreteResponse.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert compiled['Compiled'] and not compiled['ProductionChanged']
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
known=172967.56279474715/120000
probes=int(re.findall(r'проб: (\d+)',outcome)[0]);assert probes<=96
proof=dict(native=native,original_capacity_checks=dict(passed=passed,failed=failed),
    native_has_peak=flags[0]=='True',native_peak=float(flags[1]),native_accuracy=float(flags[2]),
    design_verified=flags[3]=='True',steel_event_point=int(flags[4]),probes=probes,
    accepted_states=len(frame['fields']),last_control_strain=float(frame['fields'][-1,0]),
    largest_accepted_lambda=float(max(frame['lambdas'])),known_uniaxial_maximum=known,
    largest_accepted_vs_known_difference=abs(float(max(frame['lambdas']))-known),
    source_raw_history=raw,only_postpeak_load_prediction_floor_changed=True,
    original_two_elastic_initial_states_exact=True,original96_budget_and_final_checks=True,
    independent_all_fields_pending=True,physical_capacity_certified=False,entire_goal_completed=False)
(root/'CapacityScaledGrowthProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('734 NATIVE',native,'HAS_PEAK',proof['native_has_peak'],'STEEL_EVENT',proof['steel_event_point'],
    'PROBES',probes,'STATES',proof['accepted_states'],'LAST_EPS',proof['last_control_strain'])
