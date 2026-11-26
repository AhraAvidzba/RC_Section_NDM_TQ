"""ДЛЯ ТЕСТОВ. Связывает отклонённую поправку, её малые пробы и фактическое пересечение поверхности бетона №5."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
capture=base/'Cold911NewtonRejectedCapture_20261126';damping=base/'Cold912ExtendedSpatialDamping_20261126';scan=base/'Cold913RejectedNewtonStressScan_20261126'
a=json.loads((capture/'NewtonRejectedCaptureProof.json').read_bytes());b=json.loads((damping/'ExtendedSpatialDampingProof.json').read_bytes());c=json.loads((scan/'RejectedNewtonStressScanProof.json').read_bytes())
assert a['recorded_rejected_directions']==2 and not b['lost_original_cases'] and not b['gained_original_cases']
assert c['actual_rejected_iterate_reconstructed'] and c['accepted_history_native_guard_unchanged']
assert all(x['all_local_returns_available'] for s in c['segments'] for x in s['samples'])
validate=run_path(str(here/'Validate-TestHardPrepared.py'));rows=validate['TestRows'](scan/'Native/ArcScan.csv')
yield5={int(r[1]):float(r[6]) for r in rows if r[0]=='Y' and float(r[2])==1. and int(r[3])==5 and int(r[4])==1}
assert yield5[2]<0 and yield5[4]>0 and yield5[5]<0
near=next(x for x in c['endpoint_comparisons'] if x['fraction']==1/256)
relative=abs(near['row53_actual_direction_MPa']/near['row53_predicted_direction_MPa']-1)
assert relative<.002 and c['endpoint_comparisons'][1]['row53_actual_direction_MPa']<0
assert c['endpoint_comparisons'][1]['largest_row53_points'][0]['point']==5
baseline=base/'Cold909ResidualArcPredictorBaseline_20261126';same=[]
for path in (baseline/'Native').glob('Arc32.Case*'):
    if path.is_file():assert path.read_bytes()==(damping/'Native'/path.name).read_bytes();same.append(path.name)
assert len(same)==40
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];later=[]
for path in sorted((damping/'Native').glob('NewtonRejected.*.bin')):
    r=R(path.read_bytes());r.position=0;meta,field,history,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
    A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();n=int(r.VectorTest()[0]);samples=[]
    for _ in range(n):
        info,trial,actual=[r.VectorTest() for _ in range(3)];samples.append(info.tolist())
    assert r.position==len(r.raw) and n==13 and end[-1]==1/4096
    later.append(dict(serial=int(path.stem.split('.')[-1]),meta=meta.tolist(),failure=end.tolist(),valid_samples=samples))
assert len(later)==2 and later[0]['meta'][1]==3 and later[1]['meta'][1]==4
reduction=1-later[0]['failure'][1]/a['records'][0]['norm_MPa']
result=dict(actual_failing_step_not_accepted_state_predictor=True,
    exact_linear_system_maximum_backward_error=max(x['linear_backward_error'] for x in a['records']),
    current_norm_MPa=a['records'][0]['norm_MPa'],near_fraction=1/256,
    near_direction_relative_difference=relative,near_direction_MPa=near['row53_actual_direction_MPa'],
    matrix_direction_MPa=near['row53_predicted_direction_MPa'],
    finite_fraction1_over128_direction_MPa=c['endpoint_comparisons'][1]['row53_actual_direction_MPa'],
    real_concrete_point5_station_plus_trial_yield_MPa={str(k):yield5[k] for k in (2,4,5)},
    active_surface_crossing_fraction_bracket=[1/256,1/128],
    local_material_returns_all_available=True,accepted_material_history_unchanged=True,
    base_residual_reconstruction_error_MPa=c['actual_capture_base_maximum_residual_difference_MPa'],
    local_norm_reduction_from_five_more_halvings=reduction,later_rejected_directions=later,
    all40_root_and_operator_exports_exact_baseline=same,
    additional_arc_cases_solved=0,original_arc_cases_lost=0,
    extended_damping_not_selected_as_a_solution=True,cold_for_unhelpful_damping_not_run=True,
    material_local_active_set_transition_is_not_physical_failure=True,
    no_general_cause_of_all_NumFail_claimed=True,physical_capacity_certified=False,whole_goal_completed=False)
(scan/'ActualNewtonKinkDiagnosisProof.json').write_text(json.dumps(result,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ACTUAL MATERIAL SURFACE KINK VERIFIED;NEAR DIRECTION ERROR',relative,'EXTENDED DAMPING LOCAL REDUCTION',reduction,'NO NEW ARC CASES')
