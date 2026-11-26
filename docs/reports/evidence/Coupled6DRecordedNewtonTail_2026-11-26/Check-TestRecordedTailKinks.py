"""ДЛЯ ТЕСТОВ. Локализует пробные поверхности вдоль фактической поправки54; смена ветви не считается разрушением."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent
scope=run_path(str(here/'Check-TestRecordedTailStressScan.py'))
root=scope['root'];s=scope['scope'];current=scope['current'];delta=scope['delta'];yield_function=s['TestYield']
crossings=[]
for point in range(1,44):
    for side in (-1,0,1):
        before=yield_function(point,side,current)
        after=yield_function(point,side,current+delta[:60])
        if not before<0<after:continue
        lo=0.;hi=1.
        for _ in range(60):
            mid=(lo+hi)/2
            if mid==lo or mid==hi:break
            if yield_function(point,side,current+mid*delta[:60])>0:hi=mid
            else:lo=mid
        crossings.append(dict(point=point,station_z_over_L=side*.001,side=side,
            fraction_lower=lo,fraction_upper=hi,trial_surface_before_MPa=before,trial_surface_full_step_MPa=after))
crossings.sort(key=lambda x:x['fraction_upper'])
assert crossings[0]['point']==5 and crossings[0]['side']==0
assert 1/1024<crossings[0]['fraction_lower']<crossings[0]['fraction_upper']<1/256
near=next(x for x in scope['items'] if x['fraction']==1/65536)
finite=next(x for x in scope['items'] if x['fraction']==1/128)
relative=abs(near['row55_actual_direction_MPa']/near['row55_predicted_direction_MPa']-1)
assert relative<1e-5 and finite['row55_actual_direction_MPa']<0
assert finite['largest_row55_points'][0]['point']==5
assert all(v['all_local_returns_available'] for segment in scope['proof']['segments'] for v in segment['samples'])
proof=dict(actual_frozen_record_pair=[54,55],accepted_prefix_states=25,
    first_trial_surface_crossing=crossings[0],all_incoming_to_outgoing_trial_surface_crossings=crossings,
    original_minimum_fraction=1/128,original_minimum_fraction_beyond_first_crossing=True,
    near_fraction=1/65536,near_direction_relative_difference=relative,
    near_direction_MPa=near['row55_actual_direction_MPa'],matrix_direction_MPa=near['row55_predicted_direction_MPa'],
    refined_matrix_same_direction_MPa=near['row55_refined_predicted_direction_MPa'],
    finite_minimum_fraction_direction_MPa=finite['row55_actual_direction_MPa'],
    finite_row55_point5_contribution_MPa=finite['largest_row55_points'][0]['derivative_MPa'],
    local_material_returns_all_available=True,accepted_material_history_unchanged=True,
    numerical_crossing_is_not_physical_capacity=True,no_all_NumFail_cause_claim=True,
    no_new_NDS_search=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'RecordedTailKinksProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FIRST CROSSING',crossings[0],'NEAR RELATIVE DIRECTION ERROR',relative)
