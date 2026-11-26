"""ДЛЯ ТЕСТОВ. Измеряет обе односторонние границы напряжений вокруг первой смены ветви в точной истории25."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;file=here/'Check-TestRecordedTailStressScan.py'
code=file.read_text('utf-8-sig')
for old,new in [
 ('Cold926RecordedTailStressScan_20261126','Cold928RecordedTailBoundaryScan_20261126'),
 ('[18.];predictions=[r.VectorTest() for _ in range(18)]','[32.];predictions=[r.VectorTest() for _ in range(32)]'),
 ('list(range(1,19))','list(range(1,33))'),('len(rows)==18*8*129*2','len(rows)==32*8*129*2'),
 ('key[0] in range(1,19)','key[0] in range(1,33)')]:
    assert old in code,old;code=code.replace(old,new)
scope={'__file__':str(file),'__name__':'recorded_tail_boundary_check_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];s=scope['scope'];cases=s['cases'];stresses=s['stresses'];assembled=scope['assembled']
assert all(item['all_local_returns_available'] for segment in scope['proof']['segments'] for item in segment['samples'])
items=[]
for k in range(4,11):
    lower=cases[18+2*(k-4)];upper=cases[19+2*(k-4)]
    a,b=lower['case'],upper['case'];sigma_a=stresses[a,1.,5,0];sigma_b=stresses[b,1.,5,0]
    items.append(dict(fraction_half_width=10.**(-k),lower_fraction=lower['newton_fraction'],upper_fraction=upper['newton_fraction'],
        maximum_point5_center_stress_difference_MPa=float(max(abs(sigma_b-sigma_a))),
        row55_residual_difference_MPa=float(assembled[b,1.]['residual'][54]-assembled[a,1.]['residual'][54]),
        trial_surface_lower_MPa=s['yields'][a,1.,5,0][1],trial_surface_upper_MPa=s['yields'][b,1.,5,0][1],
        lower_sigma_MPa=sigma_a.tolist(),upper_sigma_MPa=sigma_b.tolist()))
proof=dict(native=scope['proof']['native'],accepted_history_native_guard_unchanged=True,
    actual257_compiled_sources_exact925=True,all_material_returns_available=True,
    no_new_NDS_or_material_history_acceptance=True,boundary_pairs=items,
    continuity_judgment_requires_convergence_of_pair_differences=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'RecordedTailBoundaryScanProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in items:print('WIDTH',item['fraction_half_width'],'STRESS DIFF',item['maximum_point5_center_stress_difference_MPa'],
    'ROW55 DIFF',item['row55_residual_difference_MPa'],'YIELD',item['trial_surface_lower_MPa'],item['trial_surface_upper_MPa'])
