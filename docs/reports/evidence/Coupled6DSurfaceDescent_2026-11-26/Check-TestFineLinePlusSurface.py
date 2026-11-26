"""ДЛЯ ТЕСТОВ. Собирает невязки нового отказа929 и односторонние напряжения положительной станции точки5."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;file=here/'Check-TestRecordedTailStressScan.py'
code=file.read_text('utf-8-sig')
for old,new in [('Cold926RecordedTailStressScan_20261126','Cold930FineLinePlusSurface_20261126'),
 ('[18.];predictions=[r.VectorTest() for _ in range(18)]','[38.];predictions=[r.VectorTest() for _ in range(38)]'),
 ('list(range(1,19))','list(range(1,39))'),('len(rows)==18*8*129*2','len(rows)==38*8*129*2'),
 ('key[0] in range(1,19)','key[0] in range(1,39)'),('matched==len(probes)==8','matched==len(probes)==17')]:
    assert old in code,old;code=code.replace(old,new)
scope={'__file__':str(file),'__name__':'fine_line_plus_surface_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];s=scope['scope'];cases=s['cases'];stresses=s['stresses'];assembled=scope['assembled']
assert all(x['all_local_returns_available'] for segment in scope['proof']['segments'] for x in segment['samples'])
items=[]
for k in range(6,13):
    lower=cases[24+2*(k-6)];upper=cases[25+2*(k-6)]
    a,b=lower['case'],upper['case'];sa=stresses[a,1.,5,1];sb=stresses[b,1.,5,1]
    items.append(dict(fraction_half_width=10.**(-k),lower_fraction=lower['newton_fraction'],upper_fraction=upper['newton_fraction'],
        maximum_point5_plus_stress_difference_MPa=float(max(abs(sb-sa))),
        row55_residual_difference_MPa=float(assembled[b,1.]['residual'][54]-assembled[a,1.]['residual'][54]),
        trial_yield_lower_MPa=s['yields'][a,1.,5,1][1],trial_yield_upper_MPa=s['yields'][b,1.,5,1][1],
        lower_sigma_MPa=sa.tolist(),upper_sigma_MPa=sb.tolist()))
proof=scope['proof'];proof.update(boundary_pairs=items,actual_source_record_pair929=[1,2],
    accepted25_history_unchanged=True,all_material_returns_available=True,
    central_surface_continuity_not_substituted_for_plus_station=True)
(root/'FineLinePlusSurfaceProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in items:print('WIDTH',item['fraction_half_width'],'PLUS STRESS DIFF',item['maximum_point5_plus_stress_difference_MPa'],
    'ROW55 DIFF',item['row55_residual_difference_MPa'],'YIELD',item['trial_yield_lower_MPa'],item['trial_yield_upper_MPa'])
