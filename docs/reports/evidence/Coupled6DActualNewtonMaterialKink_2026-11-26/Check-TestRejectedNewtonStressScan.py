"""ДЛЯ ТЕСТОВ. Собирает отклик в действительной неудачной точке и выделяет локальные источники изменения направления."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;file=here/'Check-TestCold32OutgoingStressScan.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold906Cold32OutgoingStressScan_20261126','Cold913RejectedNewtonStressScan_20261126'),
 ('[3.];predictions=[r.VectorTest() for _ in range(3)]','[13.];predictions=[r.VectorTest() for _ in range(13)]'),
 ("[c['case'] for c in cases]==[1,2,3]","[c['case'] for c in cases]==list(range(1,14))"),
 ('len(rows)==3*8*129*2','len(rows)==13*8*129*2'),("key[0] in (1,2,3)","key[0] in range(1,14)"),
 ('Cold32OutgoingStressScanProof.json','RejectedNewtonStressScanProof.json')]:
    assert old in code,old;code=code.replace(old,new)
code=code[:code.index("print('READONLY LOCAL SCAN VERIFIED;")]
namespace={'__file__':str(file),'__name__':'actual_rejected_newton_stress_test'}
exec(compile(code,str(file),'exec'),namespace)
root=namespace['root'];parent=namespace['parent'];R=namespace['R'];assembled=namespace['assembled'];cases=namespace['cases']
r=R((root/'Input/ActualRejectedNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,delta=[r.VectorTest() for _ in range(6)]
A=np.array([r.VectorTest() for _ in range(61)]);end=r.VectorTest();count=int(r.VectorTest()[0]);captured=[]
for _ in range(count):captured.append([r.VectorTest() for _ in range(3)])
assert r.position==len(r.raw)
np.testing.assert_array_equal(accepted,namespace['accepted'])
np.testing.assert_array_equal(current,namespace['predictions'][1])
assert np.array_equal(current,accepted+(current-accepted)), 'Иначе интерполяция уровня1 требует отдельного учёта округления.'
basevalue=assembled[2,1.]
capture_error=float(max(abs(basevalue['residual']-residual)))
assert capture_error<2e-11 # Независимая сборка из 15-значного текстового экспорта; не изменение допуска приёмки НДС.
predicted=A@delta;items=[];maximum_captured_probe_error=0.
for case in cases[2:]:
    f=case['newton_fraction'];value=assembled[case['case'],1.]
    derivative=(value['residual']-basevalue['residual'])/f
    point=(value['points']-basevalue['points'])[:,52]/f;order=np.argsort(abs(point))[::-1][:8]
    matches=[(info,actual) for info,_,actual in captured if info[0]==f]
    if matches:maximum_captured_probe_error=max(maximum_captured_probe_error,float(max(abs(value['residual']-matches[0][1]))))
    field=namespace['predictions'][case['case']-1]
    direction_strain=np.einsum('spcd,d->spc',namespace['op'],field-current)
    items.append(dict(fraction=f,independent_physical_residual_MPa=value['physical_norm'],
        maximum_strain_increment=float(np.max(abs(direction_strain))),
        row53_actual_direction_MPa=float(derivative[52]),row53_predicted_direction_MPa=float(predicted[52]),
        row53_real_concrete_direction_MPa=float(sum((value['concrete']-basevalue['concrete'])[:25,52])/f),
        row53_subtracted_concrete_direction_MPa=float(-sum((value['concrete']-basevalue['concrete'])[25:,52])/f),
        row53_steel_direction_MPa=float(sum((value['steel']-basevalue['steel'])[:,52])/f),
        row53_other_direction_MPa=float((value['other'][52]-basevalue['other'][52])/f),
        largest_row53_points=[dict(point=int(p)+1,kind='real_concrete' if p<25 else 'steel_minus_replaced_concrete',
            derivative_MPa=float(point[p]),yield_before={str(s):namespace['yields'][2,1.,int(p)+1,s] for s in (-1,0,1)},
            yield_after={str(s):namespace['yields'][case['case'],1.,int(p)+1,s] for s in (-1,0,1)}) for p in order]))
assert maximum_captured_probe_error<2e-11
proof=namespace['proof'];proof.update(actual_rejected_iterate_reconstructed=True,
    actual_capture_base_maximum_residual_difference_MPa=capture_error,
    actual_capture_probes_maximum_residual_difference_MPa=maximum_captured_probe_error,
    read_only_endpoints_from_actual_Newton_correction=True,endpoint_comparisons=items,
    numerical_no_capacity_claim=True)
(root/'RejectedNewtonStressScanProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ACTUAL FAILED ITERATE REBUILT;BASE ERROR',capture_error,'CAPTURED PROBE ERROR',maximum_captured_probe_error)
for item in items:print('FRACTION',item['fraction'],'NORM',item['independent_physical_residual_MPa'],'D53',item['row53_actual_direction_MPa'],'TOP',item['largest_row53_points'][0]['point'])
