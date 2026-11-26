"""ДЛЯ ТЕСТОВ. Воспроизводит отклик фактического отказа54/55 и выделяет бетонные вклады вдоль обеих сторон поправки."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json
import numpy as np

here=Path(__file__).resolve().parent
file=here/'Check-TestCold32OutgoingStressScan.py'
code=file.read_text('utf-8-sig')
replacements=[
 ('Cold906Cold32OutgoingStressScan_20261126','Cold926RecordedTailStressScan_20261126'),
 ('Cold900Cold32OutgoingJacobian_20261126','Cold925RecordedTailSystem_20261126'),
 ('assert len(fields)==32','assert len(fields)==25'),
 ('[3.];predictions=[r.VectorTest() for _ in range(3)]','[18.];predictions=[r.VectorTest() for _ in range(18)]'),
 ("[c['case'] for c in cases]==[1,2,3]","[c['case'] for c in cases]==list(range(1,19))"),
 ('len(rows)==3*8*129*2','len(rows)==18*8*129*2'),
 ('key[0] in (1,2,3)','key[0] in range(1,19)'),
 ("int(row[1])==32","int(row[1])==25"),
 ('[:,52]','[:,54]'),('[52]','[54]'),('[:25,52]','[:25,54]'),('[25:,52]','[25:,54]'),
 ('row53','row55'),
 ('immutable_actual_compiled_book900_reused','immutable_actual_compiled_book925_reused'),
 ('actual_components=256','actual_components=257'),
 ('source32_history_and_stresses_exact','source25_history_and_stresses_exact'),
 ('Cold32OutgoingStressScanProof.json','RecordedTailStressScanProof.json')]
for old,new in replacements:
    assert old in code,old
    code=code.replace(old,new)
code=code[:code.index("print('READONLY LOCAL SCAN VERIFIED;")]
scope={'__file__':str(file),'__name__':'recorded_tail_stress_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];parent=scope['parent'];R=scope['R'];assembled=scope['assembled']
prep=json.loads((root/'Preparation.json').read_bytes())
for name,expected in prep['inputs'].items():
    assert sha256((root/name).read_bytes()).hexdigest().upper()==expected,name
for path in (parent/'tests').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),str(path)
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257

def TestReadCapture(path):
    """Читает записанные до линейного решения матрицу и направление, затем фактически вычисленные пробы."""
    reader=R(path.read_bytes());reader.position=0
    values=[reader.VectorTest() for _ in range(6)]
    matrix=np.array([reader.VectorTest() for _ in range(61)])
    end=reader.VectorTest();count=int(reader.VectorTest()[0]);probes=[]
    for _ in range(count):probes.append([reader.VectorTest() for _ in range(3)])
    assert reader.position==len(reader.raw)
    return values,matrix,end,probes

(meta,current,accepted,residual,rhs,delta),A,end,probes=TestReadCapture(root/'Input/ActualRejectedNewton.bin')
(meta_ref,current_ref,accepted_ref,residual_ref,rhs_ref,delta_ref),Aref,end_ref,probes_ref=TestReadCapture(root/'Input/ActualRefinedNewton.bin')
for left,right in ((current,current_ref),(accepted,accepted_ref),(residual,residual_ref),(rhs,rhs_ref)):
    np.testing.assert_array_equal(left,right)
np.testing.assert_array_equal(accepted,scope['accepted'])
np.testing.assert_array_equal(current,scope['predictions'][1])
interpolated=accepted+(current-accepted)
interpolation_error=abs(interpolated-current)
interpolation_bound=2*np.finfo(float).eps*(abs(accepted)+abs(current-accepted))
assert np.all(interpolation_error<=interpolation_bound)  # Операции вычитания и сложения дают до двух округлений.
interpolation_strain_error=float(np.max(abs(np.einsum('spcd,d->spc',scope['op'],interpolated-current))))
base=assembled[2,1.]
base_error=float(max(abs(base['residual']-residual)))
assert base_error<2e-11  # Только воспроизведение из 15-значного CSV; допуск принятия НДС не меняется.
predicted=A@delta;refined_predicted=Aref@delta
items=[];probe_error=0.;matched=0
for case in scope['cases'][2:]:
    fraction=case['newton_fraction'];value=assembled[case['case'],1.]
    derivative=(value['residual']-base['residual'])/fraction
    point=(value['points']-base['points'])[:,54]/fraction
    order=np.argsort(abs(point))[::-1][:8]
    matching=[(info,field,actual) for info,field,actual in probes if info[0]==fraction]
    if matching:
        matched+=1
        np.testing.assert_array_equal(matching[0][1],scope['predictions'][case['case']-1])
        probe_error=max(probe_error,float(max(abs(value['residual']-matching[0][2]))))
    direction_strain=np.einsum('spcd,d->spc',scope['op'],scope['predictions'][case['case']-1]-current)
    items.append(dict(fraction=fraction,independent_physical_residual_MPa=value['physical_norm'],
        maximum_strain_increment=float(np.max(abs(direction_strain))),
        row55_actual_direction_MPa=float(derivative[54]),
        row55_predicted_direction_MPa=float(predicted[54]),
        row55_refined_predicted_direction_MPa=float(refined_predicted[54]),
        row55_real_concrete_direction_MPa=float(sum((value['concrete']-base['concrete'])[:25,54])/fraction),
        row55_subtracted_concrete_direction_MPa=float(-sum((value['concrete']-base['concrete'])[25:,54])/fraction),
        row55_steel_direction_MPa=float(sum((value['steel']-base['steel'])[:,54])/fraction),
        row55_other_direction_MPa=float((value['other'][54]-base['other'][54])/fraction),
        largest_row55_points=[dict(point=int(p)+1,kind='real_concrete' if p<25 else 'steel_minus_replaced_concrete',
            derivative_MPa=float(point[p]),
            yield_before={str(s):scope['yields'][2,1.,int(p)+1,s] for s in (-1,0,1)},
            yield_after={str(s):scope['yields'][case['case'],1.,int(p)+1,s] for s in (-1,0,1)}) for p in order]))
assert matched==len(probes)==8
assert probe_error<2e-11
proof=scope['proof']
proof.update(actual_rejected_iterate_reconstructed=True,
    actual_capture_base_maximum_residual_difference_MPa=base_error,
    actual_capture_probes_maximum_residual_difference_MPa=probe_error,
    actual_capture_probes_matched=matched,endpoint_comparisons=items,
    endpoint_interpolation_bitexact=bool(np.array_equal(interpolated,current)),
    endpoint_interpolation_maximum_field_error=float(max(interpolation_error)),
    endpoint_interpolation_maximum_strain_error=interpolation_strain_error,
    endpoint_interpolation_error_checked_against_actual_recorded_residuals=True,
    source25_final_prefix_only=True,all257_sources_preserved=True,
    frozen_material_history_unchanged=True,read_only_endpoints_from_actual_Newton_correction=True,
    cached_point11_not_assumed_cause=True,numerical_no_capacity_claim=True)
(root/'RecordedTailStressScanProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ACTUAL TAIL REBUILT;BASE ERROR',base_error,'PROBE ERROR',probe_error)
for item in items:
    print('F',item['fraction'],'NORM',item['independent_physical_residual_MPa'],
        'D55',item['row55_actual_direction_MPa'],'PRED',item['row55_predicted_direction_MPa'],
        'TOP',[(p['point'],p['derivative_MPa']) for p in item['largest_row55_points'][:3]])
