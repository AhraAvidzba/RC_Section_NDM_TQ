"""ДЛЯ ТЕСТОВ. Проверяет фактическую невязку, шесть сил и исходное управление прогнозов обеих сторон поверхности."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;file=here/'Check-TestCold32OutgoingStressScan.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold906Cold32OutgoingStressScan_20261126','Cold932SurfaceCorrectionScan_20261126'),
 ('Cold900Cold32OutgoingJacobian_20261126','Cold931SurfaceMatrices_20261126'),
 ('assert len(fields)==32','assert len(fields)==25'),
 ('[3.];predictions=[r.VectorTest() for _ in range(3)]','[87.];predictions=[r.VectorTest() for _ in range(87)]'),
 ("[c['case'] for c in cases]==[1,2,3]","[c['case'] for c in cases]==list(range(1,88))"),
 ('len(rows)==3*8*129*2','len(rows)==87*8*129*2'),('key[0] in (1,2,3)','key[0] in range(1,88)'),
 ("int(row[1])==32","int(row[1])==25"),('[:,52]','[:,54]'),('[52]','[54]'),('[:25,52]','[:25,54]'),('[25:,52]','[25:,54]'),
 ('row53','row55'),('immutable_actual_compiled_book900_reused','immutable_actual_compiled_book931_reused'),
 ('actual_components=256','actual_components=257'),('source32_history_and_stresses_exact','source25_history_and_stresses_exact'),
 ('Cold32OutgoingStressScanProof.json','SurfaceCorrectionScanProof.json')]:
    assert old in code,old;code=code.replace(old,new)
code=code[:code.index("print('READONLY LOCAL SCAN VERIFIED;")]
scope={'__file__':str(file),'__name__':'surface_correction_scan_test'}
exec(compile(code,str(file),'exec'),scope)
root=scope['root'];R=scope['R'];r=R((root/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,old_delta=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);expected=np.array([r.VectorTest() for _ in range(61)])
assert r.position==len(r.raw)
np.testing.assert_array_equal(current,scope['predictions'][1])
items=[];groups={}
for case,field in zip(scope['cases'],scope['predictions']):
    c=case['case'];value=scope['assembled'].get((c,1.))
    if value is None:
        items.append(dict(case=c,name=case['name'],material_available=False));continue
    lam=scope['factor']+case['lambda_increment']
    control=-rhs[-1]+original[-1,:60]@(field-current)+original[-1,-1]*(lam-meta[2])
    gauge=meta[4]*max(abs(value['residual'][56:]))
    norm=float(max(value['physical_norm'],value['six_force_norm'],abs(control),gauge))
    source=scope['source'];bound=float(source['options'][3]+source['options'][4]*max(abs((scope['L']@(source['offset']+lam*source['direction']))[:56])))
    item=dict(case=c,name=case['name'],material_available=True,lambda_value=float(lam),total_norm_MPa=norm,
        physical_norm_MPa=value['physical_norm'],six_force_norm_MPa=value['six_force_norm'],
        gauge_scaled_norm_MPa=float(gauge),control_residual_MPa=float(control),original_bound_MPa=bound,
        satisfies_original_equilibrium=bool(norm<=bound+1e-12))
    if c>2:
        f=case['correction_fraction'];key=str(case['matrix_index'])+'_'+case['mode']
        item.update(matrix_index=case['matrix_index'],mode=case['mode'],fraction=f,
            improves_original_merit=bool(norm<max(abs(residual[:56]))*(1-1e-4*f)))
        groups.setdefault(key,[]).append(item)
    items.append(item)
proof=scope['proof'];proof.update(endpoint_results=items,original_force_gauge_scalar_and_six_force_norm_checked=True,
    independently_equilibrated_endpoints=sum(x.get('satisfies_original_equilibrium',False) for x in items[2:]),
    improving_endpoints=sum(x.get('improves_original_merit',False) for x in items[2:]),
    no_native_NDS_or_material_history_acceptance=True,physical_capacity_certified=False)
(root/'SurfaceCorrectionScanProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for key,group in groups.items():
    best=min(group,key=lambda x:x['total_norm_MPa']);good=[x for x in group if x['improves_original_merit']]
    print(key,'BEST',best['total_norm_MPa'],'AT',best['fraction'],'IMPROVING',[(x['fraction'],x['total_norm_MPa']) for x in good])
print('ORIGINAL EQUILIBRIA',proof['independently_equilibrated_endpoints'])
