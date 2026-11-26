"""ДЛЯ ТЕСТОВ. Проверяет прогнозы активной ветви по полным локальным напряжениям, шести усилиям, связям и точному исходному управлению."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;file=here/'Check-TestCold32OutgoingStressScan.py';code=file.read_text('utf-8-sig')
for old,new in [('Cold906Cold32OutgoingStressScan_20261126','Cold916ProspectiveAffineRoots_20261126'),
 ('[3.];predictions=[r.VectorTest() for _ in range(3)]','[7.];predictions=[r.VectorTest() for _ in range(7)]'),
 ("[c['case'] for c in cases]==[1,2,3]","[c['case'] for c in cases]==list(range(1,8))"),
 ('len(rows)==3*8*129*2','len(rows)==7*8*129*2'),("key[0] in (1,2,3)","key[0] in range(1,8)"),
 ('Cold32OutgoingStressScanProof.json','ProspectiveAffineRootsProof.json')]:
    assert old in code,old;code=code.replace(old,new)
code=code[:code.index("print('READONLY LOCAL SCAN VERIFIED;")]
namespace={'__file__':str(file),'__name__':'prospective_affine_roots_local_test'}
exec(compile(code,str(file),'exec'),namespace)
root=namespace['root'];R=namespace['R'];r=R((root/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,old_delta=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);assert r.position==len(r.raw)
results=[]
for case,field in zip(namespace['cases'],namespace['predictions']):
    c=case['case'];value=namespace['assembled'][c,1.]
    factor=namespace['factor']+case['lambda_increment']
    control=-rhs[-1]+original[-1,:60]@(field-current)+original[-1,-1]*(factor-meta[2])
    bound=float(.0001*(namespace['source']['options'][3]+namespace['source']['options'][4]*max(abs((namespace['L']@(namespace['source']['offset']+factor*namespace['source']['direction']))[:56]))))
    gauge=float(meta[4]*max(abs(value['residual'][56:])))
    norm=max(value['physical_norm'],value['six_force_norm'],abs(control),gauge)
    strain=np.einsum('spcd,d->spc',namespace['op'],field)
    bars=np.einsum('spc,pc->sp',strain[:,25:],namespace['axis'][25:])
    long=namespace['axis'][25:,2]==1
    util=float(max((abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)).ravel()))
    results.append(dict(case=c,name=case['name'],lambda_value=float(factor),physical_weak_norm_MPa=value['physical_norm'],
        six_force_norm_MPa=value['six_force_norm'],scalar_control_residual_MPa=float(control),
        gauge_scaled_norm_MPa=gauge,bound_MPa=bound,total_norm_MPa=float(norm),
        maximum_three_station_steel_utilization=util,independently_equilibrated=bool(norm<=bound+1e-12 and util<=1)))
proof=namespace['proof'];proof.update(exact_captured_control_checked_independently=True,
    affine_branch_matrix_and_seed_residual_used=True,endpoint_results=results,
    independently_equilibrated_affine_predictions=sum(x['independently_equilibrated'] for x in results[2:]),
    no_native_Newton_or_history_acceptance=True,physical_capacity_certified=False)
(root/'ProspectiveAffineRootsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in results:print(item['name'],'WEAK',item['physical_weak_norm_MPa'],'CONTROL',item['scalar_control_residual_MPa'],'TOTAL',item['total_norm_MPa'],'BOUND',item['bound_MPa'],'EQUILIBRIUM',item['independently_equilibrated'])
