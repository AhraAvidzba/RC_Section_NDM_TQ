"""ДЛЯ ТЕСТОВ. Независимо проверяет возврат к неизменённым61 исходным уравнениям, включая линейную цель и шесть усилий."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold941OriginalControlReturn_20261126'
template=here/'Check-TestRecordedYieldCoordinate.py';code=template.read_text('utf-8-sig')
cut="r=R((root/'Native/RecordedYieldCoordinate.bin').read_bytes())";assert code.count(cut)==1
code=code[:code.index(cut)]
old="root=base/'Cold936RecordedYieldCoordinate_20261126'";assert code.count(old)==1
code=code.replace(old,"root=base/'Cold941OriginalControlReturn_20261126'",1)
scope={'__file__':str(template),'__name__':'original_control_return_independent_test'}
exec(compile(code,str(template),'exec'),scope)
R=scope['R'];meta=scope['meta'];source=scope['source'];projection=scope['projection'];op=scope['op'];axis=scope['axis'];L=scope['L']
r=R((root/'Native/OriginalControlReturn.bin').read_bytes());assert r.raw[:16]==b'NDMORIGCTRL00001'
assert r.VectorTest().tolist()==[2.]
cases=[];fields=[]
for index in range(1,3):
    info=r.VectorTest();assert len(info)==11 and info[0]==index and info[6]<=meta[7]
    item=dict(case=index,native_solved=bool(info[1]),material_ready=bool(info[2]),lambda_value=float(info[3]),
        total_native_residual_MPa=float(info[4]),initial_native_bound_MPa=float(info[5]),
        steps=int(info[6]),calls=int(info[7]),jacobians=int(info[8]),stop_code=int(info[10]),
        independently_solves_original_control=False)
    if info[2]:
        field,values,forces=r.VectorTest(),r.VectorTest(),r.VectorTest();stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,code=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
            assert code==0 and key not in seen and len(sigma)==6;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        assembled=scope['TestAssemble'](field,info[3],stress)
        physical_error=float(max(abs(values[:56]-assembled['residual'][:56])))
        gauge_error=float(max(abs(values[56:60]-meta[8]*assembled['residual'][56:])))
        force_error=float(max(abs(forces-assembled['six_force_residual_vector'])))
        control=float(meta[8]*(projection@field-meta[4]));control_error=float(abs(control-values[-1]))
        assert max(physical_error,gauge_error,force_error,control_error)<2e-11
        bound=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+info[3]*source['direction']))[:56])))
        norm=float(max(assembled['physical_norm'],assembled['six_force_norm'],meta[8]*max(abs(assembled['residual'][56:])),abs(control)))
        assert abs(norm-info[4])<2e-11
        strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
        verified=bool(norm<=bound+1e-12 and steel<=1 and info[3]>0)
        item.update(independently_solves_original_control=verified,total_independent_residual_MPa=norm,
            original_final_load_bound_MPa=bound,original_linear_control_residual_MPa=control,
            six_force_residual_MPa=assembled['six_force_norm'],all_material_returns_available=True,
            physical_reconstruction_difference_MPa=physical_error,gauge_reconstruction_difference_MPa=gauge_error,
            force_reconstruction_difference_MPa=force_error,control_reconstruction_difference_MPa=control_error,
            maximum_steel_strain_utilization=steel)
        if info[1]:assert verified,item
        fields.append(field)
    cases.append(item)
assert r.position==len(r.raw)
proof=dict(native=scope['native'],actual_component_count=257,all256_other_source_components_preserved=True,
    source_history25_and_materials_unchanged=True,original_linear_projection_target_and_six_force_path_unchanged=True,
    initial_native940_material_root_residual_bitexact_by_guard=True,original30_steps=True,
    no_intermediate_material_history_acceptance=True,cases=cases,
    original_equilibria_verified=sum(c['independently_solves_original_control'] for c in cases),
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
if len(fields)==2:
    proof['algorithm_agreement']=dict(lambda_difference=abs(cases[0]['lambda_value']-cases[1]['lambda_value']),
        maximum_three_station_strain_difference=float(abs(np.einsum('spcd,d->spc',op,fields[0]-fields[1])).max()))
(root/'OriginalControlReturnProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for case in cases:print('CASE',case['case'],'NATIVE',case['native_solved'],'VERIFIED ORIGINAL',case['independently_solves_original_control'],
    'LAMBDA',case['lambda_value'],'STEPS',case['steps'],'R',case.get('total_independent_residual_MPa'))
