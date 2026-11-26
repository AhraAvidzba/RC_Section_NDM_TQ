"""ДЛЯ ТЕСТОВ. Независимо проверяет равновесия подхода953 и неизменённый совместный критерий половинного шага."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;script=here/'Check-TestDetectedLoadingBoundary.py';code=script.read_text('utf-8-sig')
cut="r=R((root/'Native/DetectedLoadingBoundary.bin').read_bytes())";assert code.count(cut)==1;code=code[:code.index(cut)]
assert code.count('Cold952DetectedLoadingBoundary_20261126')==2
code=code.replace('Cold952DetectedLoadingBoundary_20261126','Cold953DetectedEventApproach_20261126')
env={'__file__':str(script),'__name__':'detected_event_approach_test'};exec(compile(code,str(script),'exec'),env)
inner=env['scope'];root=env['root'];R=inner['R'];op=inner['op'];coords=inner['coords'];source=inner['source'];L=inner['L'];E=inner['meta'][8]
def TestNorm(delta):
    """Вычисляет независимую среднюю тензорную норму по фактическим трём станциям и всем точкам."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))
r=R((root/'Native/ResolvedEventApproach.bin').read_bytes());assert r.raw[:16]==b'NDMEVAPPROACH001'
info,effort,accepted,face,projection=[r.VectorTest() for _ in range(5)];assert len(info)==15 and len(effort)==8
assert info[14]==2 and effort[0]+info[14]<=source['options'][2]
np.testing.assert_array_equal(accepted,source['fields'][-1]);half=r.VectorTest() if info[0] else None
states=[]
records=[('event',face,info[3],info[9])]+([('half_approach',half,info[4],info[10])] if info[0] else [])
for name,field,lam,target in records:
    stress=np.zeros((3,43,6));seen=set()
    for _ in range(129):
        point,side,error=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
        assert error==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
    a=inner['TestAssemble'](field,lam,stress);control=float(E*(projection@field-target))
    goal=source['offset']+lam*source['direction'];bound=float(source['options'][3]+source['options'][4]*max(abs((L@goal)[:56])))
    norm=float(max(a['physical_norm'],a['six_force_norm'],E*max(abs(a['residual'][56:])),abs(control)))
    assert norm<=bound+2e-11,(name,norm,bound)
    strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],inner['axis'][25:]);long=inner['axis'][25:,2]==1
    steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)));assert steel<=1
    states.append(dict(name=name,lambda_value=float(lam),independent_total_residual_MPa=norm,final_bound_MPa=bound,
        control_residual_MPa=control,steel_strain_utilization=steel,point18_mid_trial_yield_MPa=inner['TestYield'](18,0,field)))
assert r.position==len(r.raw)
length=TestNorm(face-accepted);assert abs(length-info[8])<2e-10*length
strain=np.einsum('spcd,d->spc',op,face-accepted)
expected_projection=np.einsum('spcd,spc,p,c->d',op,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2])*length)
assert max(abs(expected_projection-projection))<2e-10
verified=False;field_error=None;lambda_error=None;bound=None
if info[0]:
    field_error=TestNorm(face-(2*half-accepted))/length;lambda_error=float(abs(2*info[4]-info[2]-info[3]))
    alpha=info[11];radius=float(np.sqrt(length**2+alpha**2*(info[3]-info[2])**2))
    allowance=(.005*radius)**2-(field_error*length)**2;noise=8*source['options'][4]*(1+abs(info[2]))
    bound=max(noise,float(np.sqrt(allowance)/alpha) if allowance>0 else noise)
    joint=float(np.sqrt((field_error*length)**2+(alpha*lambda_error)**2)/radius)
    assert abs(field_error-info[5])<2e-10 and abs(lambda_error-info[6])<1e-12 and abs(bound-info[7])<1e-12 and abs(joint-info[12])<2e-10
    verified=bool(field_error<=.005 and lambda_error<=bound and info[4]>0)
assert bool(info[1])==verified
proof=dict(native=inner['native'],actual_components=257,all256_other_components_preserved=True,
    numerical_classes_exact952=True,event_equilibrium_from_generic_detector=True,
    native_half_solved=bool(info[0]),native_half_verified=bool(info[1]),independent_half_verified=verified,
    field_error=field_error,lambda_error=lambda_error,lambda_bound=bound,states=states,
    event_steps=2,half_steps=int(effort[0]),combined_steps=int(effort[0]+2),total_budget=int(source['options'][2]),
    original_force_tolerances_and_half_policy_preserved=True,no_material_history_acceptance=True,
    physical_capacity_certified=False,cold_gain_claimed=False,full_original_target_not_claimed_solved=True)
(root/'DetectedEventApproachProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('EVENT APPROACH HALF SOLVED',proof['native_half_solved'],'VERIFIED',verified,'FIELD ERROR',field_error,'LAMBDA ERROR',lambda_error,'BOUND',bound,'TOTAL',proof['combined_steps'])
for state in states:print(state['name'],'LAMBDA',state['lambda_value'],'R',state['independent_total_residual_MPa'],'F18',state['point18_mid_trial_yield_MPa'])
