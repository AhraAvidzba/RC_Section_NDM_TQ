"""ДЛЯ ТЕСТОВ. Независимо проверяет короткие ступени, половинные критерии, бюджеты, фактическое принятие и восстановление25."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold958BoundedIncomingSteps_20261126';parent=base/'Cold957FourPartEventHistory_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
baseline_validation,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(fields)==25
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestAssemble(');end=helper.index('summary=[];',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
begin=helper.index('def TestYield(');end=helper.index('def TestAssemble(',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']

def TestReadStress(reader):
    """Читает все129 фактических пробных тензоров и отклоняет недоступный локальный возврат."""
    stress=np.zeros((3,43,6));seen=set()
    for _ in range(129):
        point,side,error=reader.VectorTest();key=int(point),int(side)
        assert error==0 and key not in seen;seen.add(key)
        stress[int(side)+1,int(point)-1]=reader.VectorTest()
    return stress

def TestReadHistory(reader):
    """Читает принятые состояния бетона, доли энергии, напряжения и осевое напряжение стали."""
    saved={};stress=np.zeros((3,43,6));metrics=np.zeros((3,43,3))
    for _ in range(129):
        point,side=reader.VectorTest();key=int(point),int(side)
        assert key not in saved
        state,metric,sigma=[reader.VectorTest() for _ in range(3)]
        assert len(state)==10 and len(metric)==3 and len(sigma)==6
        saved[key]=state;stress[int(side)+1,int(point)-1]=sigma;metrics[int(side)+1,int(point)-1]=metric
    return dict(states=saved,stress=stress,metrics=metrics)

def TestNorm(delta):
    """Собирает исходную тензорную метрику по трём станциям и всем43 точкам."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))

def TestEquilibrium(field,lam,stress,projection=None,target=None):
    """Проверяет исходные56 физических уравнений,4 связи,6 усилий и пределы двух семейств стали."""
    a=TestAssemble(field,lam,stress);bound=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+lam*source['direction']))[:56])))
    norm=float(max(a['physical_norm'],a['six_force_norm'],E*max(abs(a['residual'][56:]))))
    if projection is not None:norm=max(norm,abs(float(E*(projection@field-target))))
    strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
    assert norm<=bound+2e-11 and steel<=1,(norm,bound,steel)
    return dict(lambda_value=float(lam),residual_MPa=norm,bound_MPa=bound,six_force_residual_MPa=a['six_force_norm'],steel_utilization=steel)

r=R((root/'Native/BoundedIncoming.bin').read_bytes());assert r.raw[:16]==b'NDMINCOMING00001'
context=r.VectorTest();reference=float(r.VectorTest()[0]);original=r.VectorTest();initial=TestReadHistory(r)
np.testing.assert_array_equal(original,source['fields'][-1]);assert reference==source['lambdas'][-1]
assert context[:3].tolist()==[source['options'][3],source['options'][4],source['options'][2]]
previous_field=original;previous_history=initial;previous_lambda=reference;previous_revision=None
accepted_count=0;cycles=[];acceptances=[];all_pairs=0
tag=r.VectorTest()
while tag[0]==1:
    assert len(tag)==10 and int(tag[1])==len(cycles)+1 and tag[2]==previous_lambda and tag[9]==accepted_count
    cycle_number=int(tag[1]);accepted=r.VectorTest();event=r.VectorTest();history=TestReadHistory(r)
    np.testing.assert_array_equal(accepted,previous_field)
    if previous_revision is not None:assert tag[3]==previous_revision
    TestEquilibrium(accepted,tag[2],history['stress'])
    if acceptances:
        np.testing.assert_allclose(history['stress'],acceptances[-1]['stress'],rtol=0,atol=2e-12)
    cycle=dict(cycle=cycle_number,reference_lambda=float(tag[2]),guide_solved=bool(tag[4]),guide_steps=int(tag[6]),attempts=[])
    if tag[4]:
        guide=r.VectorTest();guide_stress=TestReadStress(r)
        cycle['guide']=TestEquilibrium(guide,tag[5],guide_stress)
        assert len(event)==7 and event[0]==1
        states=history['states'];value=TestYield(int(event[1]),int(event[2]),guide);assert abs(value)<2e-10
        cycle['guide'].update(trial_yield_MPa=value,point=int(event[1]),side=int(event[2]))
        movement=guide-accepted;length=TestNorm(movement)
        strain=np.einsum('spcd,d->spc',op,movement)
        projection=np.einsum('spcd,spc,p,c->d',op,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2])*length)
        coordinate=float(projection@accepted);prior_steps=int(tag[6]);attempt_number=0
        tag=r.VectorTest()
        while tag[0]==2:
            assert len(tag)==8 and tag[1]==cycle_number and tag[2]==attempt_number+1
            attempt_number+=1;fraction=float(tag[3]);target=float(tag[4]);solved=bool(tag[5]);count=int(tag[7]);effort=r.VectorTest()
            assert abs(target-(coordinate+fraction*length))<2e-12*max(abs(target),1e-8)
            assert fraction==.5**(attempt_number-1) and prior_steps<=effort[0]<=context[2]
            item=dict(attempt=attempt_number,fraction=fraction,returned=solved,total_steps=int(effort[0]),new_steps=int(effort[0]-prior_steps),pairs=[])
            prior_steps=int(effort[0])
            for _ in range(count):
                meta,base_field,full,half,p=[r.VectorTest() for _ in range(5)]
                full_stress=TestReadStress(r);half_stress=TestReadStress(r);all_pairs+=1
                np.testing.assert_array_equal(base_field,accepted);assert max(abs(p-projection))<2e-10
                full_result=TestEquilibrium(full,meta[0],full_stress,p,target)
                half_target=.5*(float(p@accepted)+target)
                half_result=TestEquilibrium(half,meta[1],half_stress,p,half_target)
                distance=TestNorm(full-accepted);field_error=TestNorm(full-(2*half-accepted))/distance
                lambda_error=float(abs(2*meta[1]-meta[2]-meta[0]));alpha=meta[10]
                radius=float(np.sqrt(distance**2+alpha**2*(meta[0]-meta[2])**2))
                allowance=(.005*radius)**2-(field_error*distance)**2;noise=float(8*context[1]*(1+abs(meta[2])))
                bound=max(noise,float(np.sqrt(allowance)/alpha) if allowance>0 else noise)
                joint=float(np.sqrt((field_error*distance)**2+(alpha*lambda_error)**2)/radius)
                assert abs(distance-meta[9])<2e-10*distance and abs(field_error-meta[6])<2e-10
                assert abs(lambda_error-meta[7])<1e-12 and abs(bound-meta[8])<1e-12 and abs(joint-meta[11])<2e-10
                shape_passed=bool(field_error<=.005 and lambda_error<=bound)
                item['pairs'].append(dict(full=full_result,half=half_result,field_error=field_error,lambda_error=lambda_error,lambda_bound=bound,shape_passed=shape_passed))
                if solved:
                    assert shape_passed
                    success=dict(field=full,stress=full_stress,projection=p,lambda_value=float(meta[0]),fraction=fraction)
            assert not solved or count>0
            cycle['attempts'].append(item);tag=r.VectorTest()
        if tag[0]==3:
            assert len(tag)==7 and tag[1]==cycle_number and cycle['attempts'][-1]['returned'] and tag[4]==accepted_count+1
            field=r.VectorTest();p=r.VectorTest();np.testing.assert_array_equal(field,success['field']);np.testing.assert_array_equal(p,success['projection'])
            assert tag[2]==success['lambda_value'] and tag[3]>0 and tag[6]>=0
            noise=float(8*context[1]*(1+abs(previous_lambda)));assert tag[2]>previous_lambda+noise+tag[6]
            accepted_count+=1;previous_field=field;previous_lambda=float(tag[2]);previous_revision=float(tag[3])
            acceptance=dict(cycle=cycle_number,lambda_value=previous_lambda,fraction=success['fraction'],native_lambda_uncertainty=float(tag[6]),
                lambda_increment_over_previous=float(tag[2]-cycle['reference_lambda']),source_force_and_half_and_steel_checks_passed=True)
            cycle['accepted']=acceptance;acceptances.append(success);cycle['reached_guide']=bool(tag[5]);tag=r.VectorTest()
    else:tag=r.VectorTest()
    cycles.append(cycle)
assert tag[0]==4 and len(tag)==5 and tag[1]==accepted_count and tag[2]==previous_lambda
final_field=r.VectorTest();final_history=TestReadHistory(r);np.testing.assert_array_equal(final_field,previous_field)
final_equilibrium=TestEquilibrium(final_field,previous_lambda,final_history['stress'])
if acceptances:np.testing.assert_allclose(final_history['stress'],acceptances[-1]['stress'],rtol=0,atol=2e-12)
final_reached=bool(tag[4]);restored_tag=r.VectorTest();assert restored_tag[0]==5
restored_field=r.VectorTest();restored=TestReadHistory(r);np.testing.assert_array_equal(restored_field,original)
for key in initial['states']:np.testing.assert_array_equal(restored['states'][key],initial['states'][key])
np.testing.assert_array_equal(restored['stress'],initial['stress']);np.testing.assert_array_equal(restored['metrics'],initial['metrics'])
assert r.position==len(r.raw)
real_keys=[(point,side) for point in range(1,26) for side in (-1,0,1)]
history_progress=dict(
    maximum_real_concrete_plastic_strain_change=float(max(np.max(abs(final_history['states'][key][2:8]-initial['states'][key][2:8])) for key in real_keys)),
    maximum_real_concrete_envelope_energy_fraction_change=float(np.max(abs(final_history['metrics'][:,:25,:2]-initial['metrics'][:,:25,:2]))),
    net_field_RMS_advance=TestNorm(final_field-original))
event_keys=[(c['guide']['point'],c['guide']['side']) for c in cycles if 'guide' in c]
if event_keys and len(set(event_keys))==1:
    key=event_keys[0]
    if np.array_equal(initial['states'][key],final_history['states'][key]):
        states=initial['states']
        history_progress.update(event_point=key[0],event_side=key[1],event_plastic_state_unchanged=True,
            initial_trial_yield_MPa=TestYield(*key,original),final_trial_yield_MPa=TestYield(*key,final_field))
proof=dict(native=native,actual_components=257,all256_other_components_preserved=True,all_numerical_sources_exact957=True,
    original25_field_concrete_state_energy_stress_and_steel_stress_restored_bitexact=True,cycles=cycles,
    total_pairs_independently_verified=all_pairs,accepted_steps=accepted_count,reached_current_material_event=final_reached,
    final_equilibrium=final_equilibrium,history_progress=history_progress,original_force_steel_shape_and_uncertainty_guards_preserved=True,
    guide_and_rejected_attempts_within_original30_per_accepted_step=True,first_known_event952_cost_counted=True,
    earlier_primary_and_predictor_cost_not_included=True,diagnostic_material_history_only=True,no_production_history_acceptance=True,
    physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(root/'BoundedIncomingStepsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BOUNDED INCOMING:',accepted_count,'ACCEPTED;',all_pairs,'PAIRS VERIFIED;REACHED EVENT',final_reached,'RESTORED25 BITEXACT')
print('ACTUAL HISTORY PROGRESS',json.dumps(history_progress,sort_keys=True))
for cycle in cycles:
    print('CYCLE',cycle['cycle'],'REFERENCE',cycle['reference_lambda'],'GUIDE',cycle.get('guide',{}).get('lambda_value'))
    for a in cycle['attempts']:print(' ATTEMPT',a['fraction'],'RETURNED',a['returned'],'STEPS',a['total_steps'],'SHAPE',[p['field_error'] for p in a['pairs']])
    if 'accepted' in cycle:print(' ACCEPTED',cycle['accepted'])
