"""ДЛЯ ТЕСТОВ. Независимо проверяет силы, последовательную историю, одинаковую конечную поверхность и точное восстановление25."""
from pathlib import Path
from runpy import run_path
import json, numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold955SequentialEventHistory_20261126';parent=base/'Cold954RefinedEventApproach_20261126'
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
r=R((root/'Native/SequentialEvent.bin').read_bytes());assert r.raw[:16]==b'NDMSEQEVENT00001'
context,projection,geometry=[r.VectorTest() for _ in range(3)];assert len(context)==7 and len(projection)==60 and len(geometry)==2
records=[]
for stage in range(6):
    info,effort,event=[r.VectorTest() for _ in range(3)];assert len(info)==6 and len(effort)==8 and info[0]==stage and info[4]<=context[4]
    trial=None
    if info[5]:
        field=r.VectorTest();lam=float(r.VectorTest()[0]);stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,error=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
            assert error==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        trial=dict(field=field,lambda_value=lam,stress=stress)
    accepted=r.VectorTest();committed={};committed_stress=np.zeros((3,43,6));metrics=np.zeros((3,43,3))
    for _ in range(129):
        point,side=r.VectorTest();key=int(point),int(side);state,metric,sigma=[r.VectorTest() for _ in range(3)]
        assert key not in committed and len(state)==10 and len(metric)==3 and len(sigma)==6
        committed[key]=state;committed_stress[int(side)+1,int(point)-1]=sigma;metrics[int(side)+1,int(point)-1]=metric
    records.append(dict(info=info,effort=effort,event=event,trial=trial,field=accepted,states=committed,stress=committed_stress,metrics=metrics))
assert r.position==len(r.raw)
for stage in (0,3,5):
    np.testing.assert_array_equal(records[stage]['field'],source['fields'][-1])
for stage in (3,5):
    for key in records[0]['states']:np.testing.assert_array_equal(records[stage]['states'][key],records[0]['states'][key])
    np.testing.assert_array_equal(records[stage]['stress'],records[0]['stress'])
    np.testing.assert_array_equal(records[stage]['metrics'],records[0]['metrics'])

def TestNorm(delta):
    """Применяет независимую тензорную меру трёх станций, без геометрического удвоения половинного приращения."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))

def TestEquilibrium(field,lam,stress,control=None):
    """Собирает все56 физических уравнений,4 связи,6 усилий и пределы двух семейств стали из фактических тензоров."""
    a=TestAssemble(field,lam,stress);bound=float(context[2]+context[3]*max(abs((L@(source['offset']+lam*source['direction']))[:56])))
    norm=float(max(a['physical_norm'],a['six_force_norm'],E*max(abs(a['residual'][56:]))))
    if control is not None:norm=max(norm,abs(float(E*(projection@field-control))))
    strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
    assert norm<=bound+2e-11 and steel<=1,(norm,bound,steel)
    return dict(lambda_value=lam,residual_MPa=norm,bound_MPa=bound,six_force_residual_MPa=a['six_force_norm'],maximum_steel_strain_utilization=steel)

results={};states=records[0]['states']
full=records[0]['trial'];assert full is not None
results['one_full']=TestEquilibrium(full['field'],full['lambda_value'],full['stress'])
selected=json.loads((root/'SourceDetectedEventProof.json').read_bytes())['selected_event'];key=int(selected[1]),int(selected[2])
results['one_full']['trial_yield_MPa']=TestYield(*key,full['field']);assert abs(results['one_full']['trial_yield_MPa'])<2e-10
for stage in (0,3,5):
    results['restored_'+str(stage)]=TestEquilibrium(records[stage]['field'],float(context[0]),records[stage]['stress'])
np.testing.assert_array_equal(records[4]['field'],full['field'])
np.testing.assert_allclose(records[4]['stress'],full['stress'],rtol=0,atol=2e-12)
results['full_committed']=TestEquilibrium(records[4]['field'],full['lambda_value'],records[4]['stress'])
half=records[1]['trial'];two=records[2]['trial'];comparison=None
if half is not None:
    results['half']=TestEquilibrium(half['field'],half['lambda_value'],half['stress'],geometry[0]-.5*geometry[1])
    np.testing.assert_array_equal(records[1]['field'],half['field'])
    np.testing.assert_allclose(records[1]['stress'],half['stress'],rtol=0,atol=2e-12)
    results['half_committed']=TestEquilibrium(records[1]['field'],half['lambda_value'],records[1]['stress'])
    strain=np.einsum('spcd,d->spc',op,half['field']);bars=200000*np.einsum('spc,pc->sp',strain[:,25:],axis[25:])
    np.testing.assert_allclose(records[1]['metrics'][:,25:,2],bars,rtol=2e-13,atol=2e-11)
if two is not None:
    assert half is not None
    event=records[2]['event'];eventkey=int(event[1]),int(event[2]);assert len(event)==7
    states=records[1]['states']
    results['two_substeps']=TestEquilibrium(two['field'],two['lambda_value'],two['stress'])
    value=TestYield(*eventkey,two['field']);assert abs(value)<2e-10
    results['two_substeps']['trial_yield_MPa']=value
    np.testing.assert_array_equal(records[2]['field'],two['field'])
    np.testing.assert_allclose(records[2]['stress'],two['stress'],rtol=0,atol=2e-12)
    results['two_committed']=TestEquilibrium(records[2]['field'],two['lambda_value'],records[2]['stress'])
    same_surface=eventkey==key and np.array_equal(records[0]['states'][key],records[1]['states'][key])
    field_error=TestNorm(two['field']-full['field'])/TestNorm(full['field']-source['fields'][-1])
    lambda_error=abs(two['lambda_value']-full['lambda_value']);noise=float(8*context[3]*(1+abs(context[0])))
    comparison=dict(same_terminal_yield_surface=same_surface,selected_event=event.tolist(),
        endpoint_field_error_relative_to_full_increment=field_error,endpoint_lambda_difference=lambda_error,
        absolute_tensor_strain_RMS_difference=TestNorm(two['field']-full['field']),
        maximum_real_concrete_tensor_strain_difference=float(np.max(abs(np.einsum('spcd,d->spc',op[:,:25],two['field']-full['field'])))),
        maximum_steel_axial_strain_difference=float(np.max(abs(np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op[:,25:],two['field']-full['field']),axis[25:])))),
        existing_lambda_noise=noise,endpoint_difference_below_existing_half_field_and_lambda_thresholds=bool(same_surface and field_error<=.005 and lambda_error<=noise),
        maximum_concrete_stress_difference_MPa=float(np.max(abs(two['stress']-full['stress']))),
        maximum_concrete_plastic_state_difference=float(max(np.max(abs(records[2]['states'][k][2:8]-records[4]['states'][k][2:8])) for k in records[0]['states'])),
        maximum_concrete_energy_fraction_difference=float(np.max(abs(records[2]['metrics'][:,:,:2]-records[4]['metrics'][:,:,:2]))),
        half_steps=int(records[1]['effort'][0]),second_event_steps=int(records[2]['effort'][0]),
        combined_steps_including_source_full_event=int(records[2]['info'][4]),earlier_predictor_and_failed_primary_effort_excluded=True)
proof=dict(native=native,actual_components=257,all256_other_components_preserved=True,numerical_sources_exact954=True,
    source_history25_validated=True,original25_field_concrete_state_energy_stress_and_steel_stress_restored_bitexact_twice=True,
    original_history_collections_not_extended=True,accepted_revision_changes_on_replay_are_intentional=True,
    half_solved=half is not None,second_event_solved=two is not None,results=results,comparison=comparison,
    diagnostic_temporary_material_history_only=True,geometric_half_guard_unchanged_and_not_claimed_passed=True,
    no_production_history_acceptance=True,physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(root/'SequentialEventHistoryProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SEQUENTIAL HISTORY: HALF',proof['half_solved'],'SECOND EVENT',proof['second_event_solved'],'RESTORED25 BITEXACT TWICE')
for name,result in results.items():print(name,'LAMBDA',result['lambda_value'],'R',result['residual_MPa'])
print('ENDPOINT COMPARISON',json.dumps(comparison,sort_keys=True))
