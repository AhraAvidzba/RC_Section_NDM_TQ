"""ДЛЯ ТЕСТОВ. Независимо проверяет шесть усилий, все уравнения и материальную координату каждого результата936."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold936RecordedYieldCoordinate_20261126';parent=base/'Cold935Point11Matrices_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,initial_residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==25
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestAssemble(');end=helper.index('summary=[];',begin)
body=helper[begin:end];old='return dict(residual=residual,concrete=concrete'
assert old in body;body=body.replace(old,'return dict(six_force_residual_vector=(actual-target)/scales,residual=residual,concrete=concrete',1)
exec(compile(body,str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
r=R((base/'Cold930FineLinePlusSurface_20261126/Native/ArcScan.States.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
begin=helper.index('def TestYield(');end=helper.index('def TestAssemble(',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
r=R((root/'Native/RecordedYieldCoordinate.bin').read_bytes());assert r.raw[:16]==b'NDMPOINTYIELD001'
assert r.VectorTest().tolist()==[4.]
cases=[]
for index in range(1,5):
    info=r.VectorTest();assert len(info)==15 and info[0]==index and info[10]<=meta[7]
    assert info[1]==(meta[2] if index<=2 else meta[9]) and info[2]==(index%2)
    item=dict(case=index,target_trial_yield_MPa=float(info[1]),complete_Newton_line_search=bool(info[2]),
        native_solved=bool(info[3]),mapped_ready=bool(info[4]),material_ready=bool(info[5]),lambda_value=float(info[6]),
        native_coordinate_residual_MPa=float(info[7]),native_start_tolerance_MPa=float(info[8]),
        native_final_total_norm_MPa=float(info[9]),steps=int(info[10]),calls=int(info[11]),jacobians=int(info[12]),
        native_best_residual_MPa=float(info[13]),stop_code=int(info[14]),independently_equilibrated=False)
    if info[5]:
        field,values,force_residual=r.VectorTest(),r.VectorTest(),r.VectorTest();stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,code=r.VectorTest();sigma=r.VectorTest();assert code==0 and len(sigma)==6
            key=int(point),int(side);assert key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        assembled=TestAssemble(field,info[6],stress)
        native_physical_error=float(max(abs(values[:56]-assembled['residual'][:56])))
        force_error=float(max(abs(force_residual-assembled['six_force_residual_vector'])))
        gauge_error=float(max(abs(values[56:60]-meta[8]*assembled['residual'][56:])))
        assert native_physical_error<2e-11 and force_error<2e-11 and gauge_error<2e-11
        bound=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+info[6]*source['direction']))[:56])))
        gauge=float(meta[8]*max(abs(assembled['residual'][56:])))
        norm=float(max(assembled['physical_norm'],assembled['six_force_norm'],gauge))
        assert abs(norm-info[9])<2e-11
        coordinate=TestYield(int(meta[0]),int(meta[1]),field)-info[1]
        coordinate_difference=float(abs(coordinate-info[7]));assert coordinate_difference<2e-10
        strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        steel_util=float(max((abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)).ravel()))
        original_linear=float(meta[8]*(projection@field-meta[4]))
        assert abs(values[-1]-original_linear)<2e-11
        progress=float(meta[8]*(projection@(field-accepted)))
        requested=float(meta[8]*(meta[4]-projection@accepted))
        strain_increment=np.einsum('spcd,d->spc',op,field-accepted)
        max_real_increment=float(np.max(abs(strain_increment[:,:25])))
        verified=bool(norm<=bound+1e-12 and abs(info[7])<=2e-12 and abs(coordinate)<=2e-12+coordinate_difference and steel_util<=1 and info[6]>0)
        item.update(independently_equilibrated=verified,all_material_returns_available=True,
            total_residual_MPa=norm,original_final_load_bound_MPa=bound,six_force_norm_MPa=assembled['six_force_norm'],
            gauge_scaled_norm_MPa=gauge,independent_trial_yield_residual_MPa=float(coordinate),
            trial_yield_reconstruction_difference_MPa=coordinate_difference,
            native_physical_reconstruction_difference_MPa=native_physical_error,
            native_gauge_reconstruction_difference_MPa=gauge_error,
            six_force_reconstruction_difference_MPa=force_error,
            maximum_three_station_steel_utilization=steel_util,original_linear_control_residual_MPa=original_linear,
            original_projection_progress_from_accepted_MPa=progress,
            original_projection_requested_progress_MPa=requested,
            original_projection_fraction_of_requested_progress=float(progress/requested),
            maximum_real_concrete_strain_increment_from_accepted=max_real_increment,
            projection_progress_exceeds_16_source_force_bounds=bool(progress>16*max(bound,info[8])),
            solves_original_linear_control=bool(verified and abs(original_linear)<=bound+1e-12))
    cases.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=257,all256_other_source_components_preserved=True,
    source_actual934_state25_and_initial_residual_bitexact_by_native_guard=True,
    material_coordinate_intentionally_replaces_original_linear_control=True,
    same_six_component_load_direction_and_all_materials=True,original30_steps_per_independent_case=True,
    accepted_history_unchanged_by_native_guard=True,cases=cases,
    independently_verified_material_coordinate_equilibria=sum(x['independently_equilibrated'] for x in cases),
    original_linear_control_equilibria=sum(x.get('solves_original_linear_control',False) for x in cases),
    no_material_acceptance_or_capacity_claim=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'RecordedYieldCoordinateProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in cases:print('CASE',item['case'],'TARGET',item['target_trial_yield_MPa'],'NATIVE',item['native_solved'],
    'VERIFIED',item['independently_equilibrated'],'LAMBDA',item['lambda_value'],'STEPS',item['steps'],
    'R',item.get('total_residual_MPa',item['native_final_total_norm_MPa']),'OLD CONTROL',item.get('original_linear_control_residual_MPa'))
