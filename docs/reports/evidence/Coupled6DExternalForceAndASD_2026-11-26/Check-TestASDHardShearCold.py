"""ДЛЯ ТЕСТОВ. Независимо собирает все уравнения и шесть усилий холодного ASD из точных сохранённых тензоров; численный останов не становится пределом."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);root=p.parse_args().root
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
contract=json.loads((root/'Input/MaterialContract.json').read_bytes())
single=contract.get('integration_mode')=='native_single_update_test'
assert ('native_ASD_single_update_test_not_closed' if single else 'closed_ASD_integration_not_identical_OpenSees_single_update') in text
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
assert not (root/'Input/Reverse1.bin').exists(), 'Cold control must have no ready equilibrium field input'
source_audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert source_audit['components']==257
assert source_audit['imported_this_pass']==(['CCoupledASDFlow','CCoupledConcreteResponse','modTestCoupledElasticState'] if single else ['modTestCoupledElasticState'])
r=R((root/'Native/ASDCold.bin').read_bytes());assert r.raw[:16]==b'NDMASDCOLD000001'
options=r.VectorTest();dof,points,concrete=map(int,options[:3]);half,absolute,relative,budget=options[3:]
target=r.VectorTest();assert dof==60 and points==43 and concrete==25 and budget==30
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/ASDCold.Operators.csv',dof,points)
op=np.array([B+s*half*gradient for s in (-1,0,1)])
area=coords[:,2];center=np.average(coords[:concrete,:2],axis=0,weights=area[:concrete]);x,y=(coords[:,:2]-center).T
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
previous=0.;steps=[];attempts=[];accepted_fields=[];all_strains=[];all_stresses=[];long=axis[concrete:,2]==1
maximum_strain_drift=0.;maximum_stress_drift=0.
while True:
    meta=r.VectorTest()
    if meta[0]==0:
        ending=meta;break
    attempt,base,factor,solved,iterations,evaluations,revision=meta
    assert attempt==len(attempts)+1 and base==previous and factor>base and solved in (0,1)
    assert 0<=iterations<=budget and evaluations>=0
    attempts.append(dict(attempt=int(attempt),base_lambda=base,target_lambda=factor,solved=bool(solved),iterations=int(iterations),function_evaluations=int(evaluations)))
    if not solved:continue
    field=r.VectorTest();reported=r.VectorTest();strain=np.zeros((3,points,6));net=strain.copy();sigma=strain.copy();seen=set()
    for _ in range(3*points):
        point,side,code=r.VectorTest();key=(int(side)+1,int(point)-1)
        assert code==0 and key not in seen;seen.add(key)
        strain[key]=r.VectorTest();net[key]=r.VectorTest();sigma[key]=r.VectorTest()
    strain_error=float(np.max(abs(strain-np.einsum('spcd,d->spc',op,field))))
    assert strain_error<1e-14
    stress_error=float(np.max(abs(net[:,:concrete]-sigma[:,:concrete])));assert stress_error==0
    maximum_strain_drift=max(maximum_strain_drift,strain_error)
    # Все стальные точки на пройденной ветви остаются упругими: это проверяется, а не предполагается.
    eps=np.einsum('spc,pc->sp',strain[:,concrete:],axis[concrete:]);steel=eps*200000.
    fy=np.where(long[None,:],np.where(eps>=0,400.,350.),240.)
    assert np.all(abs(steel)<fy)
    steel_error=float(np.max(abs(net[:,concrete:]+sigma[:,concrete:]-steel[:,:,None]*axis[None,concrete:])))
    assert steel_error<1e-10;maximum_stress_drift=max(maximum_stress_drift,steel_error)
    external=L@(factor*target)
    residual=g@field-external+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/(2*half))
    actual=np.array([area@net[1,:,2],(area*y)@net[1,:,2],(area*x)@net[1,:,2],area@(x*net[1,:,5]-y*net[1,:,4]),area@net[1,:,4],area@net[1,:,5]])
    bound=float(absolute+relative*max(abs(external[:dof-4])))
    physical=float(max(abs(residual[:dof-4])));gauge=float(30000.*max(abs(residual[dof-4:])))
    force=float(max(abs((actual-factor*target)/scales)))
    assert max(physical,gauge,force)<=bound+1e-10,(factor,physical,gauge,force,bound)
    assert max(abs((actual-reported)/scales))<1e-10
    utilization=float(np.max(abs(eps)/np.where(long[None,:],np.where(eps>=0,.025,.018),.02)))
    states=R((root/'Native'/f'ASDCold.State{len(steps)+1:03d}.bin').read_bytes())
    assert states.raw[:16]==b'NDMHISTSTATE0001' and states.VectorTest().tolist()==[revision+1]
    np.testing.assert_array_equal(states.VectorTest(),field)
    principal=np.linalg.eigvalsh(np.array([[[e[0],e[3]/2,e[4]/2],[e[3]/2,e[1],e[5]/2],[e[4]/2,e[5]/2,e[2]]] for e in strain[:,:concrete].reshape(-1,6)]))
    steps.append(dict(lambda_value=factor,physical_residual_MPa=physical,gauge_residual_MPa=gauge,six_force_residual_MPa=force,
                      original_bound_MPa=bound,steel_utilization=utilization,max_concrete_principal_tension_strain=float(principal.max()),
                      min_concrete_principal_compression_strain=float(principal.min()),iterations=int(iterations),function_evaluations=int(evaluations)))
    accepted_fields.append((factor,field));all_strains.append(strain.reshape(-1,6));all_stresses.append(sigma.reshape(-1,6));previous=factor
assert r.position==len(r.raw) and previous==ending[1] and len(steps)==ending[3]
assert len(attempts)-len(steps)==ending[4] and ending[5] in (0,1)
if ending[5]:
    assert not attempts[-1]['solved'] and (attempts[-1]['target_lambda']-previous)/2<.0003
else:
    assert previous==1.2, 'An exhausted attempt budget is not reaching the requested load'
initial=R((root/'Native/ASDCold.InitialState.bin').read_bytes());initial.VectorTest();assert np.all(initial.VectorTest()==0)
material_comparison=None
if (root/'Native/OpenSeesNativeMaterialReplay.json').exists():
    replay=json.loads((root/'Native/OpenSeesNativeMaterialReplay.json').read_bytes())
    assert replay['native_binary_sha256']==sha256(r.raw).hexdigest().upper() and not replay['section_equilibrium_solved_by_OpenSees']
    np.testing.assert_array_equal(np.array(replay['strain_histories']),np.array(all_strains))
    np.testing.assert_array_equal(np.array(replay['expected_VBA_concrete_stresses']),np.array(all_stresses))
    error=float(np.max(abs(np.array(all_stresses)-np.array(replay['observed_OpenSees_concrete_stresses']))))
    assert error==replay['maximum_absolute_stress_error_MPa'] and replay['material_tensor_match']==(error<replay['comparison_tolerance_MPa'])
    material_comparison=dict(maximum_stress_error_MPa=error,comparison_tolerance_MPa=replay['comparison_tolerance_MPa'],
                             material_tensor_match=replay['material_tensor_match'],scalar_comparisons=replay['scalar_comparisons'])
proof=dict(native=native,accepted_steps=len(steps),rejected_trials=len(attempts)-len(steps),maximum_accepted_lambda=previous,
    kernel_seconds=ending[2],all_six_resultants_and56plus4_equations_verified=True,all_steel_remained_elastic=True,
    maximum_operator_strain_error=maximum_strain_drift,maximum_independent_steel_stress_error_MPa=maximum_stress_drift,
    no_ready_equilibrium_input_and_initial_field_zero=True,source_unchanged=True,
    integration_mode='native_single_update_test' if single else 'closed_fixed_strain',
    OpenSees_material_tensor_equivalence_verified=bool(material_comparison and material_comparison['material_tensor_match']),
    material_comparison=material_comparison,
    requested_load_reached=bool(previous==1.2),numerical_stop=bool(ending[5]),
    numerical_stop_not_physical_capacity=True,physical_capacity_certified=False,whole_goal_completed=False,steps=steps,attempts=attempts)
(root/'ASDHardShearColdProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('ASD COLD INDEPENDENTLY VERIFIED',json.dumps({k:v for k,v in proof.items() if k not in ('steps','attempts')},sort_keys=True))
