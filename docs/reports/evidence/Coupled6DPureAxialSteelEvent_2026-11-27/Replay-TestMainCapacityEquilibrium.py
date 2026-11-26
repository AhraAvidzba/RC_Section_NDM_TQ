"""ДЛЯ ТЕСТОВ. Независимо пересчитывает принятый путь с настоящим ASDConcrete3D OpenSees, сталью и всеми уравнениями секционного равновесия."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
from time import perf_counter
import json,numpy as np
import openseespy.opensees as ops
import openseespywin.opensees as native

p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);root=p.parse_args().root;here=Path(__file__).resolve().parent
source=json.loads((root/'EarlySourceProof.json').read_bytes())
assert source.get('three_changed_components_exact') or source.get('changed_components_exact')
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==source['saved_book_sha256']
contract_raw=(root/'Input/MaterialContract.json').read_bytes();contract=json.loads(contract_raw)
assert contract['integration_mode']=='native_single_update_working' and contract['concrete_E_MPa']==30000 and contract['poisson']==.2
assert ops.version()=='3.8.0' and ops.NDTest is native.NDTest
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];raw=(root/'Native/MainCapacityHistory.bin').read_bytes();r=R(raw)
assert raw[:16] in (b'NDMMAINHIST00001',b'NDMMAINHIST00002',b'NDMMAINHIST00003')
header=r.VectorTest()
spatial_wire=(raw[:16]==b'NDMMAINHIST00003')
if spatial_wire:
    assert len(header)==8;component=0;base_Mx=0.
    offset=r.VectorTest();direction=r.VectorTest();assert len(offset)==len(direction)==6
    import struct
    packet=struct.unpack('<i6di',(root/'Input/CapacityLoad.bin').read_bytes());assert packet[0]==6
    prescribed=np.array(packet[1:7]);path_code=packet[-1]
    if path_code==7 or (path_code==0 and np.any(prescribed[3:])):
        expected_offset=np.zeros(6);expected_direction=prescribed
    else:
        assert path_code in range(1,7),'Auto NM requires its resolved path in the verifier.'
        expected_offset=prescribed.copy();expected_offset[path_code-1]=0
        expected_direction=np.zeros(6);expected_direction[path_code-1]=prescribed[path_code-1]
    np.testing.assert_array_equal(offset,expected_offset);np.testing.assert_array_equal(direction,expected_direction)
    dof,points,concrete,area_scale,length,half,absolute,relative=header
elif raw[:16]==b'NDMMAINHIST00001':
    assert len(header)==9;component=2
else:
    assert len(header)==10;component=int(header[-1]);assert header[-1]==component and component in (2,4)
    import struct
    component_count,prescribed_component=struct.unpack('<ii',(root/'Input/CapacityComponent.bin').read_bytes())
    assert component_count==1 and component==prescribed_component
if not spatial_wire:
    dof,points,concrete,area_scale,length,half,absolute,relative,base_Mx=header[:9]
    offset=np.zeros(6);direction=np.zeros(6);direction[component-1]=base_Mx
dof,points,concrete=map(int,(dof,points,concrete))
assert (dof,points,concrete,half)==(60,43,25,.001)
if (root/'Input/NumericalTolerances.json').exists():
    tolerance_study=json.loads((root/'Input/NumericalTolerances.json').read_bytes())
    assert tolerance_study['study_only'] and tolerance_study['production_tolerance_unchanged']
    assert (absolute,relative)==(tolerance_study['absolute_MPa'],tolerance_study['relative'])
else:assert (absolute,relative)==(.0000005,.0000005)
if spatial_wire:pass
elif (root/'Input/CapacityBase.bin').exists():
    import struct
    parameter_count,prescribed_base=struct.unpack('<id',(root/'Input/CapacityBase.bin').read_bytes())
    assert parameter_count==1 and base_Mx==prescribed_base and base_Mx>0
else:assert base_Mx==1000000.
count,=r.VectorTest();assert int(count)==count and count>1
history=[];factors=[];fractions=[]
for _ in range(int(count)):
    fraction,factor=r.VectorTest();assert 0<=fraction<=1 and factor>=0
    if not spatial_wire:assert fraction==1
    if fraction<1:assert factor==0
    field=r.VectorTest();assert len(field)==dof and np.all(np.isfinite(field));history.append(field);factors.append(factor);fractions.append(fraction)
assert r.position==len(raw)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/MainCapacity.Operators.csv',dof,points)
assert abs(area_scale*L[0,0]-1)<1e-12 and abs(length-L[0,0]/L[1,1])<1e-12
op=np.array([B+s*half*gradient for s in (-1,0,1)])
weights=coords[:,2];center=np.average(coords[:concrete,:2],axis=0,weights=weights[:concrete]);x,y=(coords[:,:2]-center).T
long=axis[concrete:,2]==1;assert sum(long)==2
command=['ASDConcrete3D',30000.,.2]
for name,rows in contract['uniaxial_tables'].items():
    for j,field in enumerate(('e','s','d')):command.extend(['-'+name+field,*(row[j] for row in rows)])
command+=['-tangent'];ops.wipe();ops.model('basic','-ndm',3,'-ndf',3)
for tag in range(1,3*points+1):ops.nDMaterial(command[0],tag,*command[1:])
plastic=np.zeros((3,points-concrete));back=plastic.copy();E=200000.;H=1500/(1-1500/E);records=[];observed=[]
force_scales=np.array([area_scale,area_scale*length,area_scale*length,area_scale*length,area_scale,area_scale])
started=perf_counter()
try:
    for index,(field,factor,fraction) in enumerate(zip(history,factors,fractions)):
        strain=np.einsum('spcd,d->spc',op,field);sigma=np.zeros_like(strain)
        for side in range(3):
            for point in range(points):
                tag=side*points+point+1;xx,yy,zz,xy,xz,yz=strain[side,point]
                ops.NDTest('SetStrain',tag,xx,yy,zz,xy,yz,xz)
                xx,yy,zz,xy,yz,xz=ops.NDTest('GetStress',tag);sigma[side,point]=[xx,yy,zz,xy,xz,yz]
                ops.NDTest('CommitState',tag)
        eps=np.einsum('spc,pc->sp',strain[:,concrete:],axis[concrete:]);trial=E*(eps-plastic);rel=trial-back
        fy_t=np.where(long[None,:],400.,240.);fy_c=np.where(long[None,:],350.,240.)
        delta=np.where(rel>fy_t,(rel-fy_t)/(E+H),np.where(rel<-fy_c,(rel+fy_c)/(E+H),0.))
        steel=trial-E*delta;plastic+=delta;back+=H*delta
        net=sigma.copy();net[:,concrete:]=steel[:,:,None]*axis[None,concrete:]-sigma[:,concrete:]
        load=fraction*offset+factor*direction;external=L@load
        residual=g@field-external+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/(2*half))
        actual=np.array([weights@net[1,:,2],(weights*y)@net[1,:,2],(weights*x)@net[1,:,2],
            weights@(x*net[1,:,5]-y*net[1,:,4]),weights@net[1,:,4],weights@net[1,:,5]])
        bound=float(absolute+relative*max(abs(external[:dof-4])))
        physical=float(max(abs(residual[:dof-4])));gauge=float(30000*max(abs(residual[dof-4:])))
        force=float(max(abs((actual-load)/force_scales)))
        usage=float(np.max(abs(eps)/np.where(long[None,:],np.where(eps>=0,.025,.018),.02)))
        records.append(dict(index=index+1,lambda_value=factor,Mx_kNm=float(actual[1]/1e6),T_kNm=float(actual[3]/1e6),kappaX_per_mm=float(field[1]/length),
            twist_per_mm=float(field[3]/length),
            physical_residual_MPa=physical,gauge_residual_MPa=gauge,six_force_residual_MPa=force,original_bound_MPa=bound,
            steel_strain_utilization=usage,concrete_min_longitudinal_strain=float(np.min(strain[:,:concrete,2])),
            concrete_max_longitudinal_strain=float(np.max(strain[:,:concrete,2]))))
        records[-1].update(resultant_errors_N_Nmm=(actual-load).tolist(),
            actual_resultants_N_Nmm=actual.tolist(),target_resultants_N_Nmm=load.tolist(),offset_fraction=fraction,
            maximum_force_error_N=float(max(abs((actual-load)[[0,4,5]]))),
            maximum_moment_error_Nmm=float(max(abs((actual-load)[[1,2,3]]))),
            reference_bound_MPa=float(.0000005*(1+max(abs(external[:dof-4])))),
            meets_reference_tolerance=bool(max(physical,gauge,force)<=.0000005*(1+max(abs(external[:dof-4])))+1e-10))
        observed.append(sigma.tolist())
        assert max(physical,gauge,force)<=bound+1e-10,(index+1,factor,physical,gauge,force,bound)
finally:ops.wipe()
proof=dict(engine_version='3.8.0',installed_binary_sha256=sha256(Path(native.__file__).read_bytes()).hexdigest().upper(),
    input_history_sha256=sha256(raw).hexdigest().upper(),material_contract_sha256=sha256(contract_raw).hexdigest().upper(),
    elapsed_seconds=perf_counter()-started,all56plus4_equations_and_six_resultants_verified=True,
    actual_OpenSees_material_replay=True,section_equilibrium_not_solved_by_OpenSees=True,full_steel_plastic_history_replayed=True,
    accepted_states=len(records),maximum_Mx_kNm=max(i['Mx_kNm'] for i in records),records=records,
    controlled_component=component,base_component_Nmm=base_Mx,maximum_T_kNm=max(i['T_kNm'] for i in records),
    offset_N_Nmm=offset.tolist(),direction_N_Nmm=direction.tolist(),maximum_lambda=max(factors),
    declared_absolute_tolerance_MPa=absolute,declared_relative_tolerance=relative,
    maximum_force_error_N=max(i['maximum_force_error_N'] for i in records),
    maximum_moment_error_Nmm=max(i['maximum_moment_error_Nmm'] for i in records),
    reference_tolerance_violated_states=sum(not i['meets_reference_tolerance'] for i in records),
    physical_capacity_certified=False,whole_goal_completed=False,observed_OpenSees_concrete_stresses=observed)
(root/'Native/OpenSeesMainCapacityEquilibrium.json').write_text(json.dumps(proof,separators=(',',':'))+'\n',encoding='utf-8')
print('MAIN CAPACITY EXTERNAL EQUILIBRIUM VERIFIED',json.dumps({k:v for k,v in proof.items() if k not in ('records','observed_OpenSees_concrete_stresses')},sort_keys=True))
