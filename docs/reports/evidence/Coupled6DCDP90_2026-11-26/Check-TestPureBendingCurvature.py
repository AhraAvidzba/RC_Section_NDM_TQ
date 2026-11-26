"""ДЛЯ ТЕСТОВ. Пересобирает равновесие кривой Mx-kappaX и подтверждает существование состояний за ранней остановкой силового нагружения."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1011PureBendingASDCurvature_20261126'))
p.add_argument('--allow-timeout-prefix',action='store_true');args=p.parse_args();root=args.root
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
timeout=native['timeout600']
assert native==dict(passed=0 if timeout else 1,failed=0,macros=0 if timeout else 1,timeout600=timeout,own_process_closed=True)
assert (args.allow_timeout_prefix and timeout) or (not timeout and 'CURVATURE_TRACE:' in text)
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];b=R((root/'Native/CurvatureBaseline.bin').read_bytes());b.position=0
baseline=b.VectorTest();assert baseline[1]==.25 and baseline[3:5].tolist()==[1.,1.]
r=R((root/'Native/PureBendingCurvature.bin').read_bytes());assert r.raw[:16]==b'NDMBENDCURV00001'
header=r.VectorTest();model,mu,kref,length,dof,points,concrete,half,absolute,relative=header
model,dof,points,concrete=map(int,(model,dof,points,concrete));assert model in (1,2) and (dof,points,concrete)==(60,43,25)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/PureBending.Operators.csv',dof,points)
assert abs(length-L[0,0]/L[1,1])<1e-12
op=np.array([B+s*half*gradient for s in (-1,0,1)])
area=coords[:,2];center=np.average(coords[:concrete,:2],axis=0,weights=area[:concrete]);x,y=(coords[:,:2]-center).T
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
long=axis[concrete:,2]==1;history=[];material_frames=[];records=[];attempts=[]

def TestReadState(reader):
    """Читает поле и полный материальный отклик всех трёх станций без округления."""
    field=reader.VectorTest();reported=reader.VectorTest();strain=np.zeros((3,points,6));net=strain.copy();sigma=strain.copy()
    for k in range(3*points):
        point,side,code=reader.VectorTest();assert point==k%points+1 and side==k//points-1 and code==0
        strain[k//points,k%points]=reader.VectorTest();net[k//points,k%points]=reader.VectorTest();sigma[k//points,k%points]=reader.VectorTest()
    return field,reported,strain,net,sigma

base_state=TestReadState(b)
for i in range(int(baseline[-1])):
    fraction,f=b.VectorTest();field=b.VectorTest();assert fraction==1.
    history.append(field);material_frames.append(dict(strain=np.einsum('spcd,d->spc',op,field).reshape(-1,6).tolist(),expected=None))
assert b.position==len(b.raw)
np.testing.assert_array_equal(history[-1],base_state[0])
material_frames[-1]['expected']=base_state[-1].reshape(-1,6).tolist()
curve_base=TestReadState(r)
for a,c in zip(base_state,curve_base):np.testing.assert_array_equal(a,c)

def TestMechanics(state,factor,target=None):
    """Проверяет слабые уравнения, контроль кривизны, шесть сил и накопленное кинематическое упрочнение стали."""
    field,reported,strain,net,sigma=state
    np.testing.assert_allclose(strain,np.einsum('spcd,d->spc',op,field),atol=1e-14,rtol=0)
    plastic=np.zeros((3,points-concrete));back=plastic.copy();steel=plastic.copy();E=200000.;H=1500/(1-1500/E)
    for saved in history:
        eps=np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op[:,concrete:],saved),axis[concrete:]);trial=E*(eps-plastic);rel=trial-back
        fy_t=np.where(long[None,:],400.,240.);fy_c=np.where(long[None,:],350.,240.)
        delta=np.where(rel>fy_t,(rel-fy_t)/(E+H),np.where(rel<-fy_c,(rel+fy_c)/(E+H),0.))
        steel=trial-E*delta;plastic+=delta;back+=H*delta
    steel_error=float(np.max(abs(net[:,concrete:]+sigma[:,concrete:]-steel[:,:,None]*axis[None,concrete:])));assert steel_error<1e-7
    assert np.max(abs(net[:,:concrete]-sigma[:,:concrete]))==0
    load=np.array([0.,factor*mu,0.,0.,0.,0.]);external=L@load
    residual=g@field-external+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/(2*half))
    actual=np.array([area@net[1,:,2],(area*y)@net[1,:,2],(area*x)@net[1,:,2],area@(x*net[1,:,5]-y*net[1,:,4]),area@net[1,:,4],area@net[1,:,5]])
    bound=float(absolute+relative*max(abs(external[:dof-4])));physical=float(max(abs(residual[:dof-4])));gauge=float(30000*max(abs(residual[dof-4:])));force=float(max(abs((actual-load)/scales)))
    control=0. if target is None else 30000*abs(field[1]-target)
    assert max(physical,gauge,force,control)<=bound+1e-10,(factor,physical,gauge,force,control,bound)
    assert max(abs((actual-reported)/scales))<1e-10
    eps=np.einsum('spc,pc->sp',strain[:,concrete:],axis[concrete:]);usage=float(np.max(abs(eps)/np.where(long[None,:],np.where(eps>=0,.025,.018),.02)))
    c_usage=float(max(-strain[:,:concrete,2].reshape(-1))/.0035)
    records.append(dict(lambda_value=factor,Mx_kNm=float(actual[1]/1e6),kappaX_per_mm=float(field[1]/length),
        physical_residual_MPa=physical,gauge_residual_MPa=gauge,six_force_residual_MPa=force,control_residual_MPa=control,
        original_bound_MPa=bound,independent_steel_stress_error_MPa=steel_error,steel_strain_utilization=usage,
        longitudinal_concrete_strain_ratio_to_planar_limit=c_usage,concrete_min_longitudinal_strain=float(np.min(strain[:,:concrete,2])),
        concrete_max_longitudinal_strain=float(np.max(strain[:,:concrete,2]))))

TestMechanics(curve_base,.25)
missing_footer=False
while True:
    if r.position==len(r.raw):
        assert timeout and args.allow_timeout_prefix
        missing_footer=True
        ending=np.array([0.,records[-1]['lambda_value'],records[-1]['kappaX_per_mm'],len(records)-1,sum(not a['solved'] for a in attempts)])
        break
    meta=r.VectorTest()
    if meta[0]==0:ending=meta;break
    attempt,target,solved,factor,iterations,evaluations,norm=meta
    assert attempt==len(attempts)+1 and solved in (0,1) and 0<=iterations<=30
    attempts.append(dict(attempt=int(attempt),target_normalized_kappa=target,solved=bool(solved),iterations=int(iterations),function_evaluations=int(evaluations),reported_norm_MPa=norm))
    if not solved:continue
    state=TestReadState(r);assert state[0][1]>history[-1][1]
    history.append(state[0]);material_frames.append(dict(strain=state[2].reshape(-1,6).tolist(),expected=state[4].reshape(-1,6).tolist()))
    TestMechanics(state,factor,target)
assert r.position==len(r.raw) and ending[3]==len(records)-1 and ending[4]==len(attempts)-ending[3]
assert ending[1]==records[-1]['lambda_value'] and abs(ending[2]-records[-1]['kappaX_per_mm'])<1e-15
assert np.all(np.diff([s['kappaX_per_mm'] for s in records])>0)
factors=np.array([s['lambda_value'] for s in records]);descending=np.flatnonzero(np.diff(factors)<-1e-5)
permitted=[s for s in records if s['steel_strain_utilization']<=1 and s['longitudinal_concrete_strain_ratio_to_planar_limit']<=1]
proof=dict(native=native,model='CDP' if model==1 else 'ASD_native',reference_Mx_kNm=mu/1e6,accepted_curve_points=len(records),rejected_trials=int(ending[4]),
    maximum_equilibrium_ratio=float(max(factors)),maximum_equilibrium_Mx_kNm=max(s['Mx_kNm'] for s in records),
    observed_load_decrease_steps=descending.tolist(),post_decrease_reascending_equilibria=bool(len(descending) and factors[-1]>min(factors[descending[0]:])+.01),
    maximum_sampled_ratio_within_planar_compression_and_steel_strain_limits=max(s['lambda_value'] for s in permitted) if permitted else None,
    all56plus4_equations_and_six_resultants_verified=True,prescribed_curvature_verified=True,
    records=records,attempts=attempts,source_unchanged=True,material_and_force_tolerances_unchanged=True,
    no_future_equilibrium_input=True,completed_prefix_only=missing_footer,whole_run_passed=not timeout,
    timeout_is_not_NumFail_or_physical_failure=timeout,uncompleted_next_trial_result_unknown=timeout,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'PureBendingCurvatureProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
if model==2:
    wire=dict(material_contract=json.loads((root/'Input/MaterialContract.json').read_bytes()),frames=material_frames,
        no_equilibrium_solve_in_external_material_replay=True,native_curve_sha256=sha256(r.raw).hexdigest().upper())
    (root/'Native/CurvatureMaterialReplayInput.json').write_text(json.dumps(wire,separators=(',',':'))+'\n',encoding='utf-8')
print('PURE BENDING CURVATURE VERIFIED',json.dumps({k:v for k,v in proof.items() if k not in ('records','attempts')},sort_keys=True))
