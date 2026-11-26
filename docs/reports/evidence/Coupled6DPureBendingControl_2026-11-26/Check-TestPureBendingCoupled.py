"""ДЛЯ ТЕСТОВ. Проверяет заданные и последние принятые чистые моменты, все уравнения и историю стали 22 холодных пространственных проб."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1006PureBendingCoupled_20261126'));root=parser.parse_args().root
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
reference=json.loads((root/'Input/PureBendingReferenceProof.json').read_bytes());spec=json.loads((root/'Input/PureBendingModels.json').read_bytes())
assert reference['planar_reference_physical_limit_verified'] and not (root/'Input/Reverse1.bin').exists()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/PureBendingCoupled.bin').read_bytes())
assert r.raw[:16]==b'NDMPURECOUP00001';header=r.VectorTest()
mu,models,ratios,dof,points,concrete,half,absolute,relative=header;models,ratios,dof,points,concrete=map(int,(models,ratios,dof,points,concrete))
codes=spec.get('model_codes',[1,2]);assert codes in ([1,2],[2])
assert mu==reference['Mx_ultimate_Nmm'] and models==len(codes) and ratios==11 and dof==60 and points==43 and concrete==25
assert 'PURE_BENDING_CASES: completed='+str(len(codes)*ratios)+'; each_case_cold=True' in text
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/PureBending.Operators.csv',dof,points)
op=np.array([B+side*half*gradient for side in (-1,0,1)])
area=coords[:,2];center=np.average(coords[:concrete,:2],axis=0,weights=area[:concrete]);x,y=(coords[:,:2]-center).T
scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
long=axis[concrete:,2]==1;E=200000.;H=1500./(1.-1500./E);cases=[];maximum_steel_error=0.;maximum_strain_error=0.
for model in codes:
    for ratio in spec['ratios']:
        meta=r.VectorTest();m,rr,target,solved,factor,seconds,rejected,frame_count=meta
        assert m==model and rr==ratio and target==mu*ratio and solved in (0,1) and 0<=factor<=1
        assert frame_count==int(frame_count) and rejected>=0
        field=r.VectorTest();reported=r.VectorTest();strain=np.zeros((3,points,6));net=strain.copy();sigma=strain.copy()
        for k in range(3*points):
            point,side,code=r.VectorTest();assert point==k%points+1 and side==k//points-1 and code==0
            strain[k//points,k%points]=r.VectorTest();net[k//points,k%points]=r.VectorTest();sigma[k//points,k%points]=r.VectorTest()
        history=[];factors=[]
        for _ in range(int(frame_count)):
            fraction,f=r.VectorTest();saved=r.VectorTest();assert fraction==1. and len(saved)==dof
            history.append(saved);factors.append(f)
        if history:
            np.testing.assert_array_equal(history[-1],field);assert factors[-1]==factor
        else:assert factor==0 and np.all(field==0)
        predicted=np.einsum('spcd,d->spc',op,field);error=float(np.max(abs(predicted-strain)));assert error<1e-14
        maximum_strain_error=max(maximum_strain_error,error)
        # Пересчитывается именно накопленная пластическая история стали, а не упругая E*epsilon.
        plastic=np.zeros((3,points-concrete));back=plastic.copy();steel=plastic.copy()
        for saved in history:
            eps=np.einsum('spc,pc->sp',np.einsum('spcd,d->spc',op[:,concrete:],saved),axis[concrete:])
            trial=E*(eps-plastic);rel=trial-back
            fy_t=np.where(long[None,:],400.,240.);fy_c=np.where(long[None,:],350.,240.)
            delta=np.where(rel>fy_t,(rel-fy_t)/(E+H),np.where(rel<-fy_c,(rel+fy_c)/(E+H),0.))
            steel=trial-E*delta;plastic+=delta;back+=H*delta
        steel_error=float(np.max(abs(net[:,concrete:]+sigma[:,concrete:]-steel[:,:,None]*axis[None,concrete:])))
        assert steel_error<1e-7,(model,ratio,steel_error);maximum_steel_error=max(maximum_steel_error,steel_error)
        assert np.max(abs(net[:,:concrete]-sigma[:,:concrete]))==0
        requested=np.array([0.,target,0.,0.,0.,0.]);actual_target=factor*requested;external=L@actual_target
        residual=g@field-external+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/(2*half))
        actual=np.array([area@net[1,:,2],(area*y)@net[1,:,2],(area*x)@net[1,:,2],area@(x*net[1,:,5]-y*net[1,:,4]),area@net[1,:,4],area@net[1,:,5]])
        bound=float(absolute+relative*max(abs(external[:dof-4])))
        physical=float(max(abs(residual[:dof-4])));gauge=float(30000.*max(abs(residual[dof-4:])));force=float(max(abs((actual-actual_target)/scales)))
        assert max(physical,gauge,force)<=bound+1e-10,(model,ratio,physical,gauge,force,bound)
        assert max(abs((actual-reported)/scales))<1e-10
        if solved:assert factor==1. and abs(actual[1]-target)<=bound*scales[1]
        eps=np.einsum('spc,pc->sp',strain[:,concrete:],axis[concrete:])
        utilization=float(np.max(abs(eps)/np.where(long[None,:],np.where(eps>=0,.025,.018),.02)))
        cases.append(dict(model='CDP' if model==1 else 'ASD_native',ratio=ratio,target_Mx_kNm=target/1e6,
            equilibrium_at_requested_moment=bool(solved),last_accepted_load_factor=factor,last_accepted_Mx_kNm=float(actual[1]/1e6),
            last_accepted_ratio_to_planar_ultimate=float(actual[1]/mu),seconds=seconds,rejected_trials=int(rejected),accepted_fields=int(frame_count),
            physical_residual_MPa=physical,gauge_residual_MPa=gauge,six_force_residual_MPa=force,original_bound_MPa=bound,
            steel_strain_utilization=utilization,maximum_longitudinal_concrete_strain=float(np.max(strain[:,:concrete,2])),
            minimum_longitudinal_concrete_strain=float(np.min(strain[:,:concrete,2])),
            accepted_history_load_factors=factors,target_not_substituted_by_last_accepted=True))
assert r.position==len(r.raw) and len(cases)==len(codes)*ratios
proof=dict(reference_Mx_kNm=mu/1e6,native=native,case_count=len(cases),cases=cases,
    maximum_independent_steel_stress_error_MPa=maximum_steel_error,maximum_operator_strain_error=maximum_strain_error,
    all56plus4_equations_and_six_resultants_verified=True,each_case_cold=True,source_unchanged=True,
    material_differences_from_planar_explicit=True,numerical_failure_not_physical_failure=True,whole_goal_completed=False)
(root/'PureBendingCoupledProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('PURE BENDING COUPLED VERIFIED',json.dumps({k:v for k,v in proof.items() if k!='cases'},sort_keys=True))
for c in cases:print(c['model'],c['ratio'],c['equilibrium_at_requested_moment'],c['last_accepted_ratio_to_planar_ultimate'],c['seconds'])
