"""ДЛЯ ТЕСТОВ. Независимо исследует смену направления при переключении локальной ветви6/+1 без решения или принятия НДС."""
from pathlib import Path
from runpy import run_path
from contextlib import redirect_stdout
from io import StringIO
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');root=base/'Cold872PointDerivativeAudit20_20261126'
with redirect_stdout(StringIO()):n=run_path(str(here/'Check-TestPointDerivativeAudit.py'))
g,L,w,v,B,gradient,axis,op,records=(n[k] for k in ('g','L','w','v','B','gradient','axis','op','records'));source=n['source'];accepted=source['fields'][-1];incoming=accepted-source['fields'][-2]
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
R=n['R'];r=R((base/'Cold869FixedMetricNextScan_20261126/Native/ArcScan.States.bin').read_bytes());states={}
for _ in range(129):
    p,s=r.VectorTest();states[int(p),int(s)]=r.VectorTest()
assert r.position==len(r.raw)

def TestPhi(point,side,field):
    """Восстанавливает пробный критерий от той же фиксированной пластической истории; это выбор ветви, не разрушение."""
    state=states[point,side];alpha,gamma=state[:2];cc,ct=state[8:]
    sigma=D@(op[side+1,point-1]@field-state[2:8]);xx,yy,zz,xy,xz,yz=sigma
    tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);trace=float(np.trace(tensor));high=float(np.linalg.eigvalsh(tensor)[-1])
    q=float(np.sqrt(1.5*np.sum((tensor-np.eye(3)*trace/3)**2)));term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
    return float((alpha*trace+q+term)/(1-alpha)-cc)

M=np.einsum('spcd,spce,p,c->de',op,op,n['coords'][:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(n['coords'][:,2]));M[56:]=0;M[:,56:]=0
length=float(np.sqrt(incoming@M@incoming));alpha=json.loads((base/'Cold867FixedMetricKernel_20261126/FixedMetricKernelProof.json').read_bytes())['independent_elastic_alpha']
results=[]
for switch in (False,True):
    tensor=np.array([[records[2,side,point]['modes'][1]['D'] for point in range(1,44)] for side in (-1,0,1)])
    if switch:tensor[2,5]=records[1,1,6]['modes'][1]['D']
    tensor[:,25:]=200000*np.einsum('pc,pd->pcd',axis[25:],axis[25:])[None,:,:,:]-tensor[:,25:]
    K=g+np.einsum('pcr,pcd,pds->rs',w,tensor[1],B)
    K+=np.einsum('pcr,pcd,pds->rs',v,(tensor[2]-tensor[0])/.002,B)+np.einsum('pcr,pcd,pds->rs',v,(tensor[2]+tensor[0])/2,gradient)
    tangent=np.linalg.solve(K,L@source['direction']);tn=float(np.sqrt(tangent@M@tangent));orient=float(incoming@M@tangent+alpha**2*(source['lambdas'][-1]-source['lambdas'][-2]))
    dl=np.sign(orient)*length/tn;movement=dl*tangent
    probes=[]
    for point in range(1,44):
        for side in (-1,0,1):
            before=TestPhi(point,side,accepted);after=TestPhi(point,side,accepted+movement/1024)
            expected=bool(records[2,side,point]['direct'][0.,1]['info'][4])
            if switch and (point,side)==(6,1):expected=False
            actual=after>1e-11*states[point,side][8]
            if expected!=actual:probes.append(dict(point=point,side=side,expected_plastic=expected,trial_predicted_plastic=bool(actual),phi_before=before,phi_after=after))
    results.append(dict(point6_unloading_tangent=switch,rank=int(np.linalg.matrix_rank(K)),signed_dlambda=float(dl),trial_mismatches=probes,
        point6_phi_after=TestPhi(6,1,accepted+movement/1024),no_equilibrium_or_acceptance=True))
proof=dict(cases=results,native_local_derivatives_reassembled=True,switched_tangent_approximation_at_accepted_state=True,physical_capacity_certified=False)
(root/'TwoBranchDirectionProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in results:print(item)
