"""ДЛЯ ТЕСТОВ. Независимо строит градиент CDP и исходящее касательное направление из фактической матрицы равновесия."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
from argparse import ArgumentParser

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--refined',action='store_true');args=p.parse_args()
root=base/('Cold850RefinedOutgoingTangentMetric_20261125' if args.refined else 'Cold849OutgoingTangentMetric_20261125')
parent=base/('Cold849OutgoingTangentMetric_20261125' if args.refined else 'Cold848NewLoadingCoordinateDiagnosis_20261125')
capture_root=base/'Cold848NewLoadingCoordinateDiagnosis_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes() and len(source['fields'])==86
for name in ('src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Solver/CCoupledNewtonSolver.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls','src/Materials/CCoupledReinforcedResponse.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
r=Reader((root/'Native/OutgoingTangent.bin').read_bytes());assert r.raw[:16]==(b'NDMOUTTAN0000002' if args.refined else b'NDMOUTTAN0000001')
recent,state=r.VectorTest(),r.VectorTest();point,side=map(int,recent[1:3]);assert point==2 and side==-1
operator=op[side+1,point-1];accepted=source['fields'][-1];movement=accepted-source['fields'][-2]
length=float(np.sqrt(movement@M@movement));load_column=L@source['direction'];results=[]
# Направление малого сошедшегося корня 848 используется только для сравнения после вычисления касательной.
capture=Reader((capture_root/'Native/YieldCoordinate.Checks.bin').read_bytes());small_field=None
for case in range(1,7):
    meta=capture.VectorTest();packet=[capture.VectorTest() for _ in range(14)]
    for number in range(1,int(meta[7])+1):
        info,seed,scales,solution,mapped=[capture.VectorTest() for _ in range(5)]
        if case==6:small_field=(seed+scales*mapped)[:60]
    if meta[2]:[capture.VectorTest() for _ in range(3)]
assert capture.position==len(capture.raw) and small_field is not None
observed=small_field-accepted;observed_length=float(np.sqrt(observed@M@observed))
sequence=[(case,level,mode) for case,level in enumerate((0,.25,.0625,.015625),1) for mode in ((0,2) if args.refined else (0,))]
for case,level,mode in sequence:
    if args.refined:
        meta,field,local_gradient=[r.VectorTest() for _ in range(3)]
        assert meta[7]==mode;ready=bool(meta[8])
        if ready:residual=r.VectorTest()
    else:
        meta,field,residual,local_gradient=[r.VectorTest() for _ in range(4)];ready=True
    K=np.column_stack([r.VectorTest() for _ in range(60)]) if ready else None
    assert meta[0]==case and meta[1]==level and meta[2]==point and meta[3]==side and meta[5]==meta[6]
    np.testing.assert_allclose(field,accepted+level*movement,rtol=1e-14,atol=1e-300)
    sigma=D@(operator@field-state[2:8]);xx,yy,zz,xy,xz,yz=sigma
    tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);eigen,vectors=np.linalg.eigh(tensor);high=eigen[-1]
    trace=float(np.trace(tensor));dev=tensor-np.eye(3)*trace/3;q=float(np.sqrt(1.5*np.sum(dev**2)))
    alpha,gamma=state[:2];cc,ct=state[8:];beta=cc/ct*(1-alpha)-(1+alpha) if high>0 else gamma
    phi=(alpha*trace+q+beta*high)/(1-alpha)-cc
    assert abs(phi-meta[4])<=3e-11
    direction=vectors[:,-1]
    projector=np.array([direction[0]**2,direction[1]**2,direction[2]**2,2*direction[0]*direction[1],2*direction[0]*direction[2],2*direction[1]*direction[2]])
    qgradient=np.array([1.5*dev[i,i]/q for i in range(3)]+[3*xy/q,3*xz/q,3*yz/q])
    independent=D.T@((alpha*np.array([1,1,1,0,0,0])+qgradient+beta*projector)/(1-alpha))
    assert np.max(abs(local_gradient-independent))/np.max(abs(independent))<=1e-9
    if not ready:
        results.append(dict(case=case,seed_level=level,mode=mode,available=False));continue
    control_gradient=independent@operator
    A=np.zeros((61,61));A[:60,:60]=K;A[:60,60]=-load_column;A[60,:60]=control_gradient
    rhs=np.zeros(61);rhs[-1]=1
    row_scale=1/np.max(abs(A),axis=1);Ar=row_scale[:,None]*A;column_scale=1/np.max(abs(Ar),axis=0)
    direction=column_scale*np.linalg.solve(Ar*column_scale[None,:],row_scale*rhs)
    linear_residual=float(np.max(abs(A@direction-rhs)))
    tangent_length=float(np.sqrt(direction[:60]@M@direction[:60]))
    predictor=direction*(length/tangent_length)
    cos=float(direction[:60]@M@observed/(tangent_length*observed_length))
    incoming_phi=float(control_gradient@movement)
    item=dict(case=case,seed_level=level,mode=mode,available=True,phi_MPa=phi,independent_gradient_relative_error=float(np.max(abs(local_gradient-independent))/np.max(abs(independent))),
        bordered_linear_residual=linear_residual,scaled_matrix_condition=float(np.linalg.cond(Ar*column_scale[None,:])),
        outgoing_to_incoming_gain=tangent_length*incoming_phi/length,outgoing_observed_direction_cosine=cos,
        equal_length_predictor_phi_increment=float(control_gradient@predictor[:60]),incoming_phi_increment=incoming_phi,
        equal_length_predictor_lambda_increment=float(predictor[-1]),predictor=predictor.tolist())
    results.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,source_states=86,working_math_unchanged=True,no_candidate_accepted=True,
           independent_CDP_gradients_verified=True,existing_refined_branch=args.refined,one_sided_tangents=results,physical_capacity_certified=False,whole_goal_completed=False)
(root/'OutgoingTangentMetricProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for item in results:print({k:v for k,v in item.items() if k!='predictor'})
