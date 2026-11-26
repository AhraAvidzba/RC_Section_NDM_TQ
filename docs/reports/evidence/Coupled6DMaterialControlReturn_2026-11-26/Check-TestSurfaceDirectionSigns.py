"""ДЛЯ ТЕСТОВ. Сопоставляет направление Ньютона с активной стороной поверхности, не объявляя линейный прогноз равновесием."""
from pathlib import Path
from runpy import run_path
import numpy as np,json
here=Path(__file__).resolve().parent;scope=run_path(str(here/'Check-TestSurfaceMatrices.py'))
root=scope['root'];R=scope['R'];source=scope['source'];base=root.parent
scan=base/'Cold930FineLinePlusSurface_20261126'
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](scan/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+t*.001*gradient for t in (-1,0,1)])
r=R((scan/'Native/ArcScan.States.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
code=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=code.index('def TestYield(');end=code.index('def TestAssemble(',begin);exec(code[begin:end],globals())

def TestSurfaceDerivative(field,direction,point=5,side=1):
    """Дифференцирует пробную поверхность по спектральному проектору, независимо от нативного вычисления градиента."""
    state=states[point,side];alpha,gamma=state[:2];cc,ct=state[8:]
    sigma=D@(op[side+1,point-1]@field-state[2:8]);ds=D@(op[side+1,point-1]@direction)
    xx,yy,zz,xy,xz,yz=sigma;tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]])
    eig,vectors=np.linalg.eigh(tensor);high=eig[-1];normal=vectors[:,-1];trace=sum(sigma[:3]);dev=tensor-np.eye(3)*trace/3
    q=np.sqrt(1.5*np.sum(dev*dev));derivative=(1.5/q)*dev
    beta=cc/ct*(1-alpha)-(1+alpha) if high>0 else gamma
    derivative+=alpha*np.eye(3)+beta*np.outer(normal,normal)
    dx,dy,dz,dxy,dxz,dyz=ds;dt=np.array([[dx,dxy,dxz],[dxy,dy,dyz],[dxz,dyz,dz]])
    value=float(np.sum(derivative*dt)/(1-alpha))
    h=1e-7/max(1.,np.max(abs(ds)))
    difference=(TestYield(point,side,field+h*direction)-TestYield(point,side,field-h*direction))/(2*h)
    error=abs(value-difference)/max(1.,abs(value));assert error<2e-6
    return value,error

items=[]
for item in scope['predictions']:
    field=item['field'];delta=item['correction'][:60];before=TestYield(5,1,field)
    derivative,error=TestSurfaceDerivative(field,delta)
    endpoint=TestYield(5,1,field+delta)
    items.append(dict(level=item['level'],mode=item['mode'],before_trial_yield_MPa=before,
        directional_trial_yield_derivative_MPa=derivative,full_prediction_trial_yield_MPa=endpoint,
        independent_derivative_relative_error=error,linearized_crossing_fraction=float(-before/derivative) if derivative else None))
proof=dict(directions=items,analytic_derivative_checked_independently=True,
    no_actual_nonlinear_root_or_capacity_claim=True,material_history_unchanged=True)
(root/'SurfaceDirectionSignsProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
for item in items:print('SIDE',item['level'],item['mode'],'F',item['before_trial_yield_MPa'],'DF',item['directional_trial_yield_derivative_MPa'],
    'FULL F',item['full_prediction_trial_yield_MPa'])
