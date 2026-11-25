"""ДЛЯ ТЕСТОВ. Независимо проверяет выбор спектрального события, градиенты, силы и неизменность основной материальной модели."""
from pathlib import Path
from runpy import run_path
import json,re
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold837ElasticPrincipalEvent_20261125';parent=base/'Cold834UnresolvedArcScan_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
assert len(source['fields'])==86
assert (root/'src/Solver/CCoupledArcContinuation.cls').read_bytes()==(parent/'src/Solver/CCoupledArcContinuation.cls').read_bytes()
assert (root/'src/Materials/CCoupledConcreteReturn.cls').read_bytes()==(parent/'src/Materials/CCoupledConcreteReturn.cls').read_bytes()
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls'):
    old=(parent/name).read_text('utf-8-sig');new=(root/name).read_text('utf-8-sig')
    assert new.startswith(old),name
    assert new[len(old):].count('Friend Function TryReadElasticPrincipalBoundary')==1
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Input/ArcScanFields.bin').read_bytes());r.position=0;assert r.VectorTest()[0]==6
ends=[r.VectorTest() for _ in range(6)];assert r.position==len(r.raw)
r=Reader((parent/'Native/ArcScan.States.bin').read_bytes());parameters={}
for _ in range(129):
    point,side=r.VectorTest();parameters[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
r=Reader((root/'Native/ElasticPrincipal.Event.bin').read_bytes());assert r.raw[:16]==b'NDMELASTIC000001'
events=[]
for case in range(1,7):
    meta,info,effort,field=[r.VectorTest() for _ in range(4)]
    assert int(meta[0])==case and len(info)==9
    events.append(dict(case=case,solved=bool(meta[1]),lambda_value=float(meta[2]),info=info.tolist(),effort=effort.tolist(),field=field))
assert r.position==len(r.raw)
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
source_equilibria,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
source_field=source['fields'][-1]

def TestTensor(s):
    """Разворачивает физические компоненты напряжений без инженерного удвоения сдвигов."""
    xx,yy,zz,xy,xz,yz=s
    return np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]])

def TestTrial(field,point,side):
    """Восстанавливает упругий прогноз и поверхность CDP от сохранённой принятой истории."""
    state=parameters[point,side];alpha,gamma=state[:2];cc,ct=state[8:]
    strain=(b[point-1]+side*.001*gradient[point-1])@field
    sigma=D@(strain-state[2:8]);tensor=TestTensor(sigma)
    ev,vectors=np.linalg.eigh(tensor);tr=float(np.trace(tensor));q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2))
    term=(cc/ct*(1-alpha)-(1+alpha))*ev[-1] if ev[-1]>0 else gamma*ev[-1]
    phi=float((alpha*tr+q+term)/(1-alpha)-cc)
    return sigma,ev,vectors,phi

def TestEarliest(case):
    """Независимо перечисляет все разрешимые спектральные пересечения на данном отрезке."""
    candidates=[];delta=ends[case-1]-source_field
    for point in range(1,44):
        for side in (-1,0,1):
            _,left,_,phi0=TestTrial(source_field,point,side);_,right,_,phi1=TestTrial(ends[case-1],point,side)
            if phi0>=0 or phi1>=0:continue
            for index in range(3):
                if abs(left[index])<=32e-12 or left[index]*right[index]>=0:continue
                lo,hi=0.,1.
                for _ in range(50):
                    mid=.5*(lo+hi);_,ev,_,phi=TestTrial(source_field+mid*delta,point,side)
                    assert phi<0
                    if (ev[index]<0)==(left[index]<0):lo=mid
                    else:hi=mid
                candidates.append(dict(point=point,side=side,principal=3-index,fraction=.5*(lo+hi)))
    assert candidates
    return min(candidates,key=lambda c:c['fraction']),candidates

rows=validate['TestRows'](root/'Native/ElasticPrincipal.Roots.csv') if (root/'Native/ElasticPrincipal.Roots.csv').exists() else []
stress={};bars={};spectral=[]
for row in rows:
    kind,case,point,side=row[:4];case,point,side=int(case),int(point),int(side)
    if kind=='K':stress[case,side,point]=np.asarray(row[4:10],float);bars[case,side,point]=float(row[10])
    elif kind=='P':spectral.append((case,point,side,int(row[4]),row[5]=='True',float(row[6]),np.asarray(row[7:13],float)))
    else:raise AssertionError(kind)
results=[];maximum_gradient_error=0.;gradient_reads=0
for event in events:
    case=event['case'];earliest,candidates=TestEarliest(case)
    info=event['info'];assert info[0]==1 and info[8]==1
    point,side,principal=map(int,(info[1],info[2],info[7]))
    assert any(c['point']==point and c['side']==side and c['principal']==principal for c in candidates)
    assert abs(info[5]-earliest['fraction'])<=1e-7,(case,info,earliest)
    assert event['effort'][0]<=15
    if not event['solved']:
        results.append(dict(case=case,solved=False,earliest=earliest,info=info,effort=event['effort']))
        continue
    field=event.pop('field');sigma,ev,_,phi=TestTrial(field,point,side)
    assert phi<0 and abs(ev[3-principal])<=6e-12,(case,ev,phi)
    net=np.zeros((3,43,6));steel_utilization=0.
    for station in (-1,0,1):
        for p in range(1,44):
            measured=stress[case,station,p]
            net[station+1,p-1]=measured if p<=25 else bars[case,station,p]*axis[p-1]-measured
            if p>25:
                strain=(b[p-1]+station*.001*gradient[p-1])@field
                axial=float(strain@axis[p-1]);limit=(.025 if axial>=0 else .018) if axis[p-1,2]==1 else .02
                steel_utilization=max(steel_utilization,abs(axial)/limit)
    assert steel_utilization<=1
    target=source['offset']+event['lambda_value']*source['direction'];external=L@target
    weak=g@field+np.einsum('pcd,pc->d',w,net[1])+np.einsum('pcd,pc->d',v,(net[2]-net[0])/.002)-external
    norm=max(float(max(abs(weak[:56]))),float(E*max(abs(weak[56:]))))
    center=np.average(coords[:25,:2],axis=0,weights=coords[:25,2]);x,y=(coords[:,:2]-center).T;weights=coords[:,2]
    s=net[1];actual=np.array([weights@s[:,2],(weights*y)@s[:,2],(weights*x)@s[:,2],weights@(x*s[:,5]-y*s[:,4]),weights@s[:,4],weights@s[:,5]])
    scales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
    force=float(max(abs((actual-target)/scales)))
    bound=.0001*(source['options'][3]+source['options'][4]*max(abs(external[:56])))
    assert norm<=bound+1e-12 and force<=bound+1e-12,(case,norm,force,bound)
    results.append(dict(case=case,solved=True,lambda_value=event['lambda_value'],selected_point=point,selected_side=side,
        selected_principal=principal,earliest=earliest,info=info,effort=event['effort'],
        eigen_zero_MPa=float(ev[3-principal]),yield_MPa=phi,weak_norm_MPa=norm,six_force_norm_MPa=force,
        inner_bound_MPa=float(bound),three_station_steel_utilization=steel_utilization))
for case,point,side,principal,available,value,measured in spectral:
    event=events[case-1]
    # Поле читается повторно из бинарного свидетельства, а не из округлённого CSV.
    raw=Reader((root/'Native/ElasticPrincipal.Event.bin').read_bytes())
    for _ in range(case):meta,info,effort,field=[raw.VectorTest() for _ in range(4)]
    _,ev,vectors,phi=TestTrial(field,point,side)
    gaps=sorted(abs(ev[3-principal]-ev[j]) for j in range(3) if j!=3-principal)
    if phi>1e-10:assert not available
    if phi < -1e-10 and gaps[0]>1e-9*max(1.,max(abs(ev))):assert available
    if not available:continue
    assert phi<0 and abs(value-ev[3-principal])<=2e-10
    eigenvector=vectors[:,3-principal];xx,yy,zz=eigenvector
    independent=np.array([xx*xx,yy*yy,zz*zz,2*xx*yy,2*xx*zz,2*yy*zz])@D
    error=float(max(abs(independent-measured)));maximum_gradient_error=max(maximum_gradient_error,error)
    assert error<=1e-5+1e-8*max(abs(independent)),(case,point,side,principal,error)
    gradient_reads+=1
proof=dict(native=native,source86_equilibria_independently_rebuilt=source_equilibria,
    original_concrete_and_steel_evaluation_methods_unchanged=True,original_arc_and_return_unchanged=True,
    no_material_history_accepted=True,earliest_sign_crossings_independently_verified=True,
    events=results,solved_events=sum(x['solved'] for x in results),
    analytic_gradients_verified=gradient_reads,maximum_analytic_gradient_error=maximum_gradient_error,
    physical_capacity_certified=False,whole_goal_completed=False,arc_integration_not_yet_done=True)
(root/'ElasticPrincipalEventProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('ELASTIC PRINCIPAL EVENTS',proof['solved_events'],'OF6;GRADIENTS',gradient_reads,'ERROR',maximum_gradient_error)
for x in results:print(x)
assert proof['solved_events']==6,'Some selected event roots remain unresolved; numerical failure is not physical failure.'
