"""ДЛЯ ТЕСТОВ. Проверяет меньший подход к событию по четвертной пробе, строгим силам, истории и сохранённым напряжениям."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold792PreEventQuarter_20261125';parent=base/'Cold791ResolvedLoadingEvent_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledArcContinuation.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==10
rd=Reader((root/'Native/PreEventQuarter.Result.bin').read_bytes());assert rd.raw[:16]==b'NDMPREEVENT00001\0'[:16]
info,face,half,quarter,projection,aq,ah,restored=[rd.VectorTest() for _ in range(8)];assert rd.position==len(rd.raw) and len(info)==15
assert info[9]==info[10]==12 and sum(info[11:14])<=30 and info[14]<=6
r=Reader((root/'Native/OperationalArc.Result.bin').read_bytes());packet=[r.VectorTest() for _ in range(15)]
status=packet[0];accepted=packet[2];parameters=packet[14];event=packet[9]
assert status[3]==0 and np.array_equal(face,packet[10]) and np.array_equal(half,packet[11])
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0


def TestLength(delta):
    """Возвращает физическую длину по операторам всех трёх станций."""
    return float(np.sqrt(max(0.,delta@M@delta)))


def TestBound(factor):
    """Восстанавливает неизменный внутренний допуск для фактической нагрузки."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))


proof=dict(native=native,all_working_methods_exact791=True,quarter_solved=bool(info[0]),half_verified=bool(info[1]),
    shared_event_primary_steps=int(sum(info[11:14])),original_event_primary=int(info[11]),quarter_primary=int(info[12]),restore_primary=int(info[13]),
    total_nonlinear_calls=int(info[14]),accepted_revision_before=int(info[9]),accepted_revision_after=int(info[10]),
    physical_capacity_certified=False,whole_goal_completed=False)
if info[0]:
    spacing=TestLength(face-accepted);length=TestLength(half-accepted)
    error=TestLength(half-(2*quarter-accepted))/length
    lambda_error=abs(2*info[2]-info[4]-info[3]);lambda_bound=max(.005*abs(info[3]-info[4]),8*source['options'][4]*.0001*(1+abs(info[4])))
    np.testing.assert_allclose(info[5:9],[spacing,error,lambda_error,lambda_bound],rtol=1e-6,atol=1e-14)
    np.testing.assert_allclose(projection,M@(face-accepted)/spacing,rtol=1e-8,atol=1e-17)
    for field,target in ((quarter,projection@face-.75*spacing),(half,projection@face-.5*spacing)):
        assert abs(projection@field-target)*source['calibration'][0]<=max(TestBound(info[2]),TestBound(info[3]))+1e-12
    E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));shear=E/(2*(1+nu))
    D=np.diag([off+2*shear]*3+[shear]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
    alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:];point,side=int(event[1]),int(event[2])

    def TestYield(field):
        """Проверяет положение относительно исходной поверхности события по принятой пластической истории."""
        xx,yy,zz,xy,xz,yz=D@((b[point-1]+side*.001*gradient[point-1])@field-plastic)
        tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=xx+yy+zz;high=np.linalg.eigvalsh(tensor)[-1]
        q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2))
        term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
        return float((alpha*tr+q+term)/(1-alpha)-cc)

    proof.update(field_half_error=float(error),lambda_half_error=float(lambda_error),lambda_half_bound=float(lambda_bound),
        half_trial_yield_MPa=TestYield(half),quarter_trial_yield_MPa=TestYield(quarter),
        event_yield_MPa=TestYield(face),half_lambda=float(info[3]),quarter_lambda=float(info[2]))
    if info[1]:
        assert error<=.005 and lambda_error<=lambda_bound and info[13]==0 and np.array_equal(half,restored)
        assert TestYield(half)<0 and TestYield(quarter)<0 and abs(TestYield(face))<=2.1e-12
        assert .5*length*E>8*max(TestBound(info[2]),TestBound(info[3]))
    scale=np.array([1/L[0,0],1/L[1,1],1/L[2,2],1/L[3,3],1/L[0,0],1/L[0,0]])
    for label,field,factor,force in (('quarter',quarter,info[2],aq),('half',half,info[3],ah)):
        if label=='half' and not info[1]:continue
        force_norm=float(max(abs((force-source['offset']-factor*source['direction'])/scale)))
        gauge=float(max(abs(g[56:]@field))*E)
        assert force_norm<=TestBound(factor)+1e-14 and gauge<=TestBound(factor)+1e-13
        proof[label+'_raw_six_residual_MPa']=force_norm;proof[label+'_gauge_MPa']=gauge
    validate=run_path(str(here/'Validate-TestHardPrepared.py'))
    for label,field in (('PreEventQuarter',quarter),('PreEventHalf',half)):
        if label=='PreEventHalf' and not info[1]:continue
        mechanics,fields=validate['TestValidate'](root,label+'.Sequence');assert len(fields)==11
        np.testing.assert_allclose(fields[:10],source['fields'],rtol=1e-14,atol=1e-300)
        np.testing.assert_allclose(fields[10],field,rtol=1e-14,atol=1e-300)
        proof[label+'_independent_all_weak_and_six_force_equilibria']=mechanics
(root/'PreEventQuarterProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('792 QUARTER',bool(info[0]),'HALF VERIFIED',bool(info[1]),'ERROR',proof.get('field_half_error'),'EVENT PRIMARY',proof['shared_event_primary_steps'],'CALLS',proof['total_nonlinear_calls'])
