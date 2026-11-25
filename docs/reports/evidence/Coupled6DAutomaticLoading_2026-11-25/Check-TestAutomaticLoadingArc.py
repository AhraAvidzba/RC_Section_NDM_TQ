"""ДЛЯ ТЕСТОВ. Независимо проверяет рабочий автоматический исходящий переход, бюджеты, равновесие и точное восстановление."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,struct
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
p.add_argument('--original-state-tolerance',action='store_true');p.add_argument('--direct-loading',action='store_true')
p.add_argument('--expect-coordinate-failure',action='store_true');a=p.parse_args();root=a.root
parent=base/'Cold818AutomaticLoadingCoordinate_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('Input/Reverse1.bin','Input/SourceColdRawHistory.bin','src/Solver/CCoupledNewtonSolver.cls',
    'src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
name='src/Materials/CCoupledConcreteResponse.cls';code=(root/name).read_text('utf-8-sig')
if a.original_state_tolerance:
    old='Private Const STATE_TOLERANCE As Double = 0.0000000005'
    new='Private Const STATE_TOLERANCE As Double = 0.0000000000005'
    assert code.count(old)==1;code=code.replace(old,new,1)
assert code==(parent/name).read_text('utf-8-sig')
count,allowance=struct.unpack('<id',(root/'Input/ArcCallAllowance.bin').read_bytes());assert count==1
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Native/OperationalArc.Result.bin').read_bytes());assert r.raw[:16]==b'NDMOPERARC000003'
packet=[r.VectorTest() for _ in range(15)]
info,projection,accepted,half_seed,full_seed,half,full,ah,af,event,face,face_half,unused,ah_face,parameters=packet
returned=actual=returned_projection=None
if info[3]:returned,actual,returned_projection=[r.VectorTest() for _ in range(3)]
effort=r.VectorTest();assert r.position==len(r.raw)
assert len(info)==30 and info[14]==info[15] and info[16]<=30 and info[17]<=30 and info[18]<=allowance
assert info[19]==(info[16] if info[28]==1 else info[16]+info[17])
assert np.array_equal(accepted,source['fields'][-1]) and info[6]==source['lambdas'][-1]
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
assert len(fields)==43;np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops;E,nu=source['calibration'][:2]
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0

def TestLength(delta):
    """Возвращает норму физического тензорного поля трёх станций без множителей связей."""
    return float(np.sqrt(max(0.,delta@M@delta)))

def TestBound(factor):
    """Восстанавливает внутренний силовой допуск по точным нагрузкам выбранной траектории."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))

def TestWitness(label,field,factor):
    """Собирает все уравнения из фактически сохранённых нативных тензоров и операторов."""
    proof,fs=validate['TestValidate'](root,'OperationalArc.'+label+'.Sequence')
    assert len(fs)==44;np.testing.assert_allclose(fs[:-1],source['fields'],rtol=1e-14,atol=1e-300)
    np.testing.assert_allclose(fs[-1],field,rtol=1e-14,atol=1e-300)
    last=proof['states'][-1]
    assert abs(last['lambda']-factor)<1e-12
    assert last['residual_MPa']<=TestBound(factor)+1e-12 and last['force_residual_MPa']<=TestBound(factor)+1e-12
    return last

witnesses={};pair=None
if a.expect_coordinate_failure:
    assert a.direct_loading and allowance==3 and not info[2] and not info[3] and not info[26]
    assert info[16]==0 and 0<info[17]<=15 and info[18]==1 and info[28]==0
    assert not (root/'Native/OperationalArc.Full.Sequence.csv').exists()
elif allowance>=5 or a.direct_loading:
    assert info[2]==info[3]==info[26]==1 and info[28]==4 and info[18]==(3 if a.direct_loading else 5)
    if a.direct_loading:
        assert info[16]==0 and info[17]<=30
        guards=(root/'Native/RecentLoading.Guards.txt').read_text('cp1251').splitlines()
        assert guards==['ZERO_REJECTED;UNLOADING_REJECTED;FORWARD_SELECTED;REVISION_UNCHANGED','RESET_CONTEXT_REJECTED;REVISION_ADVANCED']
    assert event[0]==1 and event[5]==1 and event[6]==info[14]
    assert effort[0]==0 and np.array_equal(returned,face)
    assert len(parameters)==10
    point,side=int(event[1]),int(event[2]);alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:]
    off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
    D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
    def TestYield(field):
        """Проверяет материальную координату от захваченной принятой пластической истории выбранной станции."""
        xx,yy,zz,xy,xz,yz=D@((b[point-1]+side*.001*gradient[point-1])@field-plastic)
        tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=np.trace(tensor)
        q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2));high=np.linalg.eigvalsh(tensor)[-1]
        term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
        return float((alpha*tr+q+term)/(1-alpha)-cc)
    prediction=accepted.copy();prediction[:56]+=info[8]*(accepted[:56]-source['fields'][-2,:56])
    np.testing.assert_allclose([TestYield(accepted),TestYield(prediction)],event[3:5],rtol=1e-5,atol=2e-12)
    targets=event[3]+np.array([.25,.5])*(event[4]-event[3])
    np.testing.assert_allclose([TestYield(face_half),TestYield(face)],targets,rtol=0,atol=2e-12)
    length=TestLength(face-accepted);error=TestLength(face-(2*face_half-accepted))/length
    le=abs(2*info[24]-info[6]-info[27]);lb=max(.005*abs(info[27]-info[6]),8*.0001*source['options'][4]*(1+abs(info[6])))
    assert error<=.005 and le<=lb
    np.testing.assert_allclose([info[20],info[25]],[error,error],rtol=1e-5,atol=1e-12)
    original=TestLength(prediction-accepted)
    assert TestLength(face_half-accepted)<=1.5*original and length<=3*original
    assert .5*length*E>8*max(TestBound(info[6]),TestBound(info[24]),TestBound(info[27]))
    incoming=accepted-source['fields'][-2]
    cosines=[]
    for field in (face_half,face):
        cos=float(incoming@M@(field-accepted)/(TestLength(incoming)*TestLength(field-accepted)))
        assert cos>0;cosines.append(cos)
    for label,field,factor in (('Event',face,info[27]),('EventHalf',face_half,info[24]),('Full',returned,info[23])):
        witnesses[label]=TestWitness(label,field,factor)
    np.testing.assert_allclose(actual,source['offset']+info[23]*source['direction'],rtol=0,atol=1e-6)
    pair=dict(verified=True,field_half_error=error,lambda_half_error=float(le),lambda_half_bound=float(lb),
        point=point,side=side,source_lambda=float(info[6]),half_lambda=float(info[24]),full_lambda=float(info[27]),
        targets_MPa=targets.tolist(),mechanical_direction_cosines=cosines,
        exact_restoration_zero_iterations=True,unchanged_accepted_revision=True)
else:
    assert allowance==4 and not info[2] and not info[3] and info[18]==2 and info[28]==0
    assert not info[26] and not (root/'Native/OperationalArc.Full.Sequence.csv').exists()
proof=dict(native=native,source43_equilibria_independently_rebuilt=mechanics,
    automatic_unique_recent_station=True,original_state_tolerance=a.original_state_tolerance,direct_loading_coordinate=a.direct_loading,
    numerical_coordinate_failure=a.expect_coordinate_failure,
    info=info.tolist(),event=event.tolist(),effort=effort.tolist(),call_allowance=int(allowance),
    all_primary_and_call_budgets_preserved=True,pair=pair,root_witnesses=witnesses,
    no_material_history_accepted_inside_arc=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'AutomaticLoadingArcProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'AUTOMATIC ARC',bool(info[3]),'CALLS',info[18],'PRIMARY',info[16:20].tolist(),'PAIR',pair,'NO PHYSICAL LIMIT')
