"""ДЛЯ ТЕСТОВ. Независимо проверяет рабочее уточнение до события, все кандидаты, ошибки поля/λ и фактические равновесия."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import re
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold793AdaptivePreEvent_20261125');p.add_argument('--allowance',type=int,default=12)
args=p.parse_args();root=args.root
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==10
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Native/OperationalArc.Result.bin').read_bytes());assert r.raw[:16]==b'NDMOPERARC000003'
packet=[r.VectorTest() for _ in range(15)]
info,p_arc,accepted,half_seed,full_seed,half,full,ah,af,event,face,face_half,unused,ah_face,parameters=packet
assert len(info)==30 and info[14]==info[15]==12 and info[18]<=args.allowance
assert info[16]<=30 and info[17]<=30
assert info[19]==(info[16] if info[28]==1 else info[16]+info[17])
returned=actual=returned_projection=None
if info[3]:returned,actual,returned_projection=[r.VectorTest() for _ in range(3)]
effort=r.VectorTest();assert r.position==len(r.raw)
if info[3]:assert info[2] and effort[0]==0
assert np.array_equal(accepted,source['fields'][-1]) and info[6]==source['lambdas'][-1]
prior=Reader((base/'Cold791ResolvedLoadingEvent_20261125/Native/OperationalArc.Result.bin').read_bytes())
prior_packet=[prior.VectorTest() for _ in range(15)]
for i in (1,2,3,4,5,6,7,8):assert np.array_equal(packet[i],prior_packet[i]),i
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledConcreteReturn.cls'):
    assert (root/name).read_bytes()==(base/'Cold791ResolvedLoadingEvent_20261125'/name).read_bytes(),name
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0
E=source['calibration'][0];scales=np.array([1/L[0,0],1/L[1,1],1/L[2,2],1/L[3,3],1/L[0,0],1/L[0,0]])
validate=run_path(str(here/'Validate-TestHardPrepared.py'))


def TestLength(delta):
    """Возвращает тензорную длину в общей метрике без множителей связей."""
    return float(np.sqrt(max(0.,delta@M@delta)))


def TestBound(factor):
    """Восстанавливает фактически применённый допуск, сохраняя исходную шестикомпонентную траекторию."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))


def TestForce(field,factor,force):
    """Проверяет отдельно интегралы шести усилий и уравнения однозначности поля."""
    norm=float(max(abs((force-source['offset']-factor*source['direction'])/scales)))
    gauge=float(max(abs(g[56:]@field))*E)
    assert norm<=TestBound(factor)+1e-14 and gauge<=TestBound(factor)+1e-13
    return dict(raw_six_residual_MPa=norm,gauge_MPa=gauge,bound_MPa=TestBound(factor))


def TestWitness(label,field,factor):
    """Восстанавливает слабые и интегральные уравнения из нативных тензоров диагностического повтора."""
    mechanics,fields=validate['TestValidate'](root,'OperationalArc.'+label+'.Sequence')
    assert len(fields)==11
    np.testing.assert_allclose(fields[:10],source['fields'],rtol=1e-14,atol=1e-300)
    np.testing.assert_allclose(fields[10],field,rtol=1e-14,atol=1e-300)
    assert abs(mechanics['states'][-1]['lambda']-factor)<1e-12
    assert mechanics['states'][-1]['residual_MPa']<=TestBound(factor)+1e-12
    return mechanics


point=side=None
if info[26]:
    assert len(event)==7 and len(parameters)==10
    point,side=int(event[1]),int(event[2]);alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:]
    nu=source['calibration'][1];off=E*nu/((1+nu)*(1-2*nu));shear=E/(2*(1+nu))
    D=np.diag([off+2*shear]*3+[shear]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))

    def TestYield(field):
        """Считает пробную поверхность от захваченной до принятия пластической истории выбранной станции."""
        xx,yy,zz,xy,xz,yz=D@((b[point-1]+side*.001*gradient[point-1])@field-plastic)
        tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=xx+yy+zz;high=np.linalg.eigvalsh(tensor)[-1]
        q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2))
        term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
        return float((alpha*tr+q+term)/(1-alpha)-cc)

    assert abs(TestYield(face))<=2.1e-12
    np.testing.assert_allclose([TestYield(accepted),TestYield(full)],event[3:5],rtol=1e-6,atol=2e-12)
    np.testing.assert_allclose(info[27],prior_packet[0][23],rtol=0,atol=1e-12)
    assert event[6]==12
reader=Reader((root/'Native/OperationalArc.Approaches.bin').read_bytes());assert reader.raw[:16]==b'NDMARCAPPR000001'
count=reader.VectorTest();assert len(count)==1 and count[0]==info[29]
records=[]
for number in range(1,int(count[0])+1):
    record=[reader.VectorTest() for _ in range(6)];records.append(record)
assert reader.position==len(reader.raw) and len(records)<=4
proofs=[];last=None
for number,(meta,candidate,probe,projection,ac,ap) in enumerate(records,1):
    assert len(meta)==16 and meta[0]==number and meta[1]==.5**number and meta[4]==info[6] and meta[11]==meta[12]==12
    assert meta[13]<=info[17] and meta[15]<args.allowance
    expected_candidate=face_half if number==1 else last[2]
    expected_factor=info[24] if number==1 else last[0][3]
    assert np.array_equal(candidate,expected_candidate) and meta[2]==expected_factor
    assert len(probe)==60
    length=TestLength(candidate-accepted);error=TestLength(candidate-(2*probe-accepted))/length
    le=abs(2*meta[3]-info[6]-meta[2]);lb=max(.005*abs(meta[2]-info[6]),8*source['options'][4]*.0001*(1+abs(info[6])))
    resolved=.5*length*E>8*max(TestBound(meta[2]),TestBound(meta[3]))
    progress_c=float(p_arc@(candidate-accepted)+info[12]*(meta[2]-info[6]))
    progress_p=float(p_arc@(probe-accepted)+info[12]*(meta[3]-info[6]))
    spacing=TestLength(face-accepted)
    local=length<=1.5*meta[1]*spacing and TestLength(probe-accepted)<=.75*meta[1]*spacing
    yc,yp=TestYield(candidate),TestYield(probe)
    verified=error<=.005 and le<=lb and resolved and min(meta[2:4])>0 and min(progress_c,progress_p)>0 and local and yc<0 and yp<0
    np.testing.assert_allclose(meta[5:9],[length,error,le,lb],rtol=1e-6,atol=1e-14)
    assert bool(meta[9])==bool(resolved) and bool(meta[10])==bool(verified)
    np.testing.assert_allclose(projection,M@(face-accepted)/spacing,rtol=1e-8,atol=1e-17)
    for field,fraction in ((candidate,meta[1]),(probe,.5*meta[1])):
        target=projection@face-(1-fraction)*spacing
        assert abs(projection@field-target)*E<=max(TestBound(meta[2]),TestBound(meta[3]))+1e-12
    item=dict(level=number,fraction=float(meta[1]),candidate_lambda=float(meta[2]),probe_lambda=float(meta[3]),
        field_half_error=float(error),lambda_half_error=float(le),lambda_half_bound=float(lb),verified=bool(verified),
        before_yield_candidate_MPa=yc,before_yield_probe_MPa=yp,
        candidate_force=TestForce(candidate,meta[2],ac),probe_force=TestForce(probe,meta[3],ap),
        candidate_all_equations=TestWitness('ApproachCandidate'+str(number),candidate,meta[2]),
        probe_all_equations=TestWitness('ApproachProbe'+str(number),probe,meta[3]),phase_primary=int(meta[13]),calls_so_far=int(meta[15]))
    proofs.append(item);last=[meta,candidate,probe,projection,ac,ap]
if info[3]:
    assert info[28]==3 and proofs[-1]['verified'] and not any(x['verified'] for x in proofs[:-1])
    assert np.array_equal(returned,last[1]) and info[23]==last[0][2]
    assert info[18]==last[0][15]+1 and info[17]==last[0][13] and effort[0]==0
    assert info[20]==proofs[-1]['field_half_error'] or abs(info[20]-proofs[-1]['field_half_error'])<1e-10
    returned_force=TestForce(returned,info[23],actual);returned_witness=TestWitness('Full',returned,info[23])
    length=TestLength(returned-accepted)
    np.testing.assert_allclose(returned_projection,M@(returned-accepted)/length,rtol=1e-8,atol=1e-17)
else:
    assert not any(x['verified'] for x in proofs)
    returned_force=returned_witness=None
proof=dict(native=native,source10_exact=True,pair_roots_bitidentical791=True,returned_verified_candidate=bool(info[3]),
    selected_kind=int(info[28]),event_lambda=float(info[27]),returned_lambda=float(info[23]),
    nonlinear_calls=int(info[18]),pair_primary=int(info[16]),event_primary=int(info[17]),total_primary=int(info[19]),
    allowance=args.allowance,accepted_revision_unchanged=True,approaches=proofs,returned_force=returned_force,
    returned_all_equations=returned_witness,all_approach_candidates_and_probes_independently_verified=True,
    independent_last_state_weak_check_at_inner_tolerance=True,independent_weak_rounding_allowance_MPa=1e-12,
    no_controller_history_acceptance_claimed=True,physical_capacity_certified=False,whole_goal_completed=False)
guard=root/'Native/ArcHistoryGuard.Result.txt'
if guard.exists():
    text=guard.read_text('cp1251')
    assert 'ERROR;5;CALLS;0;REVISION_BEFORE;12;REVISION_AFTER;12' in text
    assert 'Принятая материальная история не соответствует последнему полю продолжения.' in text
    proof['mismatched_history_rejected_before_any_solve_with_clear_reason']=True
if args.allowance<8:
    assert not info[3] and info[18]<args.allowance
    proof['insufficient_budget_did_not_admit_candidate']=True
(root/'AdaptivePreEventProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'RETURNED',bool(info[3]),'LEVELS',len(records),'CALLS',info[18],'EVENT PRIMARY',info[17],'LAMBDA',info[23])
for item in proofs:print('LEVEL',item['level'],'FIELD',item['field_half_error'],'LAMBDA_ERROR',item['lambda_half_error'],'BOUND',item['lambda_half_bound'],'VERIFIED',item['verified'])
