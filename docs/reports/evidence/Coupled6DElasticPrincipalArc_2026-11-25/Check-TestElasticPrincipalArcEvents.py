"""ДЛЯ ТЕСТОВ. Независимо проверяет фактические спектральные события дуговой ветви по сохранённой пластической истории выбранной станции."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold838ElasticPrincipalArc_20261125')
walk=json.loads((root/'AutomaticLoadingWalkProof.json').read_bytes());assert walk['spectral_event_integration']
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/LoadingWalk.Accepted.Sequence.Operators.csv',60,43)
rows=run_path(str(here/'Validate-TestHardPrepared.py'))['TestRows'](root/'Native/LoadingWalk.Accepted.Sequence.csv')
native_stress={(int(row[1]),int(row[2]),int(row[3])):np.array(row[4:10],float) for row in rows if row[0]=='S'}
E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))
def TestTensor(s):
    """Разворачивает фактические компоненты напряжений в симметричный тензор."""
    xx,yy,zz,xy,xz,yz=s
    return np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]])
events=[]
for attempt in walk['attempts']:
    r=Reader((root/'Native'/f"LoadingWalk.{attempt['attempt']:03}.bin").read_bytes());packet=[r.VectorTest() for _ in range(15)]
    info,event=packet[0],packet[9]
    if len(event)<9 or event[8]!=1:continue
    assert attempt['solved'] and attempt['kind']==2 and info[26]==1
    field,actual,projection=[r.VectorTest() for _ in range(3)]
    point,side,principal=map(int,(event[1],event[2],event[7]));state=packet[14];assert len(state)==10
    alpha,gamma=state[:2];cc,ct=state[8:];operator=b[point-1]+side*.001*gradient[point-1]
    def TestTrial(field):
        """Поверхность и спектр вычисляются независимо от VBA из принятой пластической деформации выбранной станции."""
        sigma=D@(operator@field-state[2:8]);t=TestTensor(sigma);ev=np.linalg.eigvalsh(t);tr=float(np.trace(t))
        q=np.sqrt(1.5*np.sum((t-np.eye(3)*tr/3)**2));term=(cc/ct*(1-alpha)-(1+alpha))*ev[-1] if ev[-1]>0 else gamma*ev[-1]
        phi=float((alpha*tr+q+term)/(1-alpha)-cc)
        return float(ev[3-principal]),phi
    bracket=packet[6] if info[0] else (packet[5] if info[1] else packet[4])
    start,phi0=TestTrial(packet[2]);end,phi1=TestTrial(bracket)
    value,phi=TestTrial(packet[10]);half,phi_half=TestTrial(packet[11])
    level,phi_level=TestTrial(packet[2]+event[5]*(bracket-packet[2]))
    assert max(phi0,phi1,phi,phi_half,phi_level)<0
    assert start*end<0 and half*start>0
    assert abs(start-event[3])<=2e-10 and abs(end-event[4])<=2e-10
    assert abs(value)<=6e-12 and abs(level)<=6e-12
    assert np.array_equal(field,packet[10])
    accepted_index=attempt['source_count']+1
    actual_zero=float(np.linalg.eigvalsh(TestTensor(native_stress[accepted_index,side,point]))[3-principal])
    assert abs(actual_zero)<=6e-12
    events.append(dict(attempt=attempt['attempt'],accepted_state=accepted_index,point=point,side=side,principal=principal,
        lambda_value=attempt['lambda_value'],trial_zero_MPa=value,actual_zero_MPa=actual_zero,
        trial_yield_MPa=phi,half_yield_MPa=phi_half,approach_sign_correct=True))
assert len(events)==7
assert {(e['point'],e['principal']) for e in events}=={(3,1),(8,1),(31,3)}
proof=dict(events=events,all_events_independently_rebuilt=True,all_endpoints_and_faces_elastic=True,
    all_half_roots_on_source_side=True,all_actual_stress_zeros_verified=True,
    different_points_and_principals_selected_without_hardcoding=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'ElasticPrincipalArcEventsProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('ALL7 SPECTRAL EVENTS VERIFIED;POINTS3/8/31;PRINCIPALS1/3;BOTH SIGN DIRECTIONS;NO PHYSICAL LIMIT')
