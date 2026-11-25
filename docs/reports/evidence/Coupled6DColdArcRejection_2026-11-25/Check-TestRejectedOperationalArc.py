"""ДЛЯ ТЕСТОВ. Независимо проверяет конкретную отвергнутую дугу, её событие и разрешимость из фактических полей и напряжений."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,default=base/'Cold789RejectedOperationalArc_20261125');args=p.parse_args();root=args.root
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==10
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](base/'Cold787ArcContextOrder_20261125/Native/Stock.RawHistory.bin')
for k in ('fields','fractions','lambdas'):assert np.array_equal(source[k],frames[-1][k][:10])
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((root/'Native/OperationalArc.Result.bin').read_bytes());assert r.raw[:16]==b'NDMOPERARC000002\0'[:16]
packet=[r.VectorTest() for _ in range(15)]
info,projection,accepted,half_seed,full_seed,half,full,ah,af,event,face,face_half,unused,ah_face,parameters=packet
returned=actual=returned_projection=None
if info[3]:returned,actual,returned_projection=[r.VectorTest() for _ in range(3)]
effort=r.VectorTest();assert r.position==len(r.raw)
assert info[14]==info[15]==event[6]==12 and np.array_equal(accepted,source['fields'][-1])
assert info[16]<=30 and info[17]<=30 and info[18]<=6 and info[19]==info[16]+info[17]
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,b,gradient=ops
op=np.array([b+s*.001*gradient for s in (-1,0,1)])
M=np.einsum('spcd,spce,p,c->de',op,op,coords[:,2],np.array([1.,1.,1.,.5,.5,.5]))/(3*sum(coords[:,2]));M[56:]=0;M[:,56:]=0


def TestLength(delta):
    """Измеряет физическую длину в общей трёхстанционной метрике."""
    return float(np.sqrt(max(0.,delta@M@delta)))


def TestBound(factor):
    """Восстанавливает фактическую внутреннюю точность при данном множителе."""
    return float(.0001*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56]))))


witnesses={}
for label,flag,field,factor in (('PairFull',info[0],full,info[4]),('PairHalf',info[1],half,info[5]),('Event',info[26],face,info[23]),
    ('EventHalf',int(len(face_half)==60),face_half,info[24]),('Full',info[3],returned,info[23])):
    if not flag:continue
    mechanics,fields=validate['TestValidate'](root,'OperationalArc.'+label+'.Sequence')
    assert len(fields)==11
    np.testing.assert_allclose(fields[:10],source['fields'],rtol=1e-14,atol=1e-300)
    np.testing.assert_allclose(fields[10],field,rtol=1e-14,atol=1e-300)
    witnesses[label]=mechanics
spacing=half_resolution=required=event_error=event_yield=half_yield=None
if info[26]:
    assert event[0] and len(parameters)==10
    point,side=int(event[1]),int(event[2]);alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:]
    E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));shear=E/(2*(1+nu))
    D=np.diag([off+2*shear]*3+[shear]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3))

    def TestYield(field):
        """Вычисляет пробную поверхность из исходной принятой пластической истории, без повторного материального возврата."""
        xx,yy,zz,xy,xz,yz=D@((b[point-1]+side*.001*gradient[point-1])@field-plastic)
        tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=xx+yy+zz;high=np.linalg.eigvalsh(tensor)[-1]
        q=np.sqrt(1.5*np.sum((tensor-np.eye(3)*tr/3)**2))
        term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
        return float((alpha*tr+q+term)/(1-alpha)-cc)

    np.testing.assert_allclose([TestYield(accepted),TestYield(full if info[0] else half)],event[3:5],rtol=1e-5,atol=2e-12)
    event_yield=TestYield(face);assert abs(event_yield)<=2.1e-12
    spacing=TestLength(face-accepted);half_resolution=.5*spacing*E;required=8*max(TestBound(info[23]),TestBound(info[24]))
    if len(face_half)==60:
        event_error=TestLength(face-(2*face_half-accepted))/spacing;half_yield=TestYield(face_half)
        np.testing.assert_allclose(event_error,info[25],rtol=1e-6,atol=1e-14)
    if info[3]:assert half_resolution>required and event_error<=.005 and half_yield<0
source_rows=validate['TestRows'](root/'Native/EnvelopeEnergy.Accepted.csv');native_source=[x for x in source_rows if x[0]=='R'][-1]
source_norm=float(native_source[2]);source_bound=TestBound(info[6])
proof=dict(native=native,exact_cold_prefix10_785_787=True,nonlinear_calls=int(info[18]),pair_primary=int(info[16]),event_primary=int(info[17]),
    returned_verified_candidate=bool(info[3]),event_selected=event.tolist(),source_parameters=parameters.tolist(),
    pair_field_error=float(info[20]),event_solved=bool(info[26]),event_field_error=event_error,
    event_lambda=float(info[23]),event_yield_MPa=event_yield,half_event_yield_MPa=half_yield,
    event_spacing=spacing,event_half_resolution_MPa=half_resolution,event_required_half_resolution_MPa=required,
    event_resolved_at_inner_tolerance=(bool(half_resolution>required) if info[26] else None),
    source_native_norm_MPa=source_norm,source_inner_bound_MPa=source_bound,source_already_passes_inner_tolerance=source_norm<=source_bound,
    independent_raw_stress_weak_and_six_force_witnesses=witnesses,
    physical_capacity_certified=False,whole_goal_completed=False)
if root.name.startswith('Cold789'):
    assert not info[3] and event[1]==29 and event[2]==-1 and event_error<=.005 and half_resolution<required
    assert source_norm<source_bound
    proof['actual_rejection_due_to_unresolved_near_zero_event_confirmed']=True
(root/'RejectedOperationalArcProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'VERIFIED RETURN',bool(info[3]),'EVENT',event[1:3].tolist(),'HALF RESOLUTION',half_resolution,'REQUIRED',required,'SOURCE INNER OK',source_norm<=source_bound)
