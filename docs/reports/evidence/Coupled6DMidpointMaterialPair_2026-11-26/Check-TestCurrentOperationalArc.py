"""ДЛЯ ТЕСТОВ. Проверяет существующую дугу, её собственный прогноз, фактические равновесия и общий сертификат без будущего корня."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=base/'Cold962CurrentOperationalArc_20261126')
parser.add_argument('--changed-arc',action='store_true')
parser.add_argument('--accuracy-scale',type=float,choices=[.0001,1.],default=.0001)
a=parser.parse_args();root=a.root;parent=base/'Cold960ResolvedFieldProgress_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
if not a.changed_arc:assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for folder in ('src','tests'):
    for path in (parent/folder).rglob('*'):
        if path.is_file() and not (a.changed_arc and path.name=='CCoupledArcContinuation.cls'):
            assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));baseline,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(fields)==25
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Native/OperationalArc.Result.bin').read_bytes());assert r.raw[:16]==b'NDMOPERARC000003'
packet=[r.VectorTest() for _ in range(15)]
info,projection,accepted,half_seed,full_seed,half,full,ah,af,event,face,face_half,unused,ah_face,parameters=packet
assert len(info)==34 and info[14]==info[15] and info[16]<=30 and info[17]<=30 and info[18]<=96
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert info[6]==source['lambdas'][-1]
returned=returned_projection=actual=None
if info[3]:returned,actual,returned_projection=[r.VectorTest() for _ in range(3)]
effort=r.VectorTest();assert r.position==len(r.raw)
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]

def TestLength(delta):
    """Восстанавливает тензорную меру всех точек трёх станций, исключая служебные множители связей."""
    strain=np.einsum('spcd,d->spc',op,delta)
    return float(np.sqrt(np.einsum('spc,spc,p,c->',strain,strain,coords[:,2],np.array([1,1,1,.5,.5,.5]))/(3*sum(coords[:,2]))))

def TestBound(lam):
    """Восстанавливает фактический допуск по явно заданному масштабу внутренней точности дуги."""
    return float(a.accuracy_scale*(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+lam*source['direction']))[:56]))))

def TestPairBound(length,error,lam,reference,lambda_error):
    """Независимо вычисляет совместную границуλ из неизменённой постоянной метрики, включая её внутренний шум."""
    alpha=info[9];radius=float(np.hypot(length,alpha*(lam-reference)))
    allowance=(.005*radius)**2-(length*error)**2
    return max(8*a.accuracy_scale*source['options'][4]*(1+abs(reference)),float(np.sqrt(allowance)/alpha) if allowance>0 else 0.)

def TestWitness(label,field,lam,arc_target=None):
    """Проверяет уравнения и сталь по фактической экспортированной истории, а для пар — также исходную дуговую строку."""
    proof,fs=validate['TestValidate'](root,'OperationalArc.'+label+'.Sequence')
    assert len(fs)==26
    np.testing.assert_allclose(fs[:25],source['fields'],rtol=1e-14,atol=1e-300)
    np.testing.assert_allclose(fs[-1],field,rtol=1e-14,atol=1e-300)
    result=proof['states'][-1];assert abs(result['lambda']-lam)<1e-12
    assert result['residual_MPa']<=TestBound(lam)+2e-12 and result['force_residual_MPa']<=TestBound(lam)+2e-12
    if arc_target is not None:
        defect=float(E*(projection@field+info[12]*lam-arc_target));assert abs(defect)<=TestBound(lam)+2e-12
        result['independent_arc_control_residual_MPa']=defect
    requested=TestBound(lam)/a.accuracy_scale
    actual_norm=max(result['residual_MPa'],result['force_residual_MPa'],abs(result.get('independent_arc_control_residual_MPa',0.)))
    result.update(original_requested_bound_MPa=requested,
        passes_original_requested_force_accuracy=bool(actual_norm<=requested+2e-12),
        passes_additional_1e_minus4_force_accuracy=bool(actual_norm<=.0001*requested+2e-12))
    return result

def TestSurface(field):
    """Восстанавливает пробную поверхность Барселоны от фактической принятой истории выбранной станции."""
    assert len(event)==7 and len(parameters)==10
    point,side=int(event[1]),int(event[2]);alpha,gamma=parameters[:2];plastic=parameters[2:8];cc,ct=parameters[8:]
    nu=source['calibration'][1];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu));D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
    xx,yy,zz,xy,xz,yz=D@(op[side+1,point-1]@field-plastic)
    tensor=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);trace=float(np.trace(tensor));high=float(np.linalg.eigvalsh(tensor)[-1])
    q=float(np.sqrt(1.5*np.sum((tensor-np.eye(3)*trace/3)**2)))
    term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
    return float((alpha*trace+q+term)/(1-alpha)-cc)

witnesses={}
for label,flag,field,lam,target in [('PairFull',info[0],full,info[4],info[13]+info[11]),
 ('PairHalf',info[1],half,info[5],info[13]+.5*info[11]),('Event',info[26],face,info[27],None),
 ('EventHalf',len(face_half)==60,face_half,info[24],None),('Full',info[3],returned,info[23],None)]:
    if flag:witnesses[label]=TestWitness(label,field,lam,target)
surface=None
if info[26]:
    value=TestSurface(face);assert abs(value)<=2e-10
    assert abs(TestSurface(accepted)-event[3])<2e-10
    surface=dict(point=int(event[1]),side=int(event[2]),yield_MPa=value,lambda_value=float(info[27]))
pair=None
if info[0] and info[1]:
    length=TestLength(full-accepted);error=TestLength(full-(2*half-accepted))/length
    le=float(abs(2*info[5]-info[6]-info[4]));lb=TestPairBound(length,error,info[4],info[6],le)
    np.testing.assert_allclose(info[21:23],[le,lb],rtol=1e-5,atol=1e-14)
    pair=dict(field_error=error,lambda_error=le,lambda_bound=lb)
reader=R((root/'Native/OperationalArc.Approaches.bin').read_bytes());assert reader.raw[:16]==b'NDMARCAPPR000001'
count=int(reader.VectorTest()[0]);assert count==info[29]
approaches=[]
for index in range(1,count+1):
    meta,candidate,probe,p,ac,ap=[reader.VectorTest() for _ in range(6)]
    assert meta[0]==index and meta[11]==meta[12]==info[14] and meta[13]<=30 and meta[15]<96
    item=dict(meta=meta.tolist(),candidate=TestWitness('ApproachCandidate'+str(index),candidate,meta[2]))
    if len(probe)==60:
        item['probe']=TestWitness('ApproachProbe'+str(index),probe,meta[3])
        face_spacing=TestLength(face-accepted);origin=float(p@accepted)
        full_control=float(E*(p@candidate-origin-meta[1]*face_spacing))
        half_control=float(E*(p@probe-origin-.5*meta[1]*face_spacing))
        assert abs(full_control)<=TestBound(meta[2])+2e-11 and abs(half_control)<=TestBound(meta[3])+2e-11
        length=TestLength(candidate-accepted);error=TestLength(candidate-(2*probe-accepted))/length
        le=float(abs(2*meta[3]-info[6]-meta[2]));lb=TestPairBound(length,error,meta[2],info[6],le)
        np.testing.assert_allclose(meta[5:9],[length,error,le,lb],rtol=1e-5,atol=1e-14)
        if meta[10]:assert error<=.005 and le<=lb and meta[9]
        item.update(field_error=error,lambda_error=le,lambda_bound=lb,
            candidate_yield_MPa=TestSurface(candidate),probe_yield_MPa=TestSurface(probe),
            candidate_control_residual_MPa=full_control,probe_control_residual_MPa=half_control)
        resolved=.5*length*E>8*max(TestBound(meta[2]),TestBound(meta[3]));assert bool(meta[9])==resolved
        if meta[10]:
            assert item['candidate_yield_MPa']<0 and item['probe_yield_MPa']<0
            assert length<=1.5*meta[1]*face_spacing+1e-15 and TestLength(probe-accepted)<=.75*meta[1]*face_spacing+1e-15
            for field,lam in ((candidate,meta[2]),(probe,meta[3])):
                assert projection@(field-accepted)+info[12]*(lam-info[6])>0
    approaches.append(item)
assert reader.position==len(reader.raw)
if info[3]:
    assert info[2] and effort[0]==0 and 'Full' in witnesses
    np.testing.assert_allclose(actual,source['offset']+info[23]*source['direction'],rtol=0,atol=1e-6)
    if info[28]==3:assert approaches and approaches[-1]['meta'][10]
proof=dict(native=native,all257_actual_components_and_saved_book_bitexact960=not a.changed_arc,
    actual25_history_unchanged_inside_arc_before_diagnostic_exports=True,no_supplied_future_guide_or_field=True,existing_OperationalArc_test_reused=True,
    internal_accuracy_scale0001_and_current_acceptance_unchanged=a.accuracy_scale==.0001,
    internal_accuracy_scale=a.accuracy_scale,info=info.tolist(),effort=effort.tolist(),
    returned_verified_candidate=bool(info[3]),pair=pair,event=surface,approaches=approaches,witnesses=witnesses,
    pair_steps=int(info[16]),event_phase_steps=int(info[17]),total_primary_steps=int(info[19]),calls=int(info[18]),
    standalone_call_allowance=96,allowance_not_remaining_original_cold_budget=True,
    no_material_history_acceptance_inside_arc=True,physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
(root/'CurrentOperationalArcProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root.name,'RETURNED',bool(info[3]),'PAIR FLAGS',info[:2].tolist(),'EVENT',surface,'PHASE STEPS',info[16:20].tolist(),'APPROACHES',len(approaches))
for name,result in witnesses.items():print(name,'LAMBDA',result['lambda'],'R',result['residual_MPa'],'BOUND_INTERNAL',TestBound(result['lambda']))
