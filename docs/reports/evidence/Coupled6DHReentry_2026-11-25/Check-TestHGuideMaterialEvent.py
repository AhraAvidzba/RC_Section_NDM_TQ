"""ДЛЯ ТЕСТОВ. Независимо проверяет первую границу нагружения и полный/половинный переход от состояния70 к H-направлению."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold841HGuideMaterialEvent_20261125';parent=base/'Cold840EnvelopeReentry_20261125'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls',
    'src/Materials/CCoupledConcreteResponse.cls','src/Materials/CCoupledReinforcedResponse.cls',
    'src/Materials/CCoupledConcreteReturn.cls','src/Solver/CCoupledMaterialConstraint.cls','src/Solver/CCoupledEnvelopeEquations.cls'):
    assert (root/name).read_bytes()==(parent/name).read_bytes(),name
old=(parent/'src/Solver/CCoupledArcContinuation.cls').read_text('utf-8-sig')
new=(root/'src/Solver/CCoupledArcContinuation.cls').read_text('utf-8-sig')
assert new.startswith(old) and new[len(old):].count('Friend Function TestMaterialEventBracket')==1
read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'];source=read(root/'Input/Reverse1.bin');full=read(parent/'Input/Reverse1.bin')
assert len(source['fields'])==70
for name in ('fields','fractions','lambdas'):np.testing.assert_array_equal(source[name],full[name][:70])
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=Reader((root/'Native/HGuideEvent.bin').read_bytes())
assert r.raw[:16]==b'NDMHFACEPAIR0001'
record,info,state,face,half,candidate,restore,effort=[r.VectorTest() for _ in range(8)];assert r.position==len(r.raw)
assert record[0]==1 and record[5]==2 and record[6]==1 and record[1]==2 and record[2]==record[3]<=30
assert restore[0]==1 and restore[1]==record[8] and restore[2]==restore[3]==record[10]==record[11]
assert effort[0]==0 and len(info)==7 and info[0]==1 and len(state)==10
np.testing.assert_array_equal(candidate,face);assert record[7]==record[8]
r=Reader((root/'Input/HMaterialGuide.bin').read_bytes());r.position=0
guide_meta,guide,arc_projection=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
r=Reader((parent/'Native/EnvelopeReentry.bin').read_bytes())
for _ in range(5):guide_source,unused,actual_guide,actual_force=[r.VectorTest() for _ in range(4)]
np.testing.assert_array_equal(guide,actual_guide);assert guide_meta[0]==guide_source[6]
cases={1:dict(field=face,lambda_value=float(record[7]),accuracy=.0001),
    2:dict(field=half,lambda_value=float(record[9]),accuracy=.0001),
    3:dict(field=candidate,lambda_value=float(record[8]),accuracy=.0001)}
mechanics,M=run_path(str(here/'TestUnacceptedMechanics.py'))['TestUnacceptedRoots'](root,source,cases)
g,L,w,v,coords,axis,b,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
point,side=map(int,info[1:3]);E,nu=source['calibration'][:2];off=E*nu/((1+nu)*(1-2*nu));G=E/(2*(1+nu))
D=np.diag([off+2*G]*3+[G]*3);D[:3,:3]+=off*(np.ones((3,3))-np.eye(3));operator=b[point-1]+side*.001*gradient[point-1]
def TestYield(field):
    """Повторяет поверхность от независимого тензора и записанного принятого пластического состояния станции."""
    alpha,gamma=state[:2];cc,ct=state[8:];s=D@(operator@field-state[2:8]);xx,yy,zz,xy,xz,yz=s
    t=np.array([[xx,xy,xz],[xy,yy,yz],[xz,yz,zz]]);tr=np.trace(t);high=np.linalg.eigvalsh(t)[-1]
    q=np.sqrt(1.5*np.sum((t-np.eye(3)*tr/3)**2));term=(cc/ct*(1-alpha)-(1+alpha))*high if high>0 else gamma*high
    return float((alpha*tr+q+term)/(1-alpha)-cc)
accepted=source['fields'][-1];factor=source['lambdas'][-1]
from_phi,to_phi,face_phi,half_phi=map(TestYield,(accepted,guide,face,half))
assert from_phi<0 and to_phi>0 and half_phi<0 and abs(face_phi)<=6e-12
assert abs(from_phi-info[3])<=2e-10 and abs(to_phi-info[4])<=2e-10
assert abs(TestYield(accepted+info[5]*(guide-accepted)))<=6e-12
def TestLength(delta):
    """Измеряет тензорное перемещение всех трёх станций без множителей жёстких связей."""
    return float(np.sqrt(max(0.,delta@M@delta)))
length=TestLength(face-accepted);field_error=TestLength(face-(2*half-accepted))/length
lambda_error=abs(2*record[9]-factor-record[7]);lambda_bound=max(.005*abs(record[7]-factor),8*.0001*source['options'][4]*(1+abs(factor)))
assert field_error<=.005 and lambda_error<=lambda_bound
assert abs(field_error-record[4])<=1e-9
progress=float(guide_meta[1]*(record[7]-factor)+arc_projection@(face-accepted))
half_progress=float(guide_meta[1]*(record[9]-factor)+arc_projection@(half-accepted))
assert progress>0 and half_progress>0
assert .5*length*E>8*max(v['bound_MPa'] for v in mechanics.values())
source_proof,fields=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields,source['fields'],rtol=1e-14,atol=1e-300)
source_h=float(guide_source[3])
assert source_h < mechanics[2]['independent_point_energy_average_MPa'] < mechanics[1]['independent_point_energy_average_MPa']
assert mechanics[1]['independent_point_energy_average_MPa']==mechanics[3]['independent_point_energy_average_MPa']
proof=dict(native=native,all70_source_equilibria_verified=source_proof,all_working_math_and_materials_exact840=True,
    selected_point=point,selected_side=side,source_lambda=float(factor),event_lambda=float(record[7]),half_lambda=float(record[9]),
    source_phi_MPa=from_phi,guide_phi_MPa=to_phi,event_phi_MPa=face_phi,half_phi_MPa=half_phi,
    field_half_error=field_error,lambda_half_error=float(lambda_error),lambda_half_bound=float(lambda_bound),
    incoming_arc_progress=progress,incoming_half_arc_progress=half_progress,mechanics=mechanics,
    event_calls=int(record[1]),primary_iterations=int(record[3]),zero_iteration_exact_restore_verified=True,
    source_H_MPa=source_h,event_reaches_new_plastic_loading_with_growing_H=True,no_candidate_history_accepted=True,
    verified_local_bridge_to_loading_boundary=True,subsequent_loading_departure_not_yet_run=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'HGuideMaterialEventProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('H-GUIDED MATERIAL EVENT VERIFIED;POINT',point,'SIDE',side,'LAMBDA',record[7],';FIELD ERROR',field_error,';EXACT RESTORE0;NO HISTORY ACCEPTED')
