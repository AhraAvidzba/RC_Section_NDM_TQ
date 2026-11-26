"""ДЛЯ ТЕСТОВ. Независимо проверяет событие961 на настоящем неравновесном прогнозе934 и неизменность257-компонентной книги."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold961FailedPredictorEvent_20261126';parent=base/'Cold960ResolvedFieldProgress_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
for folder in ('src','tests'):
    for path in (parent/folder).rglob('*'):
        if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
assert (root/'Input/Reverse1.bin').read_bytes()==(parent/'Input/Reverse1.bin').read_bytes()
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==257
origin=json.loads((root/'FailedPredictorInputOrigin.json').read_bytes())
assert sha256((root/'SourceActualRejectedNewton.bin').read_bytes()).hexdigest().upper()==origin['source_sha256']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
baseline,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(fields)==25
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
E,nu=source['calibration'][:2];G=E/(2*(1+nu));off=E*nu/((1+nu)*(1-2*nu))
D=np.diag([2*G]*3+[G]*3);D[:3,:3]+=off
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestAssemble(');end=helper.index('summary=[];',begin)
body=helper[begin:end].replace('return dict(residual=residual,concrete=concrete','return dict(six_force_residual_vector=(actual-target)/scales,residual=residual,concrete=concrete',1)
exec(compile(body,str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((base/'Cold930FineLinePlusSurface_20261126/Native/ArcScan.States.bin').read_bytes());states={}
for _ in range(129):
    point,side=r.VectorTest();states[int(point),int(side)]=r.VectorTest()
assert r.position==len(r.raw)
begin=helper.index('def TestYield(');end=helper.index('def TestAssemble(',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
r=R((root/'Input/DetectedEvent.bin').read_bytes());r.position=0
context,predictor,projection=[r.VectorTest() for _ in range(3)];assert r.position==len(r.raw)
np.testing.assert_array_equal(predictor,np.array(origin['predictor_field']))
np.testing.assert_array_equal(projection,np.array(origin['projection']))
assert context.tolist()==origin['input_context']
r=R((root/'Native/DetectedLoadingBoundary.bin').read_bytes());assert r.raw[:16]==b'NDMDETECTEV00001'
info,event,effort=[r.VectorTest() for _ in range(3)];count=int(r.VectorTest()[0]);capture=[r.VectorTest() for _ in range(count)]
assert effort[0]<=source['options'][2]
proof=dict(native=native,all257_actual_components_and_saved_book_bitexact960=True,existing_test_and_event_API_reused=True,
    starting_predictor_is_actual_failed934_iterate=True,source_residual_inf_MPa=origin['source_residual_inf_MPa'],
    no_future_equilibrium_in_consumed_event_input=True,source_history25_bitexact935=True,
    native_solved=bool(info[0]),lambda_value=float(info[1]),selected_event=event.tolist(),effort=effort.tolist(),
    capture=[c.tolist() for c in capture],independently_equilibrated=False,
    no_material_history_acceptance=True,physical_capacity_certified=False,cold_gain_claimed=False,whole_goal_completed=False)
if info[0]:
    field,values,forces=[r.VectorTest() for _ in range(3)];stress=np.zeros((3,43,6));seen=set()
    for _ in range(129):
        point,side,error=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
        assert error==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
    a=TestAssemble(field,info[1],stress)
    assert max(abs(values[:56]-a['residual'][:56]))<2e-11
    assert max(abs(values[56:60]-E*a['residual'][56:]))<2e-11
    assert max(abs(forces-a['six_force_residual_vector']))<2e-11
    norm=float(max(a['physical_norm'],a['six_force_norm'],E*max(abs(a['residual'][56:]))))
    bound=float(context[2]+context[3]*max(abs((L@(source['offset']+info[1]*source['direction']))[:56])))
    assert norm<=bound+2e-11 and abs(values[-1])<2e-11
    assert len(event)==7
    yield_value=TestYield(int(event[1]),int(event[2]),field);assert abs(yield_value)<2e-10
    strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    steel=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)));assert steel<=1
    proof.update(independently_equilibrated=True,field=field.tolist(),independent_total_residual_MPa=norm,final_bound_MPa=bound,
        independent_trial_yield_MPa=yield_value,steel_utilization=steel)
assert r.position==len(r.raw)
(root/'FailedPredictorEventProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('ACTUAL FAILED PREDICTOR EVENT',proof['native_solved'],'VERIFIED',proof['independently_equilibrated'],'LAMBDA',proof['lambda_value'],'STEPS',effort[0],'EVENT',event.tolist())
