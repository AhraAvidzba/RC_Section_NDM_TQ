"""ДЛЯ ТЕСТОВ. Проверяет воспроизведение фактической системы54/55 и все найденные корни двух режимов резерва."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold925RecordedTailSystem_20261126';parent=base/'Cold922ColdNewtonCapture_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==257
assert audit['imported_this_pass']==['modTestCoupledElasticState'] and len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
for file in (parent/'src').rglob('*'):
    if file.is_file():assert file.read_bytes()==(root/file.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,movement=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);expected=np.array([r.VectorTest() for _ in range(61)]);assert r.position==len(r.raw)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin');assert len(source['fields'])==25
assert source['fields'][-1].tobytes()==accepted.tobytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted')
assert len(fields)==25
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)])
helper=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig')
begin=helper.index('def TestAssemble(');end=helper.index('summary=[];',begin)
exec(compile(helper[begin:end],str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())
r=R((root/'Native/RecordedTailSystem.bin').read_bytes());assert r.raw[:16]==b'NDMTAILSYSTEM001'
scales=r.VectorTest();values=r.VectorTest();A=np.array([r.VectorTest() for _ in range(61)])
raw=values[:60].copy();raw[56:]/=meta[4]
base_error=float(max(abs(raw-residual)));assert base_error<1e-11
rowmax=np.maximum(np.max(abs(expected),axis=1),np.finfo(float).tiny)
matrix_error=float(max((abs(A-expected)/rowmax[:,None]).ravel()));assert matrix_error<2e-12
control_error=float(abs(values[-1]+rhs[-1]));assert control_error<1e-11
cases=[]
for index in (1,2):
    info,effort=r.VectorTest(),r.VectorTest();assert len(info)==6 and len(effort)==8 and info[0]==index
    assert info[1]==(source['lambdas'][-1] if index==1 else 0) and effort[0]<=source['options'][2]
    item=dict(case=index,reference_lambda=float(info[1]),solved=bool(info[2]),lambda_value=float(info[3]),
        internal_status=int(info[4]),result_code=int(info[5]),effort=effort.tolist())
    if info[2]:
        field,actual=r.VectorTest(),r.VectorTest();stress=np.zeros((3,43,6));seen=set()
        for _ in range(129):
            point,side,code=r.VectorTest();sigma=r.VectorTest();assert code==0 and len(sigma)==6
            key=int(point),int(side);assert key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
        assembled=TestAssemble(field,info[3],stress)
        control=-rhs[-1]+original[-1,:60]@(field-current)+original[-1,-1]*(info[3]-meta[2])
        bound=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+info[3]*source['direction']))[:56])))
        norm=max(assembled['physical_norm'],assembled['six_force_norm'],abs(control),meta[4]*max(abs(assembled['residual'][56:])))
        strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
        util=float(max((abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)).ravel()))
        assert norm<=bound+1e-12 and util<=1
        normscales=np.array([1/L[0,0],1/L[1,1],1/L[1,1],1/L[1,1],1/L[0,0],1/L[0,0]])
        assert max(abs((actual-source['offset']-info[3]*source['direction'])/normscales))<=bound
        item.update(independently_equilibrated=True,total_residual_MPa=float(norm),bound_MPa=bound,
            maximum_three_station_steel_utilization=util,scalar_control_residual_MPa=float(control))
    cases.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=257,other256_components_preserved=True,
    exact_recorded_system_reproduced_before_attempts=True,base_residual_max_error_MPa=base_error,
    refined_matrix_max_relative_row_error=matrix_error,control_max_error_MPa=control_error,cases=cases,
    reference_modes_change_full_recovery_algorithm_not_only_lambda_guard=True,
    accepted_history_unchanged_by_native_guard=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'RecordedTailSystemProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('EXACT CURRENT REPRODUCED;R ERROR',base_error,'REFINED A ERROR',matrix_error,'CONTROL ERROR',control_error)
for item in cases:print('CASE',item['case'],'REF',item['reference_lambda'],'SOLVED',item['solved'],'ITER',item['effort'][0],
    'NATIVE R',item['effort'][-1],'INDEPENDENT',item.get('total_residual_MPa'))
