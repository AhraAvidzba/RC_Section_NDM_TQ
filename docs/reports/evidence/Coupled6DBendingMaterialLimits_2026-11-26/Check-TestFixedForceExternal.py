"""ДЛЯ ТЕСТОВ. Независимо проверяет фиксированные силовые пробы Newton/SciPy, все56+4 уравнения и шесть результирующих без вывода о разрушении."""
from argparse import ArgumentParser
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct,numpy as np

here=Path(__file__).resolve().parent;parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,required=True);root=parser.parse_args().root
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
case=json.loads((root/'Input/FixedForceCase.json').read_bytes());result=json.loads((root/'Native/FixedForceExternalResult.json').read_bytes())
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert result['source_unchanged'] and result['source_sha256']==compiled['WorkbookSHA256']==sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()
assert result['own_excel_final_state']['state'] in ('Absent','Exited','Replaced')
assert result['repeat_response_bitexact'] and result['accepted_history_unchanged_bitexact']
assert (root/'Native/FixedForceAcceptedBefore.bin').read_bytes()==(root/'Native/FixedForceAfterNative.bin').read_bytes()==(root/'Native/FixedForceAcceptedAfter.bin').read_bytes()
assert case['target_lambda']==case['known_peak_lambda']*case['target_ratio_to_known_peak']
assert source['lambdas'][np.argmax(source['lambdas'])]==case['known_peak_lambda']
assert 0<=case['accepted_prefix']<=len(source['fields']) and 1<=case['seed_index']<=len(source['fields'])
g,L,w,v,coords,axis,B,gradient=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Operators.csv',60,43)
op=np.array([B+s*.001*gradient for s in (-1,0,1)]);E=source['calibration'][0]
code=(here/'Check-TestCold32OutgoingStressScan.py').read_text('utf-8-sig');begin=code.index('def TestAssemble(');end=code.index('summary=[];',begin)
body=code[begin:end].replace('return dict(residual=residual,concrete=concrete','return dict(actual_forces=actual,force_scales=scales,six_force_residual=(actual-target)/scales,residual=residual,concrete=concrete',1)
exec(compile(body,str(here/'Check-TestCold32OutgoingStressScan.py'),'exec'),globals())

def TestStress(r):
    """Читает фактические тензоры129 бетонных точек всех трёх станций."""
    stress=np.zeros((3,43,6));seen=set()
    for _ in range(129):
        point,side,code=r.VectorTest();sigma=r.VectorTest();key=int(point),int(side)
        assert code==0 and key not in seen;seen.add(key);stress[int(side)+1,int(point)-1]=sigma
    return stress

def TestMechanics(field,stress):
    """Повторяет слабые уравнения и интегрирование усилий по исходным операторам и неизменному материальному отклику."""
    factor=case['target_lambda'];assembled=TestAssemble(field,factor,stress)
    bound=float(source['options'][3]+source['options'][4]*max(abs((L@(source['offset']+factor*source['direction']))[:56])))
    norm=float(max(assembled['physical_norm'],assembled['six_force_norm'],E*max(abs(assembled['residual'][56:]))))
    strain=np.einsum('spcd,d->spc',op,field);bars=np.einsum('spc,pc->sp',strain[:,25:],axis[25:]);long=axis[25:,2]==1
    utilization=float(np.max(abs(bars)/np.where(long[None,:],np.where(bars>=0,.025,.018),.02)))
    return dict(lambda_value=factor,residual_MPa=norm,bound_MPa=bound,physical_residual_MPa=assembled['physical_norm'],
        six_force_residual_MPa=assembled['six_force_norm'],gauge_residual_MPa=float(E*max(abs(assembled['residual'][56:]))),
        steel_utilization=utilization,equilibrium=bool(norm<=bound+1e-12)),assembled

r=R((root/'Native/FixedForceNative.bin').read_bytes());assert r.raw[:16]==b'NDMFIXNATIVE0001'
meta,effort=r.VectorTest(),r.VectorTest();assert meta.tolist()==[case['target_lambda'],float(case['accepted_prefix']),meta[2]] and effort[0]<=30
native=dict(solved=bool(meta[2]),iterations=int(effort[0]),reported_residual_MPa=float(effort[7]))
if meta[2]:
    field,actual=r.VectorTest(),r.VectorTest();stress=TestStress(r);metrics,assembled=TestMechanics(field,stress)
    assert abs(metrics['residual_MPa']-effort[7])<2e-11
    assert max(abs((actual-assembled['actual_forces'])/assembled['force_scales']))<2e-11
    assert metrics['equilibrium'];native.update(metrics)
assert r.position==len(r.raw)
r=R((root/'Native/FixedForceExternal.bin').read_bytes());assert r.raw[:16]==b'NDMFIXEXTERN0001'
meta,field,values,forces=[r.VectorTest() for _ in range(4)];stress=TestStress(r);external,assembled=TestMechanics(field,stress)
assert r.position==len(r.raw) and meta[0]==case['target_lambda'] and meta[4]==case['accepted_prefix']
assert max(abs(values[:56]-assembled['residual'][:56]))<2e-11
assert max(abs(values[56:60]-E*assembled['residual'][56:]))<2e-11
assert max(abs(forces-assembled['six_force_residual']))<2e-11
assert abs(meta[1]-external['bound_MPa'])<1e-12*meta[1] and abs(meta[2]-external['residual_MPa'])<2e-11
assert bool(result['externally_found_equilibrium'])==external['equilibrium']
seed=source['fields'][case['seed_index']-1]*(case['target_lambda']/source['lambdas'][case['seed_index']-1])
expected=seed+np.asarray(result['coordinate_scales'][:60])*np.asarray(result['coordinates'])
np.testing.assert_allclose(field,expected,rtol=1e-14,atol=1e-18)
trace=(root/'Native/FixedForceFTrace.bin').read_bytes();assert trace[:16]==b'NDMFIXFTRACE0001' and struct.unpack_from('<i',trace,16)[0]==60
size=16+16*60;assert (len(trace)-20)%size==0
count=(len(trace)-20)//size;last_elapsed=-1.
for index in range(count):
    at=20+index*size;number,elapsed=struct.unpack_from('<Qd',trace,at);assert number==index+1 and elapsed>=last_elapsed;last_elapsed=elapsed
    x=np.frombuffer(trace,dtype='<f8',count=60,offset=at+16);f=np.frombuffer(trace,dtype='<f8',count=60,offset=at+16+8*60)
    assert np.all(np.isfinite(x)) and np.all(np.isfinite(f))
assert count==result['actual_function_calls']
np.testing.assert_array_equal(x,np.asarray(result['coordinates']));np.testing.assert_array_equal(f,np.asarray(result['final_residual']))
proof=dict(case=case,native=native,external=external,external_method=result.get('method','trf'),optimizer_invoked=result['optimizer_invoked'],optimizer_success=result.get('optimizer_success'),
    optimizer_message=result['optimizer_message'],external_function_calls=count,external_jacobians=result['actual_jacobians'],
    external_solver_seconds=result['solver_seconds'],elapsed_seconds=result['elapsed_seconds'],own_excel_closed=True,source_unchanged=True,
    every_probe_at_exact_fixed_lambda=True,accepted_history_unchanged_bitexact=True,independent_force_and_weak_equations_verified=True,
    same_material_law_not_independent_external_physics=True,professional_OpenSees_equivalence_not_tested=True,
    no_equilibrium_absence_proof=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'FixedForceExternalProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('FIXED FORCE VERIFIED',json.dumps(proof,sort_keys=True))
