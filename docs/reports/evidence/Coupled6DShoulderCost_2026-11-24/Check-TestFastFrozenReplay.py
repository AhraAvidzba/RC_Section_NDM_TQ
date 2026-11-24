"""ДЛЯ ТЕСТОВ. Проверяет все71 замороженных поля732 и диагностирует фактический критерий роста постпикового шага."""
from pathlib import Path
from runpy import run_path
import json,re
import numpy as np

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold733FastFrozenReplay_20261124')
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert 'sourceStates=71; mode=0;' in text
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/PostCrackSourceRawHistory.bin');frame=frames[-1]
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
data=(root/'Input/PostCrackReplay.bin').read_bytes();assert data[:16]==b'NDMPOSTREPLAY001'
reader=Reader(data);size,=reader.ReadTest('<i');assert size==60
metadata={k:reader.VectorTest() for k in ('calibration','options','offset','direction')}
count,=reader.ReadTest('<i');assert count==len(frame['fields'])==71
for i in range(count):
    fraction,factor=reader.ReadTest('<dd');field=reader.VectorTest()
    assert fraction==frame['fractions'][i] and factor==frame['lambdas'][i] and np.array_equal(field,frame['fields'][i])
assert reader.position==len(data)
phase='PostCrackRawReplay.Accepted'
fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,phase)
mod=run_path(str(here/'Validate-TestHardPrepared.py'));mechanics,actual=mod['TestValidate'](root,phase)
assert fields['states']==71 and len(actual)==71
np.testing.assert_allclose(np.asarray(actual),frame['fields'],rtol=1e-14,atol=1e-300)
for p,l in zip(mechanics['states'],frame['lambdas']):assert abs(p['lambda']-l)<=1e-14*(1+abs(l))
source=root.parent/'Cold732ProvidedJacobianCapacity_20261124'
assert (root/'Input/PostCrackSourceRawHistory.bin').read_bytes()==(source/'Native/Stock.RawHistory.bin').read_bytes()
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root/name).read_bytes()==(source/name).read_bytes(),name
name='src/Solver/CCoupledNewtonSolver.cls'
code=(root/name).read_text('utf-8-sig').replace('        modTestCoupledElasticState.TestObserveFrozenReplay index, field\n','')
assert code==(source/name).read_text('utf-8-sig')
analytical=run_path(str(here/'Check-TestShoulderAccuracy.py'))['TestAnalyticalForces'](root,phase,fields)

# Тензорная метрика совпадает с ArcPhysicalSquare: среднее трёх станций с площадями и половинными весами инженерных сдвигов.
points=fields['points'];dofs=fields['dofs'];physical=fields['physical_dofs']
b=np.zeros((points,6,dofs));gradient=b.copy();areas=np.zeros(points)
for kind,*row in mod['TestRows'](root/'Native'/(phase+'.Operators.csv')):
    if kind=='B':
        p,d=int(row[0])-1,int(row[1])-1;b[p,:,d]=np.asarray(row[2:8],float);gradient[p,:,d]=np.asarray(row[8:],float)
    elif kind=='P':areas[int(row[0])-1]=float(row[3])
weights=np.array([1.,1.,1.,.5,.5,.5])
def TestDistance(delta):
    """Восстанавливает физическую норму без весов напряжений, жёсткостей и служебных множителей."""
    d=np.zeros(dofs);d[:physical]=delta[:physical]
    c=np.einsum('pcd,d->pc',b,d);g=np.einsum('pcd,d->pc',gradient,d)
    strains=np.asarray([c+s*.001*g for s in (-1,0,1)])
    return float(np.sqrt(np.einsum('spc,p,c->',strains**2,areas,weights)/(3*sum(areas))))

growth=[];f=frame['fields'];l=frame['lambdas'];rel=float(metadata['options'][4])
for i in range(2,len(f)):
    ratio=(f[i,0]-f[i-1,0])/(f[i-1,0]-f[i-2,0])
    predicted=ratio*(l[i-1]-l[i-2]);load_error=abs(l[i]-l[i-1]-predicted)
    load_bound=max(abs(predicted)*.02,8*rel*(1+abs(l[i-1])))
    expected=ratio*(f[i-1]-f[i-2]);actual=f[i]-f[i-1];length=TestDistance(expected)
    field_error=TestDistance(actual-expected)/length if length>0 else None
    growth.append(dict(index=i+1,epsilon=float(f[i,0]),increment=float(f[i,0]-f[i-1,0]),
        reconstructed_consecutive_history_ratio=float(ratio),load_prediction_error=float(load_error),
        original_load_prediction_bound=float(load_bound),load_prediction_pass=bool(load_error<=load_bound),
        field_prediction_relative_error=field_error,original_field_prediction_bound=.02,
        field_prediction_pass=field_error is not None and field_error<=.02))
proof=dict(native=native,original_binary_input_and_exported_fields71_match=True,source_raw_history=raw,
    source_numeric_kernel_exact_except_readonly_observer=True,independent_fields=fields,analytical_forces=analytical,
    analytical_force_misses=[p for p in analytical if p['error_N']>.5],
    consecutive_final_history_growth_diagnosis=growth,
    no_new_continuation_or_capacity_search=True,physical_capacity_certified=False,entire_goal_completed=False)
(root/'FastFrozenReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('733 ALL71 EXACT FIELDS AND INDEPENDENT EQUILIBRIUM VERIFIED; ANALYTICAL MISSES',len(proof['analytical_force_misses']))
for p in growth[-7:]:print('TAIL',p['index'],'DE',p['increment'],'LOAD_OK',p['load_prediction_pass'],'FIELD_ERROR',p['field_prediction_relative_error'],'FIELD_OK',p['field_prediction_pass'])
