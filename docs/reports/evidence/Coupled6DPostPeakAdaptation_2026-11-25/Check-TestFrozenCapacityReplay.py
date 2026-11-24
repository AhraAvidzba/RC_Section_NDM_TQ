"""ДЛЯ ТЕСТОВ. Независимо проверяет все поля завершённого поиска и сохраняет аналитические расхождения."""
from argparse import ArgumentParser
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np

p=ArgumentParser(description=__doc__);p.add_argument('--root',required=True,type=Path);a=p.parse_args();root=a.root
here=Path(__file__).resolve().parent;scope=json.loads((root/'Preparation.json').read_bytes())['directed_scope']
source=Path(scope['source_root']);count=scope['source_states']
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
assert f'sourceStates={count}; mode=0;' in text
rawfile=root/'Input/PostCrackSourceRawHistory.bin'
assert rawfile.read_bytes()==(source/'Native/Stock.RawHistory.bin').read_bytes()
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](rawfile);frame=frames[-1]
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];data=(root/'Input/PostCrackReplay.bin').read_bytes()
assert data[:16]==b'NDMPOSTREPLAY001';reader=Reader(data);size,=reader.ReadTest('<i');assert size==60
metadata={k:reader.VectorTest() for k in ('calibration','options','offset','direction')}
n,=reader.ReadTest('<i');assert n==count==len(frame['fields'])
for i in range(n):
    fraction,factor=reader.ReadTest('<dd');field=reader.VectorTest()
    assert fraction==frame['fractions'][i] and factor==frame['lambdas'][i] and np.array_equal(field,frame['fields'][i])
assert reader.position==len(data)
phase='PostCrackRawReplay.Accepted'
fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,phase)
mechanics,actual=run_path(str(here/'Validate-TestHardPrepared.py'))['TestValidate'](root,phase)
assert fields['states']==len(actual)==count
np.testing.assert_allclose(np.asarray(actual),frame['fields'],rtol=1e-14,atol=1e-300)
for point,factor in zip(mechanics['states'],frame['lambdas']):assert abs(point['lambda']-factor)<=1e-14*(1+abs(factor))
for name in ('src/Materials/CCoupledConcreteResponse.cls','src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root/name).read_bytes()==(source/name).read_bytes(),name
name='src/Solver/CCoupledNewtonSolver.cls'
code=(root/name).read_text('utf-8-sig').replace('        modTestCoupledElasticState.TestObserveFrozenReplay index, field\n','')
assert code==(source/name).read_text('utf-8-sig')
analytical=run_path(str(here/'Check-TestShoulderAccuracy.py'))['TestAnalyticalForces'](root,phase,fields)
proof=dict(native=native,source_root=str(source),source_states=count,source_raw_history=raw,
    original_binary_input_and_exported_fields_match=True,all_fields_independent_weak_gauge_and_six_force_checks=True,
    source_numeric_kernel_exact_except_readonly_observer=True,independent_fields=fields,
    analytical_forces=analytical,analytical_force_misses=[point for point in analytical if point['error_N']>.5],
    no_new_continuation_or_capacity_search=True,physical_capacity_certified=False,entire_goal_completed=False)
(root/'FrozenCapacityReplayProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('ALL',count,'EXACT FIELDS INDEPENDENTLY VERIFIED; ANALYTICAL MISSES:',len(proof['analytical_force_misses']))
