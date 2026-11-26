"""ДЛЯ ТЕСТОВ. Отделяет успешное исполнение макроса от фактически подтверждённого предела и сохраняет всю трассу множителей."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
import json,numpy as np

here=Path(__file__).resolve().parent;p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True);root=p.parse_args().root
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Native/MainCapacity.bin').read_bytes())
assert r.raw[:16]==b'NDMMAINCAP000001'
values=r.VectorTest();assert len(values)==10 and np.all(np.isfinite(values))
names='elapsed_seconds has_ultimate_point lambda_capacity has_steel_limit steel_limit_lambda peak_lambda peak_accuracy has_confirmed_lower_bound confirmed_lower_lambda ultimate_Mx_Nmm'.split()
result=dict(zip(names,map(float,values)));count,=r.VectorTest();assert count==int(count) and count>1
trace=np.array([r.VectorTest() for _ in range(int(count))]);assert trace.shape==(int(count),2) and r.position==len(r.raw)
assert np.all(np.isfinite(trace)) and np.all(trace[:,0]==1)
diagnostics=(root/'Native/MainCapacity.txt').read_text('utf-8-sig')
result.update(native=native,trace=trace.tolist(),maximum_saved_lambda=float(max(trace[:,1])),
    source_unchanged=True,macro_execution_is_not_capacity_acceptance=True,section_equilibrium_not_independently_rechecked=True,
    whole_goal_completed=False,physical_capacity_certified=False,diagnostics=diagnostics)
if result['has_ultimate_point']==0 and result['has_steel_limit']==0:
    assert result['lambda_capacity']==0 and 'Status=NumFail;' in diagnostics
    assert result['has_confirmed_lower_bound']==1 and abs(result['confirmed_lower_lambda']-max(trace[:,1]))<1e-10
(root/'MainCapacityExecutionProof.json').write_text(json.dumps(result,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('MAIN CAPACITY OBSERVED',json.dumps({k:v for k,v in result.items() if k not in ('trace','diagnostics')},sort_keys=True))
