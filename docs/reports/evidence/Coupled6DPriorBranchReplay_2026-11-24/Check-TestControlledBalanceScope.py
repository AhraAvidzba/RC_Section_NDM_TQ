"""ДЛЯ ТЕСТОВ. Различает тайм-аут718, отрицательную регрессию и полное исправление трёх групп719."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
import numpy as np


def TestTwin(text):
    """Вычисляет максимум исходного полинома независимо от VBA, не меняя порога1e-4."""
    rows=re.findall(r'(?m)^TEST_TWIN_PEAK_FIELD;(\d+);([^;\r\n]+);([^;\r\n]+)',text)
    assert rows
    values=[]
    for index,epsilon,factor in rows:
        t=(float(epsilon)-.0004)/.0001
        expected=np.polyval([-218,1309,-2428,1379,450],t)/600
        assert abs(float(factor)-expected)<=1e-9
        values.append(dict(index=int(index),epsilon=float(epsilon),lambda_value=float(factor)))
    roots=np.roots([-872,3927,-4856,1379]);root,=[float(v.real) for v in roots if abs(v.imag)<1e-12 and 2<=v.real<=3]
    expected=float(np.polyval([-218,1309,-2428,1379,450],root)/600)
    actual=max(v['lambda_value'] for v in values)
    return dict(states=values,expected_maximum=expected,actual_maximum=actual,error=abs(actual-expected),original_tolerance=1e-4)


def TestCheck(base):
    """Последний снимок не заменяет завершённый результат; контроль источников и двух разных исходов остаётся явным."""
    here=Path(__file__).resolve().parent;root=base/'Cold718ControlledLoadBalance_20261124';fixed=base/'Cold719ScopedPeakRefinement_20261124'
    read=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
    capacity,capacity_text=read(root,'PostCrackTension');assert capacity['timeout600'] and capacity['macros']==0
    old,old_text=read(root,'WorkingGuideRegressions');new,new_text=read(fixed,'WorkingGuideRegressions')
    assert old==dict(passed=47,failed=1,macros=2,timeout600=False,own_process_closed=True)
    assert new==dict(passed=105,failed=0,macros=3,timeout600=False,own_process_closed=True)
    a,b=TestTwin(old_text),TestTwin(new_text)
    assert a['error']>a['original_tolerance'] and b['error']<=b['original_tolerance']
    for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Materials/CCoupledConcreteResponse.cls','tests/modTestCoupledElasticState.bas'):
        assert (root/name).read_bytes()==(fixed/name).read_bytes(),name
    frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Native/Stock.RawHistory.bin')
    checkpoint=run_path(str(here/'Read-TestContinuationCheckpoint.py'))['TestRead'](root/'Native/Stock.Checkpoints.csv')
    frame=frames[-1];assert raw['incomplete_tail_bytes']==0 and len(frame['fields'])==checkpoint['count']==65
    np.testing.assert_allclose(frame['fields'],np.asarray([v['field'] for v in checkpoint['points']]),rtol=1e-14,atol=1e-300)
    compiled=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
    assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compiled['WorkbookSHA256']
    accuracy=float(re.findall(r'погрешность локального максимума: ([0-9.E+-]+)',checkpoint['details'])[0])
    maximum=float(max(frame['lambdas']))
    proof=dict(capacity_job=capacity,capacity_final_result_not_produced=True,
        raw_history=raw,last_checkpoint=checkpoint['sequence'],source_states=65,
        last_accepted_epsilon=float(frame['fields'][-1,0]),native_local_peak=maximum,native_local_peak_accuracy=accuracy,
        native_local_peak_accuracy_meets_target=accuracy<=.001,independent_full_field_replay_pending=True,
        regression_before=old,twin_peak_before=a,regression_after=new,twin_peak_after=b,
        regression_test_module_and_expectations_exact=True,material_and_Newton718_unchanged_in719=True,
        physical_capacity_certified=False,entire_goal_completed=False)
    (fixed/'ControlledBalanceScopeProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('718 TIMEOUT; RAW65 EPS',proof['last_accepted_epsilon'],'LOCAL ACCURACY',accuracy,
        '719 REGRESSION105/0; TWIN ERROR',a['error'],'->',b['error'],'NO PHYSICAL CAPACITY')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ'))
