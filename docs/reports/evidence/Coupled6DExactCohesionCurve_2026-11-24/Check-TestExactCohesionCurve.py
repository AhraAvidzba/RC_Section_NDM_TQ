"""ДЛЯ ТЕСТОВ. Проверяет законченный опыт708, включая численный отказ после известного максимума."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import csv,json,math
import numpy as np


def TestCheck(root):
    """Независимо собирает сохранённые поля; частичная кривая не объявляется подтверждением предела."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert native==dict(passed=164,failed=1,macros=1,timeout600=False,own_process_closed=True)
    rows=list(csv.reader((root/'Matrix/Native/PostCrackTension.Curve.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    checked=[]
    for row in rows:
        e,force,reported,sigma,util=map(float,row)
        sc=30000*e if e<=.0001 else 3*math.exp(-3125*(e-.0001))
        expected=21168*sc+432*min(200000*e,400)
        assert abs(force-expected)<=.5 and abs(reported-expected)<1e-7 and abs(sigma-sc)<1e-12
        assert abs(util-e/.02)<1e-12
        checked.append(dict(epsilon=e,force_N=force,expected_N=expected,error_N=abs(force-expected)))
    assert len(checked)==18 and checked[-1]['epsilon']==.002
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root/'Matrix','PostCrackPrescribed.Accepted')
    assert fields['states']==18
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    before=read(root.parent/'Cold702PostCrackIntegration_20261124/Matrix/Native/PostCrackPrescribed.Accepted11.bin')
    after=read(root/'Matrix/Native/PostCrackPrescribed.Accepted18.bin')
    for key in ('calibration','options','offset','direction'):assert np.array_equal(before[key],after[key]),key
    for key in ('fields','fractions','lambdas'):assert np.array_equal(before[key],after[key][:11]),key
    compare=list(csv.reader((root/'Matrix/Native/ExactCohesionComparison.csv').read_text('cp1251').splitlines(),delimiter=';'))
    first=next(v for v in compare if v[0]=='ORIGINAL');second=next(v for v in compare if v[0]=='EXACT')
    assert first[1]==second[1]=='1'
    difference=max(abs(float(a)-float(b)) for a,b in zip(first[2:],second[2:]))
    assert difference<=1e-8
    progress=list(csv.reader((root/'Matrix/Native/DilationProgress.csv').read_text('cp1251').splitlines(),delimiter=';'))
    assert len(progress)==38
    for begin,end in zip(progress[::2],progress[1::2]):
        assert begin[0]=='BEGIN' and end[0]=='END' and begin[1]==end[1]
        assert all(float(b)<=float(e) for b,e in zip(begin[2:8],end[2:8]))
    assert float(progress[-1][1])==.0025 and int(progress[-1][-1])==11
    peak=172967.56279474715/120000
    proof=dict(native=native,accepted_curve=checked,independent_fields=fields,first11_native_history_exact=True,
        local_same_history_stress_difference_MPa=difference,known_uniaxial_peak_lambda=peak,
        native_lambda_at_peak=float(after['lambdas'][-1]),failed_next_strain=.0025,
        prescribed_elapsed_to_peak_s=float(progress[-3][2]),failed_step_elapsed_s=float(progress[-1][2])-float(progress[-2][2]),
        physical_capacity_certified=False,automatic_capacity_not_tested=True,full_curve_passed=False,
        end_to_end10second_target_not_met=True,entire_goal_completed=False)
    assert abs(proof['native_lambda_at_peak']-peak)<.5/120000
    path=root/'ExactCohesionCurveProof.json'
    path.write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('708: VERIFIED18FIELDS; KNOWN PEAK',proof['native_lambda_at_peak'],'NEXT STEP NUMFAIL; NO CAPACITY CERTIFICATE')
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold708ExactCohesionClosure_20261124'))
