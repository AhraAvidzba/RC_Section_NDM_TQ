"""ДЛЯ ТЕСТОВ. Проверяет исходную21-точечную кривую при прогнозе полного шага из принятой истории."""
from pathlib import Path
from runpy import run_path
import json, csv, math
import numpy as np


def TestCheck(root):
    """Проверяет силу и пространственные уравнения; изменения истории учитываются явно."""
    here=Path(__file__).resolve().parent
    native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert native==dict(passed=194,failed=0,macros=1,timeout600=False,own_process_closed=True)
    parent=root.parent/'Cold711PriorDilationCurve_20261124'
    assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(parent/'tests/modTestCoupledElasticState.bas').read_bytes()
    rows=list(csv.reader((root/'Matrix/Native/PostCrackTension.Curve.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    result=[]
    for row in rows:
        e,n,reported,sigma,util=map(float,row)
        sc=30000*e if e<=.0001 else 3*math.exp(-3125*(e-.0001))
        expected=21168*sc+432*min(200000*e,400)
        assert abs(n-expected)<=.5 and abs(reported-expected)<1e-7 and abs(sc-sigma)<1e-12 and abs(util-e/.02)<1e-12
        result.append(dict(epsilon=e,force_N=n,expected_N=expected,error_N=abs(n-expected)))
    assert len(result)==21 and result[-1]['epsilon']==.004
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root/'Matrix','PostCrackPrescribed.Accepted')
    assert fields['states']==21
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    new=read(root/'Matrix/Native/PostCrackPrescribed.Accepted21.bin');old=read(parent/'Matrix/Native/PostCrackPrescribed.Accepted21.bin')
    for key in ('calibration','options','offset','direction'): assert np.array_equal(new[key],old[key]),key
    for key in ('fields','fractions','lambdas'): assert np.array_equal(new[key][:2],old[key][:2]),key
    proof=dict(native=native,curve=result,independent_fields=fields,original_test_and_parameters_exact=True,
        original_two_elastic_fields_exact=True,new_accepted_history_after_elastic_prefix=True,
        full_original21_prescribed_curve_passed=True,physical_capacity_certified=False,
        automatic_capacity_not_tested=True,entire_goal_completed=False)
    (root/'PriorBranchClosureProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('722 ORIGINAL21 VERIFIED194/0; INDEPENDENT EQUILIBRIUM21; MAX ANALYTICAL ERROR',max(v['error_N'] for v in result),'NO AUTOMATIC CAPACITY YET')
    return proof


if __name__=='__main__': TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold722PriorBranchClosure_20261124'))
