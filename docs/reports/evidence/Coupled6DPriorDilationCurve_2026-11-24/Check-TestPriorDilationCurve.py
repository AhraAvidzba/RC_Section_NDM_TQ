"""ДЛЯ ТЕСТОВ. Проверяет всю фактически сохранённую кривую711 и независимую сборку пространственного равновесия."""
from pathlib import Path
from runpy import run_path
import csv,json,math
import numpy as np


def TestCheck(root):
    """Успех полной предписанной кривой отделён от автоматического поиска и полного ТЗ."""
    here=Path(__file__).resolve().parent;base=root.parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert not native['timeout600']
    assert (root/'tests/modTestCoupledElasticState.bas').read_bytes()==(base/'Cold708ExactCohesionClosure_20261124/tests/modTestCoupledElasticState.bas').read_bytes()
    assert (root/'src/Materials/CCoupledConcreteResponse.cls').read_bytes()==(base/'Cold710PriorDilationSeed_20261124/src/Materials/CCoupledConcreteResponse.cls').read_bytes()
    rows=list(csv.reader((root/'Matrix/Native/PostCrackTension.Curve.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    values=[]
    for row in rows:
        e,n,reported,sigma,util=map(float,row)
        sc=30000*e if e<=.0001 else 3*math.exp(-3125*(e-.0001))
        expected=21168*sc+432*min(200000*e,400)
        assert abs(n-expected)<=.5 and abs(reported-expected)<1e-7 and abs(sc-sigma)<1e-12 and abs(util-e/.02)<1e-12
        values.append(dict(epsilon=e,force_N=n,expected_N=expected,error_N=abs(n-expected)))
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root/'Matrix','PostCrackPrescribed.Accepted')
    assert fields['states']==len(values)
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    binary,=list((root/'Matrix/Native').glob('PostCrackPrescribed.Accepted*.bin'))
    actual=read(binary);old=read(base/'Cold702PostCrackIntegration_20261124/Matrix/Native/PostCrackPrescribed.Accepted11.bin')
    for key in ('calibration','options','offset','direction'):assert np.array_equal(actual[key],old[key]),key
    for key in ('fields','fractions','lambdas'):assert np.array_equal(actual[key][:11],old[key]),key
    complete=len(values)==21 and values[-1]['epsilon']==.004 and native['failed']==0
    if complete:
        assert native['passed']==194 and values[18]['force_N']<values[17]['force_N'] and values[20]['force_N']<values[18]['force_N']
    proof=dict(native=native,accepted_points=len(values),curve=values,independent_fields=fields,
        full_original21_prescribed_curve_passed=complete,original708_test_module_exact=True,
        material710_exact=True,first11_history_and_all_parameters_exact=True,
        physical_capacity_certified=False,automatic_capacity_not_tested=True,entire_goal_completed=False)
    (root/'PriorDilationCurveProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('711 NATIVE',native,'ACCEPTED',len(values),'FULL ORIGINAL21',complete,'LAST',values[-1])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold711PriorDilationCurve_20261124'))
