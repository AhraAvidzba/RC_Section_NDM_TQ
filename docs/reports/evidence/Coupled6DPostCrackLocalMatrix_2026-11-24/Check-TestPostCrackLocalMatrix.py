"""ДЛЯ ТЕСТОВ. Независимо проверяет обе предписанные посттрещинные кривые без объявления их поиском предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,math,re
import numpy as np


def TestCheck(root):
    """Проверяет нативные исходы, точное начало, независимые одноосные силы и все выгруженные уравнения."""
    here=Path(__file__).resolve().parent;results={}
    native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    reference=read(root.parent/'Cold692PostCrackArithmeticRetry_20261124/Native/PostCrackPrescribed.Accepted11.bin')
    for label in ('Matrix','Refined'):
        result,text=native(root,'Curve'+label);assert not result['timeout600']
        data=root/label
        points=[[float(v) for v in line.split(';')] for line in (data/'Native/PostCrackTension.Curve.csv').read_text('cp1251').splitlines()[1:] if line]
        errors=[]
        for epsilon,force,reported,stress,utilization in points:
            sigma=30000*epsilon if epsilon<=.0001 else 3*math.exp(-3*125*(epsilon-.0001)/.12)
            expected=21168*sigma+432*min(200000*epsilon,400)
            assert abs(reported-expected)<1e-7 and abs(stress-sigma)<1e-12 and abs(utilization-epsilon/.02)<1e-12
            assert abs(force-expected)<=.5,(label,epsilon,force,expected)
            errors.append(abs(force-expected))
        field_proof=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](data,'PostCrackPrescribed.Accepted')
        snapshots=list((data/'Native').glob('PostCrackPrescribed.Accepted*.bin'));assert len(snapshots)==1
        current=read(snapshots[0])
        for key in ('calibration','options','offset','direction'):assert np.array_equal(current[key],reference[key]),(label,key)
        for key in ('fields','fractions','lambdas'):assert np.array_equal(current[key][:2],reference[key][:2]),(label,key)
        full=result['failed']==0 and result['passed']==192 and len(points)==21 and points[-1][0]==.004
        failures=[line for line in text.splitlines() if line.startswith('FAIL:')]
        results[label]=dict(native=result,prescribed_points=len(points),last_strain=points[-1][0],
           maximum_independent_force_error_N=max(errors),full_prescribed_curve_verified=full,
           original_two_initial_fields_bits_exact=True,independent_fields=field_proof,
           failures=failures,known_analytical_later_peak_lambda=(21168*3*math.exp(-3*125*(.002-.0001)/.12)+432*400)/120000,
           complete_capacity_search_run=False,physical_capacity_certified=False)
    proof=dict(study='supplied-local-matrix-postcrack-curves',results=results,entire_goal_completed=False,
       materials_and_final_tolerances_unchanged=True,working_Newton_modified=False)
    (root/'PostCrackLocalMatrixProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    for label,result in results.items():print('CURVE',label,'NATIVE',result['native'],'POINTS',result['prescribed_points'],'LAST',result['last_strain'],'FULL',result['full_prescribed_curve_verified'])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold700PostCrackLocalMatrix_20261124'))
