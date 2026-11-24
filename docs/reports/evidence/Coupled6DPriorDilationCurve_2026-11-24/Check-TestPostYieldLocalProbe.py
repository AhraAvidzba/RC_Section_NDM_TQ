"""ДЛЯ ТЕСТОВ. Сопоставляет штатный и скалярный локальные ответы после текучести стали на одной истории18."""
from pathlib import Path
from runpy import run_path
import csv,json,math
import numpy as np


def TestCheck(root):
    """Проверяет корни каждого способа отдельно; отсутствие корня не превращается в физический отказ."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert not native['timeout600']
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    a=read(root.parent/'Cold708ExactCohesionClosure_20261124/Matrix/Native/PostCrackPrescribed.Accepted18.bin')
    b=read(root/'Matrix/Native/PostCrackPrescribed.Accepted18.bin')
    for key in ('calibration','options','offset','direction','fields','fractions','lambdas'):assert np.array_equal(a[key],b[key]),key
    rows=list(csv.reader((root/'Matrix/Native/LateralFixedIntegration.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    accepted=next(v for v in rows if v[0]=='ACCEPTED')
    dilation=math.tan(math.radians(32)*math.exp(-2*3/30))
    factor=3*dilation*math.sqrt(2/3+(.1*dilation)**2)
    ratio=(factor-1)/(factor+2)
    expected_lateral=float(accepted[3])+(.0025-.002)*ratio
    expected_stress=3*math.exp(-3125*(.0025-.0001))
    modes={}
    for mode in (-1,0,1,2,4,8,16):
        samples=[v for v in rows if v[0]=='SAMPLE' and int(v[1])==mode]
        roots=[v for v in rows if v[0]=='ROOT' and int(v[1])==mode]
        checked=[]
        for row in roots:
            stresses=np.asarray(row[6:12],float)
            assert max(abs(stresses[:2]))<=1.0001e-8
            checked.append(dict(lateral=float(row[3]),stress_MPa=stresses.tolist(),substeps=int(row[5]),
                lateral_error=abs(float(row[3])-expected_lateral),sigma_zz_error_MPa=abs(stresses[2]-expected_stress)))
        modes[mode]=dict(samples=len(samples),unavailable=sum(v[4]=='0' for v in samples),roots=checked)
    proof=dict(native=native,accepted18_history_and_parameters_exact=True,
        reference=dict(target=.0025,prior=.002,expected_lateral=expected_lateral,expected_sigma_zz_MPa=expected_stress),
        modes=modes,operational_full_step_root_found=bool(modes[-1]['roots']),scalar_full_step_root_found=bool(modes[0]['roots']),
        physical_capacity_certified=False,global_postyield_step_not_solved=True,entire_goal_completed=False)
    path=root/'PostYieldLocalProbeProof.json'
    path.write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('709 NATIVE',native,'REFERENCE',proof['reference'])
    for mode,result in modes.items():print('MODE',mode,result)
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold709PostYieldLocalProbe_20261124'))
