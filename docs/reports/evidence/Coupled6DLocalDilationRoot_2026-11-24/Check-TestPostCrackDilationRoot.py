"""ДЛЯ ТЕСТОВ. Проверяет скалярное замыкание дилатансии, сохранение старых проб и независимый одноосный инкремент."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import csv,json,re,math
import numpy as np


def TestCheck(repo,root):
    """Наличие локального корня не заменяет полный расчёт; его деформации и напряжения сверяются независимо."""
    here=Path(__file__).resolve().parent;base=root.parent;parent=base/'Cold705PostCrackDilationRoot_20261124'
    def TestRows(folder):return list(csv.reader((folder/'Matrix/Native/LateralFixedIntegration.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    old=TestRows(parent);rows=TestRows(root)
    original=[];dense_started=False
    for row in rows:
        if row[0]=='ACCEPTED':
            original.append(row)
            continue
        if row[1]=='0' and row[0]=='SAMPLE' and int(row[2])==21:dense_started=True
        if row[1]=='0' and dense_started:continue
        original.append(row)
    assert original==old,'Исходные пробы не должны заменяться плотной сеткой.'
    native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert not native['timeout600']
    roots=[v for v in rows if v[0]=='ROOT' and v[1]=='0']
    accepted=next(v for v in rows if v[0]=='ACCEPTED');prior_lateral=float(accepted[3])
    dilation=math.tan(math.radians(32)*math.exp(-2*3/30))
    factor=3*dilation*math.sqrt(2/3+(.1*dilation)**2)
    ratio=(factor-1)/(factor+2)
    expected_lateral=prior_lateral+(.0008-.0006)*ratio
    expected_stress=3*math.exp(-3125*(.0008-.0001))
    measured=[]
    for row in roots:
        stress=np.asarray([float(v) for v in row[6:12]])
        assert row[5]=='1' and max(abs(stress[:2]))<=1.0001e-8
        assert abs(float(row[3])-expected_lateral)<=1e-11
        assert abs(stress[2]-expected_stress)<=.5/21168
        assert max(abs(stress[3:]))<1e-8
        measured.append(dict(epsilon_lateral=float(row[3]),stress_MPa=stress.tolist(),full_original_steps=1))
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    a=read(base/'Cold702PostCrackIntegration_20261124/Matrix/Native/PostCrackPrescribed.Accepted11.bin')
    b=read(root/'Matrix/Native/PostCrackPrescribed.Accepted11.bin')
    for key in ('calibration','options','offset','direction','fields','fractions','lambdas'):assert np.array_equal(a[key],b[key]),key
    # Удаляется только тестовая фиксация дилатансии и начальное приближение; остальное тело полного шага должно совпасть с рабочим.
    production=(repo/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    candidate=(root/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    pattern=r'(?ms)^Private Function UpdatePoint\(.*?^End Function'
    original_method,=re.findall(pattern,production);method,=re.findall(pattern,candidate)
    original_header=original_method[:original_method.index(') As Boolean')+len(') As Boolean')]
    method=re.sub(r'(?s)^Private Function UpdatePoint\(.*?\) As Boolean',lambda _:original_header,method,count=1)
    guide='''    If useGuide Then
        tension = guideTension: compression = guideCompression: dilationRate = guideDilation
        mTension.Evaluate tension, unusedStress, ct, dt, kt
        mCompression.Evaluate compression, unusedStress, cc, dc, kc
        firstReturnReady = False
    End If
'''
    method=method.replace(guide,'',1)
    for line in ('    flowReadyTest = False\n','        If holdDilationTest Then flowResidualTest = 0#: flowReadyTest = True\n','        If apexTransition And holdDilationTest Then Exit Function\n'):method=method.replace(line,'',1)
    method=method.replace(' Or holdDilationTest Then',' Then',1)
    start=method.index('        If holdDilationTest Then\n');end=method.index('        Else\n',start)+len('        Else\n')
    method=method[:start]+method[end:]
    method=method.replace('        End If\n        tension = tension + MAXIMUM_DAMPING * deltaTension','        tension = tension + MAXIMUM_DAMPING * deltaTension',1)
    method=method.replace('If Not holdDilationTest Then dilationRate','dilationRate',1)
    assert method==original_method,'В обычном режиме формулы полного шага не должны меняться.'
    assert 'flowReadyTest = (Abs(deltaDilation) <= STATE_TOLERANCE And _' in candidate
    assert 'DilationConsistent(Abs(deltaDilation), plasticMultiplier, multiplierDivisor, YIELD_TOLERANCE_RATIO * cc / mPressureModulus)' in candidate
    proof=dict(native=native,local_original_full_step_root_found=bool(measured),roots=measured,
       original_coarse_samples_and2to64_reference_records_exact=True,
       original_full_step_default_body_exact=True,original_full_step_acceptance_preserved=True,
       accepted11_native_fields_and_parameters_exact=True,
       independent_uniaxial_reference=dict(dilation=dilation,lateral_plastic_increment_ratio=ratio,
         expected_lateral_strain=expected_lateral,expected_sigma_zz_MPa=expected_stress),
       intermediate_substeps_are_only_predictor=True,full_section_curve_still_negative=True,
       physical_capacity_certified=False,entire_goal_completed=False,
       raw_csv_sha256=sha256((root/'Matrix/Native/LateralFixedIntegration.csv').read_bytes()).hexdigest().upper())
    (root/'PostCrackDilationRootProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('SCALAR DILATION FULL-STEP ROOT',bool(measured),'ROOTS',measured,'NATIVE',native)
    return proof


if __name__=='__main__':TestCheck(Path.cwd(),Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold706PostCrackDilationDense_20261124'))
