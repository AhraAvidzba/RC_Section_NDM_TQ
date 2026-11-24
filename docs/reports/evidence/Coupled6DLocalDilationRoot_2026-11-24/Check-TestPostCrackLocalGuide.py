"""ДЛЯ ТЕСТОВ. Проверяет локальный прогноз, повторный полный шаг и неизменность нативной исходной истории."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import csv,json,re,math
import numpy as np


def TestCheck(root):
    """Отличает локальный успех от полного равновесия сечения; принятие полного шага остаётся в исходном UpdatePoint."""
    here=Path(__file__).resolve().parent;parent=root.parent/'Cold702PostCrackIntegration_20261124'
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'CurveMatrix')
    assert not native['timeout600']
    def TestRows(folder):return list(csv.reader((folder/'Matrix/Native/LateralFixedIntegration.csv').read_text('cp1251').splitlines(),delimiter=';'))[1:]
    old=TestRows(parent);rows=TestRows(root)
    assert [v for v in rows if v[0]!='ACCEPTED' and v[1]!='0']==[v for v in old if v[0]!='ACCEPTED' and v[1]!='0']
    roots=[v for v in rows if v[0]=='ROOT' and v[1]=='0']
    values=[]
    for row in roots:
        stress=[float(v) for v in row[6:12]]
        assert int(row[5])==1 and max(abs(stress[0]),abs(stress[1]))<=1.0001e-8
        values.append(dict(epsilon_lateral=float(row[3]),stress_MPa=stress,full_original_steps=1))
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    a=read(parent/'Matrix/Native/PostCrackPrescribed.Accepted11.bin');b=read(root/'Matrix/Native/PostCrackPrescribed.Accepted11.bin')
    for key in ('calibration','options','offset','direction','fields','fractions','lambdas'):assert np.array_equal(a[key],b[key]),key
    oldcode=(parent/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    code=(root/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    pattern=r'(?ms)^(?:Private|Friend|Public) (?:Function|Sub|Property Get) (\w+).*?^End (?:Function|Sub|Property)'
    before={m.group(1):m.group() for m in re.finditer(pattern,oldcode)}
    after={m.group(1):m.group() for m in re.finditer(pattern,code)}
    different=[name for name in before if before[name]!=after[name]]
    assert set(different)=={'UpdatePoint','ReadPostCrackUpdateTest'},different
    assert set(after)-set(before)=={'TestTryPostCrackLocalGuide'}
    # После удаления только начального приближения исходные уравнения, итерации и критерии полного шага должны совпадать точно.
    body=after['UpdatePoint']
    body=body.replace('ByRef updated As TConcreteHistory, ByRef apparent() As Double, ByRef virginElastic As Boolean, _\n        Optional ByVal useGuide As Boolean = False, Optional ByVal guideTension As Double = 0#, _\n        Optional ByVal guideCompression As Double = 0#, Optional ByVal guideDilation As Double = 0#)',
       'ByRef updated As TConcreteHistory, ByRef apparent() As Double, ByRef virginElastic As Boolean)')
    block='''    If useGuide Then
        tension = guideTension: compression = guideCompression: dilationRate = guideDilation
        mTension.Evaluate tension, unusedStress, ct, dt, kt
        mCompression.Evaluate compression, unusedStress, cc, dc, kc
        firstReturnReady = False
    End If
'''
    assert body.count(block)==1 and body.replace(block,'',1)==before['UpdatePoint']
    helper=after['TestTryPostCrackLocalGuide']
    assert 'UpdatePoint(strain, mHistory(1, mCommitted), result, stress, virgin, _' in helper and 'AcceptTrial' not in helper
    proof=dict(native=native,original_full_step_local_root_found=bool(values),original_full_step_roots=values,
      original_full_step_equations_and_acceptance_exact=True,all_other_material_methods_exact=True,
      fixed_subdivision2to64_records_exact=True,accepted11_and_parameters_native_bits_exact=True,
      uniaxial_reference_stress_MPa=3*math.exp(-3*125*(.0008-.0001)/.12),
      intermediate_history_not_accepted=True,full_section_curve_still_negative=True,
      physical_capacity_certified=False,entire_goal_completed=False,
      raw_csv_sha256=sha256((root/'Matrix/Native/LateralFixedIntegration.csv').read_bytes()).hexdigest().upper())
    (root/'PostCrackLocalGuideProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('LOCAL GUIDE FULL-STEP ROOT',bool(values),'ROOTS',values,'NATIVE',native)
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold704PostCrackLocalGuide_20261124'))
