"""ДЛЯ ТЕСТОВ. Проверяет фактические локальные пробы, дробления, биты принятой истории и возможный скачок отклика."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import csv,json,math
import numpy as np


def TestCheck(base):
    """Не выводит разрушение из отсутствия локального корня и не выдаёт сравнение нативного материала за независимую реализацию."""
    here=Path(__file__).resolve().parent;first=base/'Cold701PostCrackLateralRoot_20261124';second=base/'Cold702PostCrackIntegration_20261124'
    native=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
    first_native,first_text=native(first,'CurveMatrix');second_native,second_text=native(second,'CurveMatrix')
    assert not first_native['timeout600'] and not second_native['timeout600']
    source=(second/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')
    payload=(here/'TestPostCrackFixedIntegration.cls.txt').read_text('utf-8-sig')
    assert source==(first/'src/Materials/CCoupledConcreteResponse.cls').read_text('utf-8-sig')+'\n'+payload
    def TestRows(file):
        return list(csv.reader(file.read_text('cp1251').splitlines(),delimiter=';'))[1:]
    old=TestRows(first/'Matrix/Native/LateralProbe.csv')
    samples=[row for row in old if row[0]=='SAMPLE']
    assert len(samples)==322 and all(row[4]=='1' for row in samples)
    rows=TestRows(second/'Matrix/Native/LateralFixedIntegration.csv');results={}
    for divisions in (0,2,4,8,16,32,64):
        selected=[row for row in rows if row[0] in ('SAMPLE','BISECT','ROOT') and int(row[1])==divisions]
        measured=[dict(kind=row[0],index=int(row[2]),epsilon_lateral=float(row[3]),substeps=int(row[5]),
            sigma_xx=float(row[6]),sigma_yy=float(row[7]),sigma_zz=float(row[8])) for row in selected if row[4]=='1']
        roots=[v for v in measured if v['kind']=='ROOT']
        for value in roots:assert max(abs(value['sigma_xx']),abs(value['sigma_yy']))<=1.0001e-8
        negatives=[v for v in measured if v['sigma_xx']<0];positives=[v for v in measured if v['sigma_xx']>0]
        gap=None
        if negatives and positives:
            left=max(negatives,key=lambda v:v['epsilon_lateral']);right=min(positives,key=lambda v:v['epsilon_lateral'])
            if left['epsilon_lateral']<=right['epsilon_lateral']:
                gap=dict(left=left,right=right,width=right['epsilon_lateral']-left['epsilon_lateral'],stress_gap_MPa=right['sigma_xx']-left['sigma_xx'])
        results[str(divisions)]=dict(valid_points=len(measured),unavailable_points=sum(row[4]!='1' for row in selected),
            actual_substeps=sorted(set(v['substeps'] for v in measured)),roots=roots,nearest_sign_bracket=gap)
    read=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary']
    a=read(first/'Matrix/Native/PostCrackPrescribed.Accepted11.bin');b=read(second/'Matrix/Native/PostCrackPrescribed.Accepted11.bin')
    for key in ('calibration','options','offset','direction','fields','fractions','lambdas'):assert np.array_equal(a[key],b[key]),key
    verified=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](second/'Matrix','PostCrackPrescribed.Accepted')
    assert len(a['fields'])==11 and verified['states']==11
    # Одна и та же точка из двух сеток сравнивается только при точном совпадении записанной координаты.
    lookup={float(row[3]):[float(v) for v in row[5:11]] for row in samples}
    differences=[]
    for row in rows:
        if row[0]=='SAMPLE' and row[1]=='0' and row[4]=='1' and float(row[3]) in lookup:
            differences.append(max(abs(x-y) for x,y in zip(lookup[float(row[3])],[float(v) for v in row[6:12]])))
    assert len(differences)>=3 and max(differences)<1e-11
    proof=dict(native701=first_native,native702=second_native,automatic701_valid_samples=322,
       automatic701_root_found=False,accepted11_native_fields_and_calibration_bits_equal=True,
       direct_integrator_vs_interface_common_points=len(differences),maximum_stress_difference_MPa=max(differences),
       material_methods_unchanged_only_test_reader_added=True,results=results,
       independent_equilibrium=verified,uniaxial_reference_stress_MPa=3*math.exp(-3*125*(.0008-.0001)/.12),
       independent_material_integration=False,physical_capacity_certified=False,entire_goal_completed=False,
       CSV701_accepted_coordinate_is_axial_not_lateral=True,
       CSV702_accepted_coordinate_corrected_to_lateral=True,
       raw701_sha256=sha256((first/'Matrix/Native/LateralProbe.csv').read_bytes()).hexdigest().upper(),
       raw702_sha256=sha256((second/'Matrix/Native/LateralFixedIntegration.csv').read_bytes()).hexdigest().upper())
    (second/'PostCrackIntegrationProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('INTEGRATION DATA VERIFIED; PREFIX11 EXACT; DIRECT/AUTO INTERFACE MATCH')
    for key,value in results.items():print('DIVISIONS',key,'ROOTS',len(value['roots']),'SUBSTEPS',value['actual_substeps'],'UNAVAILABLE',value['unavailable_points'],'BRACKET',value['nearest_sign_bracket'])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ'))
