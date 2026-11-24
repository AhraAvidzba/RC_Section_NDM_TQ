"""ДЛЯ ТЕСТОВ. Независимо проверяет48 нативных шагов: исходные40 критериев неизменны, дополнительные8 не требуют уникального вектора."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct
import numpy as np
from scipy.linalg import svd


def TestCheck(root):
    """Сверяет регуляризованный шаг и производную с SVD, внутреннее вырожденное решение — с проекцией, KKT и радиусом."""
    here=Path(__file__).resolve().parent
    source=here/'Check-TestLeastSquaresTrustStep.py';code=source.read_text('utf-8-sig')
    assert code.count('len(metadata[\'cases\']) == 40')==1
    code=code.replace('len(metadata[\'cases\']) == 40','len(metadata[\'cases\']) == 48',1)
    code=code.replace("zip(inputs, results, metadata['cases'])","zip(inputs[:40], results[:40], metadata['cases'][:40])",1)
    imports=json.loads((root/'Preparation.json').read_bytes())['directed_scope']['actual_import_this_pass']
    assert imports in (['modTestCoupledElasticState'],['CCoupledHybridFactorization','CCoupledHybridSolver','CHybridEquationsTest','modTestCoupledElasticState','modTestCoupledHybrid'])
    code=code.replace("['CCoupledHybridFactorization', 'modTestCoupledElasticState']",repr(imports),1)
    scope={'__file__':str(source),'__name__':'rank_linear_original40_test'}
    exec(compile(code,str(source),'exec'),scope)
    first=scope['TestCheck'](root)
    read=scope['TestRead']
    inputs=read(root/'Input/TrustSteps.bin',b'NDMTRUSTFIX00001',3)
    results=read(root/'Native/TrustStepResults.bin',b'NDMTRUSTOUT00001',3)
    metadata=json.loads((root/'Input/TrustStepCases.json').read_bytes())
    original=Path(metadata['first40_original_source']);raw=(original/'Input/TrustSteps.bin').read_bytes()
    assert sha256(raw).hexdigest().upper()==metadata['first40_original_input_sha256']
    readeraw=(root/'Input/TrustSteps.bin').read_bytes()
    assert readeraw[:16]==raw[:16] and readeraw[20:len(raw)]==raw[20:]
    assert metadata['cases'][:40]==json.loads((original/'Input/TrustStepCases.json').read_bytes())['cases']
    limits=metadata['extra_tolerances'];assert limits['declared_before_native_run']
    records=[]
    for index,(data,result,case) in enumerate(zip(inputs[40:],results[40:],metadata['cases'][40:]),41):
        settings,flat,f=data;info,fixed,p=result;n=int(settings[0]);a=flat.reshape(n,n);alpha,radius=settings[1:3]
        assert info[0]==index and info[1]==n and bool(info[2]) and bool(info[4]) and len(p)==n and len(fixed)==n
        u,s,vt=svd(a,full_matrices=False);uf=u.T@f
        reference_fixed=-vt.T@((s*uf)/(s*s+alpha));norm=np.linalg.norm(reference_fixed)
        fixed_error=np.linalg.norm(fixed-reference_fixed)/max(norm,1e-300)
        derivative=-np.sum((s*uf)**2/(s*s+alpha)**3)/norm if norm else 0.
        derivative_error=abs(info[3]-derivative)/max(abs(derivative),1e-300)
        assert fixed_error<=limits['fixed_step_relative'] and derivative_error<=limits['fixed_derivative_relative'],(case['id'],'fixed',fixed_error,derivative_error)
        rank=s>np.finfo(float).eps*n*s[0]
        coefficients=np.zeros(n);coefficients[rank]=uf[rank]/s[rank]
        least=-vt.T@coefficients
        actual=a@p+f;expected=a@least+f
        projection_error=np.linalg.norm(actual-expected)/max(np.linalg.norm(f),s[0]*np.linalg.norm(least),1e-300)
        scale=s[0]*(s[0]*np.linalg.norm(p)+np.linalg.norm(f))
        kkt=np.linalg.norm(a.T@actual)/max(scale,1e-300)
        radius_error=max(0.,np.linalg.norm(p)-radius)/radius
        assert projection_error<=limits['model_relative'],(case['id'],'least squares projection',projection_error)
        assert kkt<=limits['KKT_relative'] and radius_error<=limits['radius_relative'],(case['id'],'KKT/radius',kkt,radius_error)
        assert info[5]==0 and info[6]==0,(case['id'],'interior step must not search boundary',info.tolist())
        records.append(dict(id=case['id'],fixed_step_relative=fixed_error,fixed_derivative_relative=derivative_error,
            projection_relative=projection_error,KKT_relative=kkt,radius_relative=radius_error,
            minimum_norm_coordinates_not_required=True,native_step_norm=float(np.linalg.norm(p))))
    assert len(first['cases'])==40 and len(records)==8
    proof=dict(original40=first,extra8=records,first40_binary_body_and_metadata_exact=True,
       original40_criteria_unchanged=True,extra_tolerances=limits,material_changed=False,
       physical_capacity_certified=False,entire_goal_completed=False)
    (root/'RankTrustLinearProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('NATIVE LINEAR48 VERIFIED; ORIGINAL40 UNCHANGED; SVD/DERIVATIVE/KKT/RADIUS; NOT CAPACITY')
    return proof


if __name__=='__main__':
    from argparse import ArgumentParser
    parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,default=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold698RankTrustLinear48_20261124'))
    TestCheck(parser.parse_args().root.resolve())
