"""ДЛЯ ТЕСТОВ. Проверяет неизменные40 линейных задач и8 внутренних вырожденных шагов в фактическом VBA697."""
from pathlib import Path
from hashlib import sha256
import json,struct
from runpy import run_path
import numpy as np


def TestVector(values):
    """Сохраняет Double в принятом двоичном формате без десятичного преобразования."""
    data=np.asarray(values,dtype='<f8').ravel()
    return struct.pack('<i',len(data))+data.tobytes()


def TestStage(repo,root,parent,original):
    """Переиспользует закрытую книгу697; меняет только читающий тестовый модуль, не факторизацию или материал."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    archive=root/'LinearRankCompilationAncestor';archive.mkdir()
    for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
    prep=json.loads((root/'Preparation.json').read_bytes())
    path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
    assert 'Public Function RunTestTrustStep(' not in code
    payload=(here/'TestLeastSquaresTrustMacro.bas.txt').read_text('utf-8-sig')
    payload=payload.replace('TryDampedStepTest','TryDampedStep').replace('TryLeastSquaresRadiusTest','TryLeastSquaresRadius')
    code+='\n'+payload
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    runner=(root/'Run-TestRealPostCrack.ps1').read_text('utf-8-sig')
    runner=runner.replace('modTestCoupledElasticState.RunTestCoupledPostCrackCapacity','modTestCoupledElasticState.RunTestTrustStep').replace('PostCrackTension.txt','TrustStep.txt')
    raw=(original/'Input/TrustSteps.bin').read_bytes();metadata=json.loads((original/'Input/TrustStepCases.json').read_bytes())
    assert raw[:16]==b'NDMTRUSTFIX00001' and struct.unpack_from('<i',raw,16)==(40,) and len(metadata['cases'])==40
    extra=[]
    def TestAdd(name,matrix,residual,radius=10.,damping=.3):
        extra.append(dict(id=name,matrix=np.asarray(matrix,dtype=float),residual=np.asarray(residual,dtype=float),radius=radius,damping=damping,initial=0.))
    TestAdd('interior-rank-axis',[[1.,0.],[0.,0.]],[1.,2.])
    TestAdd('interior-duplicate-columns',[[1.,1.],[2.,2.]],[1.,2.])
    TestAdd('interior-tiny-residual',[[1.,1.],[2.,2.]],[1e-12,2e-12],1.)
    TestAdd('interior-zero-jacobian',np.zeros((2,2)),[1.,2.],1.)
    TestAdd('interior-permuted-rank-one',np.outer([1.,2.,3.,4.],[0.,3.,0.,2.]),[1.,2.,-1.,4.])
    for scale in (1e-6,1e6):
        TestAdd('interior-duplicate-scale-'+str(scale),np.array([[1.,1.],[2.,2.]])*scale,np.array([1.,2.])*scale,damping=.3*scale**2)
    TestAdd('interior-rank-two',[[1.,0.,1.],[0.,1.,1.],[1.,1.,2.]],[1.,-2.,-1.])
    extended=raw[:16]+struct.pack('<i',48)+raw[20:]
    for case in extra:
        extended+=TestVector([len(case['residual']),case['damping'],case['radius'],case['initial']])+TestVector(case['matrix'])+TestVector(case['residual'])
        metadata['cases'].append({k:v for k,v in case.items() if k not in ('matrix','residual')})
    metadata['first40_original_source']=str(original)
    metadata['first40_original_input_sha256']=sha256(raw).hexdigest().upper()
    metadata['extra_tolerances']=dict(KKT_relative=2e-13,model_relative=2e-13,radius_relative=2e-12,
       fixed_step_relative=1e-7,fixed_derivative_relative=1e-7,maximum_parameter_iterations=10,
       nonunique_interior_step_compares_objective_and_KKT_not_minimum_norm_coordinates=True,
       declared_before_native_run=True)
    values={path:code.replace('\n','\r\n').encode('utf-8-sig'),
       'Compile-TestExternalPoint.ps1':compiler.replace('\n','\r\n').encode('utf-8-sig'),
       'Run-TestRankTrustLinear.ps1':runner.replace('\n','\r\n').encode('utf-8-sig'),
       'Input/TrustSteps.bin':extended,
       'Input/TrustStepCases.json':(json.dumps(metadata,indent=2,sort_keys=True)+'\n').encode('utf-8'),
       'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
       'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes()}
    for name in (Path(__file__).name,'TestLeastSquaresTrustMacro.bas.txt'):values[name]=(here/name).read_bytes()
    for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
    prep['directed_scope'].update(case='RankTrust48Linear',actual_import_this_pass=['modTestCoupledElasticState'],
        actual_solver_and_factorization_inherited_exact=True,material_unchanged=True,
        first40_binary_body_exact=True,additional_interior_rank_cases=8,physical_capacity_certified=False)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('RANK TRUST LINEAR48 STAGED; FIRST40 EXACT; ONLY TEST MODULE IMPORT')


if __name__=='__main__':
    base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
    TestStage(Path.cwd(),base/'Cold698RankTrustLinear48_20261124',base/'Cold697PostCrackRankTrust_20261124',base/'LeastSquaresTrustStepB20261121')
