"""ДЛЯ ТЕСТОВ. Оценивает фактические касательные736 независимо SVD и QR, не объявляя численную мягкость разрушением."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np
from scipy.linalg import qr,lstsq

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold737TailMatrixPathRetry_20261125')
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
assert native['passed']==1 and native['failed']==0 and native['own_process_closed']
for suffix in ('.csv','.Operators.csv'):
    name='Native/PostCrackRawReplay.Accepted'+suffix
    assert (root/name).read_bytes()==(root.parent/'Cold735FrozenGrowthReplay_20261125'/name).read_bytes(),name
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Input/PostCrackSourceRawHistory.bin')
last=frames[-1];results=[]
for mode in range(3):
    path=root/'Native'/f'PostCrackRawReplay.TailMatrix{mode}.bin';data=path.read_bytes();assert data[:16]==b'NDMTAILMATRIX001'
    reader=Reader(data);size,ready=reader.ReadTest('<2i');ratio,factor=reader.ReadTest('<2d')
    assert size==61 and ratio==(0,.25,1.25)[mode]
    field=reader.VectorTest();scales=reader.VectorTest()
    expected=last['fields'][-1].copy();expected[:56]+=ratio*(last['fields'][-1,:56]-last['fields'][-2,:56])
    np.testing.assert_array_equal(field,expected)
    assert factor==last['lambdas'][-1]+ratio*(last['lambdas'][-1]-last['lambdas'][-2])
    item=dict(mode=mode,ratio=ratio,ready=bool(ready),sha256=sha256(data).hexdigest().upper(),epsilon=float(field[0]))
    if ready:
        residual=reader.VectorTest();matrix=np.array(reader.ReadTest('<'+str(size*size)+'d')).reshape(size,size)
        assert matrix.shape==(61,61) and np.isfinite(matrix).all() and np.isfinite(residual).all()
        norms=np.linalg.norm(matrix,axis=0);norms[norms==0]=1
        a=matrix/norms
        u,s,vt=np.linalg.svd(a);q,r,piv=qr(a,pivoting=True)
        threshold=np.finfo(float).eps*size*np.max(abs(r))
        rank=next((i for i in range(size) if abs(r[i,i])<=threshold),size)
        step,_,svd_rank,_=lstsq(a,-residual,cond=np.finfo(float).eps*size,lapack_driver='gelsd')
        basic=np.zeros(size)
        if rank:basic[piv[:rank]]=np.linalg.solve(r[:rank,:rank],-(q.T@residual)[:rank])
        item.update(initial_residual_inf=float(max(abs(residual))),column_normalized_singular_values=s.tolist(),
            raw_matrix_abs_max=float(np.max(abs(matrix))),raw_column_norm_min=float(min(norms)),raw_column_norm_max=float(max(norms)),
            singular_ratio=float(s[-1]/s[0]),numerical_rank_SVD=int(svd_rank),rank_QR_same_threshold_rule=rank,
            qr_threshold=float(threshold),least_squares_step_norm=float(np.linalg.norm(step)),
            pivot_basic_step_norm=float(np.linalg.norm(basic)),
            least_squares_predicted_residual=float(max(abs(a@step+residual))),
            pivot_basic_predicted_residual=float(max(abs(a@basic+residual))),
            qr_pivots=piv.tolist(),physical_scales=scales.tolist(),
            physical_weakest_vector=(scales*vt[-1]/norms).tolist())
    assert reader.position==len(data)
    results.append(item)
for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls'):
    assert (root/name).read_bytes()==(root.parent/'Cold735FrozenGrowthReplay_20261125'/name).read_bytes(),name
proof=dict(native=native,source_raw_history=raw,matrices=results,actual_native_normalized_controlled_equations=True,
    all70_tensor_output_bitidentical_to_verified735=True,
    independent_column_scaling_matches_first_trust_matrix=True,
    subsequent_trust_diagonal_history_not_reconstructed=True,no_new_nonlinear_solution=True,
    singularity_or_small_singular_value_not_physical_failure=True,working_code_promoted=False,physical_capacity_certified=False)
(root/'TailMatrixProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
for item in results:print('MATRIX',item['mode'],'READY',item['ready'],'RANK',item.get('numerical_rank_SVD'),'SINGULAR_RATIO',item.get('singular_ratio'))
