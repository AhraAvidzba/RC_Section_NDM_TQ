"""ДЛЯ ТЕСТОВ. Разбирает исходные Double полного отклика; недоступные пробы не подменяет производной ноль."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np

here=Path(__file__).resolve().parent

def TestReadSamples(path,live=False):
    """Читает только завершённые образцы; незавершённый хвост допустим исключительно при явном наблюдении живой сессии."""
    data=path.read_bytes();assert data[:16]==b'NDMTAILDIFF00001'
    Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];reader=Reader(data)
    size,expected=reader.ReadTest('<2i');assembled=reader.VectorTest();samples=[];complete=reader.position
    while reader.position<len(data):
        start=reader.position
        try:
            step,column,sign,ready=reader.ReadTest('<d3i')
            assert ready in (0,1) and 0<=column<=size and sign in (-1,0,1)
            residual=reader.VectorTest() if ready else None
        except (EOFError,AssertionError,ValueError,IndexError):
            if not live:raise
            reader.position=start;break
        samples.append(dict(step=step,column=column,sign=sign,ready=bool(ready),residual=residual))
        complete=reader.position
    if not live:assert len(samples)==expected and complete==len(data)
    return size,expected,assembled,samples,dict(bytes=len(data),complete_bytes=complete,incomplete_tail_bytes=len(data)-complete,
        sha256=sha256(data).hexdigest().upper())

def TestSpectrum(matrix):
    """Сопоставляет ранг после независимой нормировки столбцов первой доверительной матрицы."""
    norms=np.linalg.norm(matrix,axis=0);norms[norms==0]=1
    s=np.linalg.svd(matrix/norms,compute_uv=False)
    threshold=np.finfo(float).eps*max(matrix.shape)*s[0]
    return dict(rank=int(np.count_nonzero(s>threshold)),singular_values=s.tolist(),singular_ratio=float(s[-1]/s[0]),
        raw_matrix_abs_max=float(np.max(abs(matrix))),raw_column_norms=norms.tolist())

def TestCheck(root):
    """Сохраняет наблюдаемые различия и контрактные нарушения независимо от нативного диагностического счётчика."""
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackReplay')
    assert native==dict(passed=1,failed=0,macros=1,timeout600=False,own_process_closed=True)
    assert 'sourceStates=70; mode=0;' in text
    parent=root.parent/'Cold737TailMatrixPathRetry_20261125';reference=root.parent/'Cold735FrozenGrowthReplay_20261125'
    for suffix in ('.csv','.Operators.csv'):
        name='Native/PostCrackRawReplay.Accepted'+suffix
        assert (root/name).read_bytes()==(reference/name).read_bytes(),name
    for name in ('src/Solver/CCoupledNewtonSolver.cls','src/Solver/CCoupledDeformationContinuation.cls','src/Materials/CCoupledConcreteResponse.cls'):
        assert (root/name).read_bytes()==(parent/name).read_bytes(),name
    Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
    path=root/'Native/PostCrackRawReplay.TailMatrix2.bin';data=path.read_bytes()
    assert data[:16]==b'NDMTAILMATRIX001';reader=Reader(data);size,ready=reader.ReadTest('<2i')
    ratio,factor=reader.ReadTest('<2d');field=reader.VectorTest();scales=reader.VectorTest();r=reader.VectorTest()
    a=np.array(reader.ReadTest('<'+str(size*size)+'d')).reshape(size,size)
    assert reader.position==len(data) and ready and size==61 and ratio==1.25
    n,expected,assembled,samples,raw=TestReadSamples(root/'Native/PostCrackRawReplay.TailDifferences.bin')
    assert n==size and expected==247 and np.array_equal(assembled,r)
    assert samples[0]['step']==0 and samples[0]['column']==samples[0]['sign']==0 and samples[0]['ready']
    base=samples[0]['residual'];steps=[];matrices={};residuals={}
    for step in (1e-6,1e-4):
        group=[s for s in samples if s['step']==step];assert len(group)==2*size+1
        repeated,=[s for s in group if s['column']==0];assert repeated['sign']==0 and repeated['ready']
        missing=[dict(column=s['column'],sign=s['sign']) for s in group if not s['ready']]
        item=dict(step=step,unavailable=missing,baseline_repeated_exactly=bool(np.array_equal(base,repeated['residual'])))
        if not missing:
            plus=np.column_stack([next(s['residual'] for s in group if s['column']==j and s['sign']==1) for j in range(1,size+1)])
            minus=np.column_stack([next(s['residual'] for s in group if s['column']==j and s['sign']==-1) for j in range(1,size+1)])
            central=(plus-minus)/(2*step);forward=(plus-base[:,None])/step;backward=(base[:,None]-minus)/step
            matrices[step]=central;residuals[step]=(plus,minus)
            norms=np.linalg.norm(central,axis=0);native_norms=np.linalg.norm(a,axis=0)
            item.update(central=TestSpectrum(central),forward=TestSpectrum(forward),backward=TestSpectrum(backward),
                absolute_difference_to_supplied_max=float(np.max(abs(central-a))),
                relative_frobenius_difference_to_supplied=float(np.linalg.norm(central-a)/(np.linalg.norm(central)+np.finfo(float).tiny)),
                relative_column_difference_to_supplied=(np.linalg.norm(central-a,axis=0)/np.maximum(np.maximum(norms,native_norms),np.finfo(float).tiny)).tolist(),
                forward_backward_abs_difference_max=float(np.max(abs(forward-backward))))
        steps.append(item)
    comparison=None
    if len(matrices)==2:
        comparison=dict(absolute_difference_max=float(np.max(abs(matrices[1e-6]-matrices[1e-4]))),
            relative_frobenius_difference=float(np.linalg.norm(matrices[1e-6]-matrices[1e-4])/np.linalg.norm(matrices[1e-4])))
    proof=dict(native=native,source_states=70,accepted_tensor_output_bitidentical_to_verified735=True,
        numerical_core_exact737=True,source_matrix_bitidentical737=data==(parent/'Native/PostCrackRawReplay.TailMatrix2.bin').read_bytes(),
        assembled=TestSpectrum(a),assembled_vs_pure_residual_bitidentical=bool(np.array_equal(r,base)),
        assembled_vs_pure_residual_inf_difference=float(np.max(abs(r-base))),
        source_samples=raw,sample_count=len(samples),steps=steps,finite_difference_step_comparison=comparison,
        no_new_nonlinear_solution=True,physical_capacity_certified=False,whole_goal_completed=False)
    (root/'FullDifferenceProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('738 ALL247 SAMPLES READ; ASSEMBLED RANK',proof['assembled']['rank'],'VALUE/JAC VALUE IDENTICAL',proof['assembled_vs_pure_residual_bitidentical'])
    for step in steps:print('FD STEP',step['step'],'MISSING',len(step['unavailable']),'RANK',step.get('central',{}).get('rank'),'RELATIVE MATRIX DIFFERENCE',step.get('relative_frobenius_difference_to_supplied'))
    return proof

if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold738TailFullDifferences_20261125'))
