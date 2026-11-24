"""ДЛЯ ТЕСТОВ. Сравнивает исходную равномерную оценку с фактическими остатками, не объявляя линейную оценку строгой нелинейной границей."""
from pathlib import Path
from runpy import run_path
import csv,json,struct
import numpy as np


def TestRead(path):
    """Читает точные матрицу, правую часть и компонентную невязку с полным контролем длины."""
    raw=path.read_bytes();assert raw[:16]==b'NDMADJOINTRES001'
    size,=struct.unpack_from('<i',raw,16);factor,tolerance=struct.unpack_from('<2d',raw,20)
    assert size==61 and len(raw)==36+8*(size*size+2*size)
    values=np.frombuffer(raw,dtype='<f8',offset=36).copy();assert np.all(np.isfinite(values))
    return values[:size*size].reshape(size,size),values[size*size:size*size+size],values[-size:],factor,tolerance


def TestCheck(root):
    """Побитное совпадение истории отделяет диагностическую запись от изменения численного алгоритма."""
    here=Path(__file__).resolve().parent;parent=root.parent/'Cold715PeakAdjointProbe_20261124'
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
    assert native==dict(passed=18,failed=1,macros=1,timeout600=False,own_process_closed=True)
    assert (root/'Native/Stock.RawHistory.bin').read_bytes()==(parent/'Native/Stock.RawHistory.bin').read_bytes()
    assert (root/'Native/PostCrackContinuation.Accepted.csv').read_bytes()==(parent/'Native/PostCrackContinuation.Accepted.csv').read_bytes()
    read_old=run_path(str(here/'Check-TestPeakAdjointProbe.py'))['TestReadMatrix']
    results=[]
    for index,solved,pivot,code in csv.reader((root/'Native/Adjoint.Outcomes.csv').read_text('cp1251').splitlines(),delimiter=';'):
        assert solved=='1'
        name='Adjoint.'+index+'.bin';matrix,rhs,residual,factor,tolerance=TestRead(root/'Native'/name)
        old,old_rhs,old_factor,old_tolerance=read_old(parent/'Native'/name)
        assert np.array_equal(matrix,old) and np.array_equal(rhs,old_rhs) and factor==old_factor and tolerance==old_tolerance
        raw=(root/'Native'/('Adjoint.Weights.'+index+'.bin')).read_bytes();size,=struct.unpack_from('<i',raw)
        assert size==61 and len(raw)==4+8*size
        weights=np.frombuffer(raw,dtype='<f8',offset=4).copy();assert np.all(np.isfinite(weights))
        defect=matrix@weights-rhs
        backward=float(np.max(abs(defect))/(np.linalg.norm(matrix,np.inf)*np.max(abs(weights))+np.max(abs(rhs))))
        assert backward<=128*size*np.finfo(float).eps
        floor=2e-12*(1+abs(factor));uniform_norm=max(tolerance,float(np.max(abs(residual))))
        contributions=abs(weights*residual)
        results.append(dict(index=int(index),lambda_value=factor,allowed_tolerance=tolerance,
            actual_residual_inf=float(np.max(abs(residual))),weights_l1=float(np.sum(abs(weights))),
            uniform_uncertainty=float(np.sum(abs(weights))*uniform_norm+floor),
            componentwise_linearized_estimate=float(np.sum(contributions)+floor),
            signed_linearized_estimate=float(abs(weights@residual)+floor),
            dominant_actual_contribution_row=int(np.argmax(contributions))+1,
            adjoint_backward_error=backward,nonlinear_remainder_not_bounded=True))
    proof=dict(native=native,adjoint_inputs715_exact=True,all_raw_history_and_saved_fields715_exact=True,
        results=results,physical_capacity_certified=False,
        estimator_not_adopted_in_production=True,nonlinear_remainder_not_certified=True,entire_goal_completed=False)
    (root/'ActualAdjointResidualProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    for row in results:print('716',row['index'],'lambda',row['lambda_value'],'uniform',row['uniform_uncertainty'],
        'actual_weighted',row['componentwise_linearized_estimate'],'residual',row['actual_residual_inf'])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold716ActualAdjointResidual_20261124'))
