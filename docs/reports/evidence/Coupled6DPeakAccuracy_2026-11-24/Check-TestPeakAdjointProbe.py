"""ДЛЯ ТЕСТОВ. Независимо исследует точные сопряжённые матрицы715 и совместимость строки множителя при потере ранга."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import csv,json,struct
import numpy as np


def TestReadMatrix(path):
    """Читает фактическую16-байтовую сигнатуру и скаляры Double; исходная строка VBA длиннее фиксированного поля на один знак."""
    raw=path.read_bytes();assert raw[:16]==b'NDMADJOINT0000001'[:16]
    size,=struct.unpack_from('<i',raw,16);factor,tolerance=struct.unpack_from('<2d',raw,20)
    assert size==61 and len(raw)==36+8*(size*size+size)
    values=np.frombuffer(raw,dtype='<f8',offset=36).copy();assert np.all(np.isfinite(values))
    return values[:size*size].reshape(size,size),values[size*size:],factor,tolerance


def TestAnalyse(path):
    """Масштабирует столбцы и строки как штатный LU, затем проверяет SVD-совместимость; это диагностика, не готовая оценка нелинейного предела."""
    matrix,rhs,factor,tolerance=TestReadMatrix(path)
    col=np.max(abs(matrix),axis=0);zero_cols=int(np.count_nonzero(col==0));col[col==0]=1
    scaled=matrix/col
    row=np.max(abs(scaled),axis=1);zero_rows=int(np.count_nonzero(row==0));row[row==0]=1
    scaled=scaled/row[:,None];target=rhs/row
    u,s,vh=np.linalg.svd(scaled,full_matrices=True)
    results=[]
    for threshold in (np.finfo(float).eps*len(rhs)*s[0],1e-12*s[0]):
        rank=int(np.count_nonzero(s>threshold))
        y=vh[:rank].T@((u[:,:rank].T@target)/s[:rank]);weights=y/col
        defect=matrix@weights-rhs
        results.append(dict(threshold=float(threshold),rank=rank,nullity=len(rhs)-rank,
            scaled_compatibility_inf=float(np.max(abs(scaled@y-target))),
            actual_equation_defect_inf=float(np.max(abs(defect))),
            projected_null_rhs_norm=float(np.linalg.norm(u[:,rank:].T@target)),
            adjoint_weights_l1=float(np.sum(abs(weights))),
            linearized_uncertainty_from_recorded_tolerance=float(np.sum(abs(weights))*tolerance)))
    return dict(file=path.name,sha256=sha256(path.read_bytes()).hexdigest().upper(),lambda_value=factor,
        recorded_equilibrium_tolerance=tolerance,zero_columns=zero_cols,zero_rows=zero_rows,
        scaled_singular_values=s.tolist(),svd_results=results)


def TestCheck(root):
    """Сопоставляет каждый зарегистрированный исход LU с исходной матрицей и фактическим завершением поиска."""
    here=Path(__file__).resolve().parent
    native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
    assert not native['timeout600']
    outcomes=list(csv.reader((root/'Native/Adjoint.Outcomes.csv').read_text('cp1251').splitlines(),delimiter=';'))
    records=[]
    for index,solved,pivot,code in outcomes:
        record=TestAnalyse(root/'Native'/('Adjoint.'+index+'.bin'))
        record.update(index=int(index),native_LU_solved=solved=='1',native_scaled_pivot=float(pivot),native_failure_code=int(code))
        records.append(record)
    fields=run_path(str(here/'Check-TestPostCrackFields.py'))['TestAnalyse'](root,'PostCrackContinuation.Accepted')
    proof=dict(native=native,matrices=records,independent_fields=fields,
        failures=[v['index'] for v in records if not v['native_LU_solved']],
        diagnostic_only=True,physical_capacity_certified=False,entire_goal_completed=False)
    (root/'PeakAdjointProbeProof.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    for record in records:
        print(record['index'],'LAMBDA',record['lambda_value'],'LU',record['native_LU_solved'],'PIVOT',record['native_scaled_pivot'],
            'SVD',[(v['rank'],v['actual_equation_defect_inf'],v['linearized_uncertainty_from_recorded_tolerance']) for v in record['svd_results']])
    print('715 NATIVE',native,'FAILED ADJOINTS',proof['failures'],'FIELDS',fields['states'])
    return proof


if __name__=='__main__':TestCheck(Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold715PeakAdjointProbe_20261124'))
