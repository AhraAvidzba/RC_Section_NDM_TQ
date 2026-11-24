"""ДЛЯ ТЕСТОВ. Проверяет алгебру тонкого QR блока R по независимому SVD; не подтверждает реализацию VBA или НДС."""
from pathlib import Path
import json
import numpy as np
from scipy.linalg import qr,lstsq

def TestThinRankStep(b,c):
    """Повторяет предложенные отражения без нормальных уравнений для полного строчного ранга блока B."""
    rank,n=b.shape;t=b.T.copy();diagonal=np.zeros(rank);ordered=np.zeros(n)
    for column in range(rank):
        magnitude=np.linalg.norm(t[column:,column]);assert magnitude>0
        if t[column,column]<0:magnitude=-magnitude
        t[column:,column]/=magnitude;t[column,column]+=1
        for k in range(column+1,rank):
            multiplier=t[column:,column]@t[column:,k]/t[column,column]
            t[column:,k]-=multiplier*t[column:,column]
        diagonal[column]=-magnitude
    for row in range(rank):ordered[row]=(c[row]-t[:row,row]@ordered[:row])/diagonal[row]
    for column in range(rank-1,-1,-1):
        multiplier=t[column:,column]@ordered[column:]/t[column,column]
        ordered[column:]-=multiplier*t[column:,column]
    return ordered

rng=np.random.default_rng(20261125);cases=[]
for n in (2,3,6,12,31,61):
    for rank in sorted(set((0,1,n//2,n-1,n))):
        for condition in (1.,1e3,1e6):
            if rank:
                q,_=np.linalg.qr(rng.normal(size=(n,rank)))
                u,_=np.linalg.qr(rng.normal(size=(rank,rank)))
                b=u@np.diag(np.geomspace(1.,1/condition,rank))@q.T;c=rng.normal(size=rank)
            else:b=np.empty((0,n));c=np.empty(0)
            computed=TestThinRankStep(b,c);reference=lstsq(b,c,lapack_driver='gelsd')[0]
            relative=np.linalg.norm(computed-reference)/(1+np.linalg.norm(reference))
            error=np.linalg.norm(b@computed-c)/(1+np.linalg.norm(c))
            assert relative<=2e-9 and error<=2e-9,(n,rank,condition,relative,error)
            cases.append(dict(size=n,rank=rank,condition=condition,relative_to_independent_SVD=float(relative),relative_residual=float(error)))
root=Path('tmp/material-consistency')
(root/'MinimumNormDerivationProof.json').write_text(json.dumps(dict(cases=cases,seed=20261125,
    actual_VBA_not_executed=True,nonlinear_solver_not_tested=True,physical_capacity_certified=False),indent=2)+'\n',encoding='utf-8')
print('THIN QR DERIVATION VERIFIED:',len(cases),'CASES; NOT A VBA OR NONLINEAR CAPACITY TEST')
