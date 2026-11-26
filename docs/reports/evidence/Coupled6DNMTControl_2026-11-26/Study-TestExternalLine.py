"""ДЛЯ ТЕСТОВ. Локализует материальный излом вдоль фактической поправки на неизменной принятой истории."""

import json
from pathlib import Path
import struct
from types import SimpleNamespace
import numpy as np


def TestStudy(root, residual, point_trace, initial, first):
    """Все пробные поля остаются диагностическими; неудача вдоль линии не является физическим отказом."""
    raw=(root/'Input/LineDirection.bin').read_bytes();n,=struct.unpack_from('<i',raw)
    assert len(raw)==4+8*n and len(initial)==n
    direction=np.frombuffer(raw,dtype='<f8',offset=4,count=n).copy()
    raw=(root/'Input/LineJacobian.bin').read_bytes();assert len(raw)==4+8*n*n
    matrix=np.frombuffer(raw,dtype='<f8',offset=4,count=n*n).reshape(n,n).copy()
    assert max(abs(matrix@direction+first))<1e-12
    origin=json.loads((root/'Input/LinePrediction.json').read_bytes())
    candidate_point=origin['strongest_point_subspaces'][0]['point']
    samples=[];cache={};calls=0;best=initial.copy();best_norm=float(max(abs(first)))
    report=dict(candidate_point_from_operator_projection=candidate_point,samples=samples,
        physical_failure=False,material_history_committed=False,original_force_tolerance_unchanged=True)
    def TestSave():
        """Сохраняет уже полученные наблюдения перед следующей нативной пробой."""
        (root/'Native/ExternalLineStudy.json').write_bytes((json.dumps(report,indent=2,sort_keys=True)+'\n').encode())
    def TestF(x):
        """Считает число фактических запросов исследования, не число фиктивных итераций."""
        nonlocal calls
        calls+=1
        return residual(x)
    def TestAt(x,label,alpha=None):
        """Контролирует побитовую неизменность F после читающего вывода всех точек."""
        nonlocal best,best_norm
        f=TestF(x);points=point_trace(x)
        assert np.array_equal(TestF(x),f),'Point observer changed the original residual'
        assert points.shape[1]==30
        item=dict(label=label,alpha=alpha,max_F_MPa=float(max(abs(f))),
            nonlinear_line_defect_MPa=(float(max(abs(f-first-alpha*(matrix@direction)))) if alpha is not None else None),
            point_trace_index=len(samples)+1)
        samples.append(item)
        if item['max_F_MPa']<best_norm:best=x.copy();best_norm=item['max_F_MPa']
        TestSave()
        return x.copy(),f.copy(),points.copy()
    def TestAlpha(alpha):
        """Ключ кэша — точный Double; округлённых ключей и изменяемых материальных состояний нет."""
        alpha=float(alpha)
        if alpha not in cache:cache[alpha]=TestAt(initial+alpha*direction,'line',alpha)
        return cache[alpha]
    base=TestAlpha(0.)
    original_grid=(-.015625,.000244140625,.00048828125,.0009765625,.001953125,.00390625,.0078125,.015625)
    for alpha in original_grid:TestAlpha(alpha)
    end=cache[.015625]
    assert np.array_equal(base[2][:,20:],end[2][:,20:]),'Accepted material state changed'
    changed=np.flatnonzero(np.any(base[2][:,14:17]!=end[2][:,14:17],axis=1))
    report['changed_point_branch_rows']=[dict(point=int(base[2][i,0]),side=int(base[2][i,1]),
        before=base[2][i,14:18].tolist(),after=end[2][i,14:18].tolist()) for i in changed]
    selected=[int(i) for i in changed if int(base[2][i,0])==candidate_point and int(base[2][i,1])!=0]
    report['selected_rows']=selected
    if selected:
        row=selected[0];low=0.;high=.015625;flags=base[2][row,14:17]
        for _ in range(40):
            mid=low+(high-low)/2
            trial=TestAlpha(mid)
            if np.array_equal(trial[2][row,14:17],flags):low=mid
            else:high=mid
            if np.array_equal(initial+low*direction,initial+high*direction):break
        report['event']=dict(point=int(base[2][row,0]),side=int(base[2][row,1]),
            alpha_before=low,alpha_after=high,
            flags_before=TestAlpha(low)[2][row,14:18].tolist(),
            flags_after=TestAlpha(high)[2][row,14:18].tolist(),
            F_jump_over_final_bracket_MPa=float(max(abs(TestAlpha(high)[1]-TestAlpha(low)[1]))))
        center=(low+high)/2;continuity=[]
        for delta in (1e-4,1e-5,1e-6,1e-7,1e-8,1e-9,1e-10,1e-11):
            left=TestAlpha(center-delta);right=TestAlpha(center+delta)
            continuity.append(dict(alpha_half_width=delta,
                global_F_difference_MPa=float(max(abs(right[1]-left[1]))),
                selected_point_stress_difference_MPa=float(max(abs(right[2][row,8:14]-left[2][row,8:14]))),
                selected_point_yield_before=float(left[2][row,17]),selected_point_yield_after=float(right[2][row,17]),
                normalized_coordinates_distinct=not np.array_equal(left[0],right[0])))
        report['continuity_samples']=continuity
        right_alpha=min(alpha for alpha in original_grid if alpha>high and alpha>0)
        right=TestAlpha(right_alpha);right_matrix=np.empty((n,n))
        for column in range(n):
            h=1e-9*(1+abs(right[0][column]));a=right[0].copy();b=a.copy();a[column]+=h;b[column]-=h
            right_matrix[:,column]=(TestF(a)-TestF(b))/(a[column]-b[column])
        singular=np.linalg.svd(right_matrix-matrix,compute_uv=False)
        correction=np.linalg.solve(right_matrix,-right[1])
        left_root=TestAt(initial+direction,'left_linear_root')
        right_root=TestAt(right[0]+correction,'right_linear_root')
        report['two_branch_linearization']=dict(right_alpha=right_alpha,
            left_determinant_sign=float(np.linalg.slogdet(matrix)[0]),
            right_determinant_sign=float(np.linalg.slogdet(right_matrix)[0]),
            tangent_difference_singular_values=singular.tolist(),
            left_linear_root_point_flags=left_root[2][row,14:18].tolist(),
            right_linear_root_point_flags=right_root[2][row,14:18].tolist(),
            left_linear_root_F_MPa=float(max(abs(left_root[1]))),
            right_linear_root_F_MPa=float(max(abs(right_root[1]))))
    TestSave()
    return SimpleNamespace(x=best,success=False,message='Read-only material transition line study; no nonlinear solution claimed',nfev=calls,njev=1)
