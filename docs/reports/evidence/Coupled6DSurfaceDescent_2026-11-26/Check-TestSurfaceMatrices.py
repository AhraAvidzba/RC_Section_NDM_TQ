"""ДЛЯ ТЕСТОВ. Проверяет матрицы фактической системы929 и направления двух вариантов линеаризации без принятия новых НДС."""
from pathlib import Path
from runpy import run_path
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold931SurfaceMatrices_20261126';parent=base/'Cold925RecordedTailSystem_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==256
for p in (parent/'src').rglob('*'):
    if p.is_file():assert p.read_bytes()==(root/p.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,old_delta=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);expected=np.array([r.VectorTest() for _ in range(61)])
assert r.position==len(r.raw)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==25
r=R((root/'Input/SurfaceMatrixLevels.bin').read_bytes());r.position=0;levels=r.VectorTest();assert r.position==len(r.raw)
r=R((root/'Native/SurfaceMatrices.bin').read_bytes());assert r.raw[:16]==b'NDMSURFMATRIX001'
assert r.VectorTest().tolist()==[len(levels)]
records=[];matrices=[];predictions=[];maximum_seed_roundoff=0.
for index,level in enumerate(levels):
    info,seed,scales=[r.VectorTest() for _ in range(3)];assert info[0]==level
    expected_seed=current+level*old_delta[:60]
    seed_error=abs(seed-expected_seed)
    roundscale=abs(current)+abs(level*old_delta[:60])
    assert np.all(seed_error<=4*np.finfo(float).eps*np.maximum(roundscale,np.finfo(float).tiny))
    if index==0:np.testing.assert_array_equal(seed,current)
    maximum_seed_roundoff=max(maximum_seed_roundoff,float(max(seed_error)))
    item=dict(level=float(level),matrix_available=bool(info[2]),seed_lambda=float(info[1]))
    if not info[2]:records.append(item);continue
    values=r.VectorTest();A=np.array([r.VectorTest() for _ in range(61)])
    if index==0:
        base_error=float(max(abs(values[:60]-residual)));assert base_error==0
        rowmax=np.maximum(np.max(abs(expected),axis=1),np.finfo(float).tiny)
        matrix_error=float(max((abs(A-expected)/rowmax[:,None]).ravel()));assert matrix_error<2e-12
    row_error=float(max(abs(A[-1]-original[-1])));column_error=float(max(abs(A[:,60]-original[:,60])))
    assert row_error<1e-10 and column_error<1e-12
    A[-1]=original[-1];A[:,60]=original[:,60]  # Для анализа сохраняется точно записанная строка и столбец нагрузки.
    values[-1]=-rhs[-1]+original[-1,:60]@(seed-current)+original[-1,-1]*(info[1]-meta[2])
    columns=np.max(abs(A),axis=0);rows=np.max(abs(A/columns),axis=1)
    singular=np.linalg.svd(A/columns/rows[:,None],compute_uv=False)
    item.update(seed_physical_norm_MPa=float(max(abs(values[:56]))),scaled_singular_ratio=float(singular[-1]/singular[0]),
        native_scalar_row_max_error=row_error,native_load_column_max_error=column_error,corrections=[])
    for mode,field,factor,base_residual in [('matched_seed',seed,info[1],values),('modified_current',current,meta[2],-rhs)]:
        try:correction=np.linalg.solve(A/columns/rows[:,None],-base_residual/rows)/columns
        except np.linalg.LinAlgError:
            item['corrections'].append(dict(mode=mode,linear_solved=False));continue
        backward=float(max(abs(A@correction+base_residual)/np.maximum(abs(A)@abs(correction)+abs(base_residual),np.finfo(float).tiny)))
        assert backward<=64*61*np.finfo(float).eps
        entry=dict(mode=mode,linear_solved=True,linear_backward_error=backward,lambda_increment=float(correction[-1]),
            linear_norm_MPa=float(max(abs(A@correction+base_residual))))
        item['corrections'].append(entry)
        predictions.append(dict(matrix_index=index,level=float(level),mode=mode,field=field.copy(),lambda_value=float(factor),
            correction=correction,linear_backward_error=backward))
    matrices.append(dict(level=float(level),matrix=A,seed=seed,seed_lambda=float(info[1]),residual=values))
    records.append(item)
assert r.position==len(r.raw) and all(x['matrix_available'] for x in records)
proof=dict(native=native,actual_components=257,all256_other_components_preserved=True,
    base_residual_max_error_MPa=base_error,base_refined_matrix_relative_row_error=matrix_error,
    original_scalar_row_and_load_column_preserved_in_analysis=True,records=records,
    maximum_native_vs_independent_seed_arithmetic_roundoff=maximum_seed_roundoff,
    actual_exported_native_seed_used_for_all_corrections=True,
    all_linear_corrections_independently_verified=True,material_history_unchanged_by_native_guard=True,
    corrections_are_unverified_predictions_not_equilibria=True,physical_capacity_certified=False,whole_goal_completed=False)
(root/'SurfaceMatricesProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BASE R ERROR',base_error,'REFINED A ERROR',matrix_error)
for item in records:print('LEVEL',item['level'],'R',item.get('seed_physical_norm_MPa'),'CONDITION',item.get('scaled_singular_ratio'),
    'D LAMBDA',[(x['mode'],x.get('lambda_increment')) for x in item.get('corrections',[])])
