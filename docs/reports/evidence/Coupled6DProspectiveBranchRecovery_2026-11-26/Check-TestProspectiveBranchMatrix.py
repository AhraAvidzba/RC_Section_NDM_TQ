"""ДЛЯ ТЕСТОВ. Проверяет точность исходной системы и реальные изменения невязки от касательных соседних материальных ветвей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold914ProspectiveBranchMatrix_20261126';parent=base/'Cold911NewtonRejectedCapture_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,old_delta=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);assert r.position==len(r.raw)
r=R((parent/'Native/NewtonRejected.00002.bin').read_bytes());r.position=0
ref_meta,ref_current,ref_accepted,ref_residual,ref_rhs,ref_delta=[r.VectorTest() for _ in range(6)]
refined=np.array([r.VectorTest() for _ in range(61)])
np.testing.assert_array_equal(ref_current,current);np.testing.assert_array_equal(ref_accepted,accepted);np.testing.assert_array_equal(ref_rhs,rhs)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==32
r=R((root/'Native/ProspectiveBranchMatrix.bin').read_bytes());assert r.raw[:16]==b'NDMPROSPNEWTON01'
records=[];eps=np.finfo(float).eps
for level in (0,1/256,1/128,1/64,1/32,.25,1):
    info,seed,scales=[r.VectorTest() for _ in range(3)]
    assert info[0]==level and len(scales)==61
    np.testing.assert_allclose(seed,current+level*old_delta[:60],rtol=4*eps,atol=1e-300)
    item=dict(level=level,matrix_available=bool(info[2]),linear_solved=bool(info[3]))
    if not info[2]:records.append(item);continue
    values=r.VectorTest();A=np.array([r.VectorTest() for _ in range(61)])
    np.testing.assert_array_equal(A[-1],original[-1]);np.testing.assert_array_equal(A[:,60],original[:,60])
    if level==0:
        difference=abs(A-refined);rowmax=np.maximum(abs(refined).max(axis=1),np.finfo(float).tiny)
        matrix_relative=float(max((difference/rowmax[:,None]).ravel()))
        assert matrix_relative<2e-12
        raw=values[:60].copy();raw[56:]/=meta[4]
        base_residual_error=float(max(abs(raw-residual)));assert base_residual_error<1e-12
    if not info[3]:records.append(item);continue
    delta=r.VectorTest();linear=A@delta
    denominator=abs(A)@abs(delta)+abs(rhs)
    backward=float(max(abs(linear-rhs)/np.maximum(denominator,np.finfo(float).tiny)))
    assert backward<=64*61*eps
    column=np.max(abs(A),axis=0);row=np.max(abs(A/column),axis=1)
    independent=np.linalg.solve(A/column/row[:,None],rhs/row)/column
    linear_difference=float(max(abs((independent-delta)*column))/max(abs(delta*column)))
    assert linear_difference<1e-7
    probes=[]
    for f in (1,.5,.25,.125,.0625,.03125,.015625,.0078125):
        header,trial=r.VectorTest(),r.VectorTest();assert header[0]==f
        roundscale=abs(current)+abs(f*delta[:60])
        assert np.all(abs(trial-(current+f*delta[:60]))<=8*eps*np.maximum(roundscale,np.finfo(float).tiny))
        p=dict(fraction=f,available=bool(header[1]),lambda_value=float(header[2]))
        if header[1]:
            actual=r.VectorTest();assert len(actual)==61
            control=-rhs[-1]+original[-1,:60]@(trial-current)+original[-1,-1]*(header[2]-meta[2])
            assert abs(actual[-1]-control)<=1e-19
            norm=float(max(np.max(abs(actual[:56])),abs(actual[-1])))
            assert abs(norm-header[3])<1e-20
            p.update(weak_and_control_norm_MPa=norm,gauge_scaled_norm_MPa=float(max(abs(actual[56:60]))),
                norm_reduction=float(1-norm/np.max(abs(residual[:56]))),
                norm_below_original_bound=bool(norm<=1.214547460083454e-10),
                row53_residual_MPa=float(actual[52]),control_residual_MPa=float(actual[-1]))
        probes.append(p)
    item.update(linear_backward_error=backward,independent_scaled_linear_solution_difference=linear_difference,
        lambda_increment=float(delta[-1]),probes=probes)
    records.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,other255_components_preserved=True,
    exact_original_scalar_control_row_and_load_column=True,base_matrix_relative_difference_from_actual_refined_capture=matrix_relative,
    base_residual_maximum_difference_from_actual_capture_MPa=base_residual_error,
    records=records,no_accepted_history_change_by_native_guard=True,
    prospective_directions_not_accepted_NDS=True,independent_six_force_and_local_material_root_validation_not_yet_done=True,
    physical_capacity_certified=False,whole_goal_completed=False)
(root/'ProspectiveBranchMatrixProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('BASE MATRIX ERROR',matrix_relative,'BASE RESIDUAL ERROR',base_residual_error)
for item in records:
    good=[p for p in item.get('probes',[]) if p['available']]
    print('LEVEL',item['level'],'MATRIX',item['matrix_available'],'SOLVED',item['linear_solved'],
        'BEST',min((p['weak_and_control_norm_MPa'] for p in good),default=None),
        'BEST REDUCTION',max((p['norm_reduction'] for p in good),default=None))
