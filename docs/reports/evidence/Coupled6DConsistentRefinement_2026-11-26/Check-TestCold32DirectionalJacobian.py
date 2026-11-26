"""ДЛЯ ТЕСТОВ. Сопоставляет реальные направленные разности с двумя касательными; найденные дефекты сохраняет, а не ослабляет допуски ради PASS."""
from pathlib import Path
from runpy import run_path
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold899Cold32DirectionalJacobian_20261126';parent=base/'Cold895FreshSpatialArc_20261126'
source=base/'Cold898FreshArcColdStress_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
audit=json.loads((root/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==256 and audit['imported_this_pass']==['modTestCoupledElasticState']
assert len(audit['other_actual_components_exact_parent_except_identifier_case'])==255
for path in (parent/'src').rglob('*'):
    if path.is_file():assert path.read_bytes()==(root/path.relative_to(parent)).read_bytes(),str(path)
assert (root/'Input/Reverse1.bin').read_bytes()==(source/'Input/Reverse1.bin').read_bytes()
validate=run_path(str(here/'Validate-TestHardPrepared.py'))
mechanics,fields=validate['TestValidate'](root,'EnvelopeEnergy.Accepted');assert len(fields)==32
assert (root/'Native/EnvelopeEnergy.Accepted.Operators.csv').read_bytes()==(source/'Native/EnvelopeEnergy.Accepted.Operators.csv').read_bytes()
for suffix in ('.csv',):
    rows=validate['TestRows'](root/('Native/EnvelopeEnergy.Accepted'+suffix))
    reference=validate['TestRows'](source/('Native/EnvelopeEnergy.Accepted'+suffix))
    assert [r for r in rows if r[0] in ('F','S')]==[r for r in reference if r[0] in ('F','S')]
data=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=R((root/'Native/Cold32DirectionalJacobian.bin').read_bytes());assert r.raw[:16]==b'NDMCOLD32JAC0001'
meta=r.VectorTest();accepted=r.VectorTest();movement=r.VectorTest();projection=r.VectorTest()
assert len(meta)==7 and meta[0]==data['lambdas'][-1] and meta[1]==data['lambdas'][-2] and meta[3]==30000 and meta[4]==61
np.testing.assert_array_equal(accepted,data['fields'][-1])
np.testing.assert_array_equal(movement,accepted-data['fields'][-2])
ops=run_path(str(here/'Check-TestHardHalfSteps.py'))['TestOperators'](root/'Native/EnvelopeEnergy.Accepted.Operators.csv',60,43)
g,L,w,v,coords,axis,B,gradient=ops
results=[]
for level in (0,1/64,1,8):
    info=r.VectorTest();field=r.VectorTest();scales=r.VectorTest();refined_residual=r.VectorTest();plain=r.VectorTest()
    assert info[0]==level and info[2]==info[3]==1
    np.testing.assert_array_equal(field,accepted+level*movement)
    assert info[1]==meta[0]+level*(meta[0]-meta[1]) and len(scales)==61 and min(scales)>0
    refined=np.array([r.VectorTest() for _ in range(61)]).T
    ordinary=np.array([r.VectorTest() for _ in range(61)]).T
    assert np.all(np.isfinite(refined)) and np.all(np.isfinite(ordinary))
    expected=-L@data['direction']*scales[60];expected[56:]*=meta[3]
    for matrix in (refined,ordinary):
        np.testing.assert_allclose(matrix[:60,60],expected,rtol=1e-13,atol=1e-13)
        np.testing.assert_allclose(matrix[60,:56],projection[:56]*scales[:56]*meta[3],rtol=1e-13,atol=1e-13)
    direction=np.r_[movement,meta[0]-meta[1]]/scales
    predicted={'refined':refined@direction,'ordinary':ordinary@direction}
    item=dict(level=level,seed_lambda=float(info[1]),residual_MPa=float(max(abs(plain[:56]))),
        stress_mode_residual_difference_MPa=float(max(abs(plain-refined_residual))),
        matrix_difference_max=float(np.max(abs(refined-ordinary))),
        predicted={k:dict(physical_norm_MPa=float(max(abs(v[:56]))),row53_MPa=float(v[52])) for k,v in predicted.items()},samples=[])
    for delta in (-8,-1,-1/8,-1/64,1/64,1/8,1,8):
        header=r.VectorTest();coordinate=r.VectorTest();actual=r.VectorTest()
        assert header[0]==delta and header[1]==1 and len(actual)==61
        np.testing.assert_array_equal(coordinate,delta*np.r_[movement,meta[0]-meta[1]]/scales)
        derivative=(actual-plain)/delta;sample=dict(delta=delta,actual_physical_derivative_norm_MPa=float(max(abs(derivative[:56]))),
            actual_row53_MPa=float(derivative[52]),residual_norm_MPa=float(max(abs(actual[:56]))))
        for name,values in predicted.items():
            error=derivative[:56]-values[:56];worst=int(np.argmax(abs(error)))
            sample[name]=dict(maximum_physical_error_MPa=float(max(abs(error))),worst_row=worst+1,
                row53_error_MPa=float(error[52]),relative_physical_error=float(max(abs(error))/max(max(abs(derivative[:56])),max(abs(values[:56])),1e-300)))
        item['samples'].append(sample)
    results.append(item)
assert r.position==len(r.raw)
proof=dict(native=native,actual_components=256,only_short_test_entry_added=True,all_working_sources_exact895=True,
 true_cold897_history32_and_stresses_exact898=True,independent_accepted_equilibrium=mechanics,
 accepted_revision_native_guard_unchanged=True,no_new_NDS_or_material_acceptance=True,
 actual_failed_iterate_not_captured=True,direction_is_last_accepted_secant_not_failed_Newton_correction=True,
 derivative_discrepancies_preserved=True,levels=results,physical_capacity_certified=False,whole_goal_completed=False)
(root/'Cold32DirectionalJacobianProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('READONLY COLD32 JACOBIAN AUDIT;EXACT32 STRESSES;DIRECTIONAL DISCREPANCIES RECORDED')
for item in results:
    print('LEVEL',item['level'],'MODE_RESIDUAL_DIFFERENCE',item['stress_mode_residual_difference_MPa'])
    for sample in item['samples']:
        print(' DELTA',sample['delta'],'FD53',sample['actual_row53_MPa'],'REFINED_ERROR',sample['refined']['relative_physical_error'],
              'ORDINARY_ERROR',sample['ordinary']['relative_physical_error'],'WORST',sample['refined']['worst_row'])
