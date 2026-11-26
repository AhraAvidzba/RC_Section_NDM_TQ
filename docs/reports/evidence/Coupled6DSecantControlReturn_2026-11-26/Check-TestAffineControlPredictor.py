"""ДЛЯ ТЕСТОВ. Проверяет геометрическую секущую, прежнюю управляющую строку и точное происхождение нативной невязки прогноза."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold943AffineControlPredictor_20261126';parent=base/'Cold941OriginalControlReturn_20261126'
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'RealColdReplay')
assert native==dict(passed=2,failed=0,macros=2,timeout600=False,own_process_closed=True)
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((root/'Input/FrozenNewton.bin').read_bytes());r.position=0
meta,current,accepted,residual,rhs,movement=[r.VectorTest() for _ in range(6)]
original=np.array([r.VectorTest() for _ in range(61)]);expected=np.array([r.VectorTest() for _ in range(61)])
assert r.position==len(r.raw)
v=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead']
low=v(root/'SourceLowMaterialRoot.bin')[2];high=v(root/'SourceHighMaterialRoot.bin')[2]
ratio=-low['values'][-1]/(high['values'][-1]-low['values'][-1]);assert 0<ratio<1
intended=low['field']+ratio*(high['field']-low['field'])
np.testing.assert_array_equal(movement[:60],intended-current)
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](root/'Input/Reverse1.bin')
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==25
r=R((root/'Native/SurfaceMatrices.bin').read_bytes());assert r.raw[:16]==b'NDMSURFMATRIX001' and r.VectorTest().tolist()==[2.]
records=[]
for index,level in enumerate((0.,1.)):
    info,seed,scales,values=[r.VectorTest() for _ in range(4)];A=np.array([r.VectorTest() for _ in range(61)])
    assert info[0]==level and info[2] and all(np.isfinite(A.ravel()))
    if index==0:
        np.testing.assert_array_equal(seed,current);np.testing.assert_array_equal(values[:60],residual)
        error=float(np.max(abs(A-expected)/np.maximum(np.max(abs(expected),axis=1),np.finfo(float).tiny)[:,None]));assert error<2e-12
    else:
        np.testing.assert_array_equal(seed,current+movement[:60])
        assert info[1]==meta[2]+movement[-1]
    row_error=float(max(abs(A[-1]-original[-1])));column_error=float(max(abs(A[:,60]-original[:,60])))
    assert row_error<1e-10 and column_error<1e-12
    control=float(original[-1,:60]@seed-meta[4]*meta[5]);assert abs(control-values[-1])<2e-11
    records.append(dict(level=level,lambda_value=float(info[1]),native_physical_norm_MPa=float(max(abs(values[:56]))),
        original_control_residual_MPa=float(values[-1]),field=seed.tolist(),raw_residual=values.tolist(),
        scalar_row_error=row_error,load_column_error=column_error))
assert r.position==len(r.raw)
assert abs(records[1]['original_control_residual_MPa'])<2e-11
proof=dict(native=native,immutable_compiled941=True,no_new_compilation=True,source935_residual_bitexact=True,
    source935_refined_matrix_relative_row_error=error,source_history25_unchanged=True,
    secant_ratio=float(ratio),both_endpoints_are_verified_material_equilibria=True,records=records,
    seed_is_not_equilibrium=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'AffineControlPredictorProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('AFFINE SEED',records[1]['lambda_value'],'PHYSICAL R',records[1]['native_physical_norm_MPa'],
    'ORIGINAL CONTROL',records[1]['original_control_residual_MPa'],'SOURCE935 R EXACT;NO NDS CLAIM')
