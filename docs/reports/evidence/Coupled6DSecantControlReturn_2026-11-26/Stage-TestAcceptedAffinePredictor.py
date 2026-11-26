"""ДЛЯ ТЕСТОВ. Использует принятое состояние25 и дальнее материальное равновесие для прогноза исходной линейной цели."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold941OriginalControlReturn_20261126';root=base/'Cold945AcceptedAffinePredictor_20261126'
highroot=base/'Cold938RecordedYieldAdvance_20261126'
high=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead'](highroot/'Native/RecordedYieldCoordinate.bin')[2]
assert json.loads((highroot/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][2]['independently_equilibrated']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/FrozenNewton.bin').read_bytes());r.position=0
vectors=[r.VectorTest() for _ in range(128)];assert r.position==len(r.raw)
meta,current,accepted,residual,rhs,movement=vectors[:6];A=np.array(vectors[6:67])
np.testing.assert_array_equal(accepted,source['fields'][-1]);assert len(source['fields'])==25
left=float(A[-1,:60]@accepted-meta[4]*meta[5]);right=float(high['values'][-1]);assert left<0<right
ratio=-left/(right-left);assert 0<ratio<1
seed=accepted+ratio*(high['field']-accepted);factor=source['lambdas'][-1]+ratio*(high['info'][6]-source['lambdas'][-1])
vectors[5]=np.append(seed-current,factor-meta[2])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/FrozenNewton.bin':b''.join(vector(v) for v in vectors),'Input/SurfaceMatrixLevels.bin':vector([0.,1.]),
    'SourceHighMaterialRoot.bin':(highroot/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    'AffinePredictorInput.json':json.dumps(dict(ratio=ratio,left_control_MPa=left,right_control_MPa=right,
        intended_seed_lambda=float(factor),intended_seed=seed.tolist(),low_endpoint_is_accepted25=True,
        low_endpoint_material_solve_not_required=True,seed_is_not_equilibrium=True),sort_keys=True,indent=2).encode(),
    Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestSolveOriginalControlReturn';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestReadSurfaceMatrices',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SecantSeedFromAccepted25AndOneMaterialRoot',immutable_compiled941=True,
    no_compilation=True,original935_zero_point_guards_preserved=True,original_target_unchanged=True,
    interpolation_ratio=ratio,history25_and_materials_unchanged=True,no_new_NDS_or_history_acceptance=True,
    low_material936_root_not_used=True,seed_not_claimed_equilibrium=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'RATIO',ratio,'LAMBDA SEED',factor,'LOW ENDPOINT ACCEPTED25;NO LOW AUXILIARY SOLVE')
