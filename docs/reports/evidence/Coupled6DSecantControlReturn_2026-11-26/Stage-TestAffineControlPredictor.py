"""ДЛЯ ТЕСТОВ. Строит начальное поле по секущей двух материальных равновесий и читает его невязку неизменяемой книгой941."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold941OriginalControlReturn_20261126';root=base/'Cold943AffineControlPredictor_20261126'
read=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead']
lo_root=base/'Cold936RecordedYieldCoordinate_20261126';hi_root=base/'Cold938RecordedYieldAdvance_20261126'
low=read(lo_root/'Native/RecordedYieldCoordinate.bin')[2];high=read(hi_root/'Native/RecordedYieldCoordinate.bin')[2]
for folder in (lo_root,hi_root):
    assert json.loads((folder/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][2]['independently_equilibrated']
left=float(low['values'][-1]);right=float(high['values'][-1]);assert left<0<right
ratio=-left/(right-left);assert 0<ratio<1
seed=low['field']+ratio*(high['field']-low['field']);factor=low['info'][6]+ratio*(high['info'][6]-low['info'][6])
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/FrozenNewton.bin').read_bytes());r.position=0
vectors=[r.VectorTest() for _ in range(6+122)];assert r.position==len(r.raw)
meta,current,accepted,residual,rhs,movement=vectors[:6];movement=np.append(seed-current,factor-meta[2]);vectors[5]=movement
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values={'Input/FrozenNewton.bin':b''.join(vector(v) for v in vectors),'Input/SurfaceMatrixLevels.bin':vector([0.,1.]),
    'SourceLowMaterialRoot.bin':(lo_root/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    'SourceHighMaterialRoot.bin':(hi_root/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    'AffinePredictorInput.json':json.dumps(dict(ratio=ratio,left_control_MPa=left,right_control_MPa=right,
        intended_seed_lambda=float(factor),intended_seed=seed.tolist(),seed_is_not_equilibrium=True),sort_keys=True,indent=2).encode(),
    Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestSolveOriginalControlReturn';assert runner.count(old)==1
values[name]=runner.replace(old,'modTestCoupledElasticState.TestReadSurfaceMatrices',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in values.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SecantSeedFromTwoVerifiedMaterialEquilibria',immutable_compiled941=True,
    no_compilation=True,original_source935_zero_point_guards_preserved=True,original_control_target_unchanged=True,
    interpolation_ratio=ratio,history25_and_materials_unchanged=True,no_new_NDS_or_history_acceptance=True,
    seed_not_claimed_equilibrium=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'RATIO',ratio,'LAMBDA SEED',factor,'ORIGIN935 GUARDS PRESERVED')
