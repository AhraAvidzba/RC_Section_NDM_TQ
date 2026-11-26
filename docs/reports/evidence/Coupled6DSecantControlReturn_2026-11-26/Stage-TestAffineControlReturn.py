"""ДЛЯ ТЕСТОВ. Запускает исходную систему из фактического нативного прогноза по секущей943 без ручного промежуточного материального уровня."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold941OriginalControlReturn_20261126';source=base/'Cold943AffineControlPredictor_20261126'
root=base/'Cold944AffineControlReturn_20261126'
proof=json.loads((source/'AffineControlPredictorProof.json').read_bytes());point=proof['records'][1]
assert proof['source935_residual_bitexact'] and abs(point['original_control_residual_MPa'])<2e-11
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
r=R((source/'Native/SurfaceMatrices.bin').read_bytes());assert r.VectorTest().tolist()==[2.]
for _ in range(2):
    info,seed,scales,values=[r.VectorTest() for _ in range(4)];matrix=[r.VectorTest() for _ in range(61)]
assert r.position==len(r.raw) and info[0]==1 and info[2]
np.testing.assert_array_equal(seed,np.array(point['field']));np.testing.assert_array_equal(values,np.array(point['raw_residual']))
meta[3]=info[1];current=seed;residual=values[:60]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    'SourceAffineMatrices.bin':(source/'Native/SurfaceMatrices.bin').read_bytes(),
    'SourceAffinePredictorProof.json':(source/'AffineControlPredictorProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalFull61ReturnFromSecantSeed',immutable_compiled941=True,no_new_compilation=True,
    initial_R_exact_from943_native=True,manual_material_level940_not_used=True,
    original_projection_target_six_force_path_and_history25_unchanged=True,
    original30_steps_and_force_tolerances=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXACT943 SECANT SEED;ORIGINAL61 SYSTEM;SAME COMPILED941')
