"""ДЛЯ ТЕСТОВ. Увеличивает материальные цели в неизменяемой книге936, не меняя материалы, допуски или историю."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold936RecordedYieldCoordinate_20261126';root=base/'Cold938RecordedYieldAdvance_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
vectors=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
meta=vectors[0];step=float(meta[9]);meta[2]=64*step;meta[9]=512*step
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in vectors),Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in values.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='LargerPositiveMaterialTargetsSameHistory25',
    targets_MPa=[float(meta[2]),float(meta[9])],target_multiples=[64,512],
    immutable_compiled936=True,no_new_compilation=True,source_history25_and_six_force_path_unchanged=True,
    original30_steps_and_force_tolerances=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,prep['directed_scope'])
