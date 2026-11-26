"""ДЛЯ ТЕСТОВ. Проверяет прямой возврат исходной цели из дальнего корня938, исключая промежуточный поиск940."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold941OriginalControlReturn_20261126';source=base/'Cold938RecordedYieldAdvance_20261126'
root=base/'Cold942OriginalControlFarReturn_20261126'
proof=json.loads((source/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][2]
assert proof['independently_equilibrated'] and proof['case']==3
chosen=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead'](source/'Native/RecordedYieldCoordinate.bin')[2]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
meta[3]=chosen['info'][6];current=chosen['field'];residual=chosen['values'][:60].copy();residual[56:]/=meta[8]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    'SourceMaterialRoot.bin':(source/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    'VerifiedStartingMaterialRoot.json':json.dumps(proof,sort_keys=True,indent=2).encode('utf-8'),
    Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in values.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalFull61DirectReturnFromFarMaterialRoot',immutable_compiled941=True,
    no_new_compilation=True,source938_case=3,intermediate940_root_not_used=True,
    original_projection_target_and_six_force_path_unchanged=True,source_history25_unchanged=True,
    original30_steps_per_return_case=True,source_material_solve_steps=17,combined_effort_must_be_reported=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'DIRECT938 ROOT RETURN;NO940 INTERMEDIATE;SAME COMPILED941')
