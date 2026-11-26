"""ДЛЯ ТЕСТОВ. Уменьшает приращение управляющей координаты до четверти, сохраняя штатную половинную проверку и общий бюджет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold947WorkingSecantReturn_20261126';root=base/'Cold948QuarterWorkingReturn_20261126'
p=json.loads((parent/'WorkingSecantReturnProof.json').read_bytes());assert not p['cases'][0]['native_solved']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
vectors=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
meta,current,accepted,residual,projection=vectors;full=float(meta[4]);start=float(projection@accepted)
meta[4]=start+.25*(full-start)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in vectors),
    'FullControlTargetReference.json':json.dumps(dict(accepted_coordinate=start,full_target=full,reduced_target=float(meta[4]),
        increment_fraction=.25,full_failed_target_not_claimed_solved=True),sort_keys=True,indent=2).encode(),
    Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='QuarterControlIncrementWorkingReturnSameGuards',immutable_compiled947=True,
    no_new_compilation=True,guide938_and_material_history25_unchanged=True,
    reduced_diagnostic_control_increment_fraction=.25,full_original_control_target_not_solved_by_this_test=True,
    all_force_half_uncertainty_direction_and_steel_tolerances_unchanged=True,
    guide17_steps_in_total30=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'QUARTER CONTROL TARGET',meta[4],'ORIGINAL FULL TARGET',full,'SAME ALL GUARDS')
