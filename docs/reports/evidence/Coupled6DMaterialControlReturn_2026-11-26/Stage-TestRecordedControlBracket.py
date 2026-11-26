"""ДЛЯ ТЕСТОВ. Локализует исходную линейную цель между проверенными материальными равновесиями одной истории25."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold938RecordedYieldAdvance_20261126';lowroot=base/'Cold936RecordedYieldCoordinate_20261126'
root=base/'Cold940RecordedControlBracket_20261126'
low=json.loads((lowroot/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][2]
high=json.loads((parent/'RecordedYieldCoordinateProof.json').read_bytes())['cases'][2]
assert low['independently_equilibrated'] and high['independently_equilibrated']
lo=low['original_linear_control_residual_MPa'];hi=high['original_linear_control_residual_MPa'];assert lo<0<hi
targetlo=low['target_trial_yield_MPa'];targethi=high['target_trial_yield_MPa']
middle=(targetlo+targethi)/2;secant=targetlo+(targethi-targetlo)*(-lo)/(hi-lo)
assert targetlo<middle<targethi and targetlo<secant<targethi
cases=run_path(str(here/'Read-TestRecordedYieldCoordinate.py'))['TestRead'](parent/'Native/RecordedYieldCoordinate.bin')
chosen=cases[2];info=chosen['info'];assert info[3] and info[5]
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Input/RecordedYieldSeed.bin').read_bytes());r.position=0
meta,current,accepted,residual,projection=[r.VectorTest() for _ in range(5)];assert r.position==len(r.raw)
meta[2]=middle;meta[9]=secant;meta[3]=info[6]
current=chosen['field'];residual=chosen['values'][:60].copy();residual[56:]/=meta[8]
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
bracket=dict(low=low,high=high,shared_accepted25_history=True,not_a_capacity_bracket=True)
values={'Input/RecordedYieldSeed.bin':b''.join(vector(v) for v in (meta,current,accepted,residual,projection)),
    'OriginalControlBracket.json':json.dumps(bracket,sort_keys=True,indent=2).encode('utf-8'),
    'SourceMaterialRoot.bin':(parent/'Native/RecordedYieldCoordinate.bin').read_bytes(),
    Path(__file__).name:Path(__file__).read_bytes(),
    'Read-TestRecordedYieldCoordinate.py':(here/'Read-TestRecordedYieldCoordinate.py').read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in values.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OriginalLinearTargetBracketedByMaterialEquilibria',
    targets_MPa=[middle,secant],starting_root_case938=3,starting_root_lambda=float(info[6]),
    starting_residual_exact_from_native938=True,immutable_compiled936=True,no_new_compilation=True,
    source_history25_and_six_force_path_unchanged=True,original30_steps_and_force_tolerances=True,
    no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,prep['directed_scope'])
