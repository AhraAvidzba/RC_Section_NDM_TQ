"""ДЛЯ ТЕСТОВ. Переиспользует точную силовую книгу988 для случая на1% выше сохранённого максимума, меняя только фиксированную нагрузку."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold988FixedForceExternal_20261126';root=base/'Cold989FixedForceAbove_20261126'
assert json.loads((parent/'PointTraceSourceComposition.json').read_bytes())['components']==257
for directory in (parent,):
    prep=json.loads((directory/'Preparation.json').read_bytes())
    for name in ('Run-TestFixedForceExternal.py','Run-TestExternalEnvelope.py'):
        data=(here/name).read_bytes();(directory/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
    (directory/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
case=json.loads((root/'Input/FixedForceCase.json').read_bytes());case['target_ratio_to_known_peak']=1.01
case['target_lambda']=case['known_peak_lambda']*1.01
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/FixedForceCase.bin':vector([case['target_lambda'],case['accepted_prefix'],case['seed_index']]),
    'Input/FixedForceCase.json':(json.dumps(case,sort_keys=True,indent=2)+'\n').encode(),Path(__file__).name:Path(__file__).read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='NativeVersusSciPyFixedForceAboveKnownMaximum',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact988=True,only_target_force_factor_times1point01=True,
    same_accepted_peak16_history_and_seed_scaling_rule=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'TARGET',case['target_lambda'],'EXACT988 CODE;FORCES +1% ONLY')
