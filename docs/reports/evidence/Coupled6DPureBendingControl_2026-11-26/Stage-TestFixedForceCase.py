"""ДЛЯ ТЕСТОВ. Создаёт силовой контроль из явно выбранного принятого префикса без нового VBA и без передачи будущего искомого поля."""
from argparse import ArgumentParser
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parser=ArgumentParser(description=__doc__);parser.add_argument('--root',type=Path,required=True);parser.add_argument('--ratio',type=float,required=True)
parser.add_argument('--prefix',type=int,required=True);parser.add_argument('--seed',type=int,required=True);args=parser.parse_args()
parent=base/'Cold988FixedForceExternal_20261126';root=args.root
data=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
assert 0<=args.prefix<=len(data['fields']) and 1<=args.seed<=max(args.prefix,1) and .9<=args.ratio<=1.1
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
case=json.loads((parent/'Input/FixedForceCase.json').read_bytes());case.update(target_ratio_to_known_peak=args.ratio,
    target_lambda=case['known_peak_lambda']*args.ratio,accepted_prefix=args.prefix,seed_index=args.seed,
    seed_reference_lambda=float(data['lambdas'][args.seed-1]),no_future_equilibrium_field_input=True)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/FixedForceCase.bin':vector([case['target_lambda'],args.prefix,args.seed]),
    'Input/FixedForceCase.json':(json.dumps(case,sort_keys=True,indent=2)+'\n').encode(),Path(__file__).name:Path(__file__).read_bytes(),
    'Run-TestFixedForceExternal.py':(here/'Run-TestFixedForceExternal.py').read_bytes()}
prep=json.loads((root/'Preparation.json').read_bytes())
for name,content in updates.items():(root/name).write_bytes(content);prep['inputs'][name]=sha256(content).hexdigest().upper()
prep['directed_scope']=dict(case='NativeVersusSciPyFixedForceFromPriorAcceptedPrefix',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact988=True,only_fixed_force_prefix_and_predictor_data_changed=True,
    accepted_prefix=args.prefix,seed_index=args.seed,target_ratio_to_known_maximum=args.ratio,
    no_future_equilibrium_field_input=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'FIXED FORCE',case['target_lambda'],'PREFIX/SEED',args.prefix,args.seed,'EXACT988 CODE')
