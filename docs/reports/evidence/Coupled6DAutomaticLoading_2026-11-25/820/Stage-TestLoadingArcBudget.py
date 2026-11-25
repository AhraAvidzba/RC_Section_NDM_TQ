"""ДЛЯ ТЕСТОВ. Проверяет отказ от исходящей фазы при недостатке общего бюджета вызовов без новой компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold819AutomaticLoadingArc_20261125';root=base/'Cold820LoadingArcBudget4_20261125'
assert json.loads((parent/'AutomaticLoadingArcProof.json').read_bytes())['pair']['verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
values={'Input/ArcCallAllowance.bin':struct.pack('<id',1,4.),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='LoadingArcInsufficientCallBudget4',actual_import_this_pass=['CCoupledArcContinuation','modTestCoupledElasticState'],
    actual_component_count=256,immutable_compiled_source_exact819=True,fresh_compilation_claimed=False,
    only_call_allowance_changed=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
