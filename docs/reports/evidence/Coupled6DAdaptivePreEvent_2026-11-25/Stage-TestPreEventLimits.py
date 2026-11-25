"""ДЛЯ ТЕСТОВ. Переиспользует проверенную компиляцию794 для отрицательных бюджетов6 и3 без изменения физики или ядра."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold794PreEventBudget_20261125'
assert json.loads((parent/'AdaptivePreEventProof.json').read_bytes())['returned_verified_candidate']
for name,allowance in (('Cold795PreEventSixCalls_20261125',6),('Cold796PreEventThreeCalls_20261125',3)):
    root=base/name
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    prep=json.loads((root/'Preparation.json').read_bytes());data=struct.pack('<id',1,float(allowance));relative='Input/ArcCallAllowance.bin'
    (root/relative).write_bytes(data);prep['inputs'][relative]=sha256(data).hexdigest().upper()
    data=Path(__file__).read_bytes();(root/Path(__file__).name).write_bytes(data);prep['inputs'][Path(__file__).name]=sha256(data).hexdigest().upper()
    assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
    prep['directed_scope']=dict(case='OperationalPreEventLimitedCalls',actual_import_this_pass=[],actual_component_count=256,
        immutable_compiled794_reused=True,fresh_compilation_claimed=False,explicit_call_allowance=allowance,
        original_source10_material_and_force_tolerances_exact794=True,physical_capacity_certified=False,no_working_promotion=True)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print(name,'ALLOWANCE',allowance,'COMPILED794 EXACT; NO FRESH COMPILATION')
