"""ДЛЯ ТЕСТОВ. Подготавливает один штатный холодный расчёт с проверенной автоматической исходящей координатой без новой компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold823AutomaticLoadingWalk_20261125';root=base/'Cold824AutomaticLoadingCold_20261125'
proof=json.loads((parent/'AutomaticLoadingWalkProof.json').read_bytes())
assert proof['new_accepted_equilibria']==8 and proof['all_half_roots_independently_rebuilt']
check=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative']
for label,count in [('PriorBranchMaterials',87),('OperationalArcRegressions',164)]:
    native,_=check(parent,label);assert native['passed']==count and not native['failed'] and native['own_process_closed'] and not native['timeout600']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name=Path(__file__).name
(root/name).write_bytes(Path(__file__).read_bytes());prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
prep['directed_scope']=dict(case='AutomaticLoadingColdFromZero',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    immutable_compiled_source_exact823=True,fresh_compilation_claimed=False,full_cold_from_zero=True,
    only_one_HardShear_case=True,warm_fields_not_supplied_to_stock_solver=True,
    all_primary30_global96_physical_and_full_half_criteria_unchanged=True,
    actual_saved_reopened_results_required=True,timeout_seconds600=True,not_clean_performance_measurement=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
