"""ДЛЯ ТЕСТОВ. Проверяет ранний выход по новой границе и избирательное сохранение исходящей дуги одним холодным расчётом."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold831GatedLoadingArc_20261125';root=base/'Cold832GatedLoadingCold_20261125'
proof=json.loads((parent/'RetainedArcContextProof.json').read_bytes())
assert proof['total_context_cases']==15 and proof['gated_after_verified_outgoing_step']
native,_=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](parent,'OperationalArcRegressions')
assert native==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name=Path(__file__).name
(root/name).write_bytes(Path(__file__).read_bytes());prep['inputs'][name]=sha256((root/name).read_bytes()).hexdigest().upper()
prep['directed_scope']=dict(case='GatedLoadingArcColdFromZero',actual_import_this_pass=['CCoupledDeformationContinuation','modTestCoupledElasticState'],
    actual_component_count=256,immutable_compiled_source_exact831=True,fresh_compilation_claimed=False,full_cold_from_zero=True,
    only_one_HardShear_case=True,warm_fields_not_supplied_to_stock_solver=True,
    all_primary30_global96_physical_and_full_half_criteria_unchanged=True,
    proactive_only_unique_recent_loading_boundary=True,retention_only_after_verified_outgoing_kind4=True,
    actual_saved_reopened_results_required=True,timeout_seconds600=True,not_clean_performance_measurement=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
