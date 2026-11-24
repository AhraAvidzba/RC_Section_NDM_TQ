"""ДЛЯ ТЕСТОВ. Запускает реальный холодный Both/Excel с последним численным ядром, без осевой истории или диагностического прогноза."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold740CurrentComplexStock_20261125';parent=base/'Cold739PseudoTimeProbe_20261125'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](parent,'PostCrackReplay')
assert native['own_process_closed'] and native['timeout600']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
for name in ('Run-TestHMixedStock.ps1',Path(__file__).name):
    b=(here/name).read_bytes();(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='CurrentColdComplexBothExcel',
    actual_macro='modTestCoupledElasticState.RunTestStockHTrace',
    actual_workflow='RunTestCoupledWorkbookTiming(Both,Excel)',
    source_workbook_immutable_compilation_reused=True,no_new_compilation_or_import=True,
    native_numerical_kernel_exact739=True,no_axial_history_or_external_seed_used=True,
    transverse_reinforcement_and_all_six_loads_from_existing_HardShear_fixture=True,
    original_workflow_settings_and_loadpath_preserved=True,accepted_fields_captured_original_Double=True,
    real_results_reopened_required=True,not_full_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('740 REAL COLD COMPLEX BOTH; EXACT COMPILED CORE739; NO AXIAL SEED; ORIGINAL STOCK MACRO AND REOPEN')
