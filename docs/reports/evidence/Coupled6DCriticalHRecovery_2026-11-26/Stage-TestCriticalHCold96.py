"""ДЛЯ ТЕСТОВ. Запускает тот же проверенный критический H-контроль878 с нуля; расширяется только окно наблюдения, не численные бюджеты."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold878CriticalHRecoveryPairs_20261126';root=base/'Cold880CriticalHCold96_20261126'
proof=json.loads((parent/'CriticalHRecoveryPairsProof.json').read_bytes());assert proof['verified_working_H_pairs']==1 and proof['verified_case1_full_half_stresses_and_operators_bit_identical876']
reg=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](base/'Cold879CriticalHRegressions_20261126','OperationalArcRegressions')[0]
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
driver=(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig');body,=re.findall(r'(?ms)^Public Function RunTestStockHTrace\(.*?^End Function',driver)
assert 'RunTestStockHTrace = RunTestCoupledWorkbookTiming("Both", "Excel")' in body and 'mReverse' not in body and 'Input' not in body
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());runner=(root/'Run-TestHMixedStock.ps1').read_text('utf-8-sig');old='-TimeoutSeconds 600';assert runner.count(old)==1
for name,data in {'tools/build_workbook/Run-Audit02ModeTests.ps1':(base/'Cold857SpatialDeclineRecovery_20261125/tools/build_workbook/Run-Audit02ModeTests.ps1').read_bytes(),
 'Run-TestHMixedStock.ps1':runner.replace(old,'-TimeoutSeconds 1200',1).replace('\n','\r\n').encode('utf-8-sig'),
 'PriorDirectedRegressionProof.json':(json.dumps(reg,sort_keys=True,indent=2)+'\n').encode(),
 'SourceCriticalHRecoveryPairsProof.json':(parent/'CriticalHRecoveryPairsProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}.items():
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='CriticalHTrueCold96',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],actual_imports_this_replay=[],
    exact_compiled_source878=True,full_cold_from_zero=True,warm_history_not_passed_to_stock_macro=True,
    global_probe_limit96_unchanged=True,primary_phase_iterations30_unchanged=True,wall_watchdog1200_only=True,
    full_acceptance_suite=False,saved_reopened_results_required=True,no_clean_speed_claim=True,
    generic_three_station_runtime_steel_limit_audit_open=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'ACTUAL COLD;EXACT878;96/30UNCHANGED;1200WALL WATCHDOG;RESULTS REOPEN REQUIRED')
