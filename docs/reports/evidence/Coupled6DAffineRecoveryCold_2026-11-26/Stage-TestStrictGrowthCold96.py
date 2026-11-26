"""ДЛЯ ТЕСТОВ. Запускает строгий адаптивный H-контроллер с нуля в штатном Both/Excel без передачи исследовательских полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold888StrictEnvelopeGrowth_20261126';root=base/'Cold890StrictGrowthCold96_20261126'
reg=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](base/'Cold889StrictGrowthRegressions_20261126','OperationalArcRegressions')[0]
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes());assert audit['components']==256 and audit['imported_this_pass']==['CCoupledDeformationContinuation']
controller=(parent/'src/Solver/CCoupledDeformationContinuation.cls').read_text('utf-8-sig')
assert controller.count('Private Const PROBE_LIMIT As Long = 96')==1
assert controller.count('verifyGrowth = (point.Coordinate - after > MATERIAL_STEP_GROWTH * (after - before))')==1
assert controller.count('point.Coordinate, point.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True, verifyGrowth)')==1
assert 'If mProbeCount >= PROBE_LIMIT Then' in controller
body,=re.findall(r'(?ms)^Public Function RunTestStockHTrace\(.*?^End Function',(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig'))
assert 'RunTestStockHTrace = RunTestCoupledWorkbookTiming("Both", "Excel")' in body and 'mReverse' not in body and 'Input' not in body
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());runner=(root/'Run-TestHMixedStock.ps1').read_text('utf-8-sig');old='-TimeoutSeconds 600';assert runner.count(old)==1
values={'tools/build_workbook/Run-Audit02ModeTests.ps1':(base/'Cold857SpatialDeclineRecovery_20261125/tools/build_workbook/Run-Audit02ModeTests.ps1').read_bytes(),
        'Run-TestHMixedStock.ps1':runner.replace(old,'-TimeoutSeconds 1200',1).replace('\n','\r\n').encode('utf-8-sig'),
        'PriorDirectedRegressionProof.json':(json.dumps(reg,sort_keys=True,indent=2)+'\n').encode(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='StrictGrowthTrueCold96',actual_import_this_pass=json.loads((parent/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],actual_imports_this_replay=[],
    exact_compiled_source888=True,full_cold_from_zero=True,warm_history_not_passed_to_stock_macro=True,
    global_probe_limit96_unchanged=True,primary_phase_iterations30_unchanged=True,wall_watchdog1200_only=True,
    full_acceptance_suite=False,saved_reopened_results_required=True,no_clean_speed_claim=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'ACTUAL COLD;EXACT888;SAME96/30;MANDATORY PAIR ON GROWTH;RESULTS REOPEN REQUIRED')
