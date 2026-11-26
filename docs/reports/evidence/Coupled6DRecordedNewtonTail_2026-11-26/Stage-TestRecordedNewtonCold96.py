"""ДЛЯ ТЕСТОВ. Повторяет один холодный920 с автономной записью отклонённых систем; старые исследовательские режимы не включает."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold922ColdNewtonCapture_20261126';root=base/'Cold923RecordedNewtonCold96_20261126'
audit=json.loads((parent/'PointTraceSourceComposition.json').read_bytes())
assert audit['components']==257 and audit['imported_this_pass']==['CCoupledNewtonSolver','modTestColdNewtonCapture']
reg=json.loads((base/'Cold919AffineRecoveryRegressions_20261126/DirectedRegressionProof.json').read_bytes())
assert reg==dict(passed=164,failed=0,macros=5,timeout600=False,own_process_closed=True)
name='src/Solver/CCoupledNewtonSolver.cls'
assert (parent/name).read_text('utf-8-sig').replace('modTestColdNewtonCapture.TestCapture','modTestCoupledElasticState.TestNewtonRejectedCapture')==(base/'Cold918AffineBranchRecovery_20261126'/name).read_text('utf-8-sig')
body,=re.findall(r'(?ms)^Public Function RunTestStockHTrace\(.*?^End Function',(parent/'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig'))
assert 'RunTestStockHTrace = RunTestCoupledWorkbookTiming("Both", "Excel")' in body and 'mReverse' not in body and 'Input' not in body
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());name='tools/build_workbook/Run-Audit02ModeTests.ps1'
runner=(base/'Cold920AffineRecoveryCold96_20261126'/name).read_text('utf-8-sig')
anchor='        if ($MacroArgument2) { $result = [string]$excel.Run($macroName, $MacroArgument1, $MacroArgument2) }'
assert runner.count(anchor)==1
runner=runner.replace(anchor,'''        if ($macro -ne 'modTestCoupledElasticState.RunTestStockHTrace') {throw 'Unexpected instrumented cold macro.'}
        $null=$excel.Run("'RC_Section_NDM.xlsm'!modTestColdNewtonCapture.TestSetCaptureRoot",$MacroArgument1)
        $lines.Add('INDEPENDENT_NEWTON_CAPTURE_SET: no reverse-history or other test modes enabled')
'''+anchor,1)
stock=(root/'Run-TestHMixedStock.ps1').read_text('utf-8-sig');assert stock.count('-TimeoutSeconds 600')==1
values={name:runner.replace('\n','\r\n').encode('utf-8-sig'),
 'Run-TestHMixedStock.ps1':stock.replace('-TimeoutSeconds 600','-TimeoutSeconds 1200',1).replace('\n','\r\n').encode('utf-8-sig'),
 'PriorDirectedRegressionProof.json':(json.dumps(reg,sort_keys=True,indent=2)+'\n').encode(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='RecordedNewtonTrueCold96',actual_import_this_pass=['CCoupledNewtonSolver','modTestColdNewtonCapture'],
 actual_imports_this_replay=[],full_cold_from_zero=True,warm_history_not_passed_to_stock_macro=True,
 original_solver_numerical_body918=True,all25_history_frames_must_be_bit_identical920=True,
 global_probe_limit96_unchanged=True,primary_phase_iterations30_unchanged=True,watchdog1200_only=True,
 saved_reopened_results_required=True,old_test_module_and_modes_unchanged=True,
 full_acceptance_suite=False,physical_capacity_certified=False,no_clean_speed_claim=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ACTUAL COLD;EXACT NUMERICS918;AUTONOMOUS RECORDING;SAME96/30;257 SAVED SOURCES REQUIRED')
