"""ДЛЯ ТЕСТОВ. Записывает причины отказов шести скалярных материальных попыток от проверенной истории19 неизменённым ядром."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold862ColdAcceptedStressContinuity_20261125';root=base/'Cold863Cold19YieldCoordinateCapture_20261125'
proof=json.loads((parent/'ColdAcceptedStressContinuityProof.json').read_bytes());assert proof['maximum_postcommit_concrete_stress_difference_MPa']==0
negative=json.loads((base/'Cold861Cold19AutomaticArcRestart_20261125/AutomaticLoadingWalkProof.json').read_bytes())
assert negative['new_accepted_equilibria']==0 and negative['total_calls']==12
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(root/name).read_text('utf-8-sig')
assert script.count("$null = $excel.Run(\"'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadUnresolvedArcScan\")")==1
script=script.replace('modTestCoupledElasticState.TestReadUnresolvedArcScan','modTestCoupledElasticState.TestDiagnoseNewLoadingCoordinate')
script=script.replace('READONLY_PRIVATE_SCAN','FROZEN_YIELD_DIAGNOSTIC').replace('no new equilibrium search','six independent attempts; no material acceptance')
values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
name='Run-TestRealColdReplay.ps1';script=(root/name).read_text('utf-8-sig');assert script.count('-TimeoutSeconds 600')==1
values[name]=script.replace('-TimeoutSeconds 600','-TimeoutSeconds 1200',1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'SourceColdContinuityProof.json':(parent/'ColdAcceptedStressContinuityProof.json').read_bytes(),
 'SourceArcRestartProof.json':(base/'Cold861Cold19AutomaticArcRestart_20261125/AutomaticLoadingWalkProof.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
original_scope=prep['directed_scope'];prep['directed_scope']=dict(case='Cold19YieldCoordinateCapture',
 actual_import_this_pass=original_scope['actual_import_this_pass'],actual_imports_this_replay=[],immutable_compiled_book_reused=True,
 compiled_source_root='Cold858BalancedInitialStressReplay_20261125',source_states=19,
 only_PS_runner_calls_existing_diagnostic_private_procedure=True,no_VBA_edit_or_new_compile=True,
 six_ratios=[1,.5,.25,.125,.0625,.03125],original30_primary30_recovery_and96calls_per_attempt=True,
 timeout_seconds1200=True,original_materials_and_force_tolerances_unchanged=True,
 no_future_field_or_face_input=True,no_material_acceptance=True,no_new_cold_capacity_claim=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
print(root,'SIX INDEPENDENT RATIOS;EXACT COMPILED858/HISTORY19;CAPTURE ONLY;TIMEOUT1200')
