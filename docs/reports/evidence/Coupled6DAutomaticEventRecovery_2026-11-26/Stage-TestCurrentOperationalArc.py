"""ДЛЯ ТЕСТОВ. Вызывает существующую дуговую операцию на точной истории25 с её собственным прогнозом без новой компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold960ResolvedFieldProgress_20261126';root=base/'Cold962CurrentOperationalArc_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
updates={'Input/OperationalArcRatio.bin':vector([1.]),'Input/ArcCallAllowance.bin':vector([96.]),Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';code=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunFieldProgressSteps';assert code.count(old)==1
updates[name]=code.replace(old,'modTestCoupledElasticState.TestRunOperationalArc',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingOperationalArcOnExactCurrent25History',actual_import_this_pass=[],actual_component_count=257,
    all257_actual_components_and_saved_book_bitexact960=True,existing_TestRunOperationalArc_reused=True,
    no_new_numerical_code_or_compilation=True,actual_accepted25_history_only=True,no_supplied_future_guide_or_field=True,
    arc_ratio=1.,arc_call_allowance=96,original30_per_phase_preserved=True,
    call_allowance_is_single_operation_not_remaining_budget_of_original_full_cold=True,
    internal_solve_accuracy_scale_0001_and_current_acceptance_unchanged=True,
    no_material_history_acceptance_inside_arc=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EXISTING OPERATIONAL ARC;OWN PREDICTION;25 REAL STATES;RATIO1;96 CALL ALLOWANCE;ALL CODE EXACT960')
