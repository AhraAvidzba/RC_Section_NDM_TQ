"""ДЛЯ ТЕСТОВ. Удваивает только исследовательский общий бюджет холодного продолжения для проверки физического продвижения."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold851LoadingSecantReparameterization_20261125';negative=base/'Cold853LoadingSecantCold_20261125'
root=base/'Cold854ExtendedColdBudget192_20261125'
proof=json.loads((negative/'LoadingSecantColdProof.json').read_bytes())
assert proof['cold']['capacity']=='NumFail' and proof['cold832_fields_fractions_lambdas_bit_identical']
assert proof['cold']['native']['own_process_closed']
audit=json.loads((base/'Cold852AutomaticLoadingSecantWalk_20261125/ContinuationProgressAudit.json').read_bytes())
assert audit['cold_budget']['stop_by_global_probe_limit'] and audit['cold_budget']['last_arc_restore_succeeded']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ExtendedColdBudgetCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_bytes()
old=b'Private Const PROBE_LIMIT As Long = 96';new=b'Private Const PROBE_LIMIT As Long = 192'
assert code.count(old)==1
values[name]=code.replace(old,new,1);assert values[name].replace(new,old,1)==(parent/name).read_bytes()
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,"-ModuleNames 'CCoupledDeformationContinuation'",1).replace('\n','\r\n').encode('utf-8-sig')
runner=(parent/'Run-TestHMixedStock.ps1').read_text('utf-8-sig')
old='-TimeoutSeconds 600';assert runner.count(old)==1
values['Run-TestHMixedStock.ps1']=runner.replace(old,'-TimeoutSeconds 1200',1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'Negative96ColdProof.json':(negative/'LoadingSecantColdProof.json').read_bytes(),
    'Negative96ProgressAudit.json':(base/'Cold852AutomaticLoadingSecantWalk_20261125/ContinuationProgressAudit.json').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExtendedColdBudget192AgainstExact96Control',actual_import_this_pass=['CCoupledDeformationContinuation'],actual_component_count=256,
    source_parent851=True,negative_control853_96_kept=True,only_one_controller_constant_changed=True,
    default_working_and_release_budget_not_changed=True,research_probe_limit=192,negative_probe_limit=96,
    primary_iterations_per_phase30_unchanged=True,all_material_force_half_locality_and_physical_criteria_unchanged=True,
    full_cold_from_zero=True,only_one_HardShear_case=True,warm_fields_not_supplied_to_stock_solver=True,
    timeout_seconds1200=True,actual_saved_reopened_results_required=True,not_clean_performance_measurement=True,
    not_a_speedup_or_capacity_claim=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'RESEARCH192;NEGATIVE96KEPT;ONE CONSTANT;SAME ALL6/LAWS/TOLERANCES;TIMEOUT1200')
