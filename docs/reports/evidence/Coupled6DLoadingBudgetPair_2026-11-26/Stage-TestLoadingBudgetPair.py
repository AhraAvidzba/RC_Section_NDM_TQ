"""ДЛЯ ТЕСТОВ. Выделяет полной пробе20 из общих30 после доказанного корня за17; остальные критерии пары не меняются."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold979LoadingAlgorithmComparison_20261126';root=base/'Cold980LoadingBudgetPair_20261126'
proof=json.loads((parent/'LoadingAlgorithmComparisonProof.json').read_bytes())
assert proof['cases'][0]['independently_equilibrated'] and proof['cases'][0]['steps']==17
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'LoadingBudgetCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
begin=code.index('Public Function TestRunPublishedMaterialPair()');end=code.index('\nEnd Function',begin)
body=code[begin:end];old='If caseIndex = 1 Then budget = budget \\ 2';new='If caseIndex = 1 Then budget = (2 * budget) \\ 3'
assert body.count(old)==1;body=body.replace(old,new,1)
updates={name:(code[:begin]+body+code[end:]).replace('\n','\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestSolveRecordedYieldCoordinate';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunMaterialDirectionAudit',1).replace('\n','\r\n').encode('utf-8-sig')
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='PositiveMaterialMidpointPairFull20Shared30',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all256_numerical_components_exact979=True,
    only_existing_test_full_phase_budget_changed15_to20=True,total_actual_budget30_preserved=True,
    original_positive_targets_and_failed968_seed_preserved=True,midpoint_half_and_native_publication_and_uncertainty_preserved=True,
    native_H_steel_spatial_guards_preserved=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'FULL PHASE20;TOTAL30;MIDPOINT/PUBLICATION/UNCERTAINTY/DIRECTION GUARDS UNCHANGED')
