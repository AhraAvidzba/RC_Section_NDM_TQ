"""ДЛЯ ТЕСТОВ. Проверяет сложное сочетание с ранее проверенным порядком локальных методов: обычный возврат, затем полный скалярный резерв."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold741OrdinaryFirstComplex_20261125';parent=base/'Cold740CurrentComplexStock_20261125'
reference=base/'Cold711PriorDilationCurve_20261124'
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](parent,'HBackendStock')
assert native['timeout600'] and native['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'OrdinaryFirstCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='src/Materials/CCoupledConcreteResponse.cls';data=(reference/name).read_bytes()
assert b'CanPredictDilationBranch' not in data and b'TryDilationClosure' in data
values[name]=data
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledConcreteResponse'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values['Run-TestPriorBranchHardShear.ps1']=(here/'Run-TestPriorBranchHardShear.ps1').read_bytes()
values['Run-TestPriorBranchMaterials.ps1']=(here/'Run-TestPriorBranchMaterials.ps1').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,b in values.items():(root/name).write_bytes(b);prep['inputs'][name]=sha256(b).hexdigest().upper()
prep['directed_scope'].update(case='OrdinaryFirstReturnComplexPair',actual_import_this_pass=['CCoupledConcreteResponse'],
    material_source_exact_previously_verified711=True,material_source_reference=str(reference),
    only_material_numerical_order_changed_to_reference=True,newton_controller_and_workflow_exact740=True,
    full_local_yield_flow_and_global_tolerances_unchanged=True,
    directed_material_regressions_then_real_Direct_and_Both_required=True,
    no_frozen_axial_seed_in_real_workflow=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('741 EXACT MATERIAL711 WITH LATEST NEWTON/CONTROLLER; ORDINARY FIRST, FULL CLOSURE ONLY AS FALLBACK')
