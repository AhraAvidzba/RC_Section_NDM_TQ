"""ДЛЯ ТЕСТОВ. Проверяет существующие ограничения направления перед использованием опубликованных976 корней в поиске предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold976PublishedMaterialPair_20261126';root=base/'Cold977MaterialDirectionAudit_20261126'
proof=json.loads((parent/'PublishedMaterialPairProof.json').read_bytes())
assert proof['all_eligible_fields_published_unchanged'] and proof['all_publications_have_checked_native_lambda_uncertainty']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'DirectionAuditCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
entry=(here/'TestRunMaterialDirectionAudit.bas.txt').read_text('utf-8-sig').rstrip()
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n";anchor=prefix*25+'\n\nPublic Function TestRunPublishedMaterialPair()'
assert code.count(anchor)==1
updates={name:code.replace(anchor,prefix*26+'\n\n'+entry+'\n\n\nPublic Function TestRunPublishedMaterialPair()',1).replace('\n','\r\n').encode('utf-8-sig')}
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunMaterialDirectionAudit'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig');old='modTestCoupledElasticState.TestRunPublishedMaterialPair';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunMaterialDirectionAudit',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunMaterialDirectionAudit.bas.txt':(here/'TestRunMaterialDirectionAudit.bas.txt').read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingEnvelopeAndSpatialOrientationGuardsOnPublishedMaterialRoots',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all256_numerical_components_exact976=True,only_new_test_entry_added=True,
    original_published_pair_inputs_and_solver_exact976=True,existing_TrialEnvelopeCoordinate_H_and_HasSteelPlasticIncrement_used=True,
    existing_SpatialIncrementAdvances_and_IsEnvelopeUnloading_rule_measured=True,no_material_history_acceptance=True,
    no_future_equilibrium_field_input=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'NEW READONLY DIRECTION AUDIT;256 NUMERICAL COMPONENTS EXACT976;H/STEEL/SPATIAL GUARDS;NO RECOVERY PROMOTION')
