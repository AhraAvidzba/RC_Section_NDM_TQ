"""ДЛЯ ТЕСТОВ. Проверяет влияние точности только половинного решения на штатную оценку формы, не меняя критерии приёмки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold953DetectedEventApproach_20261126';root=base/'Cold954RefinedEventApproach_20261126'
proof=json.loads((parent/'DetectedEventApproachProof.json').read_bytes());assert proof['native_half_solved'] and not proof['independent_half_verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'RefinedEventApproachCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
entry=(here/'TestRunDetectedEventApproach.bas.txt').read_text('utf-8-sig')
assert entry.count('TestRunDetectedEventApproach')==2
entry=entry.replace('TestRunDetectedEventApproach','TestRunRefinedEventApproach')
old='halfLambda, CDbl(context(3)), CDbl(context(4)), CLng(context(5) - context(6))';assert entry.count(old)==1
entry=entry.replace(old,'halfLambda, 0.0001 * CDbl(context(3)), 0.0001 * CDbl(context(4)), CLng(context(5) - context(6))',1)
entry=entry.replace("    ' Материальная история не принимается независимо от результата проверки локальности.","    ' Точнее решается только половинная проба; исходные силовые критерии, CurveLambdaBound и допуск формы не ослабляются.")
entry=entry.replace('TOTAL_DETECTED_EVENT_APPROACH_TEST','TOTAL_REFINED_EVENT_APPROACH_TEST')
updates={};name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*16+'\n\nPublic Function TestRunDetectedEventApproach()';assert code.count(anchor)==1
updates[name]=code.replace(anchor,prefix*17+'\n\n'+entry.rstrip()+'\n\n\nPublic Function TestRunDetectedEventApproach()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunRefinedEventApproach'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunDetectedEventApproach';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunRefinedEventApproach',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunRefinedEventApproach.bas.txt':entry.replace('\n','\r\n').encode('utf-8-sig')})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.bas'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OnlyHalfEquilibriumSolvedMoreAccurately',actual_import_this_pass=['modTestCoupledElasticState'],
    actual_component_count=257,all_numerical_classes_exact953=True,half_solve_tolerance_factor=.0001,
    original_force_acceptance_and_relative_lambda_noise_and_half_shape_guard_unchanged=True,
    event_source952_unchanged=True,source_history25_unchanged=True,event2_plus_half_within_total30=True,
    diagnostic_precision_change_explicit=True,no_material_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'HALF FORCE SOLVE TOLERANCE1e-4;ALL ORIGINAL ACCEPTANCE GUARDS UNCHANGED')
