"""ДЛЯ ТЕСТОВ. Различает разрешимое продвижение поля и подтверждённое увеличение нагрузки, сохраняя физические и половинные проверки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold958BoundedIncomingSteps_20261126';root=base/'Cold960ResolvedFieldProgress_20261126'
negative=json.loads((base/'Cold959AccurateIncomingSteps_20261126/AccurateIncomingStepsProof.json').read_bytes())
assert negative['accepted_steps']==1 and not negative['reached_current_material_event']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'ResolvedFieldCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='    If returnedLambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty Then GoTo RejectedReturn'
assert code.count(old)==1
new='''    Dim fieldAdvance As Double, guideAdvance As Double, fieldResolution As Double
    fieldAdvance = coordinate - acceptedCoordinate: guideAdvance = guideCoordinate - acceptedCoordinate
    fieldResolution = 16# * tolerance / mReferenceModulus ' Разрешение поля согласовано с существующим поиском материального события; нагрузка может изменяться слабее своего разрешения.
    If Abs(fieldAdvance) <= fieldResolution Or fieldAdvance * guideAdvance <= 0# Then GoTo RejectedReturn
    Dim loadIncreaseResolved As Boolean
    loadIncreaseResolved = returnedLambda > referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty
    If Not loadIncreaseResolved Then
        If Abs(fieldAdvance) > Abs(guideAdvance) + fieldResolution Then GoTo RejectedReturn
        If SpatialFieldDistance(mAccepted, field) > SpatialFieldDistance(mAccepted, guide) + fieldResolution Then GoTo RejectedReturn
        details = details & "; принят шаг продолжения по разрешимому полю; увеличение lambda отдельно не подтверждено"
    End If
    details = details & "; продвижение поля / разрешение: " & CStr(fieldAdvance) & " / " & CStr(fieldResolution) & _
        "; увеличение lambda разрешимо: " & CStr(loadIncreaseResolved) & "; неопределённость lambda: " & CStr(uncertainty)'''
code=code.replace(old,new,1);updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
prefix="' FOR TESTS: baseline entry point only; existing methods retained.\n"
anchor=prefix*21+'\n\nPublic Function TestRunBoundedIncomingSteps()';assert code.count(anchor)==1
entry=(here/'TestRunFieldProgressSteps.bas.txt').read_text('utf-8-sig').rstrip()
updates[name]=code.replace(anchor,prefix*22+'\n\n'+entry+'\n\n\nPublic Function TestRunBoundedIncomingSteps()',1).replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver','modTestCoupledElasticState' -EntryPointNames 'TestRunFieldProgressSteps'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(root/name).read_text('utf-8-sig')
old='modTestCoupledElasticState.TestRunBoundedIncomingSteps';assert runner.count(old)==1
updates[name]=runner.replace(old,'modTestCoupledElasticState.TestRunFieldProgressSteps',1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes(),
    'TestRunFieldProgressSteps.bas.txt':(here/'TestRunFieldProgressSteps.bas.txt').read_bytes(),
    'FieldProgressCriteriaBlock.txt':new.replace('\n','\r\n').encode('utf-8-sig')})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith(('.bas','.cls')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ResolvedFieldProgressWithoutClaimingResolvedLoadIncrease',actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    actual_component_count=257,all255_other_components_exact958=True,source_base958_not_accuracy_variant959=True,
    original_force_half_shape_positive_lambda_material_and_steel_checks_preserved=True,
    original_lambda_uncertainty_calculation_and_storage_preserved=True,resolved_load_increase_not_required_for_intermediate_continuation=True,
    resolved_directed_field_and_local_guide_radius_required_for_unresolved_load_progress=True,
    shared30_per_accepted_step_preserved=True,input_bitexact958=True,zero_reverse_and_exhausted_budget_guards_required=True,
    original25_restoration_required=True,no_production_history_acceptance=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'RESOLVED DIRECTED FIELD CONTINUATION;LOAD UNCERTAINTY RETAINED;NOT CAPACITY;ORIGINAL PHYSICAL/HALF GUARDS')
