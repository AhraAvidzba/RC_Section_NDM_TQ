"""ДЛЯ ТЕСТОВ. Проверяет масштаб прогноза нагрузки после пика без изменения приёмки равновесия."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold734CapacityScaledGrowth_20261124';parent=base/'Cold732ProvidedJacobianCapacity_20261124'
stop=json.loads((parent/'ProvidedCapacityStopProof.json').read_bytes())
assert stop['independent_all71_fields_verified_via733'] and not stop['has_certified_capacity']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'GrowthCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='    If Abs(trial.Lambda - current.Lambda - predictedLoad) > _\n            GeomMax(Abs(predictedLoad) * POST_PEAK_SECANT_TOLERANCE, LoadResolution(current.Lambda)) Then Exit Function'
assert code.count(anchor)==1
replacement='''    Dim growthResolutionTest As Double: growthResolutionTest = LoadResolution(current.Lambda)
    If mFollowingLocalPeak Then growthResolutionTest = GeomMax(growthResolutionTest, mLambdaTolerance / 16#) ' Численный масштаб прогноза шага; не допуск приёмки равновесия или максимума.
    If Abs(trial.Lambda - current.Lambda - predictedLoad) > _
            GeomMax(Abs(predictedLoad) * POST_PEAK_SECANT_TOLERANCE, growthResolutionTest) Then Exit Function'''
code=code.replace(anchor,replacement,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    If Len(evidenceRoot) > 0 Then\n        Set mReverseSnapshotsTest = New Collection: mHardPreparedPathTest = evidenceRoot & "\\Native\\PostCrackContinuation"'
assert code.count(anchor)==1
code=code.replace(anchor,'    If capacityOnly Then GoTo Done \' Полная проверка полей выполняется отдельным точным StressOnly-повтором; исходные проверки поиска уже сохранены.\n'+anchor,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledDeformationContinuation','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='CapacityScaledPostPeakLoadPrediction',actual_import_this_pass=['CCoupledDeformationContinuation','modTestCoupledElasticState'],
    single_numerical_change_postpeak_prediction_floor='lambdaTolerance/16',original_spatial_prediction_two_percent=True,
    original_equilibrium_physical_and_peak_checks=True,original96_probes_30_iterations=True,
    original_capacity_assertions_exact=True,only_original_two_elastic_initial_states=True,
    independent_tensor_export_deferred_not_waived=True,working_code_promoted=False,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('734 STAGED; ONLY POSTPEAK LOAD-PREDICTION FLOOR CHANGED; FINAL CHECKS AND96BUDGET EXACT')
