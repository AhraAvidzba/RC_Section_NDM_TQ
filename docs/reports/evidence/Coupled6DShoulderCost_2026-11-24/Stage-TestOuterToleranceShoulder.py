"""ДЛЯ ТЕСТОВ. Проверяет остановку внутреннего резерва с запасом к полному допуску сечения; конечная приёмка неизменна."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold731OuterToleranceShoulder_20261124';parent=base/'Cold730ProvidedJacobianShoulder_20261124'
assert json.loads((parent/'ShoulderWitnessProof.json').read_bytes())['quality_confirmed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'OuterToleranceCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='''        converged = hybrid.TrySolveLeastSquaresTrust(equations, coordinates, ordinaryCalls, ordinarySteps, _
            absoluteTolerance, solution, differenceStepTest)''';assert code.count(anchor)==1
code=code.replace(anchor,'''        Dim innerToleranceTest As Double
        SetControlledLoad initialLambda, load
        innerToleranceTest = GeomMax(absoluteTolerance, 0.5 * EquilibriumTolerance(absoluteTolerance, relativeTolerance))
        converged = hybrid.TrySolveLeastSquaresTrust(equations, coordinates, ordinaryCalls, ordinarySteps, _
            innerToleranceTest, solution, differenceStepTest)''',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='InnerStopAlignedWithOuterTolerance',actual_import_this_pass=['CCoupledNewtonSolver'],
    only_inner_stop_threshold_changed=True,inner_threshold_half_seed_full_tolerance_with_absolute_floor=True,
    full_final_equilibrium_gauge_control_and_lambda_uncertainty_acceptance_unchanged=True,
    same_original52prefix_material_and_targets=True,same_30_iterations_and_three_quality_attempts=True,
    not_production_ready_without_fallback_if_early_candidate_rejected=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('731 INNER STOP WITH 50% OUTER TOLERANCE MARGIN; ALL FINAL ACCEPTANCE AND QUALITY UNCHANGED')
