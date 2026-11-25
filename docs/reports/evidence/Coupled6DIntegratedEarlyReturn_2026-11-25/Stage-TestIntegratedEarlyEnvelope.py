"""ДЛЯ ТЕСТОВ. Изолирует сочетание раннего H и проверенного встроенного возврата, без переноса в выпуск."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold752IntegratedGuideDirected_20261125'
root = base / 'Cold754IntegratedEarlyEnvelope_20261125'
proof = json.loads((parent / 'IntegratedGuideDirectedProof.json').read_bytes())
assert proof['independent_NDS_and_original_control_verified'] and proof['total_primary_steps'] == 17
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'IntegratedEarlyCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledDeformationContinuation.cls'
code = (root / name).read_text('utf-8-sig')
old = """    If Not mReinforcement Is Nothing Then
        If mReinforcement.HasSteel Then Exit Sub ' Между изменениями бетона сталь может воспринимать нагрузку при практически постоянном H; начальная геометрическая мера сохраняется.
    End If"""
new = """    If Not mReinforcement Is Nothing Then
        If mReinforcement.HasSteel Then
            If Not mSolver.UsesSpatialHistory Or Not mSolver.SupportsHybridRecovery Then Exit Sub ' Начальная H-параметризация с арматурой проверяется только для пространственного CDP.
        End If
    End If"""
assert code.count(old) == 1
changed = code.replace(old, new, 1)
assert changed.replace(new, old, 1) == code
values = {name: changed.replace('\n', '\r\n').encode('utf-8-sig')}
compiler_helper = base / 'Cold753IntegratedColdFrozenReplay_20261125/tools/build_workbook/Update-IsolatedOptimizerCode.ps1'
helper_bytes = compiler_helper.read_bytes()
assert b'$taskCompileProbe -lt 40' in helper_bytes and b'COMPILE_LOCATION_UNAVAILABLE' in helper_bytes
values['tools/build_workbook/Update-IsolatedOptimizerCode.ps1'] = helper_bytes
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
assert compiler.count("-ModuleNames 'modTestCoupledElasticState'") == 1
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledDeformationContinuation'", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='IntegratedGuideWithEarlyGrowingEnvelope', actual_import_this_pass=['CCoupledDeformationContinuation'],
    material_Newton_and_test_module_exact752=True, only_initial_H_eligibility_changed=True,
    early_H745_without_guide_previously_failed_not_promoted=True, integrated_original_target_return_present=True,
    original_positive_H_growth_and_plateau_fallback_preserved=True, original96probes30iterations_and02_H_cap_preserved=True,
    original_force_lambda_steel_direction_acceptance_preserved=True, no_working_promotion=True,
    compiler_bounded_wait_and_diagnostic_fix_exact753=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('754 RESEARCH COMBINATION ONLY; SAME752 MATERIAL/NEWTON, EARLY H INITIAL GUARD; NO PRODUCTION PROMOTION')
