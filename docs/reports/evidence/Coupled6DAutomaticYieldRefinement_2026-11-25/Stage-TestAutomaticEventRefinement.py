"""ДЛЯ ТЕСТОВ. Соединяет проверенное уточнение события782 с автоматическим проходом779 и фактическим допуском внутренних пар."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import re

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold779AutomaticYieldArcWalk_20261125'
refined = base / 'Cold782ProjectedEventRefinement_20261125'
root = base / 'Cold784AutomaticEventRefinement_20261125'
assert json.loads((refined / 'ProjectedEventRefinementProof.json').read_bytes())['targeted_regressions']['passed'] == 105
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'AutomaticEventRefinementCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'; code = (root / name).read_text('utf-8-sig')
method, = re.findall(r'(?ms)^Private Function FollowYieldEventCoordinate\(.*?^End Function', code)
new_method, = re.findall(r'(?ms)^Private Function FollowYieldEventCoordinate\(.*?^End Function', (refined / name).read_text('utf-8-sig'))
assert new_method.count('For attempt = 0 To 2') == 1
code = code.replace(method, new_method, 1)
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
name = 'tests/modTestCoupledElasticState.bas'; code = (root / name).read_text('utf-8-sig')
anchor = '        tolerance = mReverseOptionsTest(3) + mReverseOptionsTest(4) * TestArcWalkLoadNorm(previousLambda)\n'
assert code.count(anchor) == 1
code = code.replace(anchor, '''        tolerance = 0.0001 * (mReverseOptionsTest(3) + mReverseOptionsTest(4) * GeomMax(TestArcWalkLoadNorm(previousLambda), GeomMax(TestArcWalkLoadNorm(fullLambda), TestArcWalkLoadNorm(halfLambda)))) ' Разрешение соответствует фактическому допуску обоих внутренних решений, а не более грубому пользовательскому допуску.
''', 1)
code = code.replace('Дуговое приращение меньше разрешения исходного равновесия.', 'Дуговое приращение меньше разрешения внутренних решений.')
anchor = '            verified = fieldError <= 0.005 And lambdaError <= lambdaBound And factor > 0# And halfLambda > 0#\n'
assert code.count(anchor) == 2
code = code.replace(anchor, anchor + '''            verified = verified And 0.5 * spacing * mReverseModulusTest > 8# * (absTol + relTol * GeomMax(TestArcWalkLoadNorm(factor), TestArcWalkLoadNorm(halfLambda))) ' Подход к событию должен быть разрешим при фактически применённой внутренней точности.
''')
assert max(map(len, code.splitlines())) < 1024
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')", 1).replace('\n', '\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(),
    Path(__file__).name: Path(__file__).read_bytes()})
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='AutomaticThirdDifferenceEventRefinement', actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    working_method_exact782=True, original779_source33_unchanged=True, resolution_uses_actual_inner_tolerance=True,
    residual_and_gauge_acceptance_unchanged=True, half_step_error_limit_unchanged=True,
    original96_global_and30_phase_budgets_unchanged=True, independent_small_event_uncertainty_checked783=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('784 AUTOMATIC779 WITH782 EVENT REFINEMENT; ACTUAL INNER RESOLUTION; SAME33 SOURCE/96CALL/30PRIMARY BUDGETS')
