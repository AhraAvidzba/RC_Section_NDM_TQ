"""ДЛЯ ТЕСТОВ. Добавляет третий, более локальный разностный шаг после застоя проектируемого события, не увеличивая бюджеты."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import re

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold780RejectedYieldEvent_20261125'
root = base / 'Cold782ProjectedEventRefinement_20261125'
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ProjectedRefinementCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'; code = (root / name).read_text('utf-8-sig')
method, = re.findall(r'(?ms)^Private Function FollowYieldEventCoordinate\(.*?^End Function', code)
assert method.count('For attempt = 0 To 1') == method.count('Or attempt = 1 Then Exit For') == 1
updated = method.replace('For attempt = 0 To 1', 'For attempt = 0 To 2', 1).replace('Or attempt = 1 Then Exit For', 'Or attempt = 2 Then Exit For', 1)
updated = updated.replace("' После застоя уточняется только разностная матрица от той же истории; остаток общего бюджета сохраняется.", "' При застое матрица уточняется до двух раз от той же истории; все попытки делят исходные остатки итераций и вычислений.", 1)
code = code.replace(method, updated, 1)
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'", 1).replace('\n', '\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ProjectedEventThirdDifferenceRefinement', actual_import_this_pass=['CCoupledNewtonSolver'],
    actual780_failed_history_bracket_and_tolerances_unchanged=True, difference_steps=[1e-6,1e-7,1e-8],
    original_primary_and_function_budgets_shared=True, material_and_force_gauge_yield_acceptance_unchanged=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('782 PROJECTED EVENT DIFFERENCES1e-6/1e-7/1e-8; SAME780 CASE; SHARED ORIGINAL ITERATION/FUNCTION BUDGETS')
