"""ДЛЯ ТЕСТОВ. Исправляет размер массива шести усилий при успешном восстановлении полного кандидата."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold758PositiveHalfGuard_20261125'
root = base / 'Cold759HalfGuardLoadAllocationFix_20261125'
text = (parent / 'Native/RealColdReplay.txt').read_text('utf-8-sig')
assert 'RUNTIME ERROR' in text and 'Subscript out of range' in text
closure = json.loads((parent / 'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert closure['ChildHasExited'] and closure['ChildExitCode'] == 1 and closure['Passed']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'LoadFixCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'
code = (root / name).read_text('utf-8-sig')
old = 'Dim tensorError As Double, lambdaError As Double, length As Double, lambdaBound As Double, load() As Double, tolerance As Double'
new = 'Dim tensorError As Double, lambdaError As Double, length As Double, lambdaBound As Double, load(1 To COMPONENT_COUNT) As Double, tolerance As Double'
assert code.count(old) == 1
changed = code.replace(old, new, 1)
values = {name: changed.replace('\n', '\r\n').encode('utf-8-sig')}
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='HalfGuardSuccessfulPathLoadAllocationFix', actual_import_this_pass=['CCoupledNewtonSolver'],
    only_six_component_load_array_allocation_corrected=True, material_controller_test_module_exact758=True,
    failed758_native_and_own_closure_preserved=True, original_half_consistency_and_primary_budget_unchanged=True,
    no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('759 SINGLE ARRAY ALLOCATION FIX; SAME POSITIVE758 TEST; FAILURE PRESERVED')
