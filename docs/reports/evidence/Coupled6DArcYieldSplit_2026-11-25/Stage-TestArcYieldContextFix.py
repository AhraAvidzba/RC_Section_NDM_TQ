"""ДЛЯ ТЕСТОВ. Исправляет подготовку читающей управляющей строки тестового поиска события, сохраняя рабочую математику."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold777ArcYieldSplit_20261125'
root = base / 'Cold778ArcYieldContextFix_20261125'
closure = json.loads((parent / 'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode'] == 1
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'YieldContextCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'; code = (root / name).read_text('utf-8-sig')
old = (parent / 'TestArcYieldBoundary.cls.txt').read_text('utf-8-sig')
new = (here / 'TestArcYieldBoundary.cls.txt').read_text('utf-8-sig')
assert code.endswith(old)
code = code[:-len(old)] + new
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'", 1).replace('\n', '\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(),
    'TestArcYieldBoundary.cls.txt': (here / 'TestArcYieldBoundary.cls.txt').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ArcYieldTestContextInitializationFix', actual_import_this_pass=['CCoupledNewtonSolver'],
    full_degree_control_projection_initialized=True, numerical_working_methods_exact773=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('778 TEST CONTEXT ARRAY INITIALIZED; SAME777 EVENT ALGORITHM AND PHYSICS')
