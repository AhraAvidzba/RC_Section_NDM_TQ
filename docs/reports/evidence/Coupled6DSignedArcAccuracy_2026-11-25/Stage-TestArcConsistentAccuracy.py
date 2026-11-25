"""ДЛЯ ТЕСТОВ. Согласует вложенный скалярный возврат с более строгим внешним допуском CDP, сохраняя их исходное отношение100."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold771ArcMaterialAccuracy_20261125'
root = base / 'Cold772ArcConsistentAccuracy_20261125'
closure = json.loads((parent / 'Native/RealColdReplay.txt.ProcessClosure.json').read_bytes())
assert closure['Passed'] and closure['ChildHasExited'] and closure['ChildExitCode'] == 1
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ConsistentAccuracyCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Materials/CCoupledConcreteReturn.cls'
code = (root / name).read_text('utf-8-sig')
old = 'Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000001'
assert code.count(old) == 1
code = code.replace(old, 'Private Const RELATIVE_YIELD_TOLERANCE As Double = 0.0000000000001', 1)
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledConcreteReturn'", 1).replace('\n', '\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ArcWalkWithConsistentNestedCDPAccuracy', actual_import_this_pass=['CCoupledConcreteReturn'],
    scalar_return_tolerance_before=1e-10, scalar_return_tolerance_after=1e-13,
    outer_material_tolerance=1e-11, inner_outer_ratio_matches_original=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('772 NESTED CDP RETURN1e-13 BELOW OUTER1e-11; SAME PHYSICS/WALK; ORIGINAL100 TOLERANCE RATIO')
