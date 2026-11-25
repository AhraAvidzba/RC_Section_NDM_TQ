"""ДЛЯ ТЕСТОВ. Проверяет разделение дугового шага по автоматически найденному началу пластичности без изменения рабочих методов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold773HalfFirstArcWalk_20261125'
witness = base / 'Cold776ArcPointRoute_20261125'
root = base / 'Cold777ArcYieldSplit_20261125'
assert json.loads((witness / 'ArcPointRouteProof.json').read_bytes())['all775_residual_and_stress_path_bytes_exact']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0, witness / 'Input/Reverse1.bin')
archive = root / 'YieldSplitCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'
code = (root / name).read_text('utf-8-sig') + (here / 'TestArcYieldBoundary.cls.txt').read_text('utf-8-sig')
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
name = 'tests/modTestCoupledElasticState.bas'; code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunConfirmedArcWalk\n') == 1
code = code.replace('    TestRunConfirmedArcWalk\n', '    TestRunArcYieldSplit\n', 1)
anchor = '    mSignedArcWalkingTest = True: mSignedArcRatioTest = 0.0078125'
assert code.count(anchor) == 1
code = code.replace(anchor, '    mSignedArcWalkingTest = True: mSignedArcRatioTest = 1#', 1)
code += (here / 'TestArcYieldSplit.bas.txt').read_text('utf-8-sig')
assert max(map(len, code.splitlines())) < 1024
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['Input/ArcRoots.bin'] = (witness / 'Input/ArcRoots.bin').read_bytes()
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
for item in ('TestArcYieldBoundary.cls.txt', 'TestArcYieldSplit.bas.txt', Path(__file__).name):
    values[item] = (here / item).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ArcSplitAtAutomaticallyDetectedYield', actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    numerical_working_methods_exact773=True, only_test_interface_added=True, point_and_station_not_prescribed=True,
    event_equilibrium_uses_existing_projected_solver=True, event_and_half_share30_iterations=True,
    event_strict_force_tolerance=True, subsequent_arc_starts_ratio1=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('777 AUTOMATIC FIRST LOADING-YIELD EVENT; EXISTING SOLVER; STRICT HALF CHECK; POST-EVENT ARC WALK')
