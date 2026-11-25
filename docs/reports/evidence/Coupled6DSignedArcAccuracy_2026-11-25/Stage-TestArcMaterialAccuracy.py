"""ДЛЯ ТЕСТОВ. Уточняет только локальный допуск возврата CDP для строгого автоматического дугового прохода."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold770ConfirmedArcWalk_20261125'
root = base / 'Cold771ArcMaterialAccuracy_20261125'
proof = json.loads((parent / 'ConfirmedArcWalkProof.json').read_bytes())
assert proof['new_accepted_states'] == 2 and not proof['physical_capacity_certified']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'MaterialAccuracyCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Materials/CCoupledConcreteResponse.cls'
code = (root / name).read_text('utf-8-sig')
old = 'Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000001 '
assert code.count(old) == 1
code = code.replace(old, 'Private Const YIELD_TOLERANCE_RATIO As Double = 0.00000000001 ', 1)
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
assert compiler.count("-ModuleNames 'modTestCoupledElasticState'") == 1
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledConcreteResponse'", 1).replace('\n', '\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ArcWalkWithTighterLocalCDPReturn', actual_import_this_pass=['CCoupledConcreteResponse'],
    local_yield_tolerance_ratio_before=1e-8, local_yield_tolerance_ratio_after=1e-11,
    all_material_physical_parameters_and_diagrams_unchanged=True,
    Newton_controller_testwalk_and_input_prefix_exact770=True,
    accepted_history_replayed_with_tighter_return_not_mutable_state_reuse=True,
    numerical_core_exact761=False, physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('771 LOCAL CDP YIELD TOLERANCE1e-11; SAME770 STRICT WALK/INPUT/DIAGRAMS/PHYSICAL CRITERIA')
