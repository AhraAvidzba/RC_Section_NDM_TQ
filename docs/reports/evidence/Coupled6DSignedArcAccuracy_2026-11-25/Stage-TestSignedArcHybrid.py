"""ДЛЯ ТЕСТОВ. Подключает к неизменной дуговой плоскости общий разностный резерв без смены материала или ветви управления."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold760SignedArcPair_20261125'
root = base / 'Cold761SignedArcHybrid_20261125'
proof = json.loads((parent / 'SignedArcPairProof.json').read_bytes())
assert not proof['full_solved'] and not proof['half_solved'] and proof['arc_plane_and_predictors_independently_verified']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ArcHybridCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'
code = (root / name).read_text('utf-8-sig')
old = '    If Not mControlActive Or mArcLoadProjection <> 0# Or Not SupportsHybridRecovery Then Exit Function'
new = '''    If Not mControlActive Or Not SupportsHybridRecovery Then Exit Function
    If mArcLoadProjection <> 0# And referenceLambda <> 0# Then Exit Function ' Дуговой резерв не меняет меру и не требует увеличения lambda. При referenceLambda=0 материальные резервы ниже не включаются.'''
assert code.count(old) == 1
code = code.replace(old, new, 1)
old_comment = "' Для спектральных границ ASD и псевдодуги сохраняются их собственные правила продолжения."
new_comment = "' Для псевдодуги сохраняется исходная плоскость без смены управляющей меры; спектральные границы ASD используют собственные правила."
assert code.count(old_comment) == 1
code = code.replace(old_comment, new_comment, 1)
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
old = '        consistentControl:=True, initialField:=prediction, loadProjection:=q)'
new = '        consistentControl:=True, initialField:=prediction, loadProjection:=q, allowHybridRecovery:=True, conditionHybridRecovery:=True)'
assert code.count(old) == 2
code = code.replace(old, new)
assert max(map(len, code.splitlines())) < 1024
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='SignedArcPairWithHybridFD', actual_import_this_pass=['CCoupledNewtonSolver', 'modTestCoupledElasticState'],
    source35_arcplane_seeds_and_step_exact760=True, material_and_physical_criteria_exact760=True,
    general_hybrid_allowed_for_arc_only_without_reference_lambda=True, branch_changing_material_recovery_disabled_for_arc=True,
    original_arc_load_projection_and_target_preserved=True, same_primary_pair_budget30=True,
    no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('761 ARC HYBRID FULL FD; SAME760 PLANE/HISTORY/SEEDS/30BUDGET; NO BRANCH-CHANGING RECOVERY')
