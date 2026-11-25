"""ДЛЯ ТЕСТОВ. Уточняет только точность двух дуговых решений, сохраняя их цели, прогнозы, историю и исходный критерий вывода."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold763SignedArcRatio00625_20261125'
root = base / 'Cold764SignedArcTightPair_20261125'
proof = json.loads((parent / 'SignedArcPairProof.json').read_bytes())
assert proof['full_solved'] and proof['half_solved'] and not proof['locally_confirmed_by_half_step']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'TightPairCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
old = '    fullBudget = CLng(mReverseOptionsTest(2)) \\ 2: prediction = fullSeed'
new = '''    Dim solveAbsolute As Double, solveRelative As Double, toleranceFactor As Double
    toleranceFactor = 0.0001: solveAbsolute = mReverseOptionsTest(3) * toleranceFactor: solveRelative = mReverseOptionsTest(4) * toleranceFactor
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\SignedArc.SolveTolerances.bin" For Binary Access Write As #handle
    TestWriteReverseBinaryVector handle, Array(mReverseOptionsTest(3), mReverseOptionsTest(4), solveAbsolute, solveRelative, toleranceFactor)
    Close #handle
    fullBudget = CLng(mReverseOptionsTest(2)) \\ 2: prediction = fullSeed'''
assert code.count(old) == 1
code = code.replace(old, new, 1)
for output in ('fullLambda', 'halfLambda'):
    old = output + ', mReverseOptionsTest(3), mReverseOptionsTest(4),'
    new = output + ', solveAbsolute, solveRelative,'
    assert code.count(old) == 1
    code = code.replace(old, new, 1)
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
values['Compile-TestExternalPoint.ps1'] = (here / 'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='SignedArcSameTargetsTightSolve', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_core_exact761=True, history_seeds_arc_targets_and_budget_exact763=True,
    local_solve_absolute_and_relative_tolerance_factor=.0001, original_exported_history_options_unchanged=True,
    fresh_compilation_claimed=True, no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('764 SAME763 TARGETS/SEEDS/30BUDGET; ONLY LOCAL PAIR SOLVE TOLERANCE x1e-4; SOURCE HISTORY TOLERANCE UNCHANGED')
