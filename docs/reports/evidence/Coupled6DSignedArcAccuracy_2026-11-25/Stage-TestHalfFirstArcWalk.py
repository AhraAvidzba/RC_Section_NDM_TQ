"""ДЛЯ ТЕСТОВ. Использует найденный половинный корень как прогноз полной дуги при неизменных уравнениях, материале и бюджете."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold770ConfirmedArcWalk_20261125'
root = base / 'Cold773HalfFirstArcWalk_20261125'
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'HalfFirstCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
start = code.index('    fullBudget = CLng(mReverseOptionsTest(2)) \\ 2: prediction = fullSeed', code.index('Private Sub TestRunSignedArcPair()'))
end = code.index('    If mReverseSolverTest.AcceptedRevision <> revision', start)
block = code[start:end]
divide = block.index('    halfBudget = CLng(mReverseOptionsTest(2) - firstEffort(0)): prediction = halfSeed')
full_part, half_part = block[:divide], block[divide:]
half_part = half_part.replace('halfBudget = CLng(mReverseOptionsTest(2) - firstEffort(0))', 'halfBudget = CLng(mReverseOptionsTest(2)) \\ 2', 1)
full_part = full_part.replace('fullBudget = CLng(mReverseOptionsTest(2)) \\ 2: prediction = fullSeed', 'fullBudget = CLng(mReverseOptionsTest(2) - secondEffort(0)): prediction = fullSeed', 1)
old = 'factor(4) + ratio * lambdaMove, fullLambda,'
assert full_part.count(old) == 1
full_part = full_part.replace(old, 'fullPredictionLambda, fullLambda,', 1)
middle = '''    mReverseSolverTest.DiscardStep
    Dim fullPredictionLambda As Double: fullPredictionLambda = factor(4) + ratio * lambdaMove
    If halfSolved Then
        For i = 1 To mReverseSolverTest.PhysicalDegreeCount: fullSeed(i) = 2# * half(i) - accepted(i): Next i
        fullPredictionLambda = 2# * halfLambda - factor(4) ' Половинный корень остаётся непринятым; прогноз относится к той же полной цели.
    End If
'''
code = code[:start] + half_part + middle + full_part + code[end:]
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1': (here / 'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='AutomaticArcWalkHalfFirstPredictor', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_core_exact761=True, half_root_not_committed=True, half_budget15_full_uses_remainder=True,
    full_seed_extrapolated_from_half_root=True, all_targets_arc_metrics_physical_parameters_and_acceptance_exact770=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('773 HALF FIRST THEN FULL PREDICTOR; SAME761 KERNEL/SIGNED PLANES/STRICT ACCEPTANCE/SHARED30')
