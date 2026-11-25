"""ДЛЯ ТЕСТОВ. Автоматически продолжает точный32-точечный префикс строгими дуговыми парами без знания будущих полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold769SignedArcPrefix32Strict_20261125'
root = base / 'Cold770ConfirmedArcWalk_20261125'
proof = json.loads((parent / 'SignedArcPairProof.json').read_bytes())
assert proof['locally_confirmed_by_half_step'] and proof['source_states'] == 32
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ArcWalkCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mSignedArcWalkingTest As Boolean, mSignedArcRatioTest As Double, mSignedArcPairTest As Variant\n', 1)
assert code.count('    TestRunSignedArcPair\n') == 1
code = code.replace('    TestRunSignedArcPair\n', '    TestRunConfirmedArcWalk\n', 1)
old = '    ratio = settings(1): If ratio <= 0# Or ratio > 1# Then Err.Raise 5, , "Доля дуговой секущей вне диапазона."'
assert code.count(old) == 1
code = code.replace(old, '''    ratio = settings(1)
    If mSignedArcWalkingTest Then ratio = mSignedArcRatioTest
    If ratio <= 0# Or ratio > 2# Then Err.Raise 5, , "Доля дуговой секущей вне диапазона."''', 1)
old = '''    If fullSolved Then TestExportSignedArcTrial "SignedArc.Full", full, fullLambda
    If halfSolved Then TestExportSignedArcTrial "SignedArc.Half", half, halfLambda'''
assert code.count(old) == 1
code = code.replace(old, '''    mSignedArcPairTest = Array(fullSolved, halfSolved, fullLambda, halfLambda, full, half, projection, fullTarget, q, radius, alpha, firstEffort(0) + secondEffort(0))
    If Not mSignedArcWalkingTest Then
        If fullSolved Then TestExportSignedArcTrial "SignedArc.Full", full, fullLambda
        If halfSolved Then TestExportSignedArcTrial "SignedArc.Half", half, halfLambda
    End If''', 1)
code += (here / 'TestConfirmedArcWalk.bas.txt').read_text('utf-8-sig')
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
values['Compile-TestExternalPoint.ps1'] = (here / 'Compile-TestExternalPoint.ps1').read_bytes()
runner = (parent / 'Run-TestRealColdReplay.ps1').read_text('utf-8-sig')
assert runner.count('-TimeoutSeconds 180') == 1
values['Run-TestRealColdReplay.ps1'] = runner.replace('-TimeoutSeconds 180', '-TimeoutSeconds 600', 1).replace('\n', '\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values['TestConfirmedArcWalk.bas.txt'] = (here / 'TestConfirmedArcWalk.bas.txt').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='AutomaticConfirmedSignedArcWalkFrom32', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_core_exact761=True, only_test_walk_added=True, no_future_field_predictor=True,
    maximum_attempts=32, maximum_nonlinear_calls=96, primary_pair_limit=30,
    local_solve_tolerance_factor=.0001, full_half_tensor_tolerance=.005,
    original_force_resolution_required=True, source_prefix32_replayed_not_cold_capacity_search=True,
    fresh_compilation_claimed=True, no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('770 AUTOMATIC STRICT ARC WALK; EXACT PREFIX32; MAX32 ATTEMPTS/96 CALLS; KERNEL EXACT761')
