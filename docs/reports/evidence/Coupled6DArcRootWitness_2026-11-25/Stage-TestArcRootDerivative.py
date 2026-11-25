"""ДЛЯ ТЕСТОВ. Читает фактические производные и остаточные поправки полных/половинных корней770 от одной точной истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import struct

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold765AcceptedMapContinuity_20261125'
walk = base / 'Cold770ConfirmedArcWalk_20261125'
root = base / 'Cold774ArcRootDerivative_20261125'
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
read = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary']
source = read(walk / 'Input/Reverse1.bin')
first = Reader((walk / 'Native/ArcWalk.001.bin').read_bytes())
first_vectors = [first.VectorTest() for _ in range(13)]
second = Reader((walk / 'Native/ArcWalk.002.bin').read_bytes())
v = [second.VectorTest() for _ in range(13)]
assert first.position == len(first.raw) and second.position == len(second.raw)
assert first_vectors[0][0] == first_vectors[0][1] == v[0][0] == v[0][1] == 1
prefix = {key: list(source[key]) for key in ('fields', 'fractions', 'lambdas')}
prefix['fields'].append(first_vectors[9]); prefix['fractions'].append(1.); prefix['lambdas'].append(first_vectors[0][2])
assert len(prefix['fields']) == 33
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ArcRootCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'
code = (root / name).read_text('utf-8-sig') + '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Открывает неизменную дуговую систему для чтения невязки и касательной; не ищет НДС и не принимает материалы.
Friend Function BeginArcRootTest(ByRef baseLoad() As Double, ByRef direction() As Double, ByRef projection() As Double, _
        ByVal target As Double, ByRef seed() As Double, ByVal factor As Double, ByVal loadProjection As Double) As Boolean
    If Not BeginNewtonLineTest(baseLoad, direction, projection, target, seed, factor) Then Exit Function
    mArcLoadProjection = loadProjection: mConsistentControl = True
    BeginArcRootTest = True
End Function
'''
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunAcceptedMapContinuity\n') == 1
code = code.replace('    TestRunAcceptedMapContinuity\n', '    TestRunArcRootDerivative\n', 1)
code += (here / 'TestArcRootDerivative.bas.txt').read_text('utf-8-sig')
assert max(map(len, code.splitlines())) < 1024
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
binary = bytearray()
for vector in (v[9], v[10], v[5], [v[0][2], v[0][3], v[0][10], v[0][11], v[0][9], v[0][4]]):
    binary += struct.pack('<i' + 'd' * len(vector), len(vector), *vector)
values['Input/ArcRoots.bin'] = binary
values['Input/Reverse1.bin'] = run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source, prefix)
values['Input/SourceArcWalk001.bin'] = (walk / 'Native/ArcWalk.001.bin').read_bytes()
values['Input/SourceArcWalk002.bin'] = (walk / 'Native/ArcWalk.002.bin').read_bytes()
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values['TestArcRootDerivative.bas.txt'] = (here / 'TestArcRootDerivative.bas.txt').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualFullHalfArcRootDerivative', actual_import_this_pass=['CCoupledNewtonSolver','modTestCoupledElasticState'],
    only_test_interface_added_to_numeric_kernel=True, numerical_working_methods_exact761=True,
    exact_770_prefix33_and_unaccepted_attempt2_roots=True, new_equilibrium_search=False,
    two_root_difference_steps=[1e-6,1e-8], expected_difference_samples=492,
    accepted_revision_unchanged_required=True, physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('774 ACTUAL ARC ROOT DERIVATIVES; EXACT33 HISTORY;492 READONLY SAMPLES; NO NEW NDS SEARCH')
