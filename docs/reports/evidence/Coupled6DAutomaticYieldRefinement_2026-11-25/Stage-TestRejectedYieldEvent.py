"""ДЛЯ ТЕСТОВ. Воспроизводит один отказ события779 на точной43-точечной истории и раскрывает финальные проверки кандидата."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import struct

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold778ArcYieldContextFix_20261125'
source_root = base / 'Cold779AutomaticYieldArcWalk_20261125'
root = base / 'Cold780RejectedYieldEvent_20261125'
proof = json.loads((source_root / 'ConfirmedArcWalkProof.json').read_bytes())
assert proof['accepted_states'] == 43 and proof['repeated_yield_event_recovery']
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source_root / 'Input/Reverse1.bin')
frame = {key: list(source[key]) for key in ('fields', 'fractions', 'lambdas')}
def TestAdd(path, field_index):
    """Добавляет только побитный подтверждённый полный результат, не используя десятичный CSV."""
    r = Reader(path.read_bytes())
    values = [r.VectorTest() for _ in range(10 if field_index == 5 else 13)]
    frame['fields'].append(values[field_index]); frame['fractions'].append(1.); frame['lambdas'].append(values[0][3 if field_index == 5 else 2])
TestAdd(source_root / 'Native/ArcYieldSplit.Result.bin', 5)
events = {x['attempt']: x for x in proof['event_proofs']}
for attempt in proof['attempts']:
    n = attempt['number']
    if attempt['accepted']:
        TestAdd(source_root / f'Native/ArcWalk.{n:03d}.bin', 9)
    if n in events and events[n]['accepted']:
        r = Reader((source_root / f'Native/AutoEvent.{n:03d}.Result.bin').read_bytes())
        values = [r.VectorTest() for _ in range(11)]
        frame['fields'].append(values[5]); frame['fractions'].append(1.); frame['lambdas'].append(values[0][3])
assert len(frame['fields']) == 43
rr = Reader((source_root / 'Native/ArcWalk.013.bin').read_bytes())
v = [rr.VectorTest() for _ in range(13)]
assert (v[6] == frame['fields'][-1]).all() and v[0][4] == frame['lambdas'][-1]
trajectory = run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source, frame)
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'RejectedEventCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'; code = (root / name).read_text('utf-8-sig')
anchor = '    TestLocateArcYieldBoundary = FollowYieldEventCoordinate(station, seed, mCurrentControlLambda, 0#, tolerance, maxIterations, mAccepted, radius, result, resultLambda, projection)\n'
assert code.count(anchor) == 1
diagnostics = '''    Dim diagnostic() As Double: ReDim diagnostic(0 To 16 + mDegreeCount)
    For i = 0 To 6: diagnostic(i) = eventInfo(i): Next i
    diagnostic(7) = radius: diagnostic(9) = tolerance
    Dim candidateReady As Boolean
    On Error Resume Next
    candidateReady = (LBound(result) = 1 And UBound(result) = mDegreeCount)
    Err.Clear: On Error GoTo 0
    If candidateReady Then
        diagnostic(8) = SpatialFieldDistance(mAccepted, result): diagnostic(15) = resultLambda: diagnostic(16) = 1#
        For i = 1 To mDegreeCount: diagnostic(16 + i) = result(i): Next i
        SetControlledLoad resultLambda, load
        If Evaluate(result, False, recordTrial:=False) Then
            diagnostic(10) = 1#: diagnostic(11) = EquilibriumNorm(load, result)
            diagnostic(12) = Abs(CLng(GaugesVerified(tolerance)))
        End If
        Dim strain(1 To COMPONENT_COUNT) As Double, unused(1 To COMPONENT_COUNT) As Double, value As Double
        YieldStrain station.Point, station.Side, result, strain
        If ReadTransitionValue(station, strain, value, unused, False) Then diagnostic(13) = 1#: diagnostic(14) = value
    End If
    eventInfo = diagnostic
    mMeta.DiagnosticDetails = DiagnosticSummary()
'''
code = code.replace(anchor, anchor + diagnostics, 1)
binary = bytearray()
for vector in (v[9], v[10], v[5], [v[0][2],v[0][3],v[0][10],v[0][11],v[0][9],v[0][4]]):
    binary += struct.pack('<i' + 'd' * len(vector), len(vector), *vector)
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Input/Reverse1.bin': trajectory, 'Input/ArcRoots.bin': binary,
    'Input/Source779Attempt013.bin': (source_root / 'Native/ArcWalk.013.bin').read_bytes(),
    'Input/Source779Event013.bin': (source_root / 'Native/AutoEvent.013.Result.bin').read_bytes(),
    'Compile-TestExternalPoint.ps1': compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'", 1).replace('\n', '\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='Actual779RejectedYieldEvent', actual_import_this_pass=['CCoupledNewtonSolver'],
    source_history_exact779_43_states=True, actual_rejected_attempt=13,
    event_solver_and_physical_criteria_unchanged=True, only_readonly_postcheck_added=True,
    physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('780 EXACT779 HISTORY43/ATTEMPT13; CANDIDATE RADIUS/FORCE/GAUGE/YIELD POSTCHECK; SAME EVENT SOLVER')
