"""ДЛЯ ТЕСТОВ. Повторно разделяет отклонённые дуговые шаги по обнаруженным событиям CDP при общем бюджете96 вызовов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import re

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold778ArcYieldContextFix_20261125'
root = base / 'Cold779AutomaticYieldArcWalk_20261125'
assert json.loads((parent / 'ArcYieldSplitProof.json').read_bytes())['event_verified']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'AutomaticYieldCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'; code = (root / name).read_text('utf-8-sig')
code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mArcSeedCallsTest As Long, mArcSeedPrimaryTest As Long\n', 1)
initial, = re.findall(r'(?ms)^Private Sub TestRunArcYieldSplit\(\).*?^End Sub', code)
new_initial = initial.replace('        TestRunConfirmedArcWalk\n', '        mArcSeedCallsTest = 3: mArcSeedPrimaryTest = CLng(effort(0) + halfEffort(0) + restored(0))\n        TestRunConfirmedArcWalk\n', 1)
assert new_initial != initial
code = code.replace(initial, new_initial, 1)
dynamic = initial
start = dynamic.index('    Dim inputHandle As Integer:')
end = dynamic.index('    Dim revision As Long:')
dynamic = dynamic[:start] + '''    Dim context(1 To 6) As Double: context(1) = toLambda: context(6) = fromLambda
    Dim face() As Double, accepted() As Double, previous() As Double, i As Long, eventInfo As Variant, factor As Double
    accepted = mReverseFieldsTest(mReverseFieldsTest.Count): previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
''' + dynamic[end:]
dynamic = dynamic.replace('Private Sub TestRunArcYieldSplit()', 'Private Function TestTryCurrentArcYieldSplit(ByRef ends() As Double, ByVal fromLambda As Double, ByVal toLambda As Double, ByVal suffix As String) As Variant', 1)
dynamic = dynamic.replace('"\\Native\\ArcYieldSplit.Result.bin"', '"\\Native\\" & suffix & ".Result.bin"', 1)
dynamic = dynamic.replace('"\\Native\\ArcYieldSplit.Outcome.txt"', '"\\Native\\" & suffix & ".Outcome.txt"', 1)
dynamic = dynamic.replace('magic = "NDMARCYIELDSPLIT1"', 'magic = "NDMAUTOSPLIT0001"', 1)
anchor = '    TestWriteReverseBinaryVector handle, fullActual: TestWriteReverseBinaryVector handle, halfActual\n'
assert dynamic.count(anchor) == 1
dynamic = dynamic.replace(anchor, anchor+'''    Dim concrete As CCoupledConcreteResponse, sourceState As Variant
    If eventInfo(0) Then
        Select Case CLng(eventInfo(2))
            Case -1: Set concrete = mReverseSidesTest(1)
            Case 0: Set concrete = mReverseConcreteTest
            Case 1: Set concrete = mReverseSidesTest(3)
        End Select
        sourceState = concrete.ReadNewtonYieldEventStateTest(CLng(eventInfo(1)))
    Else
        sourceState = Array(0#)
    End If
    TestWriteReverseBinaryVector handle, sourceState
''', 1)
dynamic = dynamic.replace('        TestRunConfirmedArcWalk\n', '', 1)
assert dynamic.endswith('End Sub')
dynamic = dynamic[:-len('End Sub')] + '''    Dim callCount As Long: callCount = 1
    If solved Then callCount = callCount + 1
    If verified Then callCount = callCount + 1
    TestTryCurrentArcYieldSplit = Array(Abs(CLng(verified)), callCount, effort(0) + halfEffort(0) + restored(0), factor)
End Function'''
walk, = re.findall(r'(?ms)^Private Sub TestRunConfirmedArcWalk\(\).*?^End Sub', code)
new_walk = walk.replace('magic = "NDMARCWALK000001"', 'magic = "NDMARCWALK000002"', 1)
new_walk = new_walk.replace('Array(initialCount, 32#, 96#, 30#, 0.005, 0.0001)', 'Array(initialCount, 32#, 96#, 30#, 0.005, 0.0001, mArcSeedCallsTest, mArcSeedPrimaryTest)', 1)
anchor = '    reason = "Исчерпан бюджет32 дуговых пар."\n'
assert new_walk.count(anchor) == 1
new_walk = new_walk.replace(anchor, '''    calls = mArcSeedCallsTest: primary = mArcSeedPrimaryTest
    Dim eventHandle As Integer: eventHandle = FreeFile
    Open mPolishRootTest & "\\Native\\ArcWalk.Events.bin" For Binary Access Write As #eventHandle
    Dim eventMagic As String * 16: eventMagic = "NDMARCEVENTS0001"
    Put #eventHandle, , eventMagic
'''+anchor, 1)
anchor = '    For attempt = 1 To 32\n'
new_walk = new_walk.replace(anchor, anchor+'''        If calls + 3 > 96 Then reason = "Остатка общего бюджета недостаточно для пары и восстановления.": Exit For
''', 1)
anchor = '''        If ready Then
            mSignedArcRatioTest = GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(fieldError, 0.00000001)))'''
assert new_walk.count(anchor) == 1
new_walk = new_walk.replace(anchor, '''        Dim recovered As Boolean: recovered = False
        If Not ready And calls + 3 <= 96 Then
            Dim bracket() As Double, bracketLambda As Double, haveBracket As Boolean
            If mSignedArcPairTest(0) Then
                bracket = mSignedArcPairTest(4): bracketLambda = fullLambda: haveBracket = True
            ElseIf mSignedArcPairTest(1) Then
                bracket = mSignedArcPairTest(5): bracketLambda = halfLambda: haveBracket = True
            Else
                haveBracket = False
            End If
            If haveBracket Then
                Dim parentCount As Long: parentCount = mReverseFieldsTest.Count
                Dim eventResult As Variant
                eventResult = TestTryCurrentArcYieldSplit(bracket, previousLambda, bracketLambda, "AutoEvent." & Format$(attempt, "000"))
                calls = calls + CLng(eventResult(1)): primary = primary + CLng(eventResult(2))
                recovered = CBool(eventResult(0))
                TestWriteReverseBinaryVector eventHandle, Array(attempt, parentCount, Abs(CLng(recovered)), eventResult(1), eventResult(2), mReverseFieldsTest.Count, mReverseSolverTest.AcceptedRevision, eventResult(3))
            End If
        End If
        If recovered Then
            mSignedArcRatioTest = 1#
        ElseIf ready Then
            mSignedArcRatioTest = GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(fieldError, 0.00000001)))''', 1)
new_walk = new_walk.replace('    Close #handle: mSignedArcWalkingTest = False', '    Close #handle: Close #eventHandle: mSignedArcWalkingTest = False', 1)
code = code.replace(walk, new_walk, 1) + '\n\n' + "' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Повторяет проверенную процедуру события для фактического отклонённого шага текущей истории; исходные точки и станции не задаются.\n" + dynamic + '\n'
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Compile-TestExternalPoint.ps1': (here / 'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='RepeatedAutomaticYieldArcWalk', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_working_methods_exact778=True, initial_event_same778=True, event_detection_repeats_on_actual_rejected_roots=True,
    first_event_included_in_global96_budget=True, event_local30_budget_preserved=True,
    exact_source_state_for_each_selected_event_written=True, physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('779 REPEATED AUTOMATIC YIELD EVENT SPLITS; SAME778 KERNEL; FIRST EVENT INCLUDED IN96 GLOBAL CALLS')
