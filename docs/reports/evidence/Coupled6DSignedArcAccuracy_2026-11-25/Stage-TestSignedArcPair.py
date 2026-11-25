"""ДЛЯ ТЕСТОВ. Проверяет прямую псевдодугу после максимума на точном35-точечном префиксе поздней all6-истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import struct

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold759HalfGuardLoadAllocationFix_20261125'
source = base / 'Cold753IntegratedColdFrozenReplay_20261125'
root = base / 'Cold760SignedArcPair_20261125'
history = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source / 'Input/Reverse1.bin')
assert len(history['fields']) == 36 and history['lambdas'][34] < history['lambdas'][33]
prefix = {key: history[key][:35] for key in ('fields', 'fractions', 'lambdas')}
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'SignedArcCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunPositiveHalfGuard\n') == 1
code = code.replace('    TestRunPositiveHalfGuard\n', '    TestRunSignedArcPair\n', 1)
code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет полную и половинную дуговые цели от одной принятой истории с общим бюджетом30, без возврата к старой мере.
Private Sub TestRunSignedArcPair()
    Dim accepted() As Double, previous() As Double, movement() As Double, projection() As Double
    Dim count As Long: count = mReverseFieldsTest.Count
    accepted = mReverseFieldsTest(count): previous = mReverseFieldsTest(count - 1): movement = accepted
    Dim i As Long, spacing As Double, coordinate As Double
    For i = 1 To UBound(movement): movement(i) = accepted(i) - previous(i): Next i
    If Not mReverseSolverTest.SelectSpatialDirectionControl(previous, accepted, movement, projection, coordinate, spacing) Then Err.Raise 5, , "Нет пространственной секущей дугового теста."
    Dim slopes(1 To 3) As Double, factor(1 To 4) As Double, fraction As Double, a() As Double, b() As Double
    For i = 1 To 4
        mReverseTraceTest.ReadPoint count - 4 + i, fraction, factor(i)
        If fraction <> 1# Or factor(i) <= 0# Then Err.Raise 5, , "Неполная силовая история для дугового теста."
    Next i
    For i = 1 To 3
        a = mReverseFieldsTest(count - 4 + i): b = mReverseFieldsTest(count - 3 + i)
        If factor(i + 1) = factor(i) Then Err.Raise 5, , "Нулевой прирост множителя требует отдельного масштаба дуги."
        slopes(i) = mReverseSolverTest.SpatialFieldDistance(a, b) / Abs(factor(i + 1) - factor(i))
    Next i
    Dim alpha As Double, lambdaMove As Double, secant As Double, ratio As Double, radius As Double, q As Double
    alpha = GeomMax(GeomMin(slopes(1), slopes(2)), GeomMin(GeomMax(slopes(1), slopes(2)), slopes(3)))
    lambdaMove = factor(4) - factor(3): secant = Sqr(spacing * spacing + alpha * alpha * lambdaMove * lambdaMove)
    If secant <= 0# Then Err.Raise 5, , "Нулевой дуговой масштаб."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Input\\SignedArcOptions.bin" For Binary Access Read As #handle
    Dim settings As Variant: settings = TestReadPolishVector(handle): Close #handle
    ratio = settings(1): If ratio <= 0# Or ratio > 1# Then Err.Raise 5, , "Доля дуговой секущей вне диапазона."
    radius = ratio * secant: q = alpha * alpha * lambdaMove / secant
    Dim target As Double: target = q * factor(4)
    For i = 1 To UBound(projection)
        projection(i) = projection(i) * spacing / secant
        target = target + projection(i) * accepted(i)
    Next i
    Dim fullTarget As Double, halfTarget As Double, fullSeed() As Double, halfSeed() As Double, prediction As Variant
    fullTarget = target + radius: halfTarget = target + 0.5 * radius
    fullSeed = accepted: halfSeed = accepted
    For i = 1 To mReverseSolverTest.PhysicalDegreeCount
        fullSeed(i) = accepted(i) + ratio * movement(i)
        halfSeed(i) = accepted(i) + 0.5 * ratio * movement(i)
    Next i
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim fullSolved As Boolean, halfSolved As Boolean, fullLambda As Double, halfLambda As Double
    Dim full() As Double, half() As Double, fullActual() As Double, halfActual() As Double
    Dim firstEffort As Variant, secondEffort As Variant, firstContext As Variant, secondContext As Variant
    Dim firstDetails As String, secondDetails As String, fullBudget As Long, halfBudget As Long
    fullBudget = CLng(mReverseOptionsTest(2)) \\ 2: prediction = fullSeed
    fullSolved = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, fullTarget, _
        factor(4) + ratio * lambdaMove, fullLambda, mReverseOptionsTest(3), mReverseOptionsTest(4), fullBudget, _
        consistentControl:=True, initialField:=prediction, loadProjection:=q)
    firstEffort = mReverseSolverTest.ReadPolishEffortTest: firstDetails = mReverseSolverTest.ResultMeta.DiagnosticDetails
    If fullSolved Then
        mReverseSolverTest.CopyTrial full, fullActual: firstContext = mReverseSolverTest.ReadReverseControlTest
    Else
        ReDim full(1 To 1): ReDim fullActual(1 To 1): firstContext = Array(0#, 0#, 0#, 0#)
    End If
    mReverseSolverTest.DiscardStep
    halfBudget = CLng(mReverseOptionsTest(2) - firstEffort(0)): prediction = halfSeed
    halfSolved = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, halfTarget, _
        factor(4) + 0.5 * ratio * lambdaMove, halfLambda, mReverseOptionsTest(3), mReverseOptionsTest(4), halfBudget, _
        consistentControl:=True, initialField:=prediction, loadProjection:=q)
    secondEffort = mReverseSolverTest.ReadPolishEffortTest: secondDetails = mReverseSolverTest.ResultMeta.DiagnosticDetails
    If halfSolved Then
        mReverseSolverTest.CopyTrial half, halfActual: secondContext = mReverseSolverTest.ReadReverseControlTest
    Else
        ReDim half(1 To 1): ReDim halfActual(1 To 1): secondContext = Array(0#, 0#, 0#, 0#)
    End If
    If mReverseSolverTest.AcceptedRevision <> revision Or firstEffort(0) + secondEffort(0) > mReverseOptionsTest(2) Then Err.Raise 5, , "Дуговая пара изменила историю или превысила общий бюджет."
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\SignedArc.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMSIGNEDARC0001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(fullSolved)), Abs(CLng(halfSolved)), fullLambda, halfLambda, factor(4), _
        ratio, alpha, secant, radius, q, fullTarget, halfTarget, revision, mReverseSolverTest.AcceptedRevision, mReverseOptionsTest(2), factor(3), fullBudget)
    TestWriteReverseBinaryVector handle, firstEffort: TestWriteReverseBinaryVector handle, secondEffort
    TestWriteReverseBinaryVector handle, firstContext: TestWriteReverseBinaryVector handle, secondContext
    TestWriteReverseBinaryVector handle, projection: TestWriteReverseBinaryVector handle, accepted
    TestWriteReverseBinaryVector handle, fullSeed: TestWriteReverseBinaryVector handle, halfSeed
    TestWriteReverseBinaryVector handle, full: TestWriteReverseBinaryVector handle, half
    TestWriteReverseBinaryVector handle, fullActual: TestWriteReverseBinaryVector handle, halfActual
    Close #handle
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\SignedArc.Outcome.txt" For Output As #handle
    Print #handle, "FULL;" & CStr(fullSolved) & ";HALF;" & CStr(halfSolved) & ";PRIMARY;" & CStr(firstEffort(0) + secondEffort(0)) & ";BUDGET;30"
    Print #handle, "FULL_DETAILS;" & firstDetails: Print #handle, "HALF_DETAILS;" & secondDetails
    Close #handle
    mReverseSolverTest.DiscardStep
    If fullSolved Then TestExportSignedArcTrial "SignedArc.Full", full, fullLambda
    If halfSolved Then TestExportSignedArcTrial "SignedArc.Half", half, halfLambda
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Экспортирует отдельную дуговую пробу от исходной истории; полная и половинная пробы не принимаются последовательно.
Private Sub TestExportSignedArcTrial(ByVal name As String, ByRef field() As Double, ByVal factor As Double)
    Dim originalFields As Collection, originalTrace As CCoupledLoadTrace, fields As New Collection, trace As CCoupledLoadTrace, i As Long
    Set originalFields = mReverseFieldsTest: Set originalTrace = mReverseTraceTest
    For i = 1 To originalFields.Count: fields.Add originalFields(i): Next i
    fields.Add field: Set trace = originalTrace.CopyPrefix(originalTrace.Count): trace.Add 1#, factor
    Set mReverseFieldsTest = fields: Set mReverseTraceTest = trace
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\" & name
    mPtcSuppressTest = True: TestExportFastFrozenReplay 999: mPtcSuppressTest = False
    Set mReverseFieldsTest = originalFields: Set mReverseTraceTest = originalTrace
End Sub
'''
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
values['Input/Reverse1.bin'] = run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](history, prefix)
values['Input/Source752History.bin'] = (source / 'Input/Source752History.bin').read_bytes()
values['Input/SignedArcOptions.bin'] = struct.pack('<id', 1, .5)
values['Compile-TestExternalPoint.ps1'] = (here / 'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='SignedArcPairActualLate35', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_core_exact759=True, actual_history_exact752_prefix35_after_maximum34=True,
    signed_lambda_secant_preserved=True, positive_load_metric_from_median_absolute_slopes=True,
    no_old_scalar_control_return=True, full_and_half_same_history_combined30=True,
    no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('760 SIGNED ARC PAIR ON EXACT LATE35 AFTER MAX34; SAME759 NUMERIC CORE; SHARED30; NO OLD-CONTROL RETURN')
