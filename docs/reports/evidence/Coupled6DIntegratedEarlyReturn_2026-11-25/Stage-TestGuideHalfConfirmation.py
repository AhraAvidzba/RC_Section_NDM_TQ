"""ДЛЯ ТЕСТОВ. Проверяет локальность принятого вычислительного возврата половиной исходной управляющей цели."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold752IntegratedGuideDirected_20261125'
root = base / 'Cold756GuideHalfConfirmation_20261125'
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'HalfConfirmationCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunIntegratedMaterialReturn\n') == 1
code = code.replace('    TestRunIntegratedMaterialReturn\n', '    TestRunGuideHalfConfirmation\n', 1)
code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет полный возврат и решает половинную цель от той же истории в остатке30 итераций; принятия материала нет.
Private Sub TestRunGuideHalfConfirmation()
    Dim branch As New CCoupledDeformationContinuation, revision As Long, outcome As Variant, firstEffort As Variant
    revision = mReverseSolverTest.AcceptedRevision
    outcome = branch.TestAuditEnvelopeDeparture(mReverseSolverTest, mReverseReinforcedTest, mReverseBaseTest, mReverseDirectionTest, _
        mReverseFieldsTest, mReverseTraceTest, mReverseOptionsTest)
    firstEffort = mReverseSolverTest.ReadPolishEffortTest
    If Not CBool(outcome(0)) Then Err.Raise 5, , "Не воспроизведён исходный полный возврат для проверки половины шага."
    Dim full() As Double, actual() As Double, projection() As Double, accepted() As Double
    Dim context As Variant, uncertainty As Double, consistent As Boolean
    mReverseSolverTest.CopyTrial full, actual: mReverseSolverTest.CopyAccepted accepted
    If Not mReverseSolverTest.CopyRecoveredControl(projection, uncertainty) Then Err.Raise 5, , "Нет исходной проекции возврата."
    context = mReverseSolverTest.ReadReverseControlTest
    consistent = mReverseSolverTest.RecoveredControlRequiresBranchTangent
    Dim i As Long, coordinate As Double, target As Double, halfTarget As Double, budget As Long
    For i = 1 To UBound(full): coordinate = coordinate + projection(i) * accepted(i): Next i
    target = context(1): halfTarget = 0.5 * (coordinate + target)
    budget = CLng(mReverseOptionsTest(2) - firstEffort(0))
    If budget < 1 Then Err.Raise 5, , "После полного возврата нет бюджета проверки половины шага."
    Dim seed() As Double, prediction As Variant, halfLambda As Double, reference As Double, fraction As Double
    seed = accepted
    For i = 1 To UBound(full): seed(i) = 0.5 * (accepted(i) + full(i)): Next i
    prediction = seed: mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, reference
    mReverseSolverTest.DiscardStep
    Dim solved As Boolean
    solved = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, halfTarget, _
        0.5 * (reference + outcome(2)), halfLambda, mReverseOptionsTest(3), mReverseOptionsTest(4), budget, _
        consistentControl:=consistent, initialField:=prediction, allowSmallCorrection:=True, allowHybridRecovery:=True, conditionHybridRecovery:=True)
    Dim secondEffort As Variant: secondEffort = mReverseSolverTest.ReadPolishEffortTest
    If firstEffort(0) + secondEffort(0) > mReverseOptionsTest(2) Then Err.Raise 5, , "Превышен общий бюджет полного и половинного расчётов."
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Проверка половины шага изменила принятую историю."
    Dim half() As Double, halfActual() As Double, halfContext As Variant, otherProjection() As Double, changedControl As Boolean
    If solved Then
        mReverseSolverTest.CopyTrial half, halfActual: halfContext = mReverseSolverTest.ReadReverseControlTest
        changedControl = mReverseSolverTest.CopyRecoveredControl(otherProjection)
    Else
        ReDim half(1 To 1): ReDim halfActual(1 To 1): halfContext = Array(0#, 0#, 0#, 0#)
    End If
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\GuideHalf.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMGUIDEHALF0001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(1#, Abs(CLng(solved)), outcome(2), reference, halfLambda, target, halfTarget, coordinate, _
        revision, mReverseSolverTest.AcceptedRevision, budget, Abs(CLng(consistent)), Abs(CLng(changedControl)), uncertainty)
    TestWriteReverseBinaryVector handle, firstEffort: TestWriteReverseBinaryVector handle, secondEffort
    TestWriteReverseBinaryVector handle, halfContext: TestWriteReverseBinaryVector handle, projection
    TestWriteReverseBinaryVector handle, accepted: TestWriteReverseBinaryVector handle, full
    TestWriteReverseBinaryVector handle, half: TestWriteReverseBinaryVector handle, halfActual
    Close #handle
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\GuideHalf.Outcome.txt" For Output As #handle
    Print #handle, "FULL_SOLVED;True;HALF_SOLVED;" & CStr(solved) & ";FIRST_STEPS;" & CStr(firstEffort(0)) & ";HALF_STEPS;" & CStr(secondEffort(0)) & ";BUDGET;30"
    Print #handle, "DETAILS;" & mReverseSolverTest.ResultMeta.DiagnosticDetails
    Close #handle
    mReverseSolverTest.DiscardStep
    If Not solved Then Exit Sub
    Dim originalFields As Collection, originalTrace As CCoupledLoadTrace, sequence As New Collection, trace As CCoupledLoadTrace
    Set originalFields = mReverseFieldsTest: Set originalTrace = mReverseTraceTest
    For i = 1 To originalFields.Count: sequence.Add originalFields(i): Next i
    sequence.Add half: Set trace = originalTrace.CopyPrefix(originalTrace.Count): trace.Add 1#, halfLambda
    Set mReverseFieldsTest = sequence: Set mReverseTraceTest = trace
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\GuideHalf"
    mPtcSuppressTest = True: TestExportFastFrozenReplay 999: mPtcSuppressTest = False
    Set mReverseFieldsTest = originalFields: Set mReverseTraceTest = originalTrace
End Sub
'''
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
values['Compile-TestExternalPoint.ps1'] = (here / 'Compile-TestExternalPoint.ps1').read_bytes()
values['tools/build_workbook/Update-IsolatedOptimizerCode.ps1'] = (base / 'Cold753IntegratedColdFrozenReplay_20261125/tools/build_workbook/Update-IsolatedOptimizerCode.ps1').read_bytes()
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='GuideOriginalControlHalfConfirmation', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_core_exact752=True, original_full_recovery_repeated_in_same_native_call=True,
    half_target_on_same_accepted46_history=True, combined_primary_budget30=True,
    no_material_acceptance_in_full_or_half_solves=True, no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('756 FULL RETURN THEN HALF ORIGINAL TARGET; SAME ACCEPTED46 AND COMBINED30 BUDGET; NUMERIC CORE EXACT752')
