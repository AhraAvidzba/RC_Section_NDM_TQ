"""ДЛЯ ТЕСТОВ. Проверяет успешную ветвь половинного контроля на двух реальных упругих состояниях всех6 усилий."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold757HalfGuardedMaterialReturn_20261125'
root = base / 'Cold758PositiveHalfGuard_20261125'
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'PositiveHalfCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'
code = (root / name).read_text('utf-8-sig')
code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вызывает рабочую проверку половины возврата у подготовленного реального кандидата; принятую историю не меняет.
Friend Function TestConfirmOriginalReturnHalf(ByRef full() As Double, ByVal factor As Double, ByRef projection() As Double, _
        ByVal target As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maximumSteps As Long, _
        ByVal reference As Double, ByRef details As String) As Boolean
    If mMaterialGuideReturnActive Then Err.Raise 5, , "Уже выполняется материальный возврат."
    On Error GoTo FailedTest
    mMaterialGuideReturnActive = True
    TestConfirmOriginalReturnHalf = ConfirmOriginalReturnHalf(full, factor, projection, target, True, absoluteTolerance, relativeTolerance, maximumSteps, reference, details)
    mMaterialGuideReturnActive = False
    Exit Function
FailedTest:
    Dim number As Long, source As String, reason As String
    number = Err.Number: source = Err.Source: reason = Err.Description
    mMaterialGuideReturnActive = False
    Err.Raise number, source, reason
End Function
'''
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunIntegratedMaterialReturn\n') == 1
code = code.replace('    TestRunIntegratedMaterialReturn\n', '    TestRunPositiveHalfGuard\n', 1)
code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет успешный половинный контроль, точное восстановление полного кандидата и исходной цели на упругом all6-пути.
Private Sub TestRunPositiveHalfGuard()
    Dim fields As New Collection, trace As CCoupledLoadTrace, full() As Double, accepted() As Double, movement() As Double
    Dim projection() As Double, actual() As Double, restored() As Double, context As Variant, before As Variant, after As Variant
    Dim fraction As Double, reference As Double, factor As Double, target As Double, spacing As Double, i As Long, seed As Variant
    fields.Add mReverseFieldsTest(1): Set trace = mReverseTraceTest.CopyPrefix(1)
    If Not mReverseSolverTest.ReplayAcceptedFields(fields, fields.Count) Then Err.Raise 5, , "Не восстановлен упругий префикс."
    accepted = fields(1): full = mReverseFieldsTest(2): movement = full
    For i = 1 To UBound(full): movement(i) = full(i) - accepted(i): Next i
    If Not mReverseSolverTest.SelectSpatialDirectionControl(accepted, full, movement, projection, target, spacing) Then Err.Raise 5, , "Нет меры упругого приращения."
    trace.ReadPoint 1, fraction, reference: mReverseTraceTest.ReadPoint 2, fraction, factor
    seed = full
    If Not mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, target, factor, factor, _
            mReverseOptionsTest(3), mReverseOptionsTest(4), CLng(mReverseOptionsTest(2)), consistentControl:=True, initialField:=seed) Then Err.Raise 5, , "Нет полного упругого кандидата."
    mReverseSolverTest.CopyTrial full, actual: before = mReverseSolverTest.ReadPolishEffortTest
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim details As String, passed As Boolean
    passed = mReverseSolverTest.TestConfirmOriginalReturnHalf(full, factor, projection, target, mReverseOptionsTest(3), _
        mReverseOptionsTest(4), CLng(mReverseOptionsTest(2)), reference, details)
    after = mReverseSolverTest.ReadPolishEffortTest
    If Not passed Then Err.Raise 5, , "Рабочая проверка отклонила упругий половинный шаг: " & details
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Положительный тест принял материальную историю."
    mReverseSolverTest.CopyTrial restored, actual: context = mReverseSolverTest.ReadReverseControlTest
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\PositiveHalf.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMPOSHALF000001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(1#, reference, factor, target, revision, mReverseSolverTest.AcceptedRevision)
    TestWriteReverseBinaryVector handle, before: TestWriteReverseBinaryVector handle, after
    TestWriteReverseBinaryVector handle, context: TestWriteReverseBinaryVector handle, projection
    TestWriteReverseBinaryVector handle, full: TestWriteReverseBinaryVector handle, restored: TestWriteReverseBinaryVector handle, actual
    Close #handle
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\PositiveHalf.Outcome.txt" For Output As #handle
    Print #handle, "PASSED;True;DETAILS;" & details
    Close #handle
    mReverseSolverTest.DiscardStep
    fields.Add restored: trace.Add 1#, factor
    Set mReverseFieldsTest = fields: Set mReverseTraceTest = trace
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\PositiveHalf"
    mPtcSuppressTest = True: TestExportFastFrozenReplay 999: mPtcSuppressTest = False
End Sub
'''
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
prep['directed_scope'] = dict(case='PositiveOriginalReturnHalfGuard', actual_import_this_pass=['CCoupledNewtonSolver', 'modTestCoupledElasticState'],
    working_757_half_guard_unchanged=True, real_all6_elastic_prefix_lambda025_to05=True,
    full_candidate_and_exact_original_control_must_be_restored=True, accepted_revision_must_not_change=True,
    no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('758 POSITIVE HALF GUARD ON REAL ELASTIC ALL6 PREFIX; WORKING METHODS EXACT757')
