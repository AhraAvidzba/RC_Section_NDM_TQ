"""ДЛЯ ТЕСТОВ. Читает полную H-касательную настоящего связанного сечения от32 проверенных состояний, без нового поиска НДС."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import ast,json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold746ActualHFullDifferences_20261125';parent=base/'Cold744RecoveredColdFrozenReplay_20261125'
proof=json.loads((parent/'RecoveredColdFrozenReplayProof.json').read_bytes())
assert proof['states']==32 and proof['native']['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ActualHCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
values={};prep=json.loads((root/'Preparation.json').read_bytes())
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
pattern=r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub';method,=re.findall(pattern,code)
needle='    mPolishExecutedTest = True\n';assert method.count(needle)==1
code=code.replace(method,method.replace(needle,'    TestActualHFullDifferences\n'+needle,1),1)
body='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает оператор H в прогнозе1,25 последней секущей реальной32-точечной истории; это новая читающая проба, а не точный повтор прежнего отказа.
Private Sub TestActualHFullDifferences()
    Dim previous() As Double, accepted() As Double, before As Double, after As Double
    If Not mReverseSolverTest.TryAcceptedEnvelopeSegment(previous, before, after) Then Err.Raise 5, , "Не восстановлена пара реальной H-истории."
    If after <= before Then Err.Raise 5, , "Последняя H-пара не растёт."
    mReverseSolverTest.CopyAccepted accepted
    Dim factor As Double, priorFactor As Double, fraction As Double, baseline As Double, resolution As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, factor
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count - 1, fraction, priorFactor
    If Not mReverseSolverTest.TryEnvelopeControlState(mReverseBaseTest, mReverseDirectionTest, factor, _
            mReverseOptionsTest(3), mReverseOptionsTest(4), baseline, resolution) Then Err.Raise 5, , "Нет принятой H реального контекста."
    If baseline <> after Then Err.Raise 5, , "Принятый буфер H не совпал с концом восстановленной пары."
    Dim ratio As Double: ratio = 1.25
    Dim target As Double: target = after + ratio * (after - before)
    Dim seed() As Double, degree As Long, seedFactor As Double, projection() As Double
    seed = accepted
    For degree = 1 To mReverseSolverTest.PhysicalDegreeCount
        seed(degree) = accepted(degree) + ratio * (accepted(degree) - previous(degree))
    Next degree
    seedFactor = factor + ratio * (factor - priorFactor)
    ReDim projection(1 To mReverseOperatorsTest.DegreeCount): projection(1) = 1#
    If Not mReverseSolverTest.BeginNewtonLineTest(mReverseBaseTest, mReverseDirectionTest, projection, seed(1), seed, seedFactor) Then _
        Err.Raise 5, , "Не открыт нормированный контекст реального сечения."
    Dim prepared As New CCoupledEnvelopeEquations, equations As ICoupledEquationSystem
    prepared.Initialize Nothing, mReverseOperatorsTest, mReverseReinforcedTest, mReverseSidesTest(2), mReverseSidesTest(4), _
        HARD_PREPARED_HALF_TEST, mReverseCalibrationTest(0) / mReverseCalibrationTest(3)
    Set equations = prepared.AtTarget(mReverseSolverTest, target)
    Dim zero() As Double, residual() As Double, matrix() As Double, scales As Variant, ready As Boolean
    ReDim zero(1 To mReverseOperatorsTest.DegreeCount + 1)
    ready = equations.TryEvaluate(zero, True, residual, matrix)
    scales = mReverseSolverTest.ReadNewtonLineScalesTest
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\ActualH32"
    Dim handle As Integer: handle = FreeFile
    Dim magic As String * 16: magic = "NDMTAILMATRIX001"
    Dim size As Long: size = UBound(zero)
    Dim flag As Long: flag = Abs(CLng(ready))
    Open mHardPreparedPathTest & ".Matrix.bin" For Binary Access Write As #handle
    Put #handle, , magic: Put #handle, , size: Put #handle, , flag
    Put #handle, , ratio: Put #handle, , seedFactor
    TestWriteReverseBinaryVector handle, seed: TestWriteReverseBinaryVector handle, scales
    If ready Then
        TestWriteReverseBinaryVector handle, residual
        Dim row As Long, column As Long, value As Double
        For row = 1 To size
            For column = 1 To size: value = matrix(column, row): Put #handle, , value: Next column
        Next row
    End If
    Close #handle
    handle = FreeFile
    Open mHardPreparedPathTest & ".Context.bin" For Binary Access Write As #handle
    TestWriteReverseBinaryVector handle, Array(before, after, target, baseline, resolution, prepared.RowScale, factor, priorFactor, mReverseSolverTest.AcceptedRevision)
    Close #handle
    If Not ready Then Err.Raise 5, , "Полная H-касательная прогноза реального сечения численно недоступна."
    TestWriteTailDifferenceSamples equations, zero, residual
    mReverseSolverTest.EndNewtonLineTest
End Sub
'''
tree=ast.parse((here/'Stage-TestTailFullDifferences.py').read_text('utf-8-sig'))
samples=[n.value.value for n in ast.walk(tree) if isinstance(n,ast.AugAssign) and isinstance(n.target,ast.Name)
    and n.target.id=='code' and isinstance(n.value,ast.Constant) and isinstance(n.value.value,str)
    and 'Private Sub TestWriteTailDifferenceSamples' in n.value.value]
assert len(samples)==1
if 'Private Sub TestWriteTailDifferenceSamples' in code:
    for method_name in ('TestWriteTailDifferenceSamples','TestAppendTailDifference'):
        pattern=r'(?ms)^Private Sub '+method_name+r'\(.*?^End Sub'
        actual,=re.findall(pattern,code);expected,=re.findall(pattern,samples[0])
        assert actual==expected,method_name
    code+='\n'+body
else:
    code+='\n'+body+'\n'+samples[0]
assert max(map(len,code.splitlines()))<1024
methods=re.findall(r'(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)',code)
assert len(methods)==len(set(n.casefold() for n in methods))
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Actual32CoupledHFullResidualDerivative',actual_import_this_pass=['modTestCoupledElasticState'],
    source_states=32,actual743_material_and_Newton_exact=True,all_other_numerical_classes_exact744=True,
    source_cold_history_verified_by744=True,read_only_secant_predictor_ratio=1.25,
    new_readonly_forecast_not_exact_failed_request_replay=True,all61columns_both_signs=True,
    normalized_difference_steps=[1e-6,1e-4],expected_samples=247,
    sample_writer_exact738=True,accepted_revision_unchanged_required=True,
    no_new_NDS_or_capacity_search=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('746 ACTUAL32 REAL ALL6 HISTORY; READONLY H MATRIX AND247 FULL DIFFERENCES; NO ROOT OR CAPACITY SEARCH')
