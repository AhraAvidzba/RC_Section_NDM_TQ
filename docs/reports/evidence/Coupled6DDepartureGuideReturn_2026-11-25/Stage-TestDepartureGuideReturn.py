"""ДЛЯ ТЕСТОВ. Проверяет обратный корень только как начальное приближение возврата к исходной цели, без принятия его материала."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold750DepartureGuideReturn_20261125';parent=base/'Cold749DepartureCandidateAudit_20261125'
proof=json.loads((parent/'DepartureCandidateAuditProof.json').read_bytes())
assert proof['independent_forward_direction_check_would_also_reject'] and proof['native']['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'GuideReturnCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Private mDepartureAuditFieldTest() As Double, mDepartureAuditFactorTest As Double\n'
assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''Private mDepartureGuideProjectionTest() As Double, mDepartureGuideTargetTest As Double
Private mDepartureGuideEffortTest As Variant
''',1)
pattern=r'(?ms)^Public Sub TestCaptureDepartureCandidate\(.*?^End Sub';method,=re.findall(pattern,code)
anchor='    mDepartureAuditFieldTest = field: mDepartureAuditFactorTest = factor\n';assert method.count(anchor)==1
changed=method.replace(anchor,anchor+'''    mDepartureGuideProjectionTest = oldProjection: mDepartureGuideTargetTest = oldTarget
    mDepartureGuideEffortTest = solver.ReadPolishEffortTest
''',1);code=code.replace(method,changed,1)
pattern=r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub';method,=re.findall(pattern,code)
anchor='    TestRunDepartureCandidateAudit\n';assert method.count(anchor)==1
code=code.replace(method,method.replace(anchor,anchor+'    TestReturnOriginalTargetFromDepartureGuide\n',1),1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Использует непринятый обратный корень только как прогноз; исходная46-точечная история восстанавливается до новой пробы, общий бюджет остаётся30.
Private Sub TestReturnOriginalTargetFromDepartureGuide()
    If mDepartureAuditCountTest <> 1 Then Err.Raise 5, , "Нужен один наблюдаемый вспомогательный корень."
    Dim spent As Long: spent = CLng(mDepartureGuideEffortTest(0))
    Dim available As Long: available = CLng(mReverseOptionsTest(2)) - spent
    If available < 1 Then Err.Raise 5, , "После вспомогательной пробы не осталось общего бюджета."
    If Not mReverseSolverTest.ReplayAcceptedFields(mReverseFieldsTest, mReverseFieldsTest.Count) Then Err.Raise 5, , "Не восстановлена исходная история перед возвратом к цели."
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim seed As Variant: seed = mDepartureAuditFieldTest
    Dim factor As Double, solved As Boolean, started As Double: started = Timer
    solved = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, mDepartureGuideProjectionTest, _
        mDepartureGuideTargetTest, mDepartureAuditFactorTest, factor, mReverseOptionsTest(3), mReverseOptionsTest(4), available, _
        consistentControl:=True, initialField:=seed, allowSmallCorrection:=True, allowHybridRecovery:=True, conditionHybridRecovery:=True)
    Dim seconds As Double: seconds = Timer - started: If seconds < 0# Then seconds = seconds + 86400#
    Dim effort As Variant: effort = mReverseSolverTest.ReadPolishEffortTest
    If spent + CLng(effort(0)) > CLng(mReverseOptionsTest(2)) Then Err.Raise 5, , "Возврат превысил общий бюджет."
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Возврат принял пробный материал."
    Dim field() As Double, actual() As Double, coordinate As Double, index As Long
    If solved Then
        mReverseSolverTest.CopyTrial field, actual
        For index = 1 To UBound(field): coordinate = coordinate + mDepartureGuideProjectionTest(index) * field(index): Next index
    Else
        ReDim field(1 To 1): ReDim actual(1 To 1)
    End If
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\GuideReturn.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMGUIDERET00001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(spent, available, effort(0), Abs(CLng(solved)), factor, coordinate, mDepartureGuideTargetTest, revision, mReverseSolverTest.AcceptedRevision, mDepartureAuditFactorTest, seconds)
    TestWriteReverseBinaryVector handle, mDepartureGuideEffortTest: TestWriteReverseBinaryVector handle, effort
    TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
    Close #handle
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\GuideReturn.Outcome.txt" For Output As #handle
    Print #handle, "SOLVED;" & CStr(solved) & ";SPENT;" & CStr(spent) & ";AVAILABLE;" & CStr(available) & ";USED;" & CStr(effort(0)) & ";SECONDS;" & CStr(seconds)
    Print #handle, "ORIGINAL_TARGET;" & CStr(mDepartureGuideTargetTest) & ";ACTUAL;" & CStr(coordinate)
    Print #handle, "DETAILS;" & mReverseSolverTest.ResultMeta.DiagnosticDetails
    Close #handle
    mReverseSolverTest.DiscardStep
    If Not solved Then Exit Sub
    Dim originalFields As Collection, originalTrace As CCoupledLoadTrace, sequence As New Collection, trace As CCoupledLoadTrace
    Set originalFields = mReverseFieldsTest: Set originalTrace = mReverseTraceTest
    For index = 1 To originalFields.Count: sequence.Add originalFields(index): Next index
    sequence.Add field: Set trace = originalTrace.CopyPrefix(originalTrace.Count): trace.Add 1#, factor
    Set mReverseFieldsTest = sequence: Set mReverseTraceTest = trace
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\GuideReturn"
    mPtcSuppressTest = True
    TestExportFastFrozenReplay 999
    mPtcSuppressTest = False
    Set mReverseFieldsTest = originalFields: Set mReverseTraceTest = originalTrace
End Sub
'''
assert max(map(len,code.splitlines()))<1024
methods=re.findall(r'(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)',code)
assert len(methods)==len(set(n.casefold() for n in methods))
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values['Input/ExistingGuideCandidate.bin']=(parent/'Native/Departure.Candidate1.bin').read_bytes()
values['Compile-TestExternalPoint.ps1']=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ReturnOriginalTargetFromUnacceptedDepartureGuide',actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_kernel_exact749=True,source_states=46,repeat_original_auxiliary_work_and_read_actual_spent_iterations=True,
    remaining_iterations_is30_minus_spent=True,original_material_history_restored_before_return=True,
    auxiliary_backward_root_not_in_return_history=True,original_linear_control_target_required=True,
    returned_NDS_not_controller_admission=True,existing_controlled_NDS_API_used=True,
    no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('750 SAME KERNEL; REPEAT AUXILIARY WORK THEN RETURN ORIGINAL TARGET WITH ONLY REMAINING30-STEP BUDGET')
