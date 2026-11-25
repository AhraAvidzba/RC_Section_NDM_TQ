"""ДЛЯ ТЕСТОВ. Проверяет встроенный возврат одним рабочим вызовом от точной46-точечной истории, сохраняя позднюю H-политику."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold752IntegratedGuideDirected_20261125';parent=base/'Cold751IntegratedMaterialGuide_20261125'
frozen=base/'Cold747EarlyEnvelopeFrozenReplay_20261125';audit=base/'Cold749DepartureCandidateAudit_20261125'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'IntegratedDirectedCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
pattern=r'(?ms)^Private Function TryReturnOriginalMaterialControl\(.*?^End Function';method,=re.findall(pattern,code)
old='CountReturn:\n    phaseDetails = mMeta.DiagnosticDetails\n'
new='CountReturn:\n    On Error GoTo 0\n    phaseDetails = mMeta.DiagnosticDetails\n'
assert method.count(old)==1;changed=method.replace(old,new,1)
code=code.replace(method,changed,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
wrapper,=re.findall(r'(?ms)^Friend Function TestAuditEnvelopeDeparture\(.*?^End Function',(audit/name).read_text('utf-8-sig'))
wrapper=wrapper.replace('    ReDim mProjection(1 To solver.PhysicalDegreeCount + 4)',
    '    Dim accepted() As Double: solver.CopyAccepted accepted\n    ReDim mProjection(1 To UBound(accepted))')
code+="\n' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Повторяет одну штатную процедуру резерва от заданной принятой истории, не изменяя её.\n"+wrapper+'\n'
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(frozen/name).read_text('utf-8-sig')
pattern=r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub';method,=re.findall(pattern,code)
needle='    mPolishExecutedTest = True\n';assert method.count(needle)==1
code=code.replace(method,method.replace(needle,'    TestRunIntegratedMaterialReturn\n'+needle,1),1)
old='''Private Sub TestExportFastFrozenReplay()
    Const caseIndex As Long = 0'''
new='Private Sub TestExportFastFrozenReplay(Optional ByVal caseIndex As Long = 0)'
assert code.count(old)==1;code=code.replace(old,new,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет встроенный материальный возврат одним исходным вызовом с общим бюджетом30; публикация контроллером и поиск предела не утверждаются.
Private Sub TestRunIntegratedMaterialReturn()
    Dim branch As New CCoupledDeformationContinuation, revision As Long, outcome As Variant, effort As Variant
    revision = mReverseSolverTest.AcceptedRevision
    outcome = branch.TestAuditEnvelopeDeparture(mReverseSolverTest, mReverseReinforcedTest, mReverseBaseTest, mReverseDirectionTest, _
        mReverseFieldsTest, mReverseTraceTest, mReverseOptionsTest)
    effort = mReverseSolverTest.ReadPolishEffortTest
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Встроенный резерв принял пробную материальную историю."
    If effort(0) > mReverseOptionsTest(2) Then Err.Raise 5, , "Встроенный резерв превысил общий первичный бюджет."
    Dim solved As Boolean: solved = CBool(outcome(0))
    Dim field() As Double, actual() As Double, projection() As Double, uncertainty As Double, context As Variant
    If solved Then
        mReverseSolverTest.CopyTrial field, actual
        If Not mReverseSolverTest.CopyRecoveredControl(projection, uncertainty) Then Err.Raise 5, , "Нет подтверждённой исходной меры возврата."
        context = mReverseSolverTest.ReadReverseControlTest
    Else
        ReDim field(1 To 1): ReDim actual(1 To 1): ReDim projection(1 To 1)
        context = Array(0#, 0#, 0#, 0#)
    End If
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\IntegratedReturn.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMINTEGRET00001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(solved)), outcome(2), outcome(3), revision, mReverseSolverTest.AcceptedRevision, uncertainty)
    TestWriteReverseBinaryVector handle, effort: TestWriteReverseBinaryVector handle, context
    TestWriteReverseBinaryVector handle, projection: TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
    Close #handle
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\IntegratedReturn.Outcome.txt" For Output As #handle
    Print #handle, "SOLVED;" & CStr(solved) & ";PRIMARY_STEPS;" & CStr(effort(0)) & ";PROBES;" & CStr(outcome(1)) & ";LAMBDA;" & CStr(outcome(2))
    Print #handle, "DETAILS;" & mReverseSolverTest.ResultMeta.DiagnosticDetails
    If Not solved Then Print #handle, "FAILURE;" & CStr(outcome(6))
    Close #handle
    mReverseSolverTest.DiscardStep
    If Not solved Then Exit Sub
    Dim originalFields As Collection, originalTrace As CCoupledLoadTrace, sequence As New Collection, trace As CCoupledLoadTrace, index As Long
    Set originalFields = mReverseFieldsTest: Set originalTrace = mReverseTraceTest
    For index = 1 To originalFields.Count: sequence.Add originalFields(index): Next index
    sequence.Add field: Set trace = originalTrace.CopyPrefix(originalTrace.Count): trace.Add 1#, outcome(2)
    Set mReverseFieldsTest = sequence: Set mReverseTraceTest = trace
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\IntegratedReturn"
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
values['Input/Reverse1.bin']=(frozen/'Input/Reverse1.bin').read_bytes()
values['Input/Source745History.bin']=(frozen/'Input/Source745History.bin').read_bytes()
values['Input/ExpectedOriginalControl.bin']=(audit/'Native/Departure.Candidate1.bin').read_bytes()
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
imports=['CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState']
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')').replace('\n','\r\n').encode('utf-8-sig')
values['Run-TestRealColdReplay.ps1']=(frozen/'Run-TestRealColdReplay.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='IntegratedOriginalTargetReturnDirected',actual_import_this_pass=imports,
    late_H743_policy_preserved=True,material_exact743=True,
    integrated751_helper_has_non_reentrant_error_cleanup=True,
    one_original_recovery_call_from_exact46_states=True,source_history_verified747=True,
    only_remaining_primary_steps_allowed_and_all_effort_combined=True,
    original_lambda_force_steel_direction_guards_preserved=True,
    diagnostic_sequence_not_controller_branch_admission=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('752 INTEGRATED RETURN ONE ORIGINAL CALL; ERROR CLEANUP GUARDED; EXACT46 HISTORY; LATE H RETAINED')
