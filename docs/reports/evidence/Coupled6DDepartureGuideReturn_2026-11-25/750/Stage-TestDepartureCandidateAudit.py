"""ДЛЯ ТЕСТОВ. Наблюдает материальный кандидат перед проверкой прироста lambda, не меняя эту проверку или алгоритм материала."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold749DepartureCandidateAudit_20261125';parent=base/'Cold745EarlyGrowingEnvelope_20261125'
frozen=base/'Cold747EarlyEnvelopeFrozenReplay_20261125'
proof=json.loads((frozen/'EarlyEnvelopeFrozenReplayProof.json').read_bytes())
assert proof['states']==46 and proof['native']['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'DepartureAuditCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
values['Input/Reverse1.bin']=(frozen/'Input/Reverse1.bin').read_bytes()
values['Input/Source745History.bin']=(frozen/'Input/Source745History.bin').read_bytes()
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет одну штатную процедуру резерва после H от заново восстановленных46 состояний; начальный бюджет относится к этой отдельной диагностической пробе.
Friend Function TestAuditEnvelopeDeparture(ByVal solver As CCoupledNewtonSolver, ByVal reinforcement As CCoupledReinforcedResponse, _
        ByRef offset() As Double, ByRef direction() As Double, ByVal history As Collection, ByVal trace As CCoupledLoadTrace, _
        ByVal options As Variant) As Variant
    Set mSolver = solver: Set mReinforcement = reinforcement
    mOffset = offset: mDirection = direction
    Set mTrace = trace.CopyPrefix(trace.Count): Set mCurrentFields = history
    Set mMeta = New CResultMeta
    mMaxIterations = CLng(options(2)): mAbsTol = options(3): mRelTol = options(4)
    mLambdaTolerance = options(5): mMaximumLambda = options(6)
    mEnvelopeDepartureRevision = solver.AcceptedRevision: mDepartureRecoveryRevision = -1
    mMaterialRecoveryRevision = -1: mFailedProbeRevision = -1: mLastHybridRevision = -1
    mUsesFieldControl = True: mFieldControlChanges = 1
    ReDim mProjection(1 To solver.PhysicalDegreeCount + 4)
    Dim current As TBranchPoint, trial As TBranchPoint, fraction As Double, factor As Double, index As Long
    current.HistoryIndex = history.Count
    For index = 1 To trace.Count
        trace.ReadPoint index, fraction, factor
        mBestLambda = GeomMax(mBestLambda, factor)
        If fraction = 1# And factor = 1# Then mDesignHistoryIndex = index
    Next index
    trace.ReadPoint trace.Count, fraction, current.Lambda
    Dim revision As Long: revision = solver.AcceptedRevision
    Dim solved As Boolean: solved = TryRecoverEnvelopeDeparture(history, current, trial)
    If solver.AcceptedRevision <> revision Then Err.Raise 5, , "Диагностическая процедура приняла материал."
    TestAuditEnvelopeDeparture = Array(Abs(CLng(solved)), mProbeCount, trial.Lambda, current.Lambda, revision, solver.AcceptedRevision, mProbeDetails)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Function TryDifferenceMaterialField\(.*?^End Function',code)
old='    If Not TrialLambdaUncertainty(absoluteTolerance, relativeTolerance, uncertainty) Then GoTo Restore\n'
new='''    Dim uncertaintyReadyTest As Boolean
    uncertaintyReadyTest = TrialLambdaUncertainty(absoluteTolerance, relativeTolerance, uncertainty)
    modTestCoupledElasticState.TestCaptureDepartureCandidate Me, field, projection, originalProjection, originalTarget, _
        lambda, referenceLambda, uncertainty, uncertaintyReadyTest, tolerance, EquilibriumNorm(load, field)
    If Not uncertaintyReadyTest Then GoTo Restore
'''
assert method.count(old)==1
changed=method.replace(old,new,1);assert changed.replace(new,old,1)==method
code=code.replace(method,changed,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(frozen/name).read_text('utf-8-sig')
assert 'mDepartureAuditRootTest' not in code
code=code.replace('Option Explicit\n','''Option Explicit
Private mDepartureAuditRootTest As String, mDepartureAuditCountTest As Long
Private mDepartureAuditFieldTest() As Double, mDepartureAuditFactorTest As Double
''',1)
pattern=r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub';method,=re.findall(pattern,code)
needle='    mPolishExecutedTest = True\n';assert method.count(needle)==1
code=code.replace(method,method.replace(needle,'    TestRunDepartureCandidateAudit\n'+needle,1),1)
old='''Private Sub TestExportFastFrozenReplay()
    Const caseIndex As Long = 0'''
new='Private Sub TestExportFastFrozenReplay(Optional ByVal caseIndex As Long = 0)'
assert code.count(old)==1;code=code.replace(old,new,1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет поля и оценки непосредственно перед неизменённой проверкой прироста lambda; только чтение, без новой материальной пробы.
Public Sub TestCaptureDepartureCandidate(ByVal solver As CCoupledNewtonSolver, ByRef field() As Double, ByRef projection() As Double, _
        ByRef oldProjection() As Double, ByVal oldTarget As Double, ByVal factor As Double, ByVal reference As Double, _
        ByVal uncertainty As Double, ByVal uncertaintyReady As Boolean, ByVal tolerance As Double, ByVal norm As Double)
    If Len(mDepartureAuditRootTest) = 0 Then Exit Sub
    Dim accepted() As Double: solver.CopyAccepted accepted
    mDepartureAuditCountTest = mDepartureAuditCountTest + 1
    Dim file As String: file = mDepartureAuditRootTest & "\\Native\\Departure.Candidate" & CStr(mDepartureAuditCountTest) & ".bin"
    If Len(Dir$(file)) > 0 Then Err.Raise 5, , "Диагностический кандидат уже записан."
    Dim handle As Integer: handle = FreeFile
    Open file For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMDEPARTCAND001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(factor, reference, uncertainty, Abs(CLng(uncertaintyReady)), tolerance, norm, oldTarget, solver.AcceptedRevision, mReverseCalibrationTest(0))
    TestWriteReverseBinaryVector handle, accepted: TestWriteReverseBinaryVector handle, field
    TestWriteReverseBinaryVector handle, projection: TestWriteReverseBinaryVector handle, oldProjection
    Close #handle
    mDepartureAuditFieldTest = field: mDepartureAuditFactorTest = factor
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет резерв от точного46-точечного хвоста; отдельный кандидат выгружается как диагностическая последовательность и не принимается контроллером.
Private Sub TestRunDepartureCandidateAudit()
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    mDepartureAuditCountTest = 0: mDepartureAuditRootTest = mPolishRootTest
    Dim branch As New CCoupledDeformationContinuation, outcome As Variant
    outcome = branch.TestAuditEnvelopeDeparture(mReverseSolverTest, mReverseReinforcedTest, mReverseBaseTest, mReverseDirectionTest, _
        mReverseFieldsTest, mReverseTraceTest, mReverseOptionsTest)
    mDepartureAuditRootTest = vbNullString
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Резерв изменил принятую историю."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\Departure.Outcome.txt" For Output As #handle
    Print #handle, "SOLVED;" & CStr(outcome(0)) & ";PROBES;" & CStr(outcome(1)) & ";LAMBDA;" & CStr(outcome(2)) & ";CANDIDATES;" & CStr(mDepartureAuditCountTest)
    Print #handle, "DETAILS;" & CStr(outcome(6))
    Print #handle, "ACCEPTED_REVISION_UNCHANGED;" & CStr(revision)
    Close #handle
    mReverseSolverTest.DiscardStep
    If mDepartureAuditCountTest = 0 Then Exit Sub
    Dim originalFields As Collection, originalTrace As CCoupledLoadTrace, sequence As New Collection, trace As CCoupledLoadTrace, index As Long
    Set originalFields = mReverseFieldsTest: Set originalTrace = mReverseTraceTest
    For index = 1 To originalFields.Count: sequence.Add originalFields(index): Next index
    sequence.Add mDepartureAuditFieldTest: Set trace = originalTrace.CopyPrefix(originalTrace.Count): trace.Add 1#, mDepartureAuditFactorTest
    Set mReverseFieldsTest = sequence: Set mReverseTraceTest = trace
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\Departure"
    mPtcSuppressTest = True
    TestExportFastFrozenReplay 999
    mPtcSuppressTest = False
    Set mReverseFieldsTest = originalFields: Set mReverseTraceTest = originalTrace
End Sub
'''
methods=re.findall(r'(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)',code)
assert len(methods)==len(set(n.casefold() for n in methods))
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
imports=['CCoupledNewtonSolver','CCoupledDeformationContinuation','modTestCoupledElasticState']
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')').replace('\n','\r\n').encode('utf-8-sig')
values['Run-TestRealColdReplay.ps1']=(frozen/'Run-TestRealColdReplay.ps1').read_bytes()
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='DepartureMaterialCandidateReadOnlyAudit',actual_import_this_pass=imports,
    source_states=46,original_history_exact745_verified747=True,material_model_exact745=True,
    Newton_numeric_method_exact745_except_readonly_capture=True,unchanged_small_lambda_guard=True,
    unchanged_full_equilibrium_and_uncertainty_checks=True,isolated_departure_retry_primary_iterations=30,
    fresh_replay_not_claimed_bit_exact_original_failed_call=True,diagnostic_candidate_not_branch_admission=True,
    no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('749 DEPARTURE RETRY FROM EXACT46; READ CANDIDATE BEFORE UNCHANGED LAMBDA GUARD; NO BRANCH ADMISSION')
