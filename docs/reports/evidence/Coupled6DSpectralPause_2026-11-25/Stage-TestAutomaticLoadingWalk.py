"""ДЛЯ ТЕСТОВ. Продолжает точную холодную историю общей рабочей дуговой операцией без знания будущих полей."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold821PriorityLoadingArc_20261125';root=base/'Cold823AutomaticLoadingWalk_20261125'
assert json.loads((parent/'AutomaticLoadingArcProof.json').read_bytes())['pair']['verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'LoadingWalkCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRunOperationalArc\n    TestRecentLoadingResetGuard';assert code.count(anchor)==1
code=code.replace(anchor,'    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет до восьми новых равновесий общей рабочей операцией; на ревизию разрешены четыре попытки, общий бюджет96 вызовов.
' Это численная ветвь от точного холодного префикса; физическая несущая способность этим тестом не объявляется.
Private Sub TestRunAutomaticLoadingWalk()
    Dim initialCount As Long: initialCount = mReverseFieldsTest.Count
    Dim attempt As Long, attemptsAtRevision As Long, totalCalls As Long, newPoints As Long, ratio As Double
    Dim sourceCount As Long, revision As Long, fraction As Double, reference As Double, factor As Double
    Dim arc As CCoupledArcContinuation, solved As Boolean, certificate As Variant, info As Variant
    Dim projection() As Double, field() As Double, actual() As Double, half() As Double, halfLambda As Double
    Dim records As New Collection, record As Collection, pairMeta As Variant
    Dim logHandle As Integer: logHandle = FreeFile
    Open mPolishRootTest & "\\Native\\LoadingWalk.Attempts.csv" For Output As #logHandle
    ratio = 1#
    For attempt = 1 To 32
        If newPoints >= 8 Or totalCalls + 3 > 96 Or attemptsAtRevision >= 4 Then Exit For
        sourceCount = mReverseFieldsTest.Count: revision = mReverseSolverTest.AcceptedRevision
        mReverseTraceTest.ReadPoint sourceCount, fraction, reference
        Set arc = New CCoupledArcContinuation
        arc.Initialize mReverseSolverTest, mReverseBaseTest, mReverseDirectionTest, mReverseOptionsTest(3), mReverseOptionsTest(4), CLng(mReverseOptionsTest(2))
        mReverseSolverTest.DiscardStep
        solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratio, 96 - totalCalls, factor, projection)
        totalCalls = totalCalls + arc.Calls: attemptsAtRevision = attemptsAtRevision + 1
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Рабочая операция приняла историю внутри численной ветви."
        certificate = arc.Verification
        If IsEmpty(certificate) Then Err.Raise 5, , "Не сохранён сертификат численного продолжения."
        info = certificate(0)
        TestWriteAutomaticWalkPacket attempt, certificate, solved, projection
        If solved Then
            mReverseSolverTest.CopyTrial field, actual
            If factor <= 0# Or factor > mReverseOptionsTest(0) Or mReverseReinforcedTest.TrialSteelUtilization > 1# Then
                Err.Raise 5, , "Численная ветвь вышла за область нагрузки или предел арматуры."
            End If
            Select Case CLng(info(28))
                Case 1
                    half = certificate(5): halfLambda = info(5)
                Case 2, 4
                    half = certificate(11): halfLambda = info(24)
                Case 3
                    Dim approaches As Collection, approach As Variant
                    Set approaches = arc.ApproachVerification: approach = approaches(approaches.Count)
                    pairMeta = approach(0): half = approach(2): halfLambda = pairMeta(3)
                Case Else
                    Err.Raise 5, , "Неизвестный вид принятого численного шага."
            End Select
            Set record = New Collection
            record.Add sourceCount: record.Add half: record.Add halfLambda: record.Add field: record.Add factor
            records.Add record
            mReverseSolverTest.AcceptStep: mReverseFieldsTest.Add field: mReverseTraceTest.Add 1#, factor
            newPoints = newPoints + 1: attemptsAtRevision = 0
        Else
            mReverseSolverTest.DiscardStep
        End If
        Print #logHandle, CStr(attempt) & ";" & CStr(sourceCount) & ";" & CStr(solved) & ";" & CStr(newPoints) & ";" & _
            TestOnsetNumber(ratio) & ";" & TestOnsetNumber(factor) & ";" & CStr(revision) & ";" & CStr(mReverseSolverTest.AcceptedRevision) & ";" & _
            CStr(arc.Calls) & ";" & CStr(totalCalls) & ";" & TestOnsetNumber(info(19)) & ";" & TestOnsetNumber(arc.FieldError) & ";" & CStr(info(28))
        If solved Then
            ratio = GeomMin(2#, 0.8 * Sqr(0.005 / GeomMax(arc.FieldError, 0.00000001)))
        Else
            ratio = 0.5 * ratio
        End If
    Next attempt
    Close #logHandle
    If totalCalls > 96 Then Err.Raise 5, , "Превышен общий бюджет численной ветви."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\LoadingWalk.Outcome.txt" For Output As #handle
    Print #handle, "INITIAL;" & CStr(initialCount) & ";ACCEPTED;" & CStr(newPoints) & ";CALLS;" & CStr(totalCalls) & ";PHYSICAL_CAPACITY_CERTIFIED;False"
    Close #handle
    mHardPreparedPathTest = mPolishRootTest & "\\Native\\LoadingWalk.Accepted"
    mPtcSuppressTest = True: TestExportFastFrozenReplay 999: mPtcSuppressTest = False
    Dim allFields As Collection, allTrace As CCoupledLoadTrace, prefix As Collection, index As Long, degree As Long
    Set allFields = mReverseFieldsTest: Set allTrace = mReverseTraceTest.CopyPrefix(mReverseTraceTest.Count)
    For index = 1 To records.Count
        Set record = records(index): Set prefix = New Collection
        For degree = 1 To CLng(record(1)): field = allFields(degree): prefix.Add field: Next degree
        Set mReverseFieldsTest = prefix: Set mReverseTraceTest = allTrace.CopyPrefix(prefix.Count)
        half = record(2): TestExportSignedArcTrial "LoadingWalk.Half" & Format$(index, "000"), half, CDbl(record(3))
    Next index
    Set mReverseFieldsTest = allFields: Set mReverseTraceTest = allTrace
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет полный сертификат каждой рабочей операции до принятия материала, включая точный непринятый кандидат.
Private Sub TestWriteAutomaticWalkPacket(ByVal attempt As Long, ByVal certificate As Variant, ByVal solved As Boolean, ByRef projection() As Double)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\LoadingWalk." & Format$(attempt, "000") & ".bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMOPERARC000003"
    Put #handle, , magic
    Dim vector As Variant
    For Each vector In certificate: TestWriteReverseBinaryVector handle, vector: Next vector
    Dim state As Variant, eventInfo As Variant, concrete As CCoupledConcreteResponse
    eventInfo = certificate(9): state = Array(0#)
    If eventInfo(0) Then
        Select Case CLng(eventInfo(2))
            Case -1: Set concrete = mReverseSidesTest(1)
            Case 0: Set concrete = mReverseConcreteTest
            Case 1: Set concrete = mReverseSidesTest(3)
        End Select
        state = concrete.ReadNewtonYieldEventStateTest(CLng(eventInfo(1)))
    End If
    TestWriteReverseBinaryVector handle, state
    If solved Then
        Dim field() As Double, actual() As Double
        mReverseSolverTest.CopyTrial field, actual
        TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual: TestWriteReverseBinaryVector handle, projection
    End If
    TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
    Close #handle
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='WorkingAutomaticArcWalkFromExactCold43',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_methods_and_materials_exact821=True,maximum_new_points8=True,maximum_attempts32=True,maximum_calls96=True,
    attempts_per_revision4=True,all_pairs_and_exact_restoration_checked_by_working_arc=True,
    no_future_field_or_hardcoded_station=True,not_cold_capacity_search=True,
    numerical_branch_not_final_physical_admissibility=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
