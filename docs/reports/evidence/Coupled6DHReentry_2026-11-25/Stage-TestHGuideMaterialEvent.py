"""ДЛЯ ТЕСТОВ. Проверяет существующий решатель материальной границы на H-направлении от последнего активного состояния70."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold840EnvelopeReentry_20261125';root=base/'Cold841HGuideMaterialEvent_20261125'
assert json.loads((parent/'EnvelopeReentryProof.json').read_bytes())['no_reentry_pair_accepted']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
r=Reader((parent/'Native/EnvelopeReentry.bin').read_bytes())
for _ in range(5):meta,effort,guide,actual=[r.VectorTest() for _ in range(4)]
assert meta[0]==5 and meta[1]==1 and meta[2]==.0001
old=base/'Cold833ContinuedOutgoingWalk_20261125'
attempts=json.loads((old/'AutomaticLoadingWalkProof.json').read_bytes())['attempts']
entry=next(item for item in attempts if item['source_count']==70)
r=Reader((old/'Native'/f"LoadingWalk.{entry['attempt']:03}.bin").read_bytes())
packet=[r.VectorTest() for _ in range(15)];assert packet[0][6]==source['lambdas'][69]
np.testing.assert_array_equal(packet[2],source['fields'][69])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'HGuideMaterialEventCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source,
    dict(fields=source['fields'][:70],fractions=source['fractions'][:70],lambdas=source['lambdas'][:70]))
values['Input/HMaterialGuide.bin']=vector([meta[6],packet[0][12]])+vector(guide)+vector(packet[1])
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Даёт проверяемый вход к действующей обработке материального события от явно заданного диагностического отрезка.
' Не меняет уравнения, критерии половинного подхода или историю; точное восстановление выполняет вызывающий тест.
Friend Function TestMaterialEventBracket(ByRef bracket() As Double, ByVal reference As Double, ByVal bracketLambda As Double, _
        ByRef arcProjection() As Double, ByVal q As Double, ByRef candidate() As Double, ByRef candidateLambda As Double, _
        ByRef eventInfo As Variant, ByRef face() As Double, ByRef half() As Double, ByRef record As Variant) As Boolean
    mCalls = 0: mPrimary = 0: mError = -1#: Set mApproaches = New Collection
    Dim accepted() As Double: mSolver.CopyAccepted accepted
    Dim revision As Long: revision = mSolver.AcceptedRevision
    Dim phase As Long, actualFace() As Double, actualHalf() As Double, halfLambda As Double, error As Double, solved As Boolean
    Dim eventLambda As Double, kind As Long, verified As Boolean
    ReDim candidate(1 To 1): ReDim face(1 To 1): ReDim half(1 To 1)
    verified = TryEvent(accepted, bracket, reference, bracketLambda, arcProjection, q, candidate, candidateLambda, phase, _
        eventInfo, face, half, actualFace, actualHalf, halfLambda, error, solved, 96, eventLambda, kind)
    If mSolver.AcceptedRevision <> revision Or phase > mMaximumIterations Or mCalls > 96 Then Err.Raise 5, , "Проба материального отрезка нарушила историю или бюджет."
    record = Array(Abs(CLng(verified)), mCalls, mPrimary, phase, error, kind, Abs(CLng(solved)), eventLambda, candidateLambda, halfLambda, revision, mSolver.AcceptedRevision)
    TestMaterialEventBracket = verified
End Function
''';values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert code.count('    TestReadUnloadingAndEnvelopeReentry\n')==1
code=code.replace('    TestReadUnloadingAndEnvelopeReentry\n','    TestHGuideMaterialEvent\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет рабочую границу и половинный подход по сохранённому H-направлению; ни один найденный кандидат не принимается.
Private Sub TestHGuideMaterialEvent()
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\HMaterialGuide.bin" For Binary Access Read As #inputHandle
    Dim raw As Variant: raw = TestReadPolishVector(inputHandle)
    Dim guideLambda As Double: guideLambda = raw(1)
    Dim q As Double: q = raw(2)
    Dim guide() As Double, arcProjection() As Double, i As Long
    raw = TestReadPolishVector(inputHandle): ReDim guide(1 To UBound(raw))
    For i = 1 To UBound(raw): guide(i) = raw(i): Next i
    raw = TestReadPolishVector(inputHandle): ReDim arcProjection(1 To UBound(raw))
    For i = 1 To UBound(raw): arcProjection(i) = raw(i): Next i
    Close #inputHandle
    Dim fraction As Double, reference As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, reference
    Dim arc As New CCoupledArcContinuation
    arc.Initialize mReverseSolverTest, mReverseBaseTest, mReverseDirectionTest, mReverseOptionsTest(3), mReverseOptionsTest(4), 30
    Dim candidate() As Double, factor As Double, eventInfo As Variant, face() As Double, half() As Double, record As Variant, solved As Boolean
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim accepted() As Double: mReverseSolverTest.CopyAccepted accepted
    solved = arc.TestMaterialEventBracket(guide, reference, guideLambda, arcProjection, q, candidate, factor, eventInfo, face, half, record)
    Dim eventState As Variant: eventState = Array(0#)
    If eventInfo(0) Then
        Dim concrete As CCoupledConcreteResponse
        Select Case CLng(eventInfo(2))
            Case -1: Set concrete = mReverseSidesTest(1)
            Case 0: Set concrete = mReverseConcreteTest
            Case 1: Set concrete = mReverseSidesTest(3)
        End Select
        eventState = concrete.ReadNewtonYieldEventStateTest(CLng(eventInfo(1)))
    End If
    Dim restored As Boolean, restoredLambda As Double, restoredField() As Double, actual() As Double, projection() As Double
    Dim coordinate As Double, spacing As Double, movement() As Double, prediction As Variant
    If solved Then
        movement = candidate
        For i = 1 To UBound(movement): movement(i) = candidate(i) - accepted(i): Next i
        If mReverseSolverTest.SelectSpatialDirectionControl(accepted, candidate, movement, projection, coordinate, spacing) Then
            prediction = candidate
            restored = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, coordinate, factor, restoredLambda, _
                mReverseOptionsTest(3) * 0.0001, mReverseOptionsTest(4) * 0.0001, 30 - CLng(record(3)), consistentControl:=True, initialField:=prediction)
            If restored Then
                mReverseSolverTest.CopyTrial restoredField, actual
                restored = restoredLambda = factor And mReverseSolverTest.IterationCount = 0
                For i = 1 To UBound(candidate)
                    If restoredField(i) <> candidate(i) Then restored = False
                Next i
            End If
        End If
    End If
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\HGuideEvent.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMHFACEPAIR0001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, record: TestWriteReverseBinaryVector handle, eventInfo
    TestWriteReverseBinaryVector handle, eventState
    TestWriteReverseBinaryVector handle, face: TestWriteReverseBinaryVector handle, half: TestWriteReverseBinaryVector handle, candidate
    TestWriteReverseBinaryVector handle, Array(Abs(CLng(restored)), restoredLambda, revision, mReverseSolverTest.AcceptedRevision)
    TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
    Close #handle
    If UBound(face) > 1 Then TestWriteElasticPrincipalRoot 1, face
    If UBound(half) > 1 Then TestWriteElasticPrincipalRoot 2, half
    If solved Then TestWriteElasticPrincipalRoot 3, candidate
    mReverseSolverTest.DiscardStep
    Dim checked() As Double: mReverseSolverTest.CopyAccepted checked
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Материальный отрезок принял новую историю."
    For i = 1 To UBound(accepted)
        If accepted(i) <> checked(i) Then Err.Raise 5, , "Материальный отрезок изменил исходное принятое поле."
    Next i
End Sub
''';values[name]=code
imports=['CCoupledArcContinuation','modTestCoupledElasticState'];compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1)
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if isinstance(data,str):
        if n.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingMaterialEventOnHReentryGuide',actual_import_this_pass=sorted(imports),actual_component_count=256,
    all_working_math_and_materials_exact840=True,test_only_arc_adapter_added=True,
    exact70_prefix_from_verified113=True,guide_from_actual_strict840_case5=True,incoming_arc_plane_from_actual833_source70=True,
    working_event_half_pair_and30_iteration_criteria_unchanged=True,no_candidate_history_accepted=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
