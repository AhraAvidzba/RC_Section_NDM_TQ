"""ДЛЯ ТЕСТОВ. Читает чувствительность λ в фактическом последнем равновесии825 без повторения поиска или принятия нового состояния."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold826AutomaticColdReplay_20261125';source=base/'Cold825ExtendedLoadingWalk_20261125'
root=base/'Cold829OutgoingArcUncertainty_20261125'
proof=json.loads((source/'AutomaticLoadingWalkProof.json').read_bytes())
assert proof['new_accepted_equilibria']==26 and proof['all_half_roots_independently_rebuilt']
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
fields=list(original['fields']);lambdas=list(original['lambdas'])
for attempt in proof['attempts']:
    if not attempt['solved']:continue
    raw=(source/'Native'/('LoadingWalk.'+str(attempt['attempt']).zfill(3)+'.bin')).read_bytes()
    r=Reader(raw);assert raw[:16]==b'NDMOPERARC000004'
    packet=[r.VectorTest() for _ in range(15)];field,actual,projection=[r.VectorTest() for _ in range(3)]
    r.VectorTest();r.VectorTest();r.VectorTest();assert r.position==len(raw)
    assert packet[0][23]==attempt['lambda_value']
    fields.append(field);lambdas.append(packet[0][23])
frame=dict(fields=np.array(fields),fractions=np.ones(len(fields)),lambdas=np.array(lambdas))
assert len(fields)==69
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ArcUncertaintyCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    mPtcSuppressTest = False\n    mPolishExecutedTest = True';assert code.count(anchor)==1
code=code.replace(anchor,'    mPtcSuppressTest = False\n    TestReadOutgoingArcUncertainty\n    mPolishExecutedTest = True',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Восстанавливает уже известный корень без поправки и вызывает рабочую оценку чувствительности на предыдущей принятой истории.
' Пластическая история не принимается; отказ сопряжённой системы не становится физическим разрушением.
Private Sub TestReadOutgoingArcUncertainty()
    Dim count As Long: count = mReverseFieldsTest.Count
    If Not mReverseSolverTest.ReplayAcceptedFields(mReverseFieldsTest, count - 1) Then Err.Raise 5, , "Не восстановлена история перед известным корнем."
    Dim accepted() As Double, field() As Double, movement() As Double, projection() As Double, i As Long
    accepted = mReverseFieldsTest(count - 1): field = mReverseFieldsTest(count): movement = field
    For i = 1 To UBound(field): movement(i) = field(i) - accepted(i): Next i
    Dim coordinate As Double, spacing As Double, fraction As Double, reference As Double, before As Double
    If Not mReverseSolverTest.SelectSpatialDirectionControl(accepted, field, movement, projection, coordinate, spacing) Then Err.Raise 5, , "Нет меры известного исходящего корня."
    mReverseTraceTest.ReadPoint count, fraction, reference
    mReverseTraceTest.ReadPoint count - 1, fraction, before
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim seed As Variant: seed = field
    Dim restored As Boolean, restoredLambda As Double
    restored = mReverseSolverTest.TryDeformationStep(mReverseBaseTest, mReverseDirectionTest, projection, coordinate, reference, restoredLambda, _
        mReverseOptionsTest(3) * 0.0001, mReverseOptionsTest(4) * 0.0001, 1, consistentControl:=True, initialField:=seed)
    If Not restored Or mReverseSolverTest.IterationCount <> 0 Or restoredLambda <> reference Then Err.Raise 5, , "Известный исходящий корень потребовал новой поправки."
    Dim actual() As Double, checked() As Double: mReverseSolverTest.CopyTrial checked, actual
    For i = 1 To UBound(field)
        If field(i) <> checked(i) Then Err.Raise 5, , "Изменилось известное поле исходящей ветви."
    Next i
    Dim beforeEffort As Variant: beforeEffort = mReverseSolverTest.ReadPolishEffortTest
    Dim uncertainty As Double, available As Boolean
    mStockProbeRootTest = mPolishRootTest
    available = mReverseSolverTest.TrialLambdaUncertainty(mReverseOptionsTest(3) * 0.0001, mReverseOptionsTest(4) * 0.0001, uncertainty)
    mStockProbeRootTest = vbNullString
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Оценка чувствительности приняла материальную историю."
    Dim afterEffort As Variant: afterEffort = mReverseSolverTest.ReadPolishEffortTest
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\OutgoingUncertainty.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMARCUNCERT0001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(count, reference, before, count - 1, revision, mReverseSolverTest.AcceptedRevision, _
        Abs(CLng(restored)), Abs(CLng(available)), uncertainty, beforeEffort(0), mAdjointProbeCountTest)
    TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, projection
    TestWriteReverseBinaryVector handle, beforeEffort: TestWriteReverseBinaryVector handle, afterEffort
    Close #handle
    mReverseSolverTest.DiscardStep
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frame),
    'Input/Source825Attempts.json':(source/'AutomaticLoadingWalkProof.json').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OutgoingArcLambdaSensitivityAtActual825Endpoint',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    actual_source69_fields_rebuilt_from_raw_825_packets=True,numerical_kernel_and_materials_exact826=True,
    known_root_restored_with_zero_iterations=True,existing_TrialLambdaUncertainty_used=True,
    no_new_NDS_search_or_material_acceptance=True,adjoint_failure_not_physical_failure=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
