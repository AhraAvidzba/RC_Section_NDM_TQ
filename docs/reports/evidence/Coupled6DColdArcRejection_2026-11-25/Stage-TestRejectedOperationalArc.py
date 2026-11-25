"""ДЛЯ ТЕСТОВ. Воспроизводит первую отвергнутую рабочую дугу холодного запуска на точном префиксе10, не меняя численный код."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re,struct
import numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold787ArcContextOrder_20261125';cold=base/'Cold785OperationalArcController_20261125';root=base/'Cold789RejectedOperationalArc_20261125'
proof=json.loads((cold/'OperationalColdProof.json').read_bytes());assert proof['exact_prefix_with_warm_input']==10
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](cold/'Native/Stock.RawHistory.bin')
prefix={k:frames[-1][k][:10] for k in ('fields','fractions','lambdas')}
for k in prefix:assert np.array_equal(prefix[k],original[k][:10]),k
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'RejectedOperationalArcCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert code.count('    TestRunArcFailureContext\n')==1;code=code.replace('    TestRunArcFailureContext\n','',1)
method,=re.findall(r'(?ms)^Private Sub TestRunOperationalArc\(\).*?^End Sub',code)
anchor='    solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, 2#, 6, factor, projection)\n';assert method.count(anchor)==1
updated=method.replace(anchor,'''    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\OperationalArcRatio.bin" For Binary Access Read As #inputHandle
    Dim ratioData As Variant: ratioData = TestReadPolishVector(inputHandle): Close #inputHandle
    solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratioData(1), 6, factor, projection)
''',1)
updated=updated.replace('magic = "NDMOPERARC000001"','magic = "NDMOPERARC000002"',1)
anchor='    For i = LBound(certificate) To UBound(certificate): TestWriteReverseBinaryVector handle, certificate(i): Next i\n';assert updated.count(anchor)==1
updated=updated.replace(anchor,anchor+'''    Dim eventInfo As Variant, sourceState As Variant, response As CCoupledConcreteResponse
    eventInfo = certificate(9): sourceState = Array(0#)
    If eventInfo(0) Then
        Select Case CLng(eventInfo(2))
            Case -1: Set response = mReverseSidesTest(1)
            Case 0: Set response = mReverseConcreteTest
            Case 1: Set response = mReverseSidesTest(3)
        End Select
        sourceState = response.ReadNewtonYieldEventStateTest(CLng(eventInfo(1)))
    End If
    TestWriteReverseBinaryVector handle, sourceState
''',1)
anchor='    If solved Then TestExportSignedArcTrial "OperationalArc.Full", field, factor\n';assert updated.count(anchor)==1
updated=updated.replace(anchor,'''    If solved Then TestExportSignedArcTrial "OperationalArc.Full", field, factor
    Dim status As Variant, traceField() As Double: status = certificate(0)
    If status(0) Then
        traceField = certificate(6): TestExportSignedArcTrial "OperationalArc.PairFull", traceField, status(4)
    End If
    If status(1) Then
        traceField = certificate(5): TestExportSignedArcTrial "OperationalArc.PairHalf", traceField, status(5)
    End If
    If status(26) Then
        traceField = certificate(10): TestExportSignedArcTrial "OperationalArc.Event", traceField, status(23)
        traceField = certificate(11)
        If UBound(traceField) > 1 Then TestExportSignedArcTrial "OperationalArc.EventHalf", traceField, status(24)
    End If
''',1)
code=code.replace(method,updated,1)
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
    'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,prefix),
    'Input/OperationalArcRatio.bin':struct.pack('<id',1,1.),'Input/Source785RawHistory.bin':(cold/'Native/Stock.RawHistory.bin').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ActualColdPrefix10RejectedArc',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    source10_exact_cold785_and787=True,all_working_numerical_methods_exact787=True,ratio_from_data=1.,
    rejected_pair_and_event_full_half_native_stress_witnesses=True,source_event_parameters_captured_before_export=True,
    no_cold_capacity_search=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('789 EXACT COLD PREFIX10; ORIGINAL ARC RATIO1; PAIR/EVENT FULL/HALF RAW WITNESSES; SAME787 KERNEL')
