"""ДЛЯ ТЕСТОВ. Сравнивает полный H-шаг с двумя последовательно принятыми половинами и проверяет достижимость следующего уровня после промежуточной истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold888StrictEnvelopeGrowth_20261126';source=base/'Cold886Cold56LargerHPairs_20261126';root=base/'Cold891SequentialAcceptedH_20261126'
proof=json.loads((source/'Cold56LargerHPairsProof.json').read_bytes());assert proof['verified_pairs']==4
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SequentialHCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
entry='''Public Function TestRunSequentialAcceptedH() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' От исходной холодной истории56 сопоставляет путь +8dH и +4dH,+4dH; затем проверяет +8dH,+8dH. Истории путей независимы и восстанавливаются с начала.
    Dim original As Collection: Set original = mReverseFieldsTest
    Dim originalTrace As CCoupledLoadTrace: Set originalTrace = mReverseTraceTest
    Dim count As Long: count = original.Count
    Dim baseline As Double, resolution As Double, beforeH As Double, afterH As Double, previous() As Double
    Dim originalLambda As Double, previousLambda As Double, fraction As Double
    originalTrace.ReadPoint count, fraction, originalLambda: originalTrace.ReadPoint count - 1, fraction, previousLambda
    Dim absTol As Double, relTol As Double: absTol = 0.0001 * mReverseOptionsTest(3): relTol = 0.0001 * mReverseOptionsTest(4)
    If Not mReverseSolverTest.TryEnvelopeControlState(mReverseBaseTest, mReverseDirectionTest, originalLambda, absTol, relTol, baseline, resolution) Then Err.Raise 5, , "Нет исходной координаты H."
    If Not mReverseSolverTest.TryAcceptedEnvelopeSegment(previous, beforeH, afterH) Then Err.Raise 5, , "Нет исходной H-пары."
    If count <> 56 Or afterH <> baseline Or afterH <= beforeH Then Err.Raise 5, , "Неверная исходная история последовательных H-шагов."
    Dim delta As Double: delta = afterH - beforeH
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\SequentialAcceptedH.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMSEQUENTIALH01"
    Put #handle, , magic: TestWriteReverseBinaryVector handle, Array(count, originalLambda, previousLambda, baseline, beforeH, delta)
    Dim records As New Collection, record As Collection, prefix As Collection, trace As CCoupledLoadTrace
    Dim caseIndex As Long, i As Long, target As Double, factor As Double, previousFactor As Double, currentH As Double, priorH As Double
    Dim ratio As Double, returned As Double, probes As Long, solved As Boolean, attempted As Boolean
    Dim accepted() As Double, seed() As Double, field() As Double, actual() As Double, projection() As Double
    Dim full8 As Variant, full4 As Variant, lambda8 As Double, lambda4 As Double, witness As Variant, vector As Variant, info As Variant
    For caseIndex = 1 To 4
        Set prefix = New Collection
        For i = 1 To original.Count: prefix.Add original(i): Next i
        Set trace = originalTrace.CopyPrefix(originalTrace.Count)
        attempted = True: target = baseline + 8# * delta
        If caseIndex = 2 Then target = baseline + 4# * delta
        If caseIndex = 3 Then
            If IsEmpty(full4) Then attempted = False Else prefix.Add full4: trace.Add 1#, lambda4
        End If
        If caseIndex = 4 Then
            target = baseline + 16# * delta
            If IsEmpty(full8) Then attempted = False Else prefix.Add full8: trace.Add 1#, lambda8
        End If
        solved = False: returned = 0#: probes = 0: witness = Empty
        If attempted Then
            If Not mReverseSolverTest.ReplayAcceptedFields(prefix, prefix.Count) Then Err.Raise 5, , "Не восстановлена промежуточная история H-пути."
            mReverseSolverTest.CopyAccepted accepted
            trace.ReadPoint trace.Count, fraction, factor: trace.ReadPoint trace.Count - 1, fraction, previousFactor
            If Not mReverseSolverTest.TryEnvelopeControlState(mReverseBaseTest, mReverseDirectionTest, factor, absTol, relTol, currentH, resolution) Then Err.Raise 5, , "Нет промежуточного H."
            If Not mReverseSolverTest.TryAcceptedEnvelopeSegment(previous, priorH, afterH) Then Err.Raise 5, , "Нет промежуточной H-пары."
            ratio = (target - currentH) / (afterH - priorH): seed = accepted
            For i = 1 To mReverseSolverTest.PhysicalDegreeCount: seed(i) = accepted(i) + ratio * (accepted(i) - previous(i)): Next i
            solved = mReverseSolverTest.TryEnvelopeStep(mReverseBaseTest, mReverseDirectionTest, previous, factor, seed, _
                factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False, True)
            witness = mReverseSolverTest.TestReadEnvelopePairVerification
        End If
        TestWriteReverseBinaryVector handle, Array(caseIndex, prefix.Count, target, Abs(CLng(solved)), returned, probes, _
            mReverseSolverTest.IterationCount, Abs(CLng(Not IsEmpty(witness))), factor, currentH, priorH, Abs(CLng(attempted)))
        TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
        If solved Then
            If IsEmpty(witness) Then Err.Raise 5, , "Нет подтверждённой полной/половинной H-пары."
            For Each vector In witness: TestWriteReverseBinaryVector handle, vector: Next vector
            mReverseSolverTest.CopyTrial field, actual
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
            If caseIndex = 1 Then full8 = field: lambda8 = returned
            If caseIndex = 2 Then full4 = field: lambda4 = returned
            Set record = New Collection: record.Add caseIndex: record.Add witness: record.Add prefix: record.Add trace: records.Add record
        End If
        mReverseSolverTest.DiscardStep
    Next caseIndex
    Close #handle
    For caseIndex = 1 To records.Count
        Set record = records(caseIndex): witness = record(2): info = witness(0)
        Set mReverseFieldsTest = record(3): Set mReverseTraceTest = record(4)
        field = witness(2): TestExportSignedArcTrial "SequentialH.Case" & CStr(record(1)) & "Full", field, CDbl(info(1))
        field = witness(3): TestExportSignedArcTrial "SequentialH.Case" & CStr(record(1)) & "Half", field, CDbl(info(2))
    Next caseIndex
    Set mReverseFieldsTest = original: Set mReverseTraceTest = originalTrace
    If Not mReverseSolverTest.ReplayAcceptedFields(original, original.Count) Then Err.Raise 5, , "Не восстановлена исходная история после H-сравнения."
    TestRunSequentialAcceptedH = "TOTAL_SEQUENTIAL_ACCEPTED_H_TEST: passed=1; failed=0; numerical outcomes recorded; paths independently replayed; not a capacity result"
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunSequentialAcceptedH'",1)
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected sequential-H setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunSequentialAcceptedH"))
        Save-Progress
''',1);values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),'Input/SourceColdRawHistory.bin':(source/'Input/SourceColdRawHistory.bin').read_bytes(),
               'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SequentialAcceptedH',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_sources_exact888=True,source_cold_history56_exact=True,independent_history_paths=True,
    full_half_pair_gates_and30iterations_unchanged=True,no_physical_capacity_claim=True,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'COMPARE +8 WITH ACCEPTED +4,+4;TRY ACCEPTED +8,+8;ALL PATHS FROM EXACT COLD56')
