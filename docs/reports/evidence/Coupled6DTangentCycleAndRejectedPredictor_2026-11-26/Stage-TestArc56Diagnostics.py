"""ДЛЯ ТЕСТОВ. Воспроизводит рабочие дуговые пробы на холодном состоянии56 и сохраняет контекст выбора касательной, H-направления и события."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold888StrictEnvelopeGrowth_20261126';source=base/'Cold886Cold56LargerHPairs_20261126';root=base/'Cold893Arc56Diagnostics_20261126'
assert json.loads((source/'Cold56LargerHPairsProof.json').read_bytes())['verified_pairs']==4
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'Arc56CompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
entry='''Public Function TestRunArc56Diagnostics() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Повторяет восемь независимых рабочих дуговых проб; материал не принимается. Полный контекст сохраняется также при неудаче.
    If mReverseFieldsTest.Count <> 56 Then Err.Raise 5, , "Нужна точная холодная история56."
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\Arc56Diagnostics.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMARC56DIAG0001"
    Put #handle, , magic
    Dim logHandle As Integer: logHandle = FreeFile
    Open mPolishRootTest & "\\Native\\Arc56Diagnostics.txt" For Output As #logHandle
    Dim index As Long, ratio As Double, allowance As Long, solved As Boolean, factor As Double, projection() As Double
    Dim arc As CCoupledArcContinuation, revision As Long, certificate As Variant, vector As Variant, context As Variant
    Dim guide As Variant, loading As Variant, prediction As Variant, selected As Variant, group As Long
    Dim records As New Collection, record As Collection, field() As Double, actual() As Double
    For index = 1 To 8
        ratio = 0.5 ^ (index - 1): allowance = 12
        If index = 7 Then ratio = 1#: allowance = 3
        If index = 8 Then ratio = 1#: allowance = 4
        Set arc = New CCoupledArcContinuation
        arc.Initialize mReverseSolverTest, mReverseBaseTest, mReverseDirectionTest, mReverseOptionsTest(3), mReverseOptionsTest(4), CLng(mReverseOptionsTest(2))
        revision = mReverseSolverTest.AcceptedRevision: factor = 0#
        solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratio, allowance, factor, projection)
        certificate = arc.Verification
        guide = arc.EnvelopeGuideVerification: loading = arc.LoadingGuideVerification: prediction = arc.TestPredictionVerification
        If IsEmpty(certificate) Then Err.Raise 5, , "Нет контекста дуговой пробы."
        If mReverseSolverTest.AcceptedRevision <> revision Or arc.Calls > allowance Then Err.Raise 5, , "Дуговая проба изменила историю или превысила бюджет."
        TestWriteReverseBinaryVector handle, Array(index, ratio, allowance, Abs(CLng(solved)), arc.Calls, UBound(certificate) + 1, _
            Abs(CLng(Not IsEmpty(guide))), Abs(CLng(Not IsEmpty(loading))), Abs(CLng(Not IsEmpty(prediction))), factor)
        For Each vector In certificate: TestWriteReverseBinaryVector handle, vector: Next vector
        For group = 1 To 3
            Select Case group
                Case 1: selected = guide
                Case 2: selected = loading
                Case 3: selected = prediction
            End Select
            If IsEmpty(selected) Then
                TestWriteReverseBinaryVector handle, Array(0#, 0#)
            Else
                TestWriteReverseBinaryVector handle, Array(1#, UBound(selected) - LBound(selected) + 1)
                For Each vector In selected: TestWriteReverseBinaryVector handle, vector: Next vector
            End If
        Next group
        TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
        Set record = New Collection: record.Add index: record.Add certificate: record.Add solved: record.Add factor
        If solved Then
            mReverseSolverTest.CopyTrial field, actual
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual
            record.Add field: record.Add actual
        End If
        records.Add record
        Print #logHandle, "CASE;" & CStr(index) & ";RATIO;" & TestOnsetNumber(ratio) & ";ALLOWANCE;" & CStr(allowance) & ";SOLVED;" & CStr(solved)
        Print #logHandle, "DETAILS;" & arc.Details
        Print #logHandle, "SOLVER;" & mReverseSolverTest.ResultMeta.DiagnosticDetails
        mReverseSolverTest.DiscardStep
    Next index
    Close #handle: Close #logHandle
    For index = 1 To records.Count
        Set record = records(index): certificate = record(2): context = certificate(0)
        If record(3) Then
            field = record(5): TestExportSignedArcTrial "Arc56.Case" & CStr(record(1)) & "Returned", field, CDbl(record(4))
        End If
        If context(0) Then
            field = certificate(6): TestExportSignedArcTrial "Arc56.Case" & CStr(record(1)) & "Full", field, CDbl(context(4))
        End If
        If context(1) Then
            field = certificate(5): TestExportSignedArcTrial "Arc56.Case" & CStr(record(1)) & "Half", field, CDbl(context(5))
        End If
    Next index
    TestRunArc56Diagnostics = "TOTAL_ARC56_DIAGNOSTICS_TEST: passed=1; failed=0; numerical outcomes recorded; original history not accepted"
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'TestRunArc56Diagnostics'",1)
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';runner=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert runner.count(anchor)==1
runner=runner.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected arc56 setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestRunArc56Diagnostics"))
        Save-Progress
''',1);values[name]=runner.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':(source/'Input/Reverse1.bin').read_bytes(),'Input/SourceColdRawHistory.bin':(source/'Input/SourceColdRawHistory.bin').read_bytes(),
               'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Arc56Diagnostics',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all_working_sources_exact888=True,source_cold_history56_exact=True,ratios=[1,.5,.25,.125,.0625,.03125,1,1],
    allowances=[12,12,12,12,12,12,3,4],primary30_unchanged=True,no_material_history_acceptance=True,
    complete_failed_certificate_and_guides_recorded=True,physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EIGHT ACTUAL ARC PROBES;SOURCE888 UNCHANGED;FULL/HALF/LOADING/H/PREDICTION CONTEXTS')
