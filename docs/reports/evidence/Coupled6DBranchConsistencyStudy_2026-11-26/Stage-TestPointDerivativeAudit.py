"""ДЛЯ ТЕСТОВ. Добавляет читающий контроль локальной касательной и её действительных приращений на доказанной истории20."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold867FixedMetricKernel_20261126';verified=base/'Cold869FixedMetricNextScan_20261126';root=base/'Cold872PointDerivativeAudit20_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0,verified/'Input/Reverse1.bin')
archive=root/'PointDerivativeCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
R=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader'];r=R((parent/'Native/LoadingWalk.Guide003.bin').read_bytes());r.VectorTest()
for count in (3,3):
    if r.VectorTest()[0]:
        for _ in range(count):r.VectorTest()
assert r.VectorTest()[0]==1
meta,probe,tangent,movement=[r.VectorTest() for _ in range(4)];assert r.position==len(r.raw)
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector'];values['Input/PointDerivative.bin']=vector(probe)+vector(movement)
entry='''Public Function TestReadPointDerivativeAudit() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Сопоставляет две действующие касательные с прямыми разностями при неизменной принятой истории20; НДС не решается.
    Dim accepted() As Double, field() As Double, probe() As Double, movement() As Double, revision As Long
    mReverseSolverTest.CopyAccepted accepted: revision = mReverseSolverTest.AcceptedRevision
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\PointDerivative.bin" For Binary Access Read As #inputHandle
    probe = TestReadPolishVector(inputHandle): movement = TestReadPolishVector(inputHandle)
    If Seek(inputHandle) <> LOF(inputHandle) + 1 Or UBound(probe) <> 60 Or UBound(movement) <> 60 Then Err.Raise 5, , "Неверный вход проверки касательной."
    Close #inputHandle
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\PointDerivative.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMPOINTDERIV001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(2#, mReverseOperatorsTest.PointCount, revision)
    TestWriteReverseBinaryVector handle, probe: TestWriteReverseBinaryVector handle, movement
    Dim location As Long, side As Long, point As Long, i As Long, j As Long, modeIndex As Long
    Dim strain(1 To 6) As Double, slope(1 To 6) As Double, delta(1 To 6) As Double, deltaSlope(1 To 6) As Double
    Dim station(1 To 6) As Double, direction(1 To 6) As Double, perturbed(1 To 6) As Double
    Dim stress(1 To 6) As Double, tangentMatrix(1 To 6, 1 To 6) As Double, column(1 To 6) As Double
    Dim material As CCoupledConcreteResponse, response As ICoupledPointResponse, state As Variant, afterState As Variant
    Dim plastic As Boolean, principal As Long, substeps As Long, ready As Boolean, failure As ESolverFailureCode
    Dim h As Variant, sign As Variant, mode As ECoupledTangentMode
    For location = 1 To 2
        If location = 1 Then field = accepted Else field = probe
        TestWriteReverseBinaryVector handle, field
        For side = -1 To 1
            Select Case side
                Case -1: Set material = mReverseSidesTest(1)
                Case 0: Set material = mReverseConcreteTest
                Case 1: Set material = mReverseSidesTest(3)
            End Select
            Set response = material
            For point = 1 To mReverseOperatorsTest.PointCount
                mReverseOperatorsTest.PointStrain point, field, strain
                mReverseOperatorsTest.AxialGradientAtPoint point, field, slope
                mReverseOperatorsTest.PointStrain point, movement, delta
                mReverseOperatorsTest.AxialGradientAtPoint point, movement, deltaSlope
                For i = 1 To 6
                    station(i) = strain(i) + side * 0.001 * slope(i)
                    direction(i) = delta(i) + side * 0.001 * deltaSlope(i)
                Next i
                state = material.ReadNewtonYieldEventStateTest(point)
                TestWriteReverseBinaryVector handle, Array(location, side, point)
                TestWriteReverseBinaryVector handle, station: TestWriteReverseBinaryVector handle, direction
                For modeIndex = 0 To 1
                    If modeIndex = 0 Then mode = ctmFullBranchConsistent Else mode = ctmRefinedBranchConsistent
                    failure = response.EvaluatePoint(point, station, stress, tangentMatrix, mode, recordTrial:=False)
                    TestWriteReverseBinaryVector handle, Array(modeIndex, failure)
                    If failure = sfcNone Then
                        TestWriteReverseBinaryVector handle, stress
                        For j = 1 To 6
                            For i = 1 To 6: column(i) = tangentMatrix(i, j): Next i
                            TestWriteReverseBinaryVector handle, column
                        Next j
                    End If
                Next modeIndex
                For Each h In Array(0#, 0.0001, 0.001, 0.01, 0.1)
                    For Each sign In Array(-1#, 1#)
                        For i = 1 To 6: perturbed(i) = station(i) + CDbl(sign) * CDbl(h) * direction(i): Next i
                        failure = response.EvaluatePoint(point, perturbed, stress, tangentMatrix, ctmStressOnly, recordTrial:=False)
                        ready = material.TryReadConcreteBranch(point, perturbed, plastic, principal, substeps)
                        TestWriteReverseBinaryVector handle, Array(CDbl(h), CDbl(sign), failure, Abs(CLng(ready)), Abs(CLng(plastic)), principal, substeps)
                        If failure = sfcNone Then TestWriteReverseBinaryVector handle, stress
                    Next sign
                Next h
                afterState = material.ReadNewtonYieldEventStateTest(point)
                For i = LBound(state) To UBound(state)
                    If afterState(i) <> state(i) Then Err.Raise 5, , "Чтение производной изменило принятую историю точки."
                Next i
            Next point
        Next side
    Next location
    Close #handle
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Проверка касательной приняла историю."
    TestReadPointDerivativeAudit = "TOTAL_POINT_DERIVATIVE_AUDIT_TEST: passed=1; failed=0; readonly true20; no equilibrium or capacity search"
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig');anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,old+" -EntryPointNames 'TestReadPointDerivativeAudit'",1).replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Run-Audit02ModeTests.ps1';script=(base/'Cold857SpatialDeclineRecovery_20261125'/name).read_text('utf-8-sig');anchor='        else { $result = [string]$excel.Run($macroName) }';assert script.count(anchor)==1
script=script.replace(anchor,anchor+'''
        if ($macro -ne 'modTestCoupledElasticState.RunTestReversePolish') {throw 'Unexpected derivative-audit setup macro.'}
        $lines.Add([string]$excel.Run("'RC_Section_NDM.xlsm'!modTestCoupledElasticState.TestReadPointDerivativeAudit"))
        Save-Progress
''',1)
values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ReadOnlyPointDerivativeAudit20',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,all_working_methods_exact867=True,
 source20_exact=True,accepted_history_unchanged=True,one_short_public_entry_only=True,physical_capacity_certified=False,no_new_NDS=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'SHORT READONLY ENTRY;SAME MATERIAL/WORKING KERNEL867;COARSE AND REFINED LOCAL DERIVATIVES')
