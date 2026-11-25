"""ДЛЯ ТЕСТОВ. Сравнивает обычную касательную с существующей уточнённой производной выбранной ветви; рабочие классы не меняет."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold849OutgoingTangentMetric_20261125';root=base/'Cold850RefinedOutgoingTangentMetric_20261125'
assert json.loads((parent/'OutgoingTangentMetricProof.json').read_bytes())['independent_CDP_gradients_verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'RefinedOutgoingTangentCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
start=code.index('Private Sub TestReadOutgoingTangentMetric()');end=code.index('\nEnd Sub',start)
method=code[start:end]
old='Dim magic As String * 16: magic = "NDMOUTTAN0000001"';assert method.count(old)==1
method=method.replace(old,'Dim magic As String * 16: magic = "NDMOUTTAN0000002"',1)
old='''        If Not mReverseSolverTest.TestEquilibriumLinearization(field, True, True, residual, matrix) Then Err.Raise 5, , "Нет читающей линеаризации исходящей ветви."
        TestWriteReverseBinaryVector handle, Array(caseIndex + 1, levels(caseIndex), point, side, value, revision, mReverseSolverTest.AcceptedRevision)
        TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, residual
        TestWriteReverseBinaryVector handle, gradient
        ReDim column(1 To UBound(field))
        For degree = 1 To UBound(field)
            For i = 1 To UBound(field): column(i) = matrix(degree, i): Next i
            TestWriteReverseBinaryVector handle, column
        Next degree'''
assert method.count(old)==1
method=method.replace(old,'''        Dim mode As Variant, ready As Boolean, zero() As Double, projection() As Double, rawScales As Variant
        Dim normalized() As Double, normalizedResidual() As Double, fullMatrix() As Double, reference As Double, fraction As Double
        For Each mode In Array(0, 2)
            If CLng(mode) = 0 Then
                ready = mReverseSolverTest.TestEquilibriumLinearization(field, True, True, residual, matrix)
            Else
                mReverseSolverTest.DiscardStep
                mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, reference
                ReDim projection(1 To 60): ReDim zero(1 To 61)
                ready = mReverseSolverTest.BeginNewtonLineTest(mReverseBaseTest, mReverseDirectionTest, projection, 0#, field, reference)
                If ready Then
                    ready = mReverseSolverTest.TryReadBranchLinearization(zero, normalizedResidual, fullMatrix)
                    If ready Then
                        rawScales = mReverseSolverTest.ReadNewtonLineScalesTest
                        ReDim matrix(1 To 60, 1 To 60): ReDim residual(1 To 60)
                        For i = 1 To 60
                            residual(i) = normalizedResidual(i)
                            If i >= 57 Then residual(i) = residual(i) / mReverseSolverTest.ReferenceModulus
                            For degree = 1 To 60
                                matrix(degree, i) = fullMatrix(degree, i) / rawScales(degree)
                                If i >= 57 Then matrix(degree, i) = matrix(degree, i) / mReverseSolverTest.ReferenceModulus
                            Next degree
                        Next i
                    End If
                    mReverseSolverTest.EndNewtonLineTest
                End If
            End If
            TestWriteReverseBinaryVector handle, Array(caseIndex + 1, levels(caseIndex), point, side, value, revision, mReverseSolverTest.AcceptedRevision, mode, Abs(CLng(ready)))
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, gradient
            If ready Then
                TestWriteReverseBinaryVector handle, residual
                ReDim column(1 To UBound(field))
                For degree = 1 To UBound(field)
                    For i = 1 To UBound(field): column(i) = matrix(degree, i): Next i
                    TestWriteReverseBinaryVector handle, column
                Next degree
            End If
        Next mode''',1)
code=code[:start]+method+code[end:];values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExistingRefinedBranchOutgoingTangentMetric',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    exact849_prefix86=True,all_working_classes_unchanged=True,only_existing_BeginNewtonLineTest_and_TryReadBranchLinearization=True,
    read_only_linearizations_requested=8,unavailable_derivatives_recorded_without_fabrication=True,
    no_nonlinear_solution_or_history_acceptance=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'EIGHT READONLY TANGENTS;EXISTING BRANCH DERIVATIVE;WORKING CLASSES EXACT849')
