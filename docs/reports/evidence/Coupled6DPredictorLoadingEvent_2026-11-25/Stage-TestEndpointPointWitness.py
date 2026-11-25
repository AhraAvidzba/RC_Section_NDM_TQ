"""ДЛЯ ТЕСТОВ. Снимает ветви и локальные касательные четырёх определяющих бетонных точек вдоль захваченной поправки805."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold805EndpointLineWitness_20261125';root=base/'Cold806EndpointPointWitness_20261125'
proof=json.loads((parent/'EndpointLineWitnessProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['original804_arc_packet_bitidentical']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'EndpointPointCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());name='tests/modTestCoupledElasticState.bas'
code=(root/name).read_text('utf-8-sig')
method,=re.findall(r'(?ms)^Private Sub TestExportEndpointLineWitness\(.*?^End Sub',code)
anchor='    Dim index As Long, degree As Long, factor As Double\n';assert method.count(anchor)==1
code=code.replace(method,method.replace(anchor,anchor+'    TestExportEndpointPointWitness original, correction, fractions\n',1),1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает одинаковую принятую историю и два существующих режима касательной; ни одна локальная проба не принимается.
Private Sub TestExportEndpointPointWitness(ByRef field() As Double, ByRef correction() As Double, ByRef fractions As Variant)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\EndpointPoint.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMENDPOINT00001"
    Put #handle, , magic
    Dim points As Variant: points = Array(25, 21, 16, 10)
    TestWriteReverseBinaryVector handle, points: TestWriteReverseBinaryVector handle, fractions
    Dim pointValue As Variant, point As Long, side As Long, index As Long, i As Long, j As Long, k As Long
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    For Each pointValue In points
        point = CLng(pointValue)
        For side = -1 To 1
            Dim concrete As CCoupledConcreteResponse, response As ICoupledPointResponse
            Select Case side
                Case -1: Set concrete = mReverseSidesTest(1)
                Case 0: Set concrete = mReverseConcreteTest
                Case 1: Set concrete = mReverseSidesTest(3)
            End Select
            Set response = concrete
            Dim strain(1 To 6) As Double, increment(1 To 6) As Double, gradient(1 To 6) As Double
            Dim movement() As Double: ReDim movement(1 To UBound(field))
            For i = 1 To UBound(field): movement(i) = correction(i): Next i
            mReverseOperatorsTest.PointStrain point, field, strain
            mReverseOperatorsTest.AxialGradientAtPoint point, field, gradient
            For i = 1 To 6: strain(i) = strain(i) + side * 0.001 * gradient(i): Next i
            mReverseOperatorsTest.PointStrain point, movement, increment
            mReverseOperatorsTest.AxialGradientAtPoint point, movement, gradient
            For i = 1 To 6: increment(i) = increment(i) + side * 0.001 * gradient(i): Next i
            TestWriteReverseBinaryVector handle, Array(point, side, revision)
            TestWriteReverseBinaryVector handle, concrete.ReadNewtonYieldEventStateTest(point)
            TestWriteReverseBinaryVector handle, strain: TestWriteReverseBinaryVector handle, increment
            Dim modes As Variant: modes = Array(ctmFullBranchConsistent, ctmRefinedBranchConsistent)
            Dim mode As Variant, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double, flat(1 To 36) As Double
            Dim result As ESolverFailureCode
            For Each mode In modes
                result = response.EvaluatePoint(point, strain, stress, tangent, CLng(mode), False)
                TestWriteReverseBinaryVector handle, Array(CLng(mode), CLng(result))
                TestWriteReverseBinaryVector handle, stress
                k = 0
                For i = 1 To 6
                    For j = 1 To 6: k = k + 1: flat(k) = tangent(i, j): Next j
                Next i
                TestWriteReverseBinaryVector handle, flat
            Next mode
            Dim trial(1 To 6) As Double, plastic As Boolean, sign As Long, substeps As Long, branch As Boolean
            Dim value As Double, yieldReady As Boolean
            For index = LBound(fractions) To UBound(fractions)
                For i = 1 To 6: trial(i) = strain(i) + fractions(index) * increment(i): Next i
                branch = concrete.TryReadConcreteBranch(point, trial, plastic, sign, substeps)
                yieldReady = concrete.TryReadTrialYield(point, trial, value, gradient, False)
                result = response.EvaluatePoint(point, trial, stress, tangent, ctmStressOnly, False)
                TestWriteReverseBinaryVector handle, Array(fractions(index), Abs(CLng(branch)), Abs(CLng(plastic)), _
                    sign, substeps, Abs(CLng(yieldReady)), value, CLng(result))
                TestWriteReverseBinaryVector handle, stress
            Next index
        Next side
    Next pointValue
    Close #handle
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение локальных ветвей изменило принятую историю."
End Sub
'''
compiler=(here/'Compile-TestExternalPoint.ps1').read_bytes()
values={name:code.replace('\n','\r\n').encode('utf-8-sig'),'Compile-TestExternalPoint.ps1':compiler,
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():
 if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
 (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactEndpointLocalPointBranchesAndTangents',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
 all_working_code_exact805=True,source45_fields_exact805=True,four_dominant_concrete_points=[25,21,16,10],
 three_original_stations=True,original_and_refined_consistent_tangents=True,
 all_local_queries_no_trial_recording_or_acceptance=True,no_new_capacity_search=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('806 PURE POINT BRANCH/TANGENT READS; ALL WORKING CODE SAME805; ORIGINAL ARC MUST MATCH')
