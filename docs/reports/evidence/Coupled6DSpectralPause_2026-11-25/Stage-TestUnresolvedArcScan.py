"""ДЛЯ ТЕСТОВ. Читает материальные поверхности и напряжения на отклонённых дуговых полях точного86-точечного пути, не решая новые НДС."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold833ContinuedOutgoingWalk_20261125';root=base/'Cold834UnresolvedArcScan_20261125'
proof=json.loads((parent/'AutomaticLoadingWalkProof.json').read_bytes())
assert proof['new_accepted_equilibria']==17 and proof['initial_source_states']==69
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
read=run_path(str(here/'Read-TestColdRawCapture.py'));Reader=read['TestReader']
fields=list(original['fields']);lambdas=list(original['lambdas']);failed={}
for item in proof['attempts']:
    r=Reader((parent/'Native'/('LoadingWalk.'+str(item['attempt']).zfill(3)+'.bin')).read_bytes())
    assert r.raw[:16]==b'NDMOPERARC000004'
    packet=[r.VectorTest() for _ in range(15)]
    if item['solved']:
        field,actual,projection=[r.VectorTest() for _ in range(3)];fields.append(field);lambdas.append(packet[0][23])
    else:failed[item['attempt']]=packet
    r.VectorTest();r.VectorTest();r.VectorTest();assert r.position==len(r.raw)
assert len(fields)==86
cases=[('Full28',failed[28][6]),('Half28',failed[28][5]),('Full31',failed[31][6]),('Half31',failed[31][5])]
for ratio in (2.,.25):
    prediction=fields[-1].copy();prediction[:56]+=ratio*(fields[-1][:56]-fields[-2][:56])
    cases.append(('Predictor'+str(ratio),prediction))
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
frame=dict(fields=np.array(fields),fractions=np.ones(len(fields)),lambdas=np.array(lambdas))
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'UnresolvedScanCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n';assert code.count(anchor)==1
code=code.replace(anchor,'    TestReadUnresolvedArcScan\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает пробную поверхность и бетонный отклик вдоль заданных диагностических отрезков; поля не объявляются принятой ветвью.
Private Sub TestReadUnresolvedArcScan()
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim source() As Double: mReverseSolverTest.CopyAccepted source
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\ArcScanFields.bin" For Binary Access Read As #inputHandle
    Dim header As Variant: header = TestReadPolishVector(inputHandle)
    Dim caseIndex As Long, point As Long, side As Long, i As Long, level As Variant
    Dim prediction() As Double, field() As Double, value As Double, available As Boolean
    Dim strain(1 To 6) As Double, slope(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim material As CCoupledConcreteResponse, response As ICoupledPointResponse, state As Variant, beforeState As Variant, failure As ESolverFailureCode
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ArcScan.csv" For Output As #handle
    Dim stateHandle As Integer: stateHandle = FreeFile
    Open mPolishRootTest & "\\Native\\ArcScan.States.bin" For Binary Access Write As #stateHandle
    Dim magic As String * 16: magic = "NDMARCSCAN000001"
    Put #stateHandle, , magic
    For side = -1 To 1
        Select Case side
            Case -1: Set material = mReverseSidesTest(1)
            Case 0: Set material = mReverseConcreteTest
            Case 1: Set material = mReverseSidesTest(3)
        End Select
        For point = 1 To mReverseOperatorsTest.PointCount
            TestWriteReverseBinaryVector stateHandle, Array(point, side)
            TestWriteReverseBinaryVector stateHandle, material.ReadNewtonYieldEventStateTest(point)
        Next point
    Next side
    Close #stateHandle
    For caseIndex = 1 To CLng(header(1))
        state = TestReadPolishVector(inputHandle): ReDim prediction(1 To UBound(state))
        For i = 1 To UBound(state): prediction(i) = state(i): Next i
        For Each level In Array(0#, 0.03125, 0.0625, 0.125, 0.25, 0.5, 0.75, 1#)
            field = source
            For i = 1 To UBound(field): field(i) = source(i) + CDbl(level) * (prediction(i) - source(i)): Next i
            For side = -1 To 1
                Select Case side
                    Case -1: Set material = mReverseSidesTest(1)
                    Case 0: Set material = mReverseConcreteTest
                    Case 1: Set material = mReverseSidesTest(3)
                End Select
                Set response = material
                For point = 1 To mReverseOperatorsTest.PointCount
                    beforeState = material.ReadNewtonYieldEventStateTest(point)
                    available = mReverseSolverTest.TryReadLoadingSurface(point, side, field, value)
                    Print #handle, "Y;" & CStr(caseIndex) & ";" & TestOnsetNumber(CDbl(level)) & ";" & CStr(point) & ";" & CStr(side) & ";" & _
                        CStr(available) & ";" & TestOnsetNumber(value)
                    mReverseOperatorsTest.PointStrain point, field, strain
                    mReverseOperatorsTest.AxialGradientAtPoint point, field, slope
                    For i = 1 To 6: strain(i) = strain(i) + side * 0.001 * slope(i): Next i
                    failure = response.EvaluatePoint(point, strain, stress, tangent, ctmStressOnly, recordTrial:=False)
                    Print #handle, "S;" & CStr(caseIndex) & ";" & TestOnsetNumber(CDbl(level)) & ";" & CStr(point) & ";" & CStr(side) & ";" & _
                        CStr(failure) & TestHardPreparedVector(stress)
                    state = material.ReadNewtonYieldEventStateTest(point)
                    For i = LBound(state) To UBound(state)
                        If state(i) <> beforeState(i) Then Err.Raise 5, , "Чтение отрезка изменило принятую материальную историю."
                    Next i
                Next point
            Next side
        Next level
    Next caseIndex
    If Seek(inputHandle) <> LOF(inputHandle) + 1 Then Err.Raise 5, , "Лишние данные отрезков материального сканирования."
    Close #inputHandle: Close #handle
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Сканирование приняло новое поле."
    Dim after() As Double: mReverseSolverTest.CopyAccepted after
    For i = 1 To UBound(source)
        If after(i) <> source(i) Then Err.Raise 5, , "Сканирование изменило исходное принятое поле."
    Next i
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frame),
    'Input/ArcScanFields.bin':vector([len(cases)])+b''.join(vector(f) for _,f in cases),
    'Input/ArcScanCases.json':(json.dumps([dict(case=i+1,name=n) for i,(n,_) in enumerate(cases)],indent=2)+'\n').encode(),
    'Input/Source833Proof.json':(parent/'AutomaticLoadingWalkProof.json').read_bytes(),
    'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='Frozen86StateUnresolvedArcMaterialScan',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    all86source_fields_exact833=True,all_working_methods_and_materials_exact833=True,
    six_diagnostic_segments_eight_fractions_all129_stations=True,
    scalar_yield_and_stress_only_read_without_accept=True,no_new_global_NDS_search=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
