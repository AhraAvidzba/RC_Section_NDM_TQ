"""ДЛЯ ТЕСТОВ. Читает односторонние линеаризации исходящей ветви, не ослабляя ограничений расстояния и приёмки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold848NewLoadingCoordinateDiagnosis_20261125';root=base/'Cold849OutgoingTangentMetric_20261125'
proof=json.loads((parent/'NewLoadingCoordinateDiagnosisProof.json').read_bytes())
assert all(p['coordinate_records'][0]['stage']=='RadiusAndLambda' for p in proof['cases'][1:])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'OutgoingTangentCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='    TestDiagnoseNewLoadingCoordinate\n';assert code.count(old)==1
code=code.replace(old,'    TestReadOutgoingTangentMetric\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет линеаризации в принятой точке и с нагружающей стороны для независимой оценки направления и метрики исходящей ветви.
Private Sub TestReadOutgoingTangentMetric()
    Dim accepted() As Double, previous() As Double, field() As Double
    mReverseSolverTest.CopyAccepted accepted: previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim levels As Variant: levels = Array(0#, 0.25, 0.0625, 0.015625)
    Dim probe() As Double: probe = accepted
    Dim i As Long
    For i = 1 To UBound(probe): probe(i) = 2# * accepted(i) - previous(i): Next i
    Dim recent As Variant
    If Not mReverseSolverTest.TryReadRecentLoadingDeparture(probe, recent) Then Err.Raise 5, , "Не восстановлена новая граница нагружения истории86."
    Dim point As Long, side As Long: point = CLng(recent(1)): side = CLng(recent(2))
    Dim response As CCoupledReinforcedResponse, concrete As CCoupledConcreteResponse
    Select Case side
        Case -1: Set response = mReverseSidesTest(2): Set concrete = mReverseSidesTest(1)
        Case 0: Set response = mReverseReinforcedTest: Set concrete = mReverseConcreteTest
        Case 1: Set response = mReverseSidesTest(4): Set concrete = mReverseSidesTest(3)
    End Select
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\OutgoingTangent.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMOUTTAN0000001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, recent
    TestWriteReverseBinaryVector handle, concrete.ReadNewtonYieldEventStateTest(point)
    Dim caseIndex As Long, residual() As Double, matrix() As Double, strain(1 To 6) As Double, slope(1 To 6) As Double
    Dim value As Double, gradient(1 To 6) As Double, column() As Double, degree As Long
    For caseIndex = 0 To UBound(levels)
        field = accepted
        For i = 1 To UBound(field): field(i) = accepted(i) + CDbl(levels(caseIndex)) * (accepted(i) - previous(i)): Next i
        mReverseOperatorsTest.PointStrain point, field, strain: mReverseOperatorsTest.AxialGradientAtPoint point, field, slope
        For i = 1 To 6: strain(i) = strain(i) + side * 0.001 * slope(i): Next i
        If Not response.TryReadTrialYield(point, strain, value, gradient, True) Then Err.Raise 5, , "Нет читающей производной новой материальной координаты."
        If Not mReverseSolverTest.TestEquilibriumLinearization(field, True, True, residual, matrix) Then Err.Raise 5, , "Нет читающей линеаризации исходящей ветви."
        TestWriteReverseBinaryVector handle, Array(caseIndex + 1, levels(caseIndex), point, side, value, revision, mReverseSolverTest.AcceptedRevision)
        TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, residual
        TestWriteReverseBinaryVector handle, gradient
        ReDim column(1 To UBound(field))
        For degree = 1 To UBound(field)
            For i = 1 To UBound(field): column(i) = matrix(degree, i): Next i
            TestWriteReverseBinaryVector handle, column
        Next degree
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Чтение исходящей касательной приняло историю."
    Next caseIndex
    Close #handle
    Dim checked() As Double: mReverseSolverTest.CopyAccepted checked
    For i = 1 To UBound(accepted)
        If accepted(i) <> checked(i) Then Err.Raise 5, , "Чтение исходящей касательной изменило принятое поле."
    Next i
End Sub
''';values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
values.update({'Compile-TestExternalPoint.ps1':(here/'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='OneSidedOutgoingTangentMetric',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    exact848_prefix86=True,all_working_math_materials_and_acceptance_unchanged=True,read_only_linearizations=4,
    no_nonlinear_solution_or_history_acceptance=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'FOUR READONLY TANGENTS;NO ACCEPTANCE CHANGE')
