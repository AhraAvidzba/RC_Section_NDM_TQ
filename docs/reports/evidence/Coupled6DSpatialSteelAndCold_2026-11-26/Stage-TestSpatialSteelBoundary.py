"""ДЛЯ ТЕСТОВ. Проверяет локализацию стального предела, когда его достигает боковая станция при Qx обоих знаков."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold881SpatialSteelLimits_20261126';root=base/'Cold884SpatialSteelBoundary_20261126'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'SpatialBoundaryCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
entry='''Public Function RunTestSpatialSteelBoundary() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Использует настоящее равновесие трёх станций и силовую локализацию при Qx обоих знаков; линейный бетон является аналитическим контролем численного алгоритма.
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 35#, 10#, 12#, 100#, "TestSideLimit"
    Dim steel As New CCoupledSteelModel, models As Object
    steel.Initialize 200000#, 1#, 1#, 199000#, 199000#, 0.0001, 0.00013
    Set models = CreateObject("Scripting.Dictionary"): models.Add "TestSideLimit", steel
    Dim operators As New CCoupledElasticState
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModels:=models
    Dim concrete(-1 To 1) As CTestCoupledPointResponse, material(-1 To 1) As CCoupledReinforcedResponse
    Dim response(-1 To 1) As ICoupledPointResponse, side As Long, i As Long, sign As Variant
    For side = -1 To 1
        Set concrete(side) = New CTestCoupledPointResponse: concrete(side).TestInitialize operators.PointCount
        Set material(side) = New CCoupledReinforcedResponse: material(side).Initialize operators, concrete(side), models
        Set response(side) = material(side)
    Next side
    Dim solver As New CCoupledNewtonSolver: solver.Initialize operators, response(0), 30000#, response(1), response(-1)
    Dim field() As Double, actual() As Double, boundary() As Double, correction() As Double
    Dim ratio As Double, central As Double, point As Long, station As Long, localPoint As Long, projected As Boolean
    Dim uncertainty As ECoupledSteelUncertainty, path As CLoadPathVector, branch As CCoupledLoadContinuation
    Dim low As Double, high As Double, onOffset As Boolean, history As Collection, meta As CResultMeta
    For Each sign In Array(-1#, 1#)
        solver.ResetHistory
        If Not solver.TryStep(0#, 0#, 0#, 0#, CDbl(sign) * 30000000#, 0#, 0.00000001, 0.00000001, 60) Then _
            Err.Raise 5, , "Не найден контрольный силовой корень: " & solver.ResultMeta.DiagnosticDetails
        solver.CopyTrial field, actual
        ratio = material(0).ReadSteelLimit(False, csuNone, point, station)
        central = material(0).LocalSteelLimit(False, csuNone, localPoint)
        report = report & "SPATIAL_STEEL_SEED: sign=" & CStr(sign) & "; ratio=" & CStr(ratio) & "; center=" & CStr(central) & "; side=" & CStr(station) & vbCrLf
        TestRecord station <> 0 And ratio > central, "side governs actual force equilibrium " & sign, passed, failed, report
        boundary = field
        projected = material(0).ProjectSteelBoundary(operators, point, boundary, station)
        Dim projection() As Double, strain As Double
        operators.CopySteelStrainProjection point, projection, station * 0.001
        strain = 0#: For i = 1 To UBound(projection): strain = strain + projection(i) * boundary(i): Next i
        TestExpectNear steel.StrainUtilization(strain), 1#, 0.00000000001, "side boundary projection " & sign, passed, failed, report
        ReDim correction(1 To operators.DegreeCount)
        For i = 1 To operators.DegreeCount: correction(i) = 0.000000000001: Next i
        projected = material(0).SteelBoundaryWithinCorrection(operators, boundary, correction, point, uncertainty, side)
        TestRecord projected And side = station And uncertainty = csuLongitudinal, "side boundary uncertainty " & sign, passed, failed, report
        solver.DiscardStep
        Set path = New CLoadPathVector
        path.Initialize 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, CDbl(sign) * 30000000#, 0#, 0#
        Set branch = New CCoupledLoadContinuation
        branch.Initialize solver, path, 4, 64, 60, 0.00000001, 0.00000001, True
        branch.TraceSteelLimit material(0), 1#, 4#, 0.0001
        Set meta = branch.ResultMeta
        Dim found As Boolean: found = branch.CopySteelLimit(low, high, onOffset, point)
        report = report & "SPATIAL_STEEL_LIMIT: sign=" & CStr(sign) & "; found=" & CStr(found) & "; low=" & CStr(low) & "; high=" & CStr(high) & "; " & meta.DiagnosticDetails & vbCrLf
        TestRecord found And Not onOffset And high > low And high - low <= 0.0001, "side limit localized with fixed tolerance " & sign, passed, failed, report
        If found Then
            TestRecord meta.InternalStatus = rsCheckFailed, "side limit is physical failure " & sign, passed, failed, report
            If Not branch.CopyLimitState(history, actual) Then Err.Raise 5, , "Не сохранено допустимое плечо бокового предела."
            If Not solver.ReplayAcceptedFields(history, history.Count) Then Err.Raise 5, , "Не повторяется допустимое плечо бокового предела."
            ratio = material(0).ReadSteelLimit(True, csuNone, point, station)
            central = material(0).LocalSteelLimit(True, csuNone, localPoint)
            TestRecord ratio <= 1# And ratio > 0.99 And station <> 0 And central < ratio, "admissible bracket checks all stations " & sign, passed, failed, report
        End If
    Next sign
    RunTestSpatialSteelBoundary = report & "TOTAL_SPATIAL_STEEL_BOUNDARY_TEST: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestSpatialSteelBoundary = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1);values[name]=code
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'modTestCoupledElasticState' -EntryPointNames 'RunTestSpatialSteelBoundary'",1)
runner=(parent/'Run-TestSpatialSteelLimits.ps1').read_text('utf-8-sig').replace('RunTestSpatialSteelLimits','RunTestSpatialSteelBoundary').replace('SpatialSteel.txt','SpatialBoundary.txt')
values['Run-TestSpatialSteelBoundary.ps1']=runner
for name,code in values.items():
 data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
for name,data in {'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()}.items():
 (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='SpatialSteelBoundary',actual_import_this_pass=['modTestCoupledElasticState'],actual_component_count=256,
    only_test_entry_added=True,all_working_sources_exact881=True,actual_three_station_force_equilibrium=True,opposite_Qx_signs=True,
    no_physical_benchmark_capacity_claim=True,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'ONE TEST ENTRY;REAL THREE-STATION FORCE ROOT AND SIDE STEEL BOUNDARY LOCALIZATION')
