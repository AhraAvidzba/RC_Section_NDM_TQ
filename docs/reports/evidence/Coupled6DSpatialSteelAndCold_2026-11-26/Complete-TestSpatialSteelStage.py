"""ДЛЯ ТЕСТОВ. Добавляет направленный нативный тест трёх станций в отдельную подготовленную копию."""
from pathlib import Path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent
root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold881SpatialSteelLimits_20261126')
prep=json.loads((root/'Preparation.json').read_bytes());values={}
entry='''Public Function RunTestSpatialSteelLimits() As String
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Проверяет боковой максимум, разные знаковые пределы и классы, независимость принятой истории и станционную проекцию.
    Dim passed As Long, failed As Long, report As String
    On Error GoTo Failed
    Dim section As CSectionModel: Set section = TestBuildSection(TestFixtureBox(180#, 120#))
    section.AddRebarElement 0#, 0#, 12#, 100#, "TestLong"
    Dim builder As New CSectionDefinitionBuilder, layout As New CTransverseRebarLayout
    builder.Initialize "TestSpatialSteelLimits"
    layout.AddGroup builder.LinePath("TestBranch", -40#, 25#, 50#, 25#), 8#, 80#, "TestHoop"
    Dim models As Object: Set models = CreateObject("Scripting.Dictionary")
    Dim steel As New CCoupledSteelModel, hoop As New CCoupledSteelModel
    steel.Initialize 200000#, 50#, 40#, 1000#, 1000#, 0.0006, 0.00025
    hoop.Initialize 200000#, 240#, 240#, 0#, 0#, 0.003, 0.002
    models.Add "TestLong", steel: models.Add "TestHoop", hoop
    Dim operators As New CCoupledElasticState
    operators.PrepareModel section, 30000#, 0#, 1, 1, 5, steelModels:=models, transverse:=layout
    Dim concrete(-1 To 1) As CTestCoupledPointResponse, material(-1 To 1) As CCoupledReinforcedResponse
    Dim responses(-1 To 1) As ICoupledPointResponse, solver As New CCoupledNewtonSolver
    Dim side As Long, point As Long, i As Long, scenario As Long, critical As Long, station As Long
    For side = -1 To 1
        Set concrete(side) = New CTestCoupledPointResponse: concrete(side).TestInitialize operators.PointCount
        Set material(side) = New CCoupledReinforcedResponse: material(side).Initialize operators, concrete(side), models
        Set responses(side) = material(side)
    Next side
    solver.Initialize operators, responses(0), 30000#, responses(1), responses(-1)
    Dim tensor(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim axial As Double, ratio As Double, expected As Double
    For scenario = 1 To 3
        For side = -1 To 1
            For point = 1 To operators.PointCount
                For i = 1 To 6: tensor(i) = 0#: Next i
                If point > operators.ConcretePointCount Then
                    If operators.TransverseGroupAt(point) = 0 Then
                        axial = 0.0002
                        If side = 1 Then axial = 0.00066
                        If scenario > 1 And side = -1 Then axial = -0.000275
                        If scenario > 1 And side = 1 Then axial = 0.00015
                        tensor(3) = axial
                    Else
                        tensor(1) = 0.0009
                        If scenario = 3 Then tensor(1) = 0.0033
                    End If
                End If
                If responses(side).EvaluatePoint(point, tensor, stress, tangent) <> sfcNone Then Err.Raise 5, , "Не получен тестовый отклик станции."
            Next point
        Next side
        ratio = material(0).ReadSteelLimit(False, csuLongitudinal, critical, station)
        expected = 1#: If scenario > 1 Then expected = -1#
        TestExpectNear ratio, 1.1, 0.000000000001, "longitudinal signed limit scenario " & scenario, passed, failed, report
        TestRecord station = expected And critical = operators.ConcretePointCount + 1, "governing longitudinal station " & scenario, passed, failed, report
        ratio = material(0).ReadSteelLimit(False, csuTransverse, critical, station)
        expected = 0.3: If scenario = 3 Then expected = 1.1
        TestExpectNear ratio, expected, 0.000000000001, "independent hoop class scenario " & scenario, passed, failed, report
        TestRecord station = 0 And operators.TransverseGroupAt(critical) > 0, "hoop central tie " & scenario, passed, failed, report
        If scenario = 1 Then
            For side = -1 To 1: responses(side).AcceptTrial: Next side
        Else
            TestExpectNear material(0).AcceptedSteelUtilization, 1.1, 0.000000000001, "trial leaves accepted utilization unchanged", passed, failed, report
            TestExpectNear material(0).AcceptedStationSteelStress(operators.ConcretePointCount + 1, 1), 50.41, 0.000000001, "trial leaves accepted side stress unchanged", passed, failed, report
        End If
    Next scenario
    Dim field() As Double, projection() As Double, gradient(1 To 6) As Double, projected As Double
    ReDim field(1 To operators.DegreeCount)
    For i = 1 To operators.DegreeCount: field(i) = i * 0.000001: Next i
    For point = operators.ConcretePointCount + 1 To operators.PointCount
        For side = -1 To 1
            operators.CopySteelStrainProjection point, projection, side * 0.001
            projected = 0#: For i = 1 To UBound(projection): projected = projected + projection(i) * field(i): Next i
            operators.PointStrain point, field, tensor
            For i = 1 To 6: gradient(i) = 0#: Next i
            operators.AxialGradientAtPoint point, field, gradient
            Dim axis(1 To 6) As Double: operators.SteelProjection point, axis
            expected = 0#: For i = 1 To 6: expected = expected + (tensor(i) + side * 0.001 * gradient(i)) * axis(i): Next i
            TestExpectNear projected, expected, 0.00000000000001, "station projection " & point & "/" & side, passed, failed, report
        Next side
    Next point
    RunTestSpatialSteelLimits = report & "TOTAL_SPATIAL_STEEL_LIMITS_TEST: passed=" & passed & "; failed=" & failed
    Exit Function
Failed:
    RunTestSpatialSteelLimits = report & "RUNTIME ERROR: " & Err.Source & " " & Err.Description
End Function
'''
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='Private Const HARD_PREPARED_HALF_TEST As Double = 0.001\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+"' FOR TESTS: baseline entry point only; existing methods retained.\n\n\n"+entry,1)
values[name]=code
name='tools/build_workbook/Update-IsolatedOptimizerCode.ps1';code=(root/name).read_text('utf-8-sig')
begin=code.index("        if ($file.BaseName -eq 'modTestCoupledElasticState') {")
end=code.index("        if ($EntryPointNames.Count -and $file.BaseName -eq 'modTestCoupledElasticState') {",begin)
code=code[:begin]+code[end:];values[name]=code
modules=prep['directed_scope']['actual_import_this_pass']+['modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
compiler=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join("'"+m+"'" for m in modules)+") -EntryPointNames 'RunTestSpatialSteelLimits'")
values['Compile-TestExternalPoint.ps1']=compiler
runner='''# ДЛЯ ТЕСТОВ: одна нативная проверка станционных пределов стали в частной книге.
param([string]$Root=$PSScriptRoot)
$ErrorActionPreference='Stop'
Push-Location -LiteralPath $Root
try {
    . './Guard-IsolatedOffice.ps1'
    Assert-IsolatedOfficeReady 'SpatialSteelLimits'
    & './tools/build_workbook/Run-FocusedKaniWatchdog.ps1' -SourceWorkbook 'docs/regression/Diagnostic.xlsm' -ReportPath 'Native/SpatialSteel.txt' `
        -Macro 'modTestCoupledElasticState.RunTestSpatialSteelLimits' -TimeoutSeconds 180 -Mode No -FailFast
    if($LASTEXITCODE -ne 0){throw ('Spatial steel test exited '+$LASTEXITCODE)}
} finally {Pop-Location}
'''
values['Run-TestSpatialSteelLimits.ps1']=runner
for name,code in values.items():
    data=code.replace('\n','\r\n').encode('utf-8-sig');(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']['actual_import_this_pass']=modules
prep['inputs'][Path(__file__).name]=sha256(Path(__file__).read_bytes()).hexdigest().upper();(root/Path(__file__).name).write_bytes(Path(__file__).read_bytes())
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('SIX EXPLICIT IMPORTS;ONE SMALL TEST ENTRY;THREE STATIONS AND DIFFERENT STEEL CLASSES')
