"""ДЛЯ ТЕСТОВ. Локализует остановку исходящей материальной координаты от точной истории86, не меняя условия допуска."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,numpy as np

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold846AutomaticHGuideWalk_20261125';root=base/'Cold848NewLoadingCoordinateDiagnosis_20261125'
proof=json.loads((parent/'AutomaticLoadingWalkProof.json').read_bytes());assert proof['new_accepted_equilibria']==16 and proof['all_new_H_strictly_growing']
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
source=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
fields=list(source['fields']);factors=list(source['lambdas'])
for attempt in proof['attempts']:
    if not attempt['solved']:continue
    r=Reader((parent/'Native'/f"LoadingWalk.{attempt['attempt']:03}.bin").read_bytes());packet=[r.VectorTest() for _ in range(15)]
    fields.append(r.VectorTest());factors.append(packet[0][23])
assert len(fields)==86 and factors[-1]==1.099981906664804
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'NewLoadingCoordinateCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
values['Input/Reverse1.bin']=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](source,
    dict(fields=np.array(fields),fractions=np.ones(86),lambdas=np.array(factors)))
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mMaterialEventDetails As String';assert code.count(anchor)==1
code=code.replace(anchor,'Private mYieldCoordinateCaptureTest As Collection \' ДЛЯ ТЕСТОВ: выключенная по умолчанию запись уже вычисленных проверок корня.\n'+anchor,1)
start=code.index('Private Function FollowYieldEventCoordinate(');end=code.index('\nEnd Function',start)
method=code[start:end];original=method
anchor='    Dim scalar As ICoupledScalarConstraint: Set scalar =';assert method.count(anchor)==1
method=method.replace(anchor,'    Dim testStage As String: testStage = "Constraint"\n'+anchor,1)
for old,new in [
('    If Not solver.HasBestPoint Then GoTo Finished','    testStage = "BestPoint"\n    If Not solver.HasBestPoint Then GoTo Finished'),
('    If Not projected.TryMap(solution, mapped) Then GoTo Finished','    testStage = "Map"\n    If Not projected.TryMap(solution, mapped) Then GoTo Finished'),
('    If Not scalar.TryEvaluate(mapped, True, value, gradient) Then GoTo Finished','    testStage = "ScalarRead"\n    If Not scalar.TryEvaluate(mapped, True, value, gradient) Then GoTo Finished'),
('    If Abs(value) > YIELD_ROOT_TOLERANCE Or Abs(solution(projected.Pivot)) > tolerance Then GoTo Finished','    testStage = "ScalarAndPivot"\n    If Abs(value) > YIELD_ROOT_TOLERANCE Or Abs(solution(projected.Pivot)) > tolerance Then GoTo Finished'),
('    If resultLambda <= 0# Or SpatialFieldDistance(anchor, result) > radius Then GoTo Finished','    testStage = "RadiusAndLambda"\n    If resultLambda <= 0# Or SpatialFieldDistance(anchor, result) > radius Then GoTo Finished'),
('    If Not Evaluate(result, False, recordTrial:=False) Then GoTo Finished','    testStage = "MaterialRead"\n    If Not Evaluate(result, False, recordTrial:=False) Then GoTo Finished'),
('    If EquilibriumNorm(load, result) > tolerance Or Not GaugesVerified(tolerance) Then GoTo Finished','    testStage = "EquilibriumAndGauge"\n    If EquilibriumNorm(load, result) > tolerance Or Not GaugesVerified(tolerance) Then GoTo Finished'),
('    FollowYieldEventCoordinate = True','    FollowYieldEventCoordinate = True\n    testStage = "Success"')]:
    assert method.count(old)==1;method=method.replace(old,new,1)
capture='''    If Not mYieldCoordinateCaptureTest Is Nothing Then
        Dim testSolution As Variant, testMapped As Variant, testInfo As Variant
        testSolution = Array(0#): testMapped = Array(0#)
        If Not solver Is Nothing Then
            If solver.HasBestPoint Then testSolution = solution
        End If
        On Error Resume Next
        If UBound(mapped) > 0 Then testMapped = mapped
        On Error GoTo Failed
        Dim testStop As Long, testBest As Double, testSteps As Long, testPivot As Long
        If Not solver Is Nothing Then testStop = solver.StopCode: testBest = solver.BestResidual: testSteps = solver.Steps
        If Not projected Is Nothing Then testPivot = projected.Pivot
        testInfo = Array(station.Point, station.Side, station.PrincipalIndex, Abs(CLng(station.ElasticPrincipal)), initialLambda, target, tolerance, _
            radius, resultLambda, Abs(CLng(solved)), Abs(CLng(FollowYieldEventCoordinate)), testStop, testBest, testSteps, testPivot, value, mLastNorm)
        mYieldCoordinateCaptureTest.Add Array(testStage, testInfo, mHybridSeed, mHybridScales, testSolution, testMapped)
    End If
'''
assert method.count('Finished:\n    mHybridActive = False')==1
method=method.replace('Finished:\n    mHybridActive = False','Finished:\n'+capture+'    mHybridActive = False',1)
# Точное удаление читающей инструментации должно вернуть исходную функцию побайтно по тексту.
restored=method.replace('    Dim testStage As String: testStage = "Constraint"\n','',1).replace(capture,'',1)
import re
restored=re.sub(r'^    testStage = "[A-Za-z]+"\n','',restored,flags=re.M)
assert restored==original
code=code[:start]+method+code[end:]
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Включает запись только уже вычисленных значений материального решателя; история и допуски не меняются.
Friend Sub TestStartYieldCoordinateCapture()
    Set mYieldCoordinateCaptureTest = New Collection
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает снимки проверок текущей диагностической пробы без повторного вычисления материала.
Friend Property Get TestYieldCoordinateCapture() As Collection
    Set TestYieldCoordinateCapture = mYieldCoordinateCaptureTest
End Property
''';values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='    TestRunContinuedOutgoingGuard\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n';assert code.count(old)==1
code=code.replace(old,'    TestDiagnoseNewLoadingCoordinate\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Повторяет шесть независимых исходящих попыток от одной принятой истории86 и записывает причины отбрасывания корней.
Private Sub TestDiagnoseNewLoadingCoordinate()
    Dim ratios As Variant: ratios = Array(1#, 0.5, 0.25, 0.125, 0.0625, 0.03125)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\YieldCoordinate.Checks.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMYIELDCHECK001"
    Put #handle, , magic
    Dim logHandle As Integer: logHandle = FreeFile
    Open mPolishRootTest & "\\Native\\YieldCoordinate.Checks.csv" For Output As #logHandle
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim accepted() As Double: mReverseSolverTest.CopyAccepted accepted
    Dim caseIndex As Long, arc As CCoupledArcContinuation, solved As Boolean, factor As Double, projection() As Double
    Dim records As Collection, record As Variant, info As Variant, i As Long, raw As Variant, field() As Double
    For caseIndex = 0 To UBound(ratios)
        mReverseSolverTest.DiscardStep: mReverseSolverTest.TestStartYieldCoordinateCapture
        Set arc = New CCoupledArcContinuation
        arc.Initialize mReverseSolverTest, mReverseBaseTest, mReverseDirectionTest, mReverseOptionsTest(3), mReverseOptionsTest(4), CLng(mReverseOptionsTest(2))
        solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, ratios(caseIndex), 96, factor, projection)
        Set records = mReverseSolverTest.TestYieldCoordinateCapture
        TestWriteReverseBinaryVector handle, Array(caseIndex + 1, ratios(caseIndex), Abs(CLng(solved)), factor, arc.Calls, revision, mReverseSolverTest.AcceptedRevision, records.Count, arc.FieldError)
        Dim certificate As Variant: certificate = arc.Verification
        For Each raw In certificate: TestWriteReverseBinaryVector handle, raw: Next raw
        For i = 1 To records.Count
            record = records(i)
            Print #logHandle, CStr(caseIndex + 1) & ";" & CStr(i) & ";" & CStr(record(0))
            Dim j As Long
            For j = 1 To 5: TestWriteReverseBinaryVector handle, record(j): Next j
            raw = record(4)
            If UBound(raw) = 61 Then
                Dim seed As Variant, scales As Variant, mapped As Variant
                mapped = record(5): seed = record(2): scales = record(3)
                If UBound(mapped) = 61 Then
                    ReDim field(1 To 60)
                    For j = 1 To 60: field(j) = seed(j) + scales(j) * mapped(j): Next j
                    TestWriteElasticPrincipalRoot (caseIndex + 1) * 10 + i, field
                End If
            End If
        Next i
        If solved Then
            Dim result() As Double, actual() As Double
            mReverseSolverTest.CopyTrial result, actual
            TestWriteReverseBinaryVector handle, result: TestWriteReverseBinaryVector handle, actual: TestWriteReverseBinaryVector handle, projection
            TestWriteElasticPrincipalRoot (caseIndex + 1) * 10 + 9, result
        End If
        mReverseSolverTest.DiscardStep
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Диагностика исходящего корня приняла историю."
        Dim checked() As Double: mReverseSolverTest.CopyAccepted checked
        For i = 1 To UBound(accepted)
            If accepted(i) <> checked(i) Then Err.Raise 5, , "Диагностика исходящего корня изменила принятое поле."
        Next i
    Next caseIndex
    Close #handle: Close #logHandle
End Sub
''';values[name]=code
imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1)
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():
    if isinstance(data,str):
        if name.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='NewLoadingCoordinatePostcheckDiagnosis',actual_import_this_pass=imports,actual_component_count=256,
    exact846_accepted_prefix86=True,all_acceptance_math_unchanged=True,working_arc_controller_and_materials_unchanged=True,
    diagnostic_only_ratios=[1,.5,.25,.125,.0625,.03125],no_candidate_accepted=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'EXACT86;NO ACCEPTANCE CHANGES;POSTCHECK REASONS RECORDED')
