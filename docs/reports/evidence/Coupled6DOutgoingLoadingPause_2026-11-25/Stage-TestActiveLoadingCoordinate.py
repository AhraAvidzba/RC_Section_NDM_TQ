"""ДЛЯ ТЕСТОВ. Ищет исходящее равновесие точного холодного конца через существующий решатель координаты CDP; история не принимается."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold813ActualRemainingEndpoint_20261125';root=base/'Cold814ActiveLoadingCoordinate_20261125'
proof=json.loads((parent/'ColdEndpointArcProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['source_states']==43 and not proof['returned_verified_candidate']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ActiveCoordinateCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Использует существующую материальную координату с положительной целью на одной принятой истории; равновесие проверяется рабочим ядром.
Friend Function TestSolveActiveLoadingCoordinate(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByVal point As Long, ByVal side As Long, ByRef seed() As Double, ByVal initialLambda As Double, _
        ByVal target As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maxIterations As Long, ByRef result() As Double, ByRef resultLambda As Double) As Boolean
    If mCandidateReady Or mHybridActive Then Err.Raise 5, , "Не завершена предыдущая проба тестовой координаты."
    ConfigureControlledLoad baseLoad, direction: ResetSolveState
    mControlActive = False: mArcLoadProjection = 0#: mConsistentControl = False
    ReDim mControlProjection(1 To mDegreeCount): mTargetControl = 0#
    SetControlledLoad initialLambda, baseLoad
    Dim station As TYieldStation: station.Point = point: station.Side = side
    Dim tolerance As Double: tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    Dim projection() As Double, radius As Double: radius = 3# * SpatialFieldDistance(mAccepted, seed)
    If radius <= 0# Or target <= 0# Or maxIterations < 1 Then Exit Function
    TestSolveActiveLoadingCoordinate = FollowYieldEventCoordinate(station, seed, initialLambda, target, tolerance, _
        maxIterations, mAccepted, radius, result, resultLambda, projection)
End Function
'''
# SetControlledLoad получает отдельный локальный вектор: входная постоянная часть траектории неизменна.
code=code.replace('    SetControlledLoad initialLambda, baseLoad\n','    Dim load(1 To COMPONENT_COUNT) As Double: SetControlledLoad initialLambda, load\n',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRunOperationalArc\n';assert code.count(anchor)==1
code=code.replace(anchor,'    TestRunActiveLoadingCoordinate\n',1)
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Проверяет пару исходящих материальных координат без требования роста нагрузки и без принятия нового поля.
Private Sub TestRunActiveLoadingCoordinate()
    Dim accepted() As Double, previous() As Double, halfSeed() As Double, fullSeed() As Double
    accepted = mReverseFieldsTest(mReverseFieldsTest.Count): previous = mReverseFieldsTest(mReverseFieldsTest.Count - 1)
    halfSeed = accepted: fullSeed = accepted
    Dim i As Long, fraction As Double, reference As Double, before As Double
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count, fraction, reference
    mReverseTraceTest.ReadPoint mReverseTraceTest.Count - 1, fraction, before
    For i = 1 To mReverseSolverTest.PhysicalDegreeCount
        halfSeed(i) = accepted(i) + 0.5 * (accepted(i) - previous(i))
        fullSeed(i) = accepted(i) + accepted(i) - previous(i)
    Next i
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ActiveLoading.Yields.csv" For Output As #handle
    Dim point As Long, side As Long, left As Double, middle As Double, right As Double, available As Boolean
    For point = 1 To mReverseOperatorsTest.PointCount
        For side = -1 To 1
            available = mReverseSolverTest.TryReadLoadingSurface(point, side, accepted, left)
            available = mReverseSolverTest.TryReadLoadingSurface(point, side, halfSeed, middle) And available
            available = mReverseSolverTest.TryReadLoadingSurface(point, side, fullSeed, right) And available
            Print #handle, "Y;" & CStr(point) & ";" & CStr(side) & ";" & CStr(available) & ";" & _
                TestOnsetNumber(left) & ";" & TestOnsetNumber(middle) & ";" & TestOnsetNumber(right)
        Next side
    Next point
    Close #handle
    point = 25: side = 1
    If Not mReverseSolverTest.TryReadLoadingSurface(point, side, accepted, left) Then Err.Raise 5, , "Нет исходной поверхности координаты."
    If Not mReverseSolverTest.TryReadLoadingSurface(point, side, fullSeed, right) Then Err.Raise 5, , "Нет прогнозной поверхности координаты."
    Dim halfTarget As Double, fullTarget As Double
    halfTarget = left + 0.25 * (right - left): fullTarget = left + 0.5 * (right - left)
    Dim sourceState As Variant, concrete As CCoupledConcreteResponse
    Set concrete = mReverseSidesTest(3): sourceState = concrete.ReadNewtonYieldEventStateTest(point)
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    mReverseSolverTest.DiscardStep
    Dim half() As Double, full() As Double, halfLambda As Double, fullLambda As Double, sh As Boolean, sf As Boolean
    Dim halfEffort As Variant, fullEffort As Variant, seed() As Double
    halfEffort = Array(0#, 0#, 0#, 0#, 0#, 0#, revision, 0#): fullEffort = halfEffort
    ReDim half(1 To 1): ReDim full(1 To 1)
    If halfTarget > 0# And fullTarget > halfTarget Then
        seed = accepted
        For i = 1 To mReverseSolverTest.PhysicalDegreeCount: seed(i) = accepted(i) + 0.25 * (fullSeed(i) - accepted(i)): Next i
        sh = mReverseSolverTest.TestSolveActiveLoadingCoordinate(mReverseBaseTest, mReverseDirectionTest, point, side, seed, _
            reference + 0.25 * (reference - before), halfTarget, mReverseOptionsTest(3) * 0.0001, _
            mReverseOptionsTest(4) * 0.0001, 15, half, halfLambda)
        halfEffort = mReverseSolverTest.ReadPolishEffortTest: mReverseSolverTest.DiscardStep
        seed = halfSeed
        If sh Then
            For i = 1 To mReverseSolverTest.PhysicalDegreeCount: seed(i) = 2# * half(i) - accepted(i): Next i
        End If
        sf = mReverseSolverTest.TestSolveActiveLoadingCoordinate(mReverseBaseTest, mReverseDirectionTest, point, side, seed, _
            reference + 0.5 * (reference - before), fullTarget, mReverseOptionsTest(3) * 0.0001, _
            mReverseOptionsTest(4) * 0.0001, 30 - CLng(halfEffort(0)), full, fullLambda)
        fullEffort = mReverseSolverTest.ReadPolishEffortTest: mReverseSolverTest.DiscardStep
    End If
    If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Проба координаты приняла историю."
    handle = FreeFile
    Open mPolishRootTest & "\\Native\\ActiveLoading.Result.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMACTLOAD000001"
    Put #handle, , magic
    TestWriteReverseBinaryVector handle, Array(point, side, reference, before, left, right, halfTarget, fullTarget, _
        Abs(CLng(sh)), Abs(CLng(sf)), halfLambda, fullLambda, revision, mReverseSolverTest.AcceptedRevision)
    TestWriteReverseBinaryVector handle, accepted: TestWriteReverseBinaryVector handle, halfSeed: TestWriteReverseBinaryVector handle, fullSeed
    TestWriteReverseBinaryVector handle, sourceState: TestWriteReverseBinaryVector handle, half: TestWriteReverseBinaryVector handle, full
    TestWriteReverseBinaryVector handle, halfEffort: TestWriteReverseBinaryVector handle, fullEffort
    Close #handle
    If sh Then TestExportSignedArcTrial "ActiveLoading.Half", half, halfLambda
    If sf Then TestExportSignedArcTrial "ActiveLoading.Full", full, fullLambda
End Sub
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",
 '-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
 'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
 if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
 (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ExactCold808OutgoingLoadingCoordinatePair',actual_import_this_pass=imports,actual_component_count=256,
 all43source_fields_and_materials_exact813=True,working_methods_unchanged=True,
 existing_FollowYieldEventCoordinate_reused=True,pair_shared_primary30=True,
 no_requirement_that_lambda_increases=True,all_physical_parameters_and_force_gauge_tolerances_unchanged=True,
 all_point_trial_yield_values_exported=True,no_material_history_acceptance=True,
 no_capacity_search_or_cold_workflow=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('814 EXISTING CDP COORDINATE AT25/+1; SHARED30 FULL/HALF; NO ACCEPTED STEP OR PHYSICAL LIMIT')
