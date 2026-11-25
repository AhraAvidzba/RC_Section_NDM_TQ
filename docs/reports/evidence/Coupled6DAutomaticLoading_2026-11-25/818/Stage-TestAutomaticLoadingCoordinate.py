"""ДЛЯ ТЕСТОВ. Проверяет автоматический выбор недавно достигнутой поверхности и существующую пару исходящей координаты."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold815ActiveLoadingNeighborhood_20261125';root=base/'Cold818AutomaticLoadingCoordinate_20261125'
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'AutomaticLoadingCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mAcceptedRevision As Long';assert code.count(anchor)==1
code=code.replace(anchor,'''Private mRecentLoadingPoint As Long, mRecentLoadingSide As Long, mRecentLoadingCount As Long
Private mRecentLoadingRevision As Long, mRecentLoadingFrom As Double, mRecentLoadingTo As Double
Private mAcceptedRevision As Long''',1)
anchor='    ReDim mAccepted(1 To mDegreeCount)\n    mAcceptedRevision = mAcceptedRevision + 1';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'\n    mRecentLoadingPoint = 0: mRecentLoadingCount = 0: mRecentLoadingRevision = -1',1)
anchor='    Dim acceptedField() As Double: acceptedField = mCandidate\n    PrepareAcceptedEnvelopeSegment';assert code.count(anchor)==1
code=code.replace(anchor,'    Dim acceptedField() As Double: acceptedField = mCandidate\n    PrepareRecentLoadingBoundary acceptedField\n    PrepareAcceptedEnvelopeSegment',1)
anchor='        PrepareAcceptedEnvelopeSegment\n        AcceptMaterialTrials\n        mAccepted = field';assert code.count(anchor)==1
code=code.replace(anchor,'        PrepareRecentLoadingBoundary field\n'+anchor,1)
header="' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Использует существующую материальную координату"
assert code.count(header)==1
code=code.split(header)[0].rstrip()+'\n'
code+='''
' Сохраняет единственную новую границу на неизменной истории до её принятия; одновременные границы не выбираются произвольно.
Private Sub PrepareRecentLoadingBoundary(ByRef field() As Double)
    mRecentLoadingPoint = 0: mRecentLoadingCount = 0: mRecentLoadingRevision = -1
    If Not mUseSpatialHistory Then Exit Sub
    If Not TypeOf mResponse Is CCoupledReinforcedResponse Then Exit Sub
    If Not TypeOf mPlusResponse Is CCoupledReinforcedResponse Then Exit Sub
    If Not TypeOf mMinusResponse Is CCoupledReinforcedResponse Then Exit Sub
    Dim point As Long, side As Long, left As Double, right As Double
    For point = 1 To mPointCount
        For side = -1 To 1
            If TryReadLoadingSurface(point, side, mAccepted, left) Then
                If left < -16# * YIELD_ROOT_TOLERANCE Then
                    If TryReadLoadingSurface(point, side, field, right) Then
                        If Abs(right) <= 16# * YIELD_ROOT_TOLERANCE Then
                            mRecentLoadingCount = mRecentLoadingCount + 1
                            mRecentLoadingPoint = point: mRecentLoadingSide = side
                            mRecentLoadingFrom = left: mRecentLoadingTo = right
                        End If
                    End If
                End If
            End If
        Next side
    Next point
    If mRecentLoadingCount = 1 Then mRecentLoadingRevision = mAcceptedRevision + 1 Else mRecentLoadingPoint = 0
End Sub

' Читает единственную только что принятую границу и положительное продвижение прогноза по её пробной поверхности.
' Смена ревизии, разгрузка и неразличимый прогноз не создают исходящую координату.
Friend Function TryReadRecentLoadingDeparture(ByRef prediction() As Double, ByRef info As Variant) As Boolean
    info = Array(0#, mRecentLoadingPoint, mRecentLoadingSide, mRecentLoadingFrom, mRecentLoadingTo, mRecentLoadingCount, mRecentLoadingRevision)
    If mRecentLoadingPoint = 0 Or mRecentLoadingRevision <> mAcceptedRevision Then Exit Function
    Dim current As Double, target As Double
    If Not TryReadLoadingSurface(mRecentLoadingPoint, mRecentLoadingSide, mAccepted, current) Then Exit Function
    If Abs(current) > 16# * YIELD_ROOT_TOLERANCE Then Exit Function
    If Not TryReadLoadingSurface(mRecentLoadingPoint, mRecentLoadingSide, prediction, target) Then Exit Function
    If target - current <= 32# * YIELD_ROOT_TOLERANCE Then Exit Function
    info = Array(1#, mRecentLoadingPoint, mRecentLoadingSide, current, target, mRecentLoadingCount, mRecentLoadingRevision)
    TryReadRecentLoadingDeparture = True
End Function

' Решает положительную координату недавно достигнутой поверхности от одной принятой истории в заранее заданной окрестности.
' Кандидат остаётся непринятым; вызывающий отдельно проверяет пару, разрешение, направление и точное восстановление.
Friend Function TryRecentLoadingCoordinate(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef seed() As Double, ByVal initialLambda As Double, ByVal target As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        ByVal referenceRadius As Double, ByRef result() As Double, ByRef resultLambda As Double) As Boolean
    If mCandidateReady Or mHybridActive Then Err.Raise 5, , "Перед исходящей координатой завершите предыдущую пробу."
    If mRecentLoadingPoint = 0 Or mRecentLoadingRevision <> mAcceptedRevision Then Exit Function
    If target <= 0# Or referenceRadius <= 0# Or maxIterations < 1 Then Exit Function
    ConfigureControlledLoad baseLoad, direction: ResetSolveState
    mControlActive = False: mArcLoadProjection = 0#: mConsistentControl = False
    ReDim mControlProjection(1 To mDegreeCount): mTargetControl = 0#
    Dim load(1 To COMPONENT_COUNT) As Double: SetControlledLoad initialLambda, load
    Dim station As TYieldStation: station.Point = mRecentLoadingPoint: station.Side = mRecentLoadingSide
    Dim tolerance As Double: tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    Dim projection() As Double
    TryRecentLoadingCoordinate = FollowYieldEventCoordinate(station, seed, initialLambda, target, tolerance, _
        maxIterations, mAccepted, referenceRadius, result, resultLambda, projection)
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    If Not mFrozenReplayActiveTest Then Exit Sub\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'''    Dim recentInfo As Variant, recentReady As Boolean, recentHandle As Integer
    recentReady = mReverseSolverTest.TryReadRecentLoadingDeparture(field, recentInfo)
    recentHandle = FreeFile
    Open mPolishRootTest & "\\Native\\RecentLoading.History.csv" For Append As #recentHandle
    Print #recentHandle, CStr(index) & TestHardPreparedVector(recentInfo)
    Close #recentHandle
''',1)
anchor='    point = 25: side = 1';assert code.count(anchor)==1
code=code.replace(anchor,'''    Dim recent As Variant
    If Not mReverseSolverTest.TryReadRecentLoadingDeparture(fullSeed, recent) Then
        Err.Raise 5, , "Нет единственной недавно принятой границы: " & TestHardPreparedVector(recent)
    End If
    point = CLng(recent(1)): side = CLng(recent(2))
    Dim markerHandle As Integer: markerHandle = FreeFile
    Open mPolishRootTest & "\\Native\\RecentLoading.Selection.csv" For Output As #markerHandle
    Print #markerHandle, "RECENT" & TestHardPreparedVector(recent)
    Close #markerHandle''',1)
anchor='mReverseSolverTest.TestSolveActiveLoadingCoordinate(mReverseBaseTest, mReverseDirectionTest, point, side, seed,'
assert code.count(anchor)==2
code=code.replace(anchor,'mReverseSolverTest.TryRecentLoadingCoordinate(mReverseBaseTest, mReverseDirectionTest, seed,')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledNewtonSolver','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='AutomaticRecentLoadingCoordinate',actual_import_this_pass=imports,actual_component_count=256,
    boundary_detected_before_material_commit_in_AcceptStep_and_Replay=True,no_point_or_side_hardcoded_in_driver=True,
    unique_new_boundary_required=True,revision_bound_context=True,all43_source_fields_exact815=True,
    full_and_half_physical_force_gauge_tolerances_unchanged=True,primary_pair30_unchanged=True,
    reference_neighborhood_fixed_before_solving=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
