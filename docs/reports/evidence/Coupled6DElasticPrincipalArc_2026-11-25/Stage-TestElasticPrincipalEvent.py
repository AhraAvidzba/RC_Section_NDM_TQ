"""ДЛЯ ТЕСТОВ. Готовит общий упругий спектральный переход CDP на точной истории 86 состояний без изменения закона материала."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from zipfile import ZipFile
import json,re

here=Path(__file__).resolve().parent;repo=Path.cwd();base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold834UnresolvedArcScan_20261125';root=base/'Cold837ElasticPrincipalEvent_20261125'
assert json.loads((parent/'UnresolvedArcScanProof.json').read_bytes())['spectral_event_not_yet_supported_by_CDP_arc_fallback']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ElasticPrincipalCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
def TestReplace(code,old,new):
    """Меняет только единственное ожидаемое место исходника."""
    assert code.count(old)==1,old[:150]
    return code.replace(old,new,1)

name='src/Materials/CCoupledConcreteResponse.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' Читает главное пробное напряжение CDP только на упругой ветви от неизменной пластической истории.
' При отрицательной поверхности нагружения фактический тензор отличается положительным скаляром; нули совпадают.
' Градиент простого собственного значения записан по инженерным деформациям; кратные значения не дифференцируются.
Friend Function TryReadElasticPrincipalBoundary(ByVal point As Long, ByRef strain() As Double, ByVal principalIndex As Long, _
        ByRef value As Double, ByRef gradient() As Double, Optional ByVal needGradient As Boolean = True) As Boolean
    If principalIndex < 1 Or principalIndex > 3 Then Exit Function
    Dim yieldValue As Double, unused(1 To 6) As Double
    If Not TryReadTrialYield(point, strain, yieldValue, unused, False) Then Exit Function
    If yieldValue >= 0# Then Exit Function
    Dim c As TConcreteHistory: c = mHistory(point, mCommitted)
    Dim effective(1 To 6) As Double, i As Long, j As Long, k As Long, ex As Double, ey As Double, ez As Double
    ex = strain(1) - c.Plastic(1): ey = strain(2) - c.Plastic(2): ez = strain(3) - c.Plastic(3)
    For i = 1 To 3
        effective(i) = mElastic(i, 1) * ex + mElastic(i, 2) * ey + mElastic(i, 3) * ez
        effective(i + 3) = mElastic(i + 3, i + 3) * (strain(i + 3) - c.Plastic(i + 3))
    Next i
    Dim trace As Double, q As Double, high As Double, middle As Double, low As Double, other1 As Double, other2 As Double
    CoupledStressInvariants effective, trace, q, high, middle, low
    Select Case principalIndex
        Case 1: value = high: other1 = middle: other2 = low
        Case 2: value = middle: other1 = high: other2 = low
        Case 3: value = low: other1 = high: other2 = middle
    End Select
    If Not needGradient Then TryReadElasticPrincipalBoundary = True: Exit Function
    Dim stressScale As Double: stressScale = GeomMax(1#, GeomMax(Abs(high), Abs(low)))
    If Abs(value - other1) <= 0.000000000001 * stressScale Or Abs(value - other2) <= 0.000000000001 * stressScale Then Exit Function
    Dim tensor(1 To 3, 1 To 3) As Double, projector(1 To 3, 1 To 3) As Double, divisor As Double, x As Double
    For i = 1 To 3: tensor(i, i) = effective(i): Next i
    tensor(1, 2) = effective(4): tensor(2, 1) = effective(4)
    tensor(1, 3) = effective(5): tensor(3, 1) = effective(5)
    tensor(2, 3) = effective(6): tensor(3, 2) = effective(6)
    divisor = (value - other1) * (value - other2)
    For i = 1 To 3
        For j = 1 To 3
            x = -(other1 + other2) * tensor(i, j)
            For k = 1 To 3: x = x + tensor(i, k) * tensor(k, j): Next k
            If i = j Then x = x + other1 * other2
            projector(i, j) = x / divisor
        Next j
    Next i
    Dim derivative(1 To 6) As Double
    For i = 1 To 3: derivative(i) = projector(i, i): Next i
    derivative(4) = 2# * projector(1, 2): derivative(5) = 2# * projector(1, 3): derivative(6) = 2# * projector(2, 3)
    For j = 1 To 6
        gradient(j) = 0#
        For i = 1 To 6: gradient(j) = gradient(j) + derivative(i) * mElastic(i, j): Next i
    Next j
    TryReadElasticPrincipalBoundary = True
End Function
'''
values[name]=code
name='src/Materials/CCoupledReinforcedResponse.cls';code=(root/name).read_text('utf-8-sig')
code+='''
' Передаёт упругую знаковую границу CDP бетонной составляющей; усилия стали и вытеснение бетона здесь не меняются.
Friend Function TryReadElasticPrincipalBoundary(ByVal point As Long, ByRef strain() As Double, ByVal principalIndex As Long, _
        ByRef value As Double, ByRef gradient() As Double, Optional ByVal needGradient As Boolean = True) As Boolean
    If Not TypeOf mConcrete Is CCoupledConcreteResponse Then Exit Function
    Dim concrete As CCoupledConcreteResponse: Set concrete = mConcrete
    TryReadElasticPrincipalBoundary = concrete.TryReadElasticPrincipalBoundary(point, strain, principalIndex, value, gradient, needGradient)
End Function
''';values[name]=code
name='src/Solver/CCoupledMaterialConstraint.cls'
source=root/name
if not source.is_file():source=repo/name
code=source.read_text('utf-8-sig')
# Проверяем взятые исходники по фактической родительской книге, включая добавленные ранее проверки стали.
helper,body,reader=run_path(str(here/'Inspect-TestResolutionVBA.py'))['TestHelpers'](repo)
with ZipFile(parent/'docs/regression/Diagnostic.xlsm') as z:_,actual=reader(helper['TestContainer'](z.read('xl/vbaProject.bin')))
for component,path in [('CCoupledMaterialConstraint',source),('CCoupledReinforcedResponse',parent/'src/Materials/CCoupledReinforcedResponse.cls')]:
    expected,_=helper['TestMVTCompilerBody'](body,path.read_text('utf-8-sig'),component)
    assert expected==body(actual[component],component)[0],component
code=TestReplace(code,'Private mStation As Double','Private mElasticPrincipal As Boolean\nPrivate mStation As Double')
code=TestReplace(code,'ByVal principal As Long, ByRef seed() As Double, ByRef scales() As Double, ByVal target As Double)',
    'ByVal principal As Long, ByRef seed() As Double, ByRef scales() As Double, ByVal target As Double, _\n        Optional ByVal elasticPrincipal As Boolean = False)')
code=TestReplace(code,'    mPoint = point: mStation = station: mPrincipal = principal: mTarget = target',
    '    If elasticPrincipal And principal = 0 Then Err.Raise 5, "CCoupledMaterialConstraint", "Для спектральной границы нужен номер главного напряжения."\n    mElasticPrincipal = elasticPrincipal\n    mPoint = point: mStation = station: mPrincipal = principal: mTarget = target')
code=TestReplace(code,'    If mPrincipal = 0 Then\n        read =',
    '    If mElasticPrincipal Then\n        read = mResponse.TryReadElasticPrincipalBoundary(mPoint, strain, mPrincipal, value, materialGradient, withGradient)\n    ElseIf mPrincipal = 0 Then\n        read =')
values[name]=code
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,'    FromValue As Double\n    ToValue As Double\nEnd Type',
    '    FromValue As Double\n    ToValue As Double\n    ElasticPrincipal As Boolean \' Спектральная координата CDP допустима только при отрицательной поверхности нагружения.\nEnd Type')
code=TestReplace(code,'ByVal target As Double) As ICoupledScalarConstraint',
    'ByVal target As Double, Optional ByVal elasticPrincipal As Boolean = False) As ICoupledScalarConstraint')
code=TestReplace(code,'principal, mHybridSeed, mHybridScales, target\n',
    'principal, mHybridSeed, mHybridScales, target, elasticPrincipal\n')
code=TestReplace(code,'    If station.PrincipalIndex > 0 Then\n        ReadTransitionValue =',
    '    If station.ElasticPrincipal Then\n        ReadTransitionValue = response.TryReadElasticPrincipalBoundary(station.Point, strain, station.PrincipalIndex, value, gradient, needGradient)\n    ElseIf station.PrincipalIndex > 0 Then\n        ReadTransitionValue =')
old='Set scalar = CreateMaterialConstraint(station.Point, station.Side, 0, target)'
code=TestReplace(code,old,'Set scalar = CreateMaterialConstraint(station.Point, station.Side, station.PrincipalIndex, target, station.ElasticPrincipal)')
code=TestReplace(code,"' Решает силовые уравнения на поверхности CDP имеющимся адаптером и полной разностной касательной.",
    "' Решает силовые уравнения на выбранной материальной границе существующим адаптером и полной разностной касательной.")
code+='''
' На неизменной истории выбирает ближайший разрешимый ноль главного напряжения упругого CDP, не принимая найденное поле.
' Граница ищется по всем точкам, трём станциям и трём главным напряжениям; пластический возврат этим путём не заменяется.
Friend Function TryElasticPrincipalBoundary(ByRef baseLoad() As Double, ByRef direction() As Double, ByRef ends() As Double, _
        ByVal fromLambda As Double, ByVal toLambda As Double, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maxIterations As Long, ByRef result() As Double, ByRef resultLambda As Double, ByRef eventInfo As Variant) As Boolean
    If mCandidateReady Or mHybridActive Then Err.Raise 5, , "Перед спектральным событием завершите пробный контекст."
    ConfigureControlledLoad baseLoad, direction: ResetSolveState
    mControlActive = False: mArcLoadProjection = 0#: mConsistentControl = False
    ReDim mControlProjection(1 To mDegreeCount): mTargetControl = 0#
    Dim station As TYieldStation, candidate As TYieldStation, nearest As Double, fraction As Double, selected As Boolean
    Dim point As Long, side As Long, principal As Long: nearest = 1#
    Dim eventRadius As Double: eventRadius = SpatialFieldDistance(mAccepted, ends)
    Dim resolution As Double
    resolution = 16# * GeomMax(ControlledLoadTolerance(baseLoad, direction, fromLambda, absoluteTolerance, relativeTolerance), _
        ControlledLoadTolerance(baseLoad, direction, toLambda, absoluteTolerance, relativeTolerance))
    Dim strain(1 To 6) As Double, unused(1 To 6) As Double, left As Double, right As Double
    For point = 1 To mPointCount
        For side = -1 To 1
            For principal = 1 To 3
                candidate.Point = point: candidate.Side = side: candidate.PrincipalIndex = principal: candidate.ElasticPrincipal = True
                YieldStrain point, side, mAccepted, strain
                If ReadTransitionValue(candidate, strain, left, unused, False) Then
                    YieldStrain point, side, ends, strain
                    If ReadTransitionValue(candidate, strain, right, unused, False) Then
                        If Abs(left) > 16# * YIELD_ROOT_TOLERANCE And right <> 0# And (left < 0#) <> (right < 0#) Then
                            candidate.FromValue = left: candidate.ToValue = right
                            If LocateYieldCoordinate(candidate, mAccepted, ends, 0#, fraction) Then
                                If fraction * eventRadius * mReferenceModulus > resolution Then
                                    If Not selected Or fraction < nearest Then station = candidate: nearest = fraction: selected = True
                                End If
                            End If
                        End If
                    End If
                End If
            Next principal
        Next side
    Next point
    eventInfo = Array(Abs(CLng(selected)), station.Point, station.Side, station.FromValue, station.ToValue, nearest, mAcceptedRevision, station.PrincipalIndex, 1#)
    If Not selected Or eventRadius <= 0# Then Exit Function
    Dim seed() As Double: seed = mAccepted
    Dim i As Long, projection() As Double, load(1 To 6) As Double, tolerance As Double
    For i = 1 To mDegreeCount: seed(i) = mAccepted(i) + nearest * (ends(i) - mAccepted(i)): Next i
    SetControlledLoad fromLambda + nearest * (toLambda - fromLambda), load
    tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    TryElasticPrincipalBoundary = FollowYieldEventCoordinate(station, seed, mCurrentControlLambda, 0#, tolerance, maxIterations, _
        mAccepted, eventRadius, result, resultLambda, projection)
End Function

' Читает упругую спектральную координату в переданном поле без принятия истории; область применимости проверяет материал.
Friend Function TryReadElasticPrincipalSurface(ByVal point As Long, ByVal side As Long, ByVal principal As Long, _
        ByRef field() As Double, ByRef value As Double) As Boolean
    If point < 1 Or point > mPointCount Or Abs(side) > 1 Or principal < 1 Or principal > 3 Then Err.Raise 5, , "Неверная спектральная координата."
    If LBound(field) <> 1 Or UBound(field) <> mDegreeCount Then Err.Raise 5, , "Поле спектральной координаты не соответствует сечению."
    Dim station As TYieldStation, strain(1 To 6) As Double, unused(1 To 6) As Double
    station.Point = point: station.Side = side: station.PrincipalIndex = principal: station.ElasticPrincipal = True
    YieldStrain point, side, field, strain
    TryReadElasticPrincipalSurface = ReadTransitionValue(station, strain, value, unused, False)
End Function
''';values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,'    TestReadUnresolvedArcScan\n','    TestReadElasticPrincipalEvent\n')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Решает спектральную границу на шести сохранённых отрезках от одной истории; результаты записываются без принятия.
Private Sub TestReadElasticPrincipalEvent()
    Dim source() As Double: mReverseSolverTest.CopyAccepted source
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim count As Long: count = mReverseTraceTest.Count
    Dim reference As Double, fraction As Double: mReverseTraceTest.ReadPoint count, fraction, reference
    Dim inputHandle As Integer: inputHandle = FreeFile
    Open mPolishRootTest & "\\Input\\ArcScanFields.bin" For Binary Access Read As #inputHandle
    Dim header As Variant: header = TestReadPolishVector(inputHandle)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ElasticPrincipal.Event.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMELASTIC000001"
    Put #handle, , magic
    Dim caseIndex As Long, i As Long, raw As Variant, ends() As Double, result() As Double, resultLambda As Double, info As Variant, solved As Boolean
    Dim target As Variant: target = TestReadElasticPrincipalTargets()
    For caseIndex = 1 To CLng(header(1))
        raw = TestReadPolishVector(inputHandle): ReDim ends(1 To UBound(raw))
        For i = 1 To UBound(raw): ends(i) = raw(i): Next i
        solved = mReverseSolverTest.TryElasticPrincipalBoundary(mReverseBaseTest, mReverseDirectionTest, ends, reference, target(caseIndex), _
            mReverseOptionsTest(3) * 0.0001, mReverseOptionsTest(4) * 0.0001, 15, result, resultLambda, info)
        TestWriteReverseBinaryVector handle, Array(caseIndex, Abs(CLng(solved)), resultLambda)
        TestWriteReverseBinaryVector handle, info
        TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
        If solved Then
            TestWriteReverseBinaryVector handle, result
            TestWriteElasticPrincipalRoot caseIndex, result
        Else
            TestWriteReverseBinaryVector handle, Array(0#)
        End If
        mReverseSolverTest.DiscardStep
        If mReverseSolverTest.AcceptedRevision <> revision Then Err.Raise 5, , "Спектральное событие приняло новую историю."
        Dim after() As Double: mReverseSolverTest.CopyAccepted after
        For i = 1 To UBound(source)
            If after(i) <> source(i) Then Err.Raise 5, , "Спектральное событие изменило исходное принятое поле."
        Next i
    Next caseIndex
    Close #handle: Close #inputHandle
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает только начальные оценки lambda сохранённых отрезков, без изменения нагрузок и критериев решения.
Private Function TestReadElasticPrincipalTargets() As Variant
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Input\\ElasticPrincipalTargets.bin" For Binary Access Read As #handle
    TestReadElasticPrincipalTargets = TestReadPolishVector(handle)
    Close #handle
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает тензоры и производные нового корня от исходной истории, не принимая материал даже для диагностического вывода.
Private Sub TestWriteElasticPrincipalRoot(ByVal caseIndex As Long, ByRef field() As Double)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\ElasticPrincipal.Roots.csv" For Append As #handle
    Dim point As Long, side As Long, principal As Long, i As Long, failure As ESolverFailureCode
    Dim concrete As CCoupledConcreteResponse, steel As CCoupledReinforcedResponse, response As ICoupledPointResponse
    Dim strain(1 To 6) As Double, slope(1 To 6) As Double, stress(1 To 6) As Double, tangent(1 To 6, 1 To 6) As Double
    Dim gradient(1 To 6) As Double, value As Double, available As Boolean, bar As Double, axial As Double, barTangent As Double, before As Variant, after As Variant
    For side = -1 To 1
        Select Case side
            Case -1: Set concrete = mReverseSidesTest(1): Set steel = mReverseSidesTest(2)
            Case 0: Set concrete = mReverseConcreteTest: Set steel = mReverseReinforcedTest
            Case 1: Set concrete = mReverseSidesTest(3): Set steel = mReverseSidesTest(4)
        End Select
        Set response = concrete
        For point = 1 To mReverseOperatorsTest.PointCount
            before = concrete.ReadNewtonYieldEventStateTest(point)
            mReverseOperatorsTest.PointStrain point, field, strain: mReverseOperatorsTest.AxialGradientAtPoint point, field, slope
            For i = 1 To 6: strain(i) = strain(i) + side * 0.001 * slope(i): Next i
            failure = response.EvaluatePoint(point, strain, stress, tangent, ctmStressOnly, recordTrial:=False)
            If failure <> 0 Then Err.Raise 5, , "Нет читающего бетонного отклика найденной спектральной границы."
            bar = 0#
            If point > mReverseOperatorsTest.ConcretePointCount Then steel.ReadSteelBranch point, strain, axial, bar, barTangent
            Print #handle, "K;" & CStr(caseIndex) & ";" & CStr(point) & ";" & CStr(side) & TestHardPreparedVector(stress) & ";" & TestOnsetNumber(bar)
            For principal = 1 To 3
                available = concrete.TryReadElasticPrincipalBoundary(point, strain, principal, value, gradient, True)
                Print #handle, "P;" & CStr(caseIndex) & ";" & CStr(point) & ";" & CStr(side) & ";" & CStr(principal) & ";" & CStr(available) & _
                    ";" & TestOnsetNumber(value) & TestHardPreparedVector(gradient)
            Next principal
            after = concrete.ReadNewtonYieldEventStateTest(point)
            For i = LBound(before) To UBound(before)
                If before(i) <> after(i) Then Err.Raise 5, , "Читающий вывод изменил принятую историю бетона."
            Next i
        Next point
    Next side
    Close #handle
End Sub
''';values[name]=code
vector=run_path(str(here/'Stage-TestColdRawReplay.py'))['TestVector']
Reader=run_path(str(here/'Read-TestColdRawCapture.py'))['TestReader']
walk=base/'Cold833ContinuedOutgoingWalk_20261125'
packets={}
for attempt in (28,31):
    r=Reader((walk/'Native'/f'LoadingWalk.{attempt:03}.bin').read_bytes())
    assert r.raw[:16]==b'NDMOPERARC000004'
    packets[attempt]=r.VectorTest()
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](parent/'Input/Reverse1.bin')
reference,before=original['lambdas'][-1],original['lambdas'][-2]
values['Input/ElasticPrincipalTargets.bin']=vector([packets[28][4],packets[28][5],packets[31][4],packets[31][5],
    reference+2*(reference-before),reference+.25*(reference-before)])
imports=['CCoupledConcreteResponse','CCoupledReinforcedResponse','CCoupledMaterialConstraint','CCoupledNewtonSolver','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=TestReplace(compiler,"-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if isinstance(data,str):
        if n.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    p=root/n;p.parent.mkdir(parents=True,exist_ok=True);p.write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
    if n.endswith(('.cls','.bas')):
        component=Path(n).stem
        if component not in prep['modules']:prep['modules'].append(component)
prep['directed_scope']=dict(case='CDPElasticPrincipalEventFromExact86',actual_import_this_pass=sorted(imports),actual_component_count=256,
    original_material_evaluation_methods_unchanged=True,ASD_principal_contract_unchanged=True,
    negative_yield_required_at_endpoints_and_material_constraint=True,earliest_resolvable_station_and_principal=True,
    existing_material_projected_solver_reused=True,no_new_history_accepted=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
