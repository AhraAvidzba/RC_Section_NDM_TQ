"""ДЛЯ ТЕСТОВ. Встраивает проверенную исходящую материальную координату в общий дуговой контроллер частного кандидата."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold818AutomaticLoadingCoordinate_20261125';root=base/'Cold819AutomaticLoadingArc_20261125'
assert json.loads((parent/'AutomaticLoadingCoordinateProof.json').read_bytes())['pair']['verified']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'AutomaticArcCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
anchor='    Dim sh As Boolean, sf As Boolean, lh As Double';assert code.count(anchor)==1
code=code.replace(anchor,'    Dim loadingPrediction() As Double: loadingPrediction = fullSeed\n'+anchor,1)
anchor='    ElseIf mCalls + 3 <= callAllowance Then\n        Dim bracket() As Double, bracketLambda As Double';assert code.count(anchor)==1
code=code.replace(anchor,'''    ElseIf mCalls + 3 <= callAllowance Then
        Dim recent As Variant
        If mSolver.TryReadRecentLoadingDeparture(loadingPrediction, recent) Then
            eventInfo = recent
            verified = TryLoadingDeparture(accepted, previous, loadingPrediction, factor(4), dl * ratio, recent, candidate, _
                candidateLambda, faceHalf, eventHalfLambda, phaseSteps)
            face = candidate: eventError = mError: eventSolved = verified: eventLambda = candidateLambda
            If verified Then selectedKind = 4
        Else
        Dim bracket() As Double, bracketLambda As Double''',1)
anchor='            eventInfo, face, faceHalf, faceActual, faceHalfActual, eventHalfLambda, eventError, eventSolved, callAllowance, eventLambda, selectedKind)\n    End If'
assert code.count(anchor)==1
code=code.replace(anchor,anchor.replace('\n    End If','\n        End If\n    End If'),1)
code+='''
' Проверяет полный и половинный выход по единственной недавно принятой границе CDP в общей заранее заданной окрестности.
' Направление проверяется по физическому полю и положительной материальной координате; входящая дуговая плоскость не навязывает рост нагрузки.
Private Function TryLoadingDeparture(ByRef accepted() As Double, ByRef previous() As Double, ByRef prediction() As Double, _
        ByVal reference As Double, ByVal predictedLambdaChange As Double, ByRef recent As Variant, ByRef full() As Double, _
        ByRef factor As Double, ByRef half() As Double, ByRef halfLambda As Double, ByRef phaseSteps As Long) As Boolean
    Dim spacing As Double: spacing = mSolver.SpatialFieldDistance(accepted, prediction)
    If spacing <= 0# Then Exit Function
    Dim halfTarget As Double, fullTarget As Double
    halfTarget = recent(3) + 0.25 * (recent(4) - recent(3))
    fullTarget = recent(3) + 0.5 * (recent(4) - recent(3))
    If halfTarget <= 0# Or fullTarget <= halfTarget Then Exit Function
    Dim seed() As Double, i As Long, solved As Boolean
    seed = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = accepted(i) + 0.25 * (prediction(i) - accepted(i)): Next i
    mCalls = mCalls + 1
    solved = mSolver.TryRecentLoadingCoordinate(mBase, mDirection, seed, reference + 0.25 * predictedLambdaChange, _
        halfTarget, mAbsolute, mRelative, mMaximumIterations \\ 2, 1.5 * spacing, half, halfLambda)
    phaseSteps = mSolver.IterationCount: mPrimary = mPrimary + phaseSteps: mSolver.DiscardStep
    If Not solved Or mMaximumIterations - phaseSteps < 1 Then Exit Function
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = 2# * half(i) - accepted(i): Next i
    mCalls = mCalls + 1
    solved = mSolver.TryRecentLoadingCoordinate(mBase, mDirection, seed, reference + 0.5 * predictedLambdaChange, _
        fullTarget, mAbsolute, mRelative, mMaximumIterations - phaseSteps, 3# * spacing, full, factor)
    phaseSteps = phaseSteps + mSolver.IterationCount: mPrimary = mPrimary + mSolver.IterationCount: mSolver.DiscardStep
    If Not solved Then Exit Function
    Dim length As Double: length = mSolver.SpatialFieldDistance(accepted, full)
    If length <= 0# Then Exit Function
    Dim doubled() As Double: doubled = half
    For i = 1 To mSolver.PhysicalDegreeCount: doubled(i) = 2# * half(i) - accepted(i): Next i
    mError = mSolver.SpatialFieldDistance(full, doubled) / length
    Dim lambdaBound As Double, advance As Double, oldLength As Double
    lambdaBound = GeomMax(FIELD_ERROR_LIMIT * Abs(factor - reference), 8# * mRelative * (1# + Abs(reference)))
    If mError > FIELD_ERROR_LIMIT Or Abs(2# * halfLambda - reference - factor) > lambdaBound Then Exit Function
    If 0.5 * length * mModulus <= 8# * GeomMax(Bound(reference), GeomMax(Bound(factor), Bound(halfLambda))) Then Exit Function
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, full, advance, oldLength) Then Exit Function
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, half, advance, oldLength) Then Exit Function
    TryLoadingDeparture = True
End Function
'''
# При неудачной первой координате результирующий массив может ещё не существовать: сертификат всегда имеет допустимый пустой вектор.
code=code.replace('    Dim loadingPrediction() As Double: loadingPrediction = fullSeed',
    '    Dim loadingPrediction() As Double: loadingPrediction = fullSeed',1)
code=code.replace('            eventInfo = recent\n            verified',
    '            eventInfo = recent: ReDim candidate(1 To 1)\n            verified',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
assert code.count('    TestRunActiveLoadingCoordinate\n')==1
code=code.replace('    TestRunActiveLoadingCoordinate\n','    TestRunOperationalArc\n',1)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');imports=['CCoupledArcContinuation','modTestCoupledElasticState']
values.update({'Compile-TestExternalPoint.ps1':compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1).replace('\n','\r\n').encode('utf-8-sig'),
    'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if n.endswith(('.cls','.bas')):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='AutomaticRecentLoadingArcRecovery',actual_import_this_pass=imports,actual_component_count=256,
    all_source43_fields_and_materials_exact818=True,unique_recent_station_from_solver_state=True,
    no_test_entry_required_by_working_arc=True,original_pair_or_selected_phase30_preserved=True,
    common_call_allowance_and_exact_restoration_required=True,no_material_acceptance_inside_arc=True,
    physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
