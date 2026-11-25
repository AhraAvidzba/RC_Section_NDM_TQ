"""ДЛЯ ТЕСТОВ. Перенастраивает материальный шаг по проверенному непринятому ориентиру, сохраняя исходные радиусы и критерии пары."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold848NewLoadingCoordinateDiagnosis_20261125';root=base/'Cold851LoadingSecantReparameterization_20261125'
proof=json.loads((parent/'NewLoadingCoordinateDiagnosisProof.json').read_bytes())
assert proof['Newton_method_exact_after_removing_readonly_instrumentation']
assert all(p['coordinate_records'][0]['stage']=='RadiusAndLambda' for p in proof['cases'][1:])
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'LoadingSecantCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
old='Private mEnvelopeGuide As Variant';assert code.count(old)==1
code=code.replace(old,old+'\nPrivate mLoadingGuide As Variant',1)
old='Set mApproaches = New Collection: mEnvelopeGuide = Empty: mNextStepRatioLimit = 2#';assert code.count(old)==1
code=code.replace(old,'Set mApproaches = New Collection: mEnvelopeGuide = Empty: mLoadingGuide = Empty: mNextStepRatioLimit = 2#',1)
old='candidateLambda, faceHalf, eventHalfLambda, phaseSteps)\n            face = candidate';assert code.count(old)==1
code=code.replace(old,'candidateLambda, faceHalf, eventHalfLambda, phaseSteps, callAllowance)\n            eventInfo = recent\n            face = candidate',1)
method,=re.findall(r'(?ms)^Private Function TryLoadingDeparture\(.*?^End Function',code)
pair=method.replace('Private Function TryLoadingDeparture(','Private Function TryLoadingDeparturePair(',1)
pair=pair.replace('ByRef phaseSteps As Long) As Boolean','ByRef phaseSteps As Long, ByVal spacing As Double) As Boolean',1)
old='    Dim spacing As Double: spacing = mSolver.SpatialFieldDistance(accepted, prediction)\n';assert pair.count(old)==1
pair=pair.replace(old,'    Dim remaining As Long: remaining = mMaximumIterations - phaseSteps\n    If remaining < 2 Then Exit Function\n',1)
pair=pair.replace('mMaximumIterations \\ 2, 1.5 * spacing','remaining \\ 2, 1.5 * spacing',1)
pair=pair.replace('    phaseSteps = mSolver.IterationCount: mPrimary = mPrimary + phaseSteps: mSolver.DiscardStep',
    '    phaseSteps = phaseSteps + mSolver.IterationCount: mPrimary = mPrimary + mSolver.IterationCount: mSolver.DiscardStep',1)
pair=pair.replace('    TryLoadingDeparture = True','    TryLoadingDeparturePair = True',1)
# Силовые, геометрические и половинные условия пары сохраняются буквально.
assert pair[pair.index('    Dim length As Double:'):pair.index('    TryLoadingDeparturePair = True')]==method[method.index('    Dim length As Double:'):method.index('    TryLoadingDeparture = True')]
wrapper='''Private Function TryLoadingDeparture(ByRef accepted() As Double, ByRef previous() As Double, ByRef prediction() As Double, _
        ByVal reference As Double, ByVal predictedLambdaChange As Double, ByRef recent As Variant, ByRef full() As Double, _
        ByRef factor As Double, ByRef half() As Double, ByRef halfLambda As Double, ByRef phaseSteps As Long, ByVal callAllowance As Long) As Boolean
    Dim spacing As Double: spacing = mSolver.SpatialFieldDistance(accepted, prediction)
    If spacing <= 0# Then Exit Function
    If TryLoadingDeparturePair(accepted, previous, prediction, reference, predictedLambdaChange, recent, full, factor, half, halfLambda, phaseSteps, spacing) Then
        TryLoadingDeparture = True: Exit Function
    End If
    If mMaximumIterations - phaseSteps < 3 Or mCalls + 4 > callAllowance Then Exit Function
    If LBound(half) <> 1 Or UBound(half) <> UBound(accepted) Or halfLambda <= 0# Then Exit Function
    Dim guide() As Double: guide = half
    Dim length As Double, advance As Double, oldLength As Double
    length = mSolver.SpatialFieldDistance(accepted, guide)
    If length <= 0# Then Exit Function
    If Not mSolver.SpatialIncrementAdvances(previous, accepted, guide, advance, oldLength) Then Exit Function
    Dim projection() As Double, coordinate As Double, movement() As Double, i As Long, seed As Variant
    movement = guide
    For i = 1 To UBound(movement): movement(i) = guide(i) - accepted(i): Next i
    If Not mSolver.SelectSpatialDirectionControl(accepted, guide, movement, projection, coordinate, oldLength) Then Exit Function
    seed = guide
    Dim verified As Boolean, guideLambda As Double, restored() As Double, actual() As Double
    mCalls = mCalls + 1
    verified = mSolver.TryDeformationStep(mBase, mDirection, projection, coordinate, halfLambda, guideLambda, _
        mAbsolute, mRelative, 1, consistentControl:=True, initialField:=seed)
    phaseSteps = phaseSteps + mSolver.IterationCount: mPrimary = mPrimary + mSolver.IterationCount
    Dim restoreIterations As Long: restoreIterations = mSolver.IterationCount
    If verified Then
        mSolver.CopyTrial restored, actual
        verified = guideLambda = halfLambda And restoreIterations = 0
        If verified Then
            For i = 1 To UBound(guide)
                If guide(i) <> restored(i) Then verified = False: Exit For
            Next i
        End If
    End If
    mSolver.DiscardStep
    If Not verified Then Exit Function
    Dim fraction As Double: fraction = spacing / length
    prediction = accepted
    For i = 1 To mSolver.PhysicalDegreeCount: prediction(i) = accepted(i) + fraction * (guide(i) - accepted(i)): Next i
    Dim redirected As Variant
    If Not mSolver.TryReadRecentLoadingDeparture(prediction, redirected) Then Exit Function
    If redirected(1) <> recent(1) Or redirected(2) <> recent(2) Or redirected(6) <> recent(6) Then Exit Function
    mLoadingGuide = Array(Array(reference, guideLambda, length, spacing, fraction, mSolver.AcceptedRevision, restoreIterations, mCalls, phaseSteps), guide, prediction)
    recent = redirected
    predictedLambdaChange = fraction * (guideLambda - reference)
    If mMaximumIterations - phaseSteps < 2 Or mCalls + 3 > callAllowance Then Exit Function
    TryLoadingDeparture = TryLoadingDeparturePair(accepted, previous, prediction, reference, predictedLambdaChange, recent, full, factor, half, halfLambda, phaseSteps, spacing)
End Function

' Решает полную и половинную материальные пробы в исходной геометрической окрестности; все повторения делят один бюджет фазы.
'''+pair
code=code.replace(method,wrapper,1)
code+='''
' Возвращает проверенный непринятый ориентир и равноудалённый прогноз; история и радиусы допуска из-за ориентира не меняются.
Friend Property Get LoadingGuideVerification() As Variant
    LoadingGuideVerification = mLoadingGuide
End Property
''';values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='        Set records = mReverseSolverTest.TestYieldCoordinateCapture';assert code.count(old)==1
code=code.replace(old,'''        Dim guideRecord As Variant: guideRecord = arc.LoadingGuideVerification
        Dim guideHandle As Integer: guideHandle = FreeFile
        Open mPolishRootTest & "\\Native\\LoadingGuide." & Format$(caseIndex + 1, "000") & ".bin" For Binary Access Write As #guideHandle
        Dim guideMagic As String * 16: guideMagic = "NDMLOADGUIDE0001"
        Put #guideHandle, , guideMagic
        TestWriteReverseBinaryVector guideHandle, Array(Abs(CLng(Not IsEmpty(guideRecord))))
        If Not IsEmpty(guideRecord) Then
            Dim guideVector As Variant
            For Each guideVector In guideRecord: TestWriteReverseBinaryVector guideHandle, guideVector: Next guideVector
        End If
        Close #guideHandle
'''+old,1);values[name]=code
imports=['CCoupledArcContinuation','modTestCoupledElasticState']
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');old="-ModuleNames 'modTestCoupledElasticState'";assert compiler.count(old)==1
values['Compile-TestExternalPoint.ps1']=compiler.replace(old,'-ModuleNames @('+','.join(repr(n) for n in imports)+')',1)
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if isinstance(data,str):
        if n.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='LoadingSecantReparameterization',actual_import_this_pass=imports,actual_component_count=256,
    exact848_prefix86=True,Newton_and_materials_unchanged=True,original_half_and_full_radii_unchanged=True,
    original_full_half_force_lambda_direction_and_resolution_criteria_unchanged=True,
    guide_requires_bit_exact_zero_iteration_force_restore=True,guide_not_accepted=True,
    additional_pair_shares_original30_phase_budget_and_all_calls_counted=True,
    no_precomputed_outgoing_field_or_tangent_input=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root,'ORIGINAL RADII/PAIR/30BUDGET;GUIDE VERIFIED ZERO ITERATIONS;NO MATERIAL CHANGE')
