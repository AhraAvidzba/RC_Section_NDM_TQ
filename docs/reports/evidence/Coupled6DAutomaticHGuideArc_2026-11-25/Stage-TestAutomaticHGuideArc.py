"""ДЛЯ ТЕСТОВ. Добавляет автоматическое H-направление на разгрузке дуги и проверяет его от точной истории70 без заданного будущего поля."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold842HReentryWalk_20261125';root=base/'Cold844AutomaticHGuideArc_20261125'
assert json.loads((parent/'HReentryWalkProof.json').read_bytes())['all_new_states_have_strictly_growing_H']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'AutomaticHGuideArcCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
def TestReplace(code,old,new):
    """Меняет единственный ожидаемый фрагмент без поиска приблизительного совпадения."""
    assert code.count(old)==1,old[:160]
    return code.replace(old,new,1)
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=TestReplace(code,'Private mApproaches As Collection','Private mApproaches As Collection\nPrivate mEnvelopeGuide As Variant')
code=TestReplace(code,'    Set mApproaches = New Collection\n    If mSolver Is Nothing',
    '    Set mApproaches = New Collection: mEnvelopeGuide = Empty\n    If mSolver Is Nothing')
anchor='    Dim candidate() As Double, candidateLambda As Double, phaseSteps As Long, eventInfo As Variant, face() As Double, faceHalf() As Double'
code=TestReplace(code,anchor,'''    Dim unloading As Boolean
    If sf Then
        unloading = IsEnvelopeUnloading(full, factor(4), lf)
    ElseIf sh Then
        unloading = IsEnvelopeUnloading(half, factor(4), lh)
    End If
    If unloading Then verified = False
'''+anchor)
anchor='''            verified = TryEvent(accepted, bracket, factor(4), bracketLambda, projection, q, candidate, candidateLambda, phaseSteps, _
                eventInfo, face, faceHalf, faceActual, faceHalfActual, eventHalfLambda, eventError, eventSolved, callAllowance, eventLambda, selectedKind)'''
code=TestReplace(code,anchor,'''            If unloading Then
                verified = TryEnvelopeGuidedEvent(accepted, previous, factor(4), factor(3), projection, q, radius, candidate, candidateLambda, phaseSteps, _
                    eventInfo, face, faceHalf, faceActual, faceHalfActual, eventHalfLambda, eventError, eventSolved, callAllowance, eventLambda, selectedKind)
            Else
'''+anchor+'''
            End If''')
anchor='''    mDetails = "Знаковая дуга: парных итераций " & CStr(pairSteps) & "; всех основных итераций " & CStr(mPrimary) & _
        "; вызовов " & CStr(mCalls) & "; ошибка поля " & CStr(mError) & "; событие " & CStr(eventSolved) & "; восстановление " & CStr(restored)'''
code=TestReplace(code,anchor,anchor+'''
    If unloading Then mDetails = mDetails & "; разгрузочная пара не принята, проверено H-направление: " & CStr(Not IsEmpty(mEnvelopeGuide))''')
anchor='''        ByVal callAllowance As Long, ByRef eventLambda As Double, ByRef selectedKind As Long) As Boolean
    mCalls = mCalls + 1
    eventSolved = mSolver.TryLoadingYieldBoundary(mBase, mDirection, bracket, reference, bracketLambda, mAbsolute, mRelative, _
        mMaximumIterations \\ 2, face, factor, eventInfo)
    phaseSteps = mSolver.IterationCount: mPrimary = mPrimary + phaseSteps'''
replacement='''        ByVal callAllowance As Long, ByRef eventLambda As Double, ByRef selectedKind As Long, _
        Optional ByVal spentIterations As Long = 0) As Boolean
    If mMaximumIterations - spentIterations < 2 Then Exit Function
    mCalls = mCalls + 1
    eventSolved = mSolver.TryLoadingYieldBoundary(mBase, mDirection, bracket, reference, bracketLambda, mAbsolute, mRelative, _
        (mMaximumIterations - spentIterations) \\ 2, face, factor, eventInfo)
    Dim eventSteps As Long: eventSteps = mSolver.IterationCount
    phaseSteps = spentIterations + eventSteps: mPrimary = mPrimary + eventSteps'''
code=TestReplace(code,anchor,replacement)
code+='''
' Выделяет снижение нагрузки без роста принятой материальной координаты; само по себе оно не является разрушением.
' Разрешение lambda соответствует штатному допуску владельца, H читается от той же неизменной истории.
Private Function IsEnvelopeUnloading(ByRef field() As Double, ByVal reference As Double, ByVal factor As Double) As Boolean
    If factor >= reference - 8# * (mRelative / ACCURACY_SCALE) * (1# + Abs(reference)) Then Exit Function
    Dim baseline As Double, resolution As Double, coordinate As Double
    If Not mSolver.TryEnvelopeControlState(mBase, mDirection, reference, mAbsolute, mRelative, baseline, resolution) Then Exit Function
    If baseline <= 0# Then Exit Function
    If Not mSolver.TryEnvelopeCoordinateAt(field, coordinate) Then Exit Function
    IsEnvelopeUnloading = coordinate <= baseline
End Function

' На разгрузке ищет только направляющий H-корень, затем возвращает проверенную материальную границу и половинный подход.
' H-корень не принимается. Его решение и событие делят исходный бюджет фазы; дополнительный вызов остаётся для точного восстановления.
Private Function TryEnvelopeGuidedEvent(ByRef accepted() As Double, ByRef previous() As Double, ByVal reference As Double, _
        ByVal previousLambda As Double, ByRef arcProjection() As Double, ByVal q As Double, ByVal radius As Double, _
        ByRef candidate() As Double, ByRef factor As Double, ByRef phaseSteps As Long, ByRef eventInfo As Variant, _
        ByRef face() As Double, ByRef half() As Double, ByRef actualFace() As Double, ByRef actualHalf() As Double, _
        ByRef halfLambda As Double, ByRef error As Double, ByRef eventSolved As Boolean, ByVal callAllowance As Long, _
        ByRef eventLambda As Double, ByRef selectedKind As Long) As Boolean
    If mCalls + 4 > callAllowance Or mMaximumIterations < 4 Then Exit Function
    Dim baseline As Double, resolution As Double
    If Not mSolver.TryEnvelopeControlState(mBase, mDirection, reference, mAbsolute, mRelative, baseline, resolution) Then Exit Function
    If baseline <= 0# Then Exit Function
    Dim pairPrevious() As Double, beforeH As Double, afterH As Double, hasPair As Boolean
    hasPair = mSolver.TryAcceptedEnvelopeSegment(pairPrevious, beforeH, afterH)
    Dim desired As Double: desired = GeomMax(4# * resolution, 0.0001 * baseline)
    If hasPair Then
        If afterH = baseline And afterH > beforeH Then desired = GeomMax(desired, afterH - beforeH)
    End If
    desired = GeomMin(desired, GeomMax(4# * resolution, 0.02 * baseline))
    Dim ratio As Double
    If hasPair Then
        If afterH = baseline And afterH > beforeH Then ratio = GeomMin(1024#, desired / (afterH - beforeH))
    End If
    Dim seed() As Double: seed = accepted
    Dim i As Long
    For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = accepted(i) + ratio * (accepted(i) - previous(i)): Next i
    Dim initialLambda As Double: initialLambda = reference + ratio * (reference - previousLambda)
    Dim guide() As Double, actualGuide() As Double, projection() As Double, guideLambda As Double, guideH As Double
    Dim solved As Boolean, probes As Long, guideSteps As Long, revision As Long
    revision = mSolver.AcceptedRevision
    solved = mSolver.TryEnvelopeStep(mBase, mDirection, previous, reference, seed, initialLambda, baseline + desired, guideLambda, _
        projection, mAbsolute, mRelative, mMaximumIterations \\ 2, probes, False)
    guideSteps = mSolver.IterationCount: phaseSteps = guideSteps
    mCalls = mCalls + probes: mPrimary = mPrimary + guideSteps
    If solved Then
        mSolver.CopyTrial guide, actualGuide
        If Not mSolver.TryReadEnvelopeCoordinate(False, guideH) Then solved = False
        If guideH <= baseline Then solved = False
    Else
        ReDim guide(1 To 1)
    End If
    mSolver.DiscardStep
    mEnvelopeGuide = Array(Array(Abs(CLng(solved)), probes, guideSteps, baseline, desired, baseline + desired, guideLambda, _
        initialLambda, guideH, resolution, beforeH, afterH, reference, previousLambda, revision, mSolver.AcceptedRevision, 3# * radius), seed, guide)
    If mSolver.AcceptedRevision <> revision Or mCalls > callAllowance Or guideSteps > mMaximumIterations \\ 2 Then _
        Err.Raise 5, "CCoupledArcContinuation", "H-направление нарушило историю или общий бюджет."
    If Not solved Or mCalls + 3 > callAllowance Then Exit Function
    Dim verified As Boolean
    verified = TryEvent(accepted, guide, reference, guideLambda, arcProjection, q, candidate, factor, phaseSteps, eventInfo, _
        face, half, actualFace, actualHalf, halfLambda, error, eventSolved, callAllowance, eventLambda, selectedKind, guideSteps)
    If verified Then
        verified = mSolver.SpatialFieldDistance(accepted, candidate) <= 3# * radius
        If Not verified Then selectedKind = 0
    End If
    TryEnvelopeGuidedEvent = verified
End Function

' Передаёт независимую диагностическую копию H-направления; направляющий корень не является принятым состоянием.
Friend Function EnvelopeGuideVerification() As Variant
    EnvelopeGuideVerification = mEnvelopeGuide
End Function
'''
values[name]=code
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
anchor='    TestRecentLoadingDirectionGuards\n    TestRunAutomaticLoadingWalk\n    TestRecentLoadingResetGuard\n'
code=TestReplace(code,anchor,'    TestRunAutomaticHGuideArc\n')
code+='''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Вызывает рабочую дугу на состоянии70 без будущего поля при трёх остатках вызовов; неизвестная история не принимается.
Private Sub TestRunAutomaticHGuideArc()
    Dim allowances As Variant: allowances = Array(5, 6, 96)
    Dim handle As Integer: handle = FreeFile
    Open mPolishRootTest & "\\Native\\AutomaticHGuide.bin" For Binary Access Write As #handle
    Dim magic As String * 16: magic = "NDMAUTOH00000001"
    Put #handle, , magic
    Dim caseIndex As Long, allowance As Long, factor As Double, solved As Boolean, projection() As Double
    Dim arc As CCoupledArcContinuation, certificate As Variant, guide As Variant, vector As Variant, info As Variant
    Dim revision As Long: revision = mReverseSolverTest.AcceptedRevision
    Dim accepted() As Double: mReverseSolverTest.CopyAccepted accepted
    For caseIndex = 1 To 3
        allowance = CLng(allowances(caseIndex - 1))
        Set arc = New CCoupledArcContinuation
        arc.Initialize mReverseSolverTest, mReverseBaseTest, mReverseDirectionTest, mReverseOptionsTest(3), mReverseOptionsTest(4), 30
        solved = arc.TryStep(mReverseFieldsTest, mReverseTraceTest, 1#, allowance, factor, projection)
        certificate = arc.Verification: info = certificate(0): guide = arc.EnvelopeGuideVerification
        TestWriteReverseBinaryVector handle, Array(caseIndex, allowance, Abs(CLng(solved)), factor, revision, mReverseSolverTest.AcceptedRevision)
        For Each vector In certificate: TestWriteReverseBinaryVector handle, vector: Next vector
        Dim state As Variant: state = Array(0#)
        Dim eventInfo As Variant: eventInfo = certificate(9)
        If eventInfo(0) Then
            Dim concrete As CCoupledConcreteResponse
            Select Case CLng(eventInfo(2))
                Case -1: Set concrete = mReverseSidesTest(1)
                Case 0: Set concrete = mReverseConcreteTest
                Case 1: Set concrete = mReverseSidesTest(3)
            End Select
            state = concrete.ReadNewtonYieldEventStateTest(CLng(eventInfo(1)))
        End If
        TestWriteReverseBinaryVector handle, state
        Dim field() As Double, actual() As Double, half() As Double
        If solved Then
            mReverseSolverTest.CopyTrial field, actual
            TestWriteReverseBinaryVector handle, field: TestWriteReverseBinaryVector handle, actual: TestWriteReverseBinaryVector handle, projection
            TestWriteElasticPrincipalRoot 10 * caseIndex + 1, field
            half = certificate(11): TestWriteElasticPrincipalRoot 10 * caseIndex + 2, half
        Else
            TestWriteReverseBinaryVector handle, Array(0#): TestWriteReverseBinaryVector handle, Array(0#): TestWriteReverseBinaryVector handle, Array(0#)
            field = certificate(6)
            If UBound(field) > 1 Then TestWriteElasticPrincipalRoot 10 * caseIndex + 1, field
            half = certificate(5)
            If UBound(half) > 1 Then TestWriteElasticPrincipalRoot 10 * caseIndex + 2, half
        End If
        TestWriteReverseBinaryVector handle, mReverseSolverTest.ReadPolishEffortTest
        TestWriteReverseBinaryVector handle, Array(Abs(CLng(Not IsEmpty(guide))))
        If Not IsEmpty(guide) Then
            For Each vector In guide: TestWriteReverseBinaryVector handle, vector: Next vector
            Dim guideInfo As Variant: guideInfo = guide(0)
            If guideInfo(0) Then
                field = guide(2): TestWriteElasticPrincipalRoot 10 * caseIndex + 3, field
            End If
        End If
        mReverseSolverTest.DiscardStep
        If mReverseSolverTest.AcceptedRevision <> revision Or arc.Calls > allowance Then Err.Raise 5, , "Автоматическое H-направление нарушило историю или остаток вызовов."
        Dim checked() As Double, i As Long: mReverseSolverTest.CopyAccepted checked
        For i = 1 To UBound(accepted)
            If checked(i) <> accepted(i) Then Err.Raise 5, , "Автоматическое H-направление изменило исходное принятое поле."
        Next i
    Next caseIndex
    Close #handle
End Sub
''';values[name]=code
values['Input/Reverse1.bin']=(base/'Cold841HGuideMaterialEvent_20261125/Input/Reverse1.bin').read_bytes()
unused=root/'Input/HMaterialGuide.bin'
if unused.exists():
    target=root/'UnusedHMaterialGuide.bin';assert target.resolve().is_relative_to(root.resolve())
    unused.rename(target);prep['inputs'].pop('Input/HMaterialGuide.bin',None)
assert not unused.exists()
imports=['CCoupledArcContinuation','modTestCoupledElasticState'];compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=TestReplace(compiler,"-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(n) for n in imports)+')')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for n,data in values.items():
    if isinstance(data,str):
        if n.endswith(('.cls','.bas')):assert max(map(len,data.splitlines()))<1024
        data=data.replace('\n','\r\n').encode('utf-8-sig')
    (root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
assert len('NDMAUTOH00000001')==16
prep['directed_scope']=dict(case='AutomaticEnvelopeGuideEventFromOriginal70',actual_import_this_pass=sorted(imports),actual_component_count=256,
    no_precomputed_future_field_input=True,original70_source_exact841=True,unloading_arc_pair_not_admitted=True,
    converged_full_or_half_can_trigger_guide_without_pair_admission=True,missing_diagnostic_half_not_read=True,
    guide_and_event_share_original30_iteration_phase=True,all_calls_including_guide_counted=True,
    accepted_event_radius_predeclared_as_three_original_arc_radii=True,original_force_half_lambda_gauge_and_steel_checks=True,
    allowance5_control_and6_and96=True,no_candidate_history_accepted=True,physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root)
