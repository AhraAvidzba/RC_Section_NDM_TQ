"""ДЛЯ ТЕСТОВ. Сохраняет бюджет для короткого предсобытийного шага после неудачи половины автоматически найденной границы."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold962CurrentOperationalArc_20261126';root=base/'Cold963IncomingEventBudget_20261126'
proof=json.loads((parent/'CurrentOperationalArcProof.json').read_bytes())
assert not proof['returned_verified_candidate'] and proof['event'] and not proof['approaches'] and proof['event_phase_steps']==30
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
ancestor=root/'IncomingEventBudgetCompilationAncestor';ancestor.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(ancestor/name)
updates={};name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
changes={
 'Private mUsesTangent As Boolean, mPredictionVerification As Variant':'Private mUsesTangent As Boolean, mPredictionVerification As Variant\nPrivate mIncomingRecoveryDetails As String \' Сведения о сокращённых полных пробах того же входящего перехода.',
 '    Set mApproaches = New Collection: mEnvelopeGuide = Empty: mLoadingGuide = Empty: mNextStepRatioLimit = 2#':'    Set mApproaches = New Collection: mEnvelopeGuide = Empty: mLoadingGuide = Empty: mNextStepRatioLimit = 2#\n    mIncomingRecoveryDetails = vbNullString',
 '    If unloading Then mDetails = mDetails & "; разгрузочная пара не принята, проверено H-направление: " & CStr(Not IsEmpty(mEnvelopeGuide))':'    If unloading Then mDetails = mDetails & "; разгрузочная пара не принята, проверено H-направление: " & CStr(Not IsEmpty(mEnvelopeGuide))\n    If Len(mIncomingRecoveryDetails) > 0 Then mDetails = mDetails & " | " & mIncomingRecoveryDetails',
 '    prediction = seed: mCalls = mCalls + 1\n    solved = mSolver.TryDeformationStep(mBase, mDirection, projection, coordinate - 0.5 * spacing':'    prediction = seed: mCalls = mCalls + 1\n    Dim halfBudget As Long: halfBudget = (mMaximumIterations - phaseSteps) \\ 4 \' Остаток нужен также сокращённому полному полю, его половине и восстановлению.\n    If halfBudget < 1 Then halfBudget = 1\n    solved = mSolver.TryDeformationStep(mBase, mDirection, projection, coordinate - 0.5 * spacing',
 '        halfLambda, mAbsolute, mRelative, mMaximumIterations - phaseSteps, consistentControl:=True, initialField:=prediction, _':'        halfLambda, mAbsolute, mRelative, halfBudget, consistentControl:=True, initialField:=prediction, _',
 '    If Not solved Then mSolver.DiscardStep: Exit Function\n    mSolver.CopyTrial half, actualHalf: mSolver.DiscardStep':'    If Not solved Then\n        mSolver.DiscardStep\n        If TryIncomingFaceApproach(accepted, face, factor, reference, projection, coordinate, spacing, arcProjection, q, _\n                callAllowance, eventInfo, phaseSteps, candidate, factor) Then\n            selectedKind = 3: TryEvent = True\n        End If\n        Exit Function\n    End If\n    mSolver.CopyTrial half, actualHalf: mSolver.DiscardStep',
 '        ByRef eventInfo As Variant, ByRef phaseSteps As Long, ByRef selected() As Double, ByRef selectedLambda As Double) As Boolean':'        ByRef eventInfo As Variant, ByRef phaseSteps As Long, ByRef selected() As Double, ByRef selectedLambda As Double, _\n        Optional ByVal initialFraction As Double = 0.5) As Boolean',
 '    candidate = initial: actualCandidate = initialActual: factor = initialLambda: fraction = 0.5':'    If initialFraction <= 0# Or initialFraction > 1# Then Err.Raise 5, "CCoupledArcContinuation", "Доля исходного предсобытийного поля вне интервала (0;1]."\n    candidate = initial: actualCandidate = initialActual: factor = initialLambda: fraction = initialFraction',
}
for old,new in changes.items():assert code.count(old)==1,(old,code.count(old));code=code.replace(old,new,1)
helper='''
' Вычисляет короткое полное равновесие от найденной границы, если половинная проба не решена; материал не принимает.
' Сокращённая точка и её половина делят остаток той же фазы, а критерии исходного предсобытийного метода сохраняются.
Private Function TryIncomingFaceApproach(ByRef accepted() As Double, ByRef face() As Double, ByVal faceFactor As Double, _
        ByVal reference As Double, ByRef projection() As Double, ByVal faceCoordinate As Double, ByVal faceSpacing As Double, _
        ByRef arcProjection() As Double, ByVal q As Double, ByVal callAllowance As Long, ByRef eventInfo As Variant, _
        ByRef phaseSteps As Long, ByRef selected() As Double, ByRef selectedLambda As Double) As Boolean
    Dim fraction As Double: fraction = 0.25
    Dim level As Long, i As Long, remaining As Long, budget As Long, solved As Boolean, steps As Long
    Dim seed() As Double, prediction As Variant, field() As Double, actual() As Double, factor As Double
    For level = 1 To 4
        remaining = mMaximumIterations - phaseSteps
        If remaining < 3 Or mCalls + 3 > callAllowance Then Exit Function
        seed = accepted
        For i = 1 To mSolver.PhysicalDegreeCount: seed(i) = accepted(i) + fraction * (face(i) - accepted(i)): Next i
        prediction = seed: budget = (remaining - 1) \\ 2 ' Часть остатка сохраняется для половины и точного восстановления кандидата.
        mCalls = mCalls + 1
        solved = mSolver.TryDeformationStep(mBase, mDirection, projection, faceCoordinate - (1# - fraction) * faceSpacing, _
            reference + fraction * (faceFactor - reference), factor, mAbsolute, mRelative, budget, _
            consistentControl:=True, initialField:=prediction, allowHybridRecovery:=True, conditionHybridRecovery:=True)
        steps = mSolver.IterationCount: phaseSteps = phaseSteps + steps: mPrimary = mPrimary + steps
        mIncomingRecoveryDetails = mIncomingRecoveryDetails & "; сокращённая полная проба " & CStr(fraction) & ": " & CStr(solved) & _
            "; итераций " & CStr(steps) & "; использовано в фазе " & CStr(phaseSteps)
        If solved Then mSolver.CopyTrial field, actual
        mSolver.DiscardStep
        If phaseSteps > mMaximumIterations Or mCalls >= callAllowance Then Err.Raise 5, "CCoupledArcContinuation", "Сокращённая проба превысила остаток фазы или вызовов."
        If solved Then
            TryIncomingFaceApproach = TryPreEventApproach(accepted, field, actual, factor, reference, projection, faceCoordinate, faceSpacing, _
                arcProjection, q, callAllowance, eventInfo, phaseSteps, selected, selectedLambda, initialFraction:=fraction)
            Exit Function
        End If
        fraction = 0.5 * fraction
    Next level
End Function
'''
anchor="' Читает исходное равновесное разрешение в принятой норме уравнений с фактическим внутренним допуском."
assert code.count(anchor)==1;code=code.replace(anchor,helper+'\n'+anchor,1)
updates[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
updates['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledArcContinuation'",1).replace('\n','\r\n').encode('utf-8-sig')
updates.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
prep=json.loads((root/'Preparation.json').read_bytes())
for name,data in updates.items():
    if name.endswith('.cls'):assert max(map(len,data.decode('utf-8-sig').splitlines()))<1024
    (root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ReserveEventPhaseForShorterIncomingStepAfterHalfFailure',actual_import_this_pass=['CCoupledArcContinuation'],
    actual_component_count=257,all256_other_components_exact962=True,test_entry_and_inputs_bitexact962=True,
    no_future_guide_or_field=True,internal_accuracy_scale0001_and_acceptance_unchanged=True,
    one_quarter_remaining_budget_for_initial_event_half=True,shortened_full_uses_half_remaining_budget_then_existing_PreEvent=True,
    original30_event_phase_and96_standalone_calls_preserved=True,no_material_history_acceptance_inside_arc=True,
    physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'EVENT HALF BUDGET RESERVED;SHORTER FULL+EXISTING PRE-EVENT CHECK;ALL ACCURACY/ACCEPTANCE UNCHANGED')
