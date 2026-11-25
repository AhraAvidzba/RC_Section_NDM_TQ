"""ДЛЯ ТЕСТОВ. Исправляет приёмку вычислительного возврата обязательной согласованностью полного и половинного шагов."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold752IntegratedGuideDirected_20261125'
root = base / 'Cold757HalfGuardedMaterialReturn_20261125'
proof = json.loads((base / 'Cold756GuideHalfConfirmation_20261125/GuideHalfConfirmationProof.json').read_bytes())
assert proof['half_solved'] and not proof['locally_confirmed_by_half_step']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'HalfGuardCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'src/Solver/CCoupledNewtonSolver.cls'
code = (root / name).read_text('utf-8-sig')
old = '''    If Abs(coordinate - target) * mReferenceModulus > tolerance Then GoTo RejectedReturn
    If Not TrialLambdaUncertainty(absoluteTolerance, relativeTolerance, uncertainty) Then GoTo RejectedReturn'''
new = '''    If Abs(coordinate - target) * mReferenceModulus > tolerance Then GoTo RejectedReturn
    If Not ConfirmOriginalReturnHalf(field, returnedLambda, projection, target, consistent, absoluteTolerance, relativeTolerance, _
            maxIterations - initialEffort.Iterations, referenceLambda, details) Then GoTo RejectedReturn
    SetControlledLoad returnedLambda, load
    If Not TrialLambdaUncertainty(absoluteTolerance, relativeTolerance, uncertainty) Then GoTo RejectedReturn'''
assert code.count(old) == 1
code = code.replace(old, new, 1)
helper = '''
' Подтверждает локальность возвращённого поля половиной исходного управляющего приращения от той же принятой истории.
' Общий остаток включает полный возврат, половину и заполнение пробных буферов; несогласованные корни не принимаются.
Private Function ConfirmOriginalReturnHalf(ByRef full() As Double, ByVal fullLambda As Double, ByRef projection() As Double, _
        ByVal target As Double, ByVal consistent As Boolean, ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, _
        ByVal maximumSteps As Long, ByVal referenceLambda As Double, ByRef details As String) As Boolean
    If Not mMaterialGuideReturnActive Or Not mCandidateReady Or mHybridActive Then Exit Function
    Dim fullEffort As TSolveEffort: ReadSolveEffort fullEffort
    Dim remaining As Long: remaining = maximumSteps - fullEffort.Iterations
    If remaining < 1 Then Exit Function
    Dim revision As Long: revision = mAcceptedRevision
    Dim base() As Double, direction() As Double, accepted() As Double, seed() As Double, prediction As Variant
    base = mBaseLoad: direction = mLoadDirection: accepted = mAccepted: seed = accepted
    Dim i As Long, coordinate As Double
    For i = 1 To mGaugeStart - 1: coordinate = coordinate + projection(i) * accepted(i): Next i
    For i = 1 To mDegreeCount: seed(i) = 0.5 * (accepted(i) + full(i)): Next i
    prediction = seed
    Dim halfTarget As Double: halfTarget = 0.5 * (coordinate + target)
    Dim halfLambda As Double, half() As Double, actual() As Double, doubled() As Double, solved As Boolean, attempted As Boolean
    Dim tensorError As Double, lambdaError As Double, length As Double, lambdaBound As Double, load() As Double, tolerance As Double
    Dim errorNumber As Long, errorSource As String, errorText As String
    On Error GoTo InterruptedHalf
    DiscardStep: attempted = True
    solved = TryDeformationStep(base, direction, projection, halfTarget, 0.5 * (referenceLambda + fullLambda), halfLambda, _
        absoluteTolerance, relativeTolerance, remaining, consistentControl:=consistent, initialField:=prediction, _
        allowSmallCorrection:=True, allowHybridRecovery:=True, conditionHybridRecovery:=True)
    If mAcceptedRevision <> revision Then Err.Raise 5, "CCoupledNewtonSolver", "Половинный возврат изменил принятую материальную историю."
    If Not solved Or mRecoveredControlReady Or halfLambda <= 0# Then GoTo CountHalf
    CopyTrial half, actual
    tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    If Abs(ControlResidual(half)) > tolerance Then GoTo CountHalf
    If Not PrincipalSteelWithinLimits(half) Then GoTo CountHalf
    length = SpatialFieldDistance(accepted, full)
    If length <= 0# Then GoTo CountHalf
    doubled = accepted
    For i = 1 To mGaugeStart - 1: doubled(i) = 2# * half(i) - accepted(i): Next i
    tensorError = SpatialFieldDistance(full, doubled) / length
    lambdaError = Abs(2# * halfLambda - referenceLambda - fullLambda)
    lambdaBound = GeomMax(YIELD_EVENT_HALF_ERROR * Abs(fullLambda - referenceLambda), 8# * relativeTolerance * (1# + Abs(referenceLambda)))
    details = details & "; полный/половинный возврат: ошибка поля " & CStr(tensorError) & "; ошибка lambda " & CStr(lambdaError) & "; допуск lambda " & CStr(lambdaBound)
    If tensorError > YIELD_EVENT_HALF_ERROR Or lambdaError > lambdaBound Then GoTo CountHalf
    DiscardStep
    mConsistentControl = consistent: mArcLoadProjection = 0#: mEnvelopeControl = False
    If Not PublishPrincipalCandidate(full, fullLambda, projection, tolerance) Then GoTo CountHalf
    mTargetControl = target ' Сохраняется точная первоначальная цель, а не её вычисленное значение на поле.
    SetControlledLoad fullLambda, load
    If Abs(ControlResidual(full)) > EquilibriumTolerance(absoluteTolerance, relativeTolerance) Then GoTo CountHalf
    ConfirmOriginalReturnHalf = True
CountHalf:
    On Error GoTo 0
    If Not ConfirmOriginalReturnHalf Then DiscardStep
    If attempted Then AddSolveEffort fullEffort
    If mIterations > maximumSteps Then Err.Raise 5, "CCoupledNewtonSolver", "Полный и половинный возврат превысили общий остаток итераций."
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorText
    Exit Function
InterruptedHalf:
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    Resume CountHalf
End Function

'''
anchor = "' Подтверждает материалом разностную пробу в действующем контексте; обычные уравнения и принятая история остаются неизменными."
assert code.count(anchor) == 1
code = code.replace(anchor, helper + anchor, 1)
assert max(map(len, code.splitlines())) < 1024
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig')}
compiler = (here / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames 'CCoupledNewtonSolver'", 1).replace('\n', '\r\n').encode('utf-8-sig')
values['tools/build_workbook/Update-IsolatedOptimizerCode.ps1'] = (base / 'Cold753IntegratedColdFrozenReplay_20261125/tools/build_workbook/Update-IsolatedOptimizerCode.ps1').read_bytes()
values['ParentCodeUpdate.json'] = (parent / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (parent / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='HalfGuardedOriginalMaterialReturn', actual_import_this_pass=['CCoupledNewtonSolver'],
    material_controller_and_test_module_exact752=True, unsound756_full_half_return_must_be_rejected=True,
    same_YIELD_EVENT_HALF_ERROR005_and_original_lambda_resolution=True, shared_primary_budget30=True,
    accepted_history_never_advanced=True, full_buffers_and_exact_original_target_republished_only_after_half_pass=True,
    no_working_promotion=True, positive_helper_specific_case_pending=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('757 ORIGINAL RETURN REQUIRES HALF CONSISTENCY; SAME30 BUDGET; REJECT DISCONNECTED756 CANDIDATE')
