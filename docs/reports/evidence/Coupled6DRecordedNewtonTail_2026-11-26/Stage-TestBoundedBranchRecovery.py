"""ДЛЯ ТЕСТОВ. Ограничивает временный корректор новой ветви тремя поправками в исходном общем бюджете без принятия материальной истории."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold925RecordedTailSystem_20261126';root=base/'Cold927BoundedBranchRecovery_20261126'
proof=json.loads((base/'Cold926RecordedTailStressScan_20261126/RecordedTailStressScanProof.json').read_bytes())
assert proof['actual_capture_probes_matched']==8 and proof['actual_capture_probes_maximum_residual_difference_MPa']<2e-11
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'BoundedBranchCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='TryProspectiveBranchStep(current, currentLambda, load, norm, absoluteTolerance, relativeTolerance)'
assert code.count(old)==1
code=code.replace(old,old[:-1]+', maxIterations)',1)
begin=code.index("' Проверяет касательную за отклонённой конечной поправкой от той же принятой истории.")
assert code[begin:].count('End Function')==1
code=code[:begin]+'''\
' Проверяет до трёх временных поправок за сменой материальной ветви от одной принятой истории.
' Внутренние точки снижают невязку своей ветви; наружный шаг принимается только при улучшении исходной нормы.
' При неудаче восстанавливается исходный контекст. Каждая дополнительная поправка расходует общий остаток итераций.
Private Function TryProspectiveBranchStep(ByRef current() As Double, ByRef currentLambda As Double, _
        ByRef load() As Double, ByVal initialNorm As Double, ByVal absoluteTolerance As Double, _
        ByVal relativeTolerance As Double, ByVal maxIterations As Long) As Boolean
    Dim oldResidual() As Double, oldMatrix() As Double, oldExternal() As Double, oldResultants() As Double
    Dim oldCandidate() As Double, oldLoad() As Double, oldNorm As Double, oldLambda As Double, oldPivot As Double
    Dim oldCode As ESolverFailureCode, oldPoint As Long, oldLastPoint As Long, oldSide As Long, oldFailures As Long
    oldResidual = mResidual: oldMatrix = mMatrix: oldExternal = mExternal: oldResultants = mTrialResultants
    oldCandidate = mCandidate: oldLoad = load: oldNorm = mLastNorm: oldLambda = mCurrentControlLambda
    oldCode = mFailureCode: oldPoint = mMaterialPoint: oldLastPoint = mLastFailedPoint
    oldSide = mLastFailedSide: oldFailures = mMaterialFailures: oldPivot = mMinimumPivot
    Dim errorNumber As Long, errorSource As String, errorReason As String
    On Error GoTo FailedRecovery
    Dim correction() As Double, trial() As Double, fraction As Double
    Dim branch() As Double, branchLambda As Double, branchNorm As Double, best() As Double, bestLambda As Double, bestNorm As Double
    Dim trialLambda As Double, trialNorm As Double, trialTolerance As Double, i As Long, localStep As Long, hasBetter As Boolean
    branch = oldCandidate: branchLambda = oldLambda
    For localStep = 1 To 3
        If localStep > 1 Then
            If mIterations + 1 >= maxIterations Then GoTo RestoreFailure ' Сохраняется место для штатной повторной проверки наружной точки.
            mIterations = mIterations + 1 ' Первая поправка уже учтена вызывающим методом вместе с отклонённым направлением.
        End If
        SetControlledLoad branchLambda, load
        If Not Evaluate(branch, True, recordTrial:=False, consistentBranch:=True, refinedBranch:=True) Then GoTo RestoreFailure
        branchNorm = EquilibriumNorm(load, branch)
        If Not SolveCorrection(branch, correction) Then GoTo RestoreFailure
        fraction = 1#: hasBetter = False: bestNorm = branchNorm
        Do While fraction >= MIN_FRACTION
            trial = branch
            For i = 1 To mDegreeCount: trial(i) = branch(i) + fraction * correction(i): Next i
            trialLambda = branchLambda + fraction * correction(mDegreeCount + 1)
            SetControlledLoad trialLambda, load
            If Evaluate(trial, False, recordTrial:=False) Then
                trialNorm = EquilibriumNorm(load, trial): trialTolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
                If trialNorm <= trialTolerance Or trialNorm < initialNorm * (1# - MIN_PROGRESS * fraction) Then
                    current = trial: currentLambda = trialLambda
                    mMinimumPivot = mSystem.MinimumPivot
                    TryProspectiveBranchStep = True
                    Exit Function
                End If
                If trialNorm < bestNorm And trialNorm < branchNorm * (1# - MIN_PROGRESS * fraction) Then
                    best = trial: bestLambda = trialLambda: bestNorm = trialNorm: hasBetter = True
                End If
            End If
            fraction = fraction * 0.5
        Loop
        If Not hasBetter Then GoTo RestoreFailure
        branch = best: branchLambda = bestLambda ' Временная точка не становится принятым состоянием материала или наружной итерацией.
    Next localStep
    GoTo RestoreFailure
FailedRecovery:
    errorNumber = Err.Number: errorSource = Err.Source: errorReason = Err.Description
RestoreFailure:
    mResidual = oldResidual: mMatrix = oldMatrix: mExternal = oldExternal: mTrialResultants = oldResultants
    mCandidate = oldCandidate: mLastNorm = oldNorm: mCurrentControlLambda = oldLambda: mMinimumPivot = oldPivot
    For i = 1 To COMPONENT_COUNT: load(i) = oldLoad(i): Next i
    mFailureCode = oldCode: mMaterialPoint = oldPoint: mLastFailedPoint = oldLastPoint
    mLastFailedSide = oldSide: mMaterialFailures = oldFailures
    If errorNumber <> 0 Then Err.Raise errorNumber, errorSource, errorReason
End Function
'''
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='BoundedUncommittedBranchCorrector',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=257,temporary_local_steps_at_most=3,extra_steps_charged_to_original_budget=True,
 original30_iterations_and_all_acceptance_unchanged=True,material_history_never_accepted=True,
 outer_merit_must_improve_original_norm=True,inner_merit_must_improve_temporary_branch_norm=True,
 actual54_current_and_refined55_matrix_guards_preserved=True,no_working_or_release_promotion=True,
 physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'BOUNDED3 TEMPORARY BRANCH CORRECTIONS;SAME30;SAME ACCEPTANCE;NO HISTORY ACCEPTANCE')
