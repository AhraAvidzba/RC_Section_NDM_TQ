"""ДЛЯ ТЕСТОВ. Включает проверенный возврат к исходной цели в материальный резерв позднего H, сохраняя бюджет и условия приёмки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold751IntegratedMaterialGuide_20261125';parent=base/'Cold743UnverifiedCumulativeRecovery_20261125'
proof=json.loads((base/'Cold750DepartureGuideReturn_20261125/DepartureGuideReturnProof.json').read_bytes())
assert proof['solved_original_target'] and proof['independent_NDS_and_original_control_verified'] and proof['total_primary_steps']==17
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'MaterialGuideCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
anchor='Private mPostCrackTangentModeTest As Long\n';assert code.count(anchor)==1
code=code.replace(anchor,anchor+'Private mMaterialGuideReturnActive As Boolean\n',1)
pattern=r'(?ms)^Private Function TryDifferenceMaterialField\(.*?^End Function';method,=re.findall(pattern,code)
guard='    details = vbNullString\n';assert method.count(guard)==1
changed=method.replace(guard,guard+'    If mMaterialGuideReturnActive Then Exit Function\n',1)
old='''        If lambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty Then GoTo Restore'''
new='''        If lambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty Then
            If TryReturnOriginalMaterialControl(field, lambda, originalProjection, originalTarget, originalConsistent, _
                    load, absoluteTolerance, relativeTolerance, maxIterations, referenceLambda, details) Then
                TryDifferenceMaterialField = True
                Exit Function
            End If
            GoTo Restore
        End If'''
assert changed.count(old)==1;changed=changed.replace(old,new,1);code=code.replace(method,changed,1)
helper='''
' Возвращает исходную управляющую цель из непринятого материального корня, не меняя принятую историю.
' Вложенный расчёт получает только остаток первичных итераций; все его затраты добавляются через общий TSolveEffort.
Private Function TryReturnOriginalMaterialControl(ByRef guide() As Double, ByVal guideLambda As Double, _
        ByRef projection() As Double, ByVal target As Double, ByVal consistent As Boolean, ByRef load() As Double, _
        ByVal absoluteTolerance As Double, ByVal relativeTolerance As Double, ByVal maxIterations As Long, _
        ByVal referenceLambda As Double, ByRef details As String) As Boolean
    If mMaterialGuideReturnActive Or mHybridActive Or Not mCandidateReady Then Exit Function
    If Not mUseSpatialHistory Or Not SupportsHybridRecovery Or mEnvelopeControl Then Exit Function
    Dim initialEffort As TSolveEffort: ReadSolveEffort initialEffort
    Dim remaining As Long: remaining = maxIterations - initialEffort.Iterations
    If remaining < 1 Then Exit Function
    Dim revision As Long: revision = mAcceptedRevision
    Dim base() As Double, direction() As Double, seed As Variant
    base = mBaseLoad: direction = mLoadDirection: seed = guide
    Dim returnedLambda As Double, solved As Boolean, attempted As Boolean
    Dim errorNumber As Long, errorSource As String, errorText As String, phaseDetails As String
    Dim returnIterations As Long, uncertainty As Double, field() As Double, actual() As Double, coordinate As Double, i As Long
    On Error GoTo InterruptedReturn
    DiscardStep ' Вспомогательный корень остаётся только прогнозом; три принятые материальные истории не продвигаются.
    mMaterialGuideReturnActive = True: attempted = True
    solved = TryDeformationStep(base, direction, projection, target, guideLambda, returnedLambda, _
        absoluteTolerance, relativeTolerance, remaining, consistentControl:=consistent, initialField:=seed, _
        allowSmallCorrection:=True, allowHybridRecovery:=True, conditionHybridRecovery:=True)
    If mAcceptedRevision <> revision Then Err.Raise 5, "CCoupledNewtonSolver", "Вычислительный возврат изменил принятую материальную историю."
    If Not solved Then GoTo RejectedReturn
    If mRecoveredControlReady Or returnedLambda <= 0# Then GoTo RejectedReturn
    CopyTrial field, actual
    For i = 1 To mGaugeStart - 1: coordinate = coordinate + projection(i) * field(i): Next i
    SetControlledLoad returnedLambda, load
    Dim tolerance As Double: tolerance = EquilibriumTolerance(absoluteTolerance, relativeTolerance)
    If Abs(coordinate - target) * mReferenceModulus > tolerance Then GoTo RejectedReturn
    If Not TrialLambdaUncertainty(absoluteTolerance, relativeTolerance, uncertainty) Then GoTo RejectedReturn
    If returnedLambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty Then GoTo RejectedReturn
    mRecoveredProjection = projection: mRecoveredControlReady = True: mRecoveredControlUncertainty = uncertainty
    mYieldRecoveryTried = True
    TryReturnOriginalMaterialControl = True
    GoTo CountReturn
RejectedReturn:
    DiscardStep
CountReturn:
    phaseDetails = mMeta.DiagnosticDetails
    If attempted Then
        returnIterations = mIterations
        AddSolveEffort initialEffort
    End If
    mMaterialGuideReturnActive = False
    details = details & "; возврат исходной цели из непринятого материального корня: " & CStr(TryReturnOriginalMaterialControl) & _
        "; итерации до / возврата: " & CStr(initialEffort.Iterations) & " / " & CStr(returnIterations) & "; общий лимит: " & CStr(maxIterations)
    If mIterations > maxIterations Then Err.Raise 5, "CCoupledNewtonSolver", "Вычислительный возврат превысил общий лимит итераций."
    If TryReturnOriginalMaterialControl Then
        mMeta.DiagnosticDetails = DiagnosticSummary() & "; исходная цель подтверждена после непринятого материального корня | " & details
    ElseIf Len(phaseDetails) > 0 Then
        details = details & " | Проба возврата: " & phaseDetails
    End If
    If errorNumber <> 0 Then
        On Error GoTo 0
        Err.Raise errorNumber, errorSource, errorText
    End If
    Exit Function
InterruptedReturn:
    errorNumber = Err.Number: errorSource = Err.Source: errorText = Err.Description
    Resume RejectedReturn
End Function

'''
anchor="' Подтверждает материалом разностную пробу в действующем контексте; обычные уравнения и принятая история остаются неизменными."
assert code.count(anchor)==1;code=code.replace(anchor,helper+anchor,1)
assert max(map(len,code.splitlines()))<1024
methods=re.findall(r'(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)',code)
assert len(methods)==len(set(n.casefold() for n in methods))
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='IntegratedMaterialGuideReturnLateH',actual_import_this_pass=['CCoupledNewtonSolver'],
    parent743_late_H_policy_preserved=True,controller_material_and_test_module_exact743=True,
    only_rejected_auxiliary_material_root_can_seed_original_target_return=True,
    original_lambda_force_gauge_steel_direction_checks_not_removed=True,
    shared_primary_budget30_and_global_probe_budget96=True,all_effort_accumulated_with_existing_TSolveEffort=True,
    auxiliary_root_never_accepted=True,return_recursion_guard=True,
    no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('751 INTEGRATED ORIGINAL TARGET RETURN; LATE H743, SAME MATERIAL, FULL COUNTERS AND30/96 BUDGETS')
