"""ДЛЯ ТЕСТОВ. Проверяет продолжение допустимого спада после численной неудачи проверки суммарного пика, без признания предела."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
root=base/'Cold743UnverifiedCumulativeRecovery_20261125';parent=base/'Cold741OrdinaryFirstComplex_20261125'
proof=json.loads((parent/'OrdinaryFirstComplexProof.json').read_bytes())
assert proof['cold741']['capacity']=='NumFail' and proof['cold741']['native']['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'CumulativeCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
def TestReplace(old,new):
    """Требует единственное совпадение контролируемого участка."""
    global code
    assert code.count(old)==1,old[:100]
    code=code.replace(old,new,1)
TestReplace('    Dim recoveryProjection() As Double\n',
    '    Dim recoveryProjection() As Double\n    Dim cumulativeLeft As TBranchPoint, cumulativeCurrent As TBranchPoint\n')
old='''        ElseIf SelectCumulativePeak(history, trial, left, current) Then
            mSolver.DiscardStep
            If VerifyPeak(history, left, current, trial, reinforced) Then
                If Not reinforced Then Exit Sub
                ResumeAfterLocalPeak history, current, trial
                increment = nominal * 0.25: minimum = increment * MIN_STEP_RATIO
            Else
                StopSearch "Суммарный спад найден, но его максимум не подтверждён повтором плеч." & " Причина: " & mProbeDetails
                Exit Sub
            End If'''
new='''        ElseIf SelectCumulativePeak(history, trial, cumulativeLeft, cumulativeCurrent) Then
            Dim cumulativeTrial As TBranchPoint, cumulativeField() As Double, cumulativeResultants() As Double
            cumulativeTrial = trial ' Уточнение плеч может заменить переданные узлы; исходный сошедшийся спад сохраняется отдельно.
            mSolver.CopyTrial cumulativeField, cumulativeResultants
            mSolver.DiscardStep
            If VerifyPeak(history, cumulativeLeft, cumulativeCurrent, trial, reinforced) Then
                If Not reinforced Then Exit Sub
                left = cumulativeLeft: current = cumulativeCurrent
                ResumeAfterLocalPeak history, current, trial
                increment = nominal * 0.25: minimum = increment * MIN_STEP_RATIO
            ElseIf reinforced And mSolver.UsesSpatialHistory And (mUsesTransverseControl Or mUsesFieldControl) Then
                trial = cumulativeTrial
                Dim cumulativeResumed As Boolean, cumulativeSeed As Variant
                cumulativeSeed = cumulativeField
                If Not ReplayUnverifiedDecline(history, left, current, trial, cumulativeSeed, cumulativeResumed) Then Exit Sub
                smoothSteps = 0
                If cumulativeResumed Then
                    lastProbeFailed = False: lastNoProgress = False: lastLoadUnresolved = False
                Else
                    lastProbeFailed = True: increment = increment * 0.5
                End If
            Else
                StopSearch "Суммарный спад найден, но его максимум не подтверждён повтором плеч." & " Причина: " & mProbeDetails
                Exit Sub
            End If'''
TestReplace(old,new)
old='''                If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
                    StopSearch "Не удалось восстановить историю возрастающей ветви."
                    Exit Sub
                End If
                If reinforced And (mUsesTransverseControl Or mUsesFieldControl) Then
                    Dim repeatSeed As Variant: repeatSeed = declinedField
                    If Probe(trial, trial.Lambda, initialField:=repeatSeed, acceptedLambda:=current.Lambda) Then
                        If UpdateTrialCoordinate(trial, current.Coordinate) And Not BeyondSteelLimit() And trial.Lambda > 0# Then
                            Dim advances As Boolean: advances = True
                            If Not mUsesEnvelopeControl And (mSolver.HasPrincipalBoundaries Or mFieldControlChanges > 0) Then _
                                advances = SpatialTrialAdvances(history, current)
                            If advances Then
                                mUnverifiedDecline = True: mUnverifiedReference = mBestLambda
                                mUnverifiedDips = mUnverifiedDips + 1
                                AcceptPoint history, trial
                                left = current: current = trial
                            Else
                                mSolver.DiscardStep: increment = increment * 0.5
                            End If
                        Else
                            mSolver.DiscardStep: increment = increment * 0.5
                        End If
                    Else
                        lastProbeFailed = True
                        increment = increment * 0.5
                    End If
                Else
                    increment = increment * 0.5
                End If'''
new='''                If reinforced And (mUsesTransverseControl Or mUsesFieldControl) Then
                    Dim repeatSeed As Variant, declineResumed As Boolean
                    repeatSeed = declinedField
                    If Not ReplayUnverifiedDecline(history, left, current, trial, repeatSeed, declineResumed) Then Exit Sub
                    If Not declineResumed Then lastProbeFailed = True: increment = increment * 0.5
                Else
                    If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
                        StopSearch "Не удалось восстановить историю возрастающей ветви."
                        Exit Sub
                    End If
                    increment = increment * 0.5
                End If'''
TestReplace(old,new)
anchor="' Выбирает H сразу при различимом росте в последних двух принятых состояниях; без такой пары сохраняет выбранную геометрическую меру."
helper='''\n' Повторяет сошедшийся спад от исходной принятой истории после неудачи проверки плеч.
' Возвращает признак восстановления истории отдельно от приёмки нового поля; неподтверждённый максимум не становится пределом.
Private Function ReplayUnverifiedDecline(ByVal history As Collection, ByRef left As TBranchPoint, _
        ByRef current As TBranchPoint, ByRef trial As TBranchPoint, ByRef prediction As Variant, ByRef resumed As Boolean) As Boolean
    resumed = False
    If Not mSolver.ReplayAcceptedFields(history, current.HistoryIndex) Then
        StopSearch "Не удалось восстановить принятую историю после численной неудачи проверки максимума."
        Exit Function
    End If
    ReplayUnverifiedDecline = True
    If Not Probe(trial, trial.Lambda, initialField:=prediction, acceptedLambda:=current.Lambda) Then Exit Function
    If Not UpdateTrialCoordinate(trial, current.Coordinate) Or BeyondSteelLimit() Or trial.Lambda <= 0# Then
        mSolver.DiscardStep
        Exit Function
    End If
    If Not mUsesEnvelopeControl And (mSolver.HasPrincipalBoundaries Or mFieldControlChanges > 0) Then
        If Not SpatialTrialAdvances(history, current) Then mSolver.DiscardStep: Exit Function
    End If
    mUnverifiedDecline = True: mUnverifiedReference = mBestLambda
    mUnverifiedDips = mUnverifiedDips + 1
    AcceptPoint history, trial
    left = current: current = trial
    resumed = True
End Function

'''
TestReplace(anchor,helper+anchor)
assert max(map(len,code.splitlines()))<1024
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig').replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledDeformationContinuation'")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
values['Run-TestWorkingGuideRegressions.ps1']=(here/'Run-TestWorkingGuideRegressions.ps1').read_bytes()
values[Path(__file__).name]=Path(__file__).read_bytes()
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='UnverifiedCumulativeDeclineRecovery',actual_import_this_pass=['CCoupledDeformationContinuation'],
    only_controller_candidate=True,material_and_Newton_exact741=True,
    cumulative_selector_uses_separate_peak_anchors=True,original_accepted_tail_preserved_on_shoulder_failure=True,
    same_replay_helper_for_single_and_cumulative_decline=True,unchanged96probe30iteration_budget=True,
    unchanged_physical_and_equilibrium_checks=True,no_working_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('743 CUMULATIVE DECLINE CAN RETRY ORIGINAL VALID FIELD; NO UNVERIFIED PEAK PROMOTION; ORIGINAL BUDGETS')
