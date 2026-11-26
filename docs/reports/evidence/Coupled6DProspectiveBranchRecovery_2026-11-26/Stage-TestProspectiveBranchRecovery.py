"""ДЛЯ ТЕСТОВ. Проверяет рабочую поправку по касательной за материальной границей с исходными критериями и бюджетом."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold911NewtonRejectedCapture_20261126';root=base/'Cold915ProspectiveBranchRecovery_20261126'
proof=json.loads((base/'Cold914ProspectiveBranchMatrix_20261126/ProspectiveBranchMatrixProof.json').read_bytes())
assert proof['base_residual_maximum_difference_from_actual_capture_MPa']==0
assert max(p['norm_reduction'] for item in proof['records'] for p in item.get('probes',[]) if p['available'])>.8
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'ProspectiveRecoveryCompilationAncestor';archive.mkdir()
for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
prep=json.loads((root/'Preparation.json').read_bytes());values={}
name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
def TestReplace(old,new):
    """Вносит единственную проверенную замену, не меняя соседний алгоритм."""
    global code
    assert code.count(old)==1,old[:100];code=code.replace(old,new,1)
anchor='Private mSystem As CCoupledLinearSystem'
TestReplace(anchor,anchor+'\nPrivate mProspectiveAttempts As Long, mProspectiveAccepted As Long')
anchor='    mIterations = 0: mEvaluations = 0: mMaterialPoint = 0: mGradientBuilds = 0'
TestReplace(anchor,anchor+'\n    mProspectiveAttempts = 0: mProspectiveAccepted = 0')
anchor='''                current, mAccepted, mResidual, Array(currentLambda, norm, tolerance, bestTrialNorm, Abs(CLng(hasValidTrial)), minimumFraction)
            If useSteelCorrection'''
TestReplace(anchor,'''                current, mAccepted, mResidual, Array(currentLambda, norm, tolerance, bestTrialNorm, Abs(CLng(hasValidTrial)), minimumFraction)
            If mControlActive And mUseSpatialHistory And hasValidTrial And mFailureCode = sfcNone And _
                    SupportsHybridRecovery And mIterations + 2 < maxIterations Then
                mIterations = mIterations + 2 ' Отклонённое направление и новая поправка расходуют тот же общий бюджет.
                mProspectiveAttempts = mProspectiveAttempts + 1
                If TryProspectiveBranchStep(current, currentLambda, load, norm, absoluteTolerance, relativeTolerance) Then
                    mProspectiveAccepted = mProspectiveAccepted + 1
                    GoTo RecomputeTangent
                End If
            End If
            If useSteelCorrection''')
anchor='    If mHybridEvaluations > 0 Then DiagnosticSummary = DiagnosticSummary & "; гибридных поправок: " & CStr(mHybridSteps) & _'
TestReplace(anchor,'''    If mProspectiveAttempts > 0 Then DiagnosticSummary = DiagnosticSummary & _
        "; проб касательной за сменой ветви: " & CStr(mProspectiveAttempts) & "; принятых поправок: " & CStr(mProspectiveAccepted)
'''+anchor)
code+='\n\n'+(here/'CandidateProspectiveBranchStep.cls.txt').read_text('utf-8-sig')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'","-ModuleNames 'CCoupledNewtonSolver'",1).replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
 'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='ProspectiveBranchModifiedNewtonRecovery',actual_import_this_pass=['CCoupledNewtonSolver'],
 actual_component_count=256,all_material_and_force_acceptance_unchanged=True,scalar_control_and_load_path_unchanged=True,
 every_recovery_attempt_consumes_two_primary_iterations=True,primary30_global96_unchanged=True,
 no_material_history_acceptance_in_recovery=True,failed_recovery_restores_original_context=True,
 no_working_or_release_promotion=True,physical_capacity_certified=False)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'PROSPECTIVE BRANCH RECOVERY;SAME FORCE/CONTROL/30/96;NO HISTORY ACCEPTANCE')
