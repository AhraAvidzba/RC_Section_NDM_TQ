"""ДЛЯ ТЕСТОВ. Включает имеющийся полный LM только для явно выбранного постпикового резерва CDP."""
from argparse import ArgumentParser
from hashlib import sha256
import json,re
from pathlib import Path
from runpy import run_path


def TestStage(repo,root,parent):
    """Изолирует численные изменения, сохраняя материал, допуски и бюджеты; оба нативных запуска раздельны."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestRealPostCrack.py'))['TestStage'](root,parent,True,True,True,True,False)
    prep=json.loads((root/'Preparation.json').read_bytes())
    values={}
    path='src/Solver/CCoupledDeformationContinuation.cls';controller=(root/path).read_text('utf-8-sig')
    anchor='    Dim solved As Boolean, recoverSameControl As Boolean, recoveryReference As Double\n'
    assert controller.count(anchor)==1
    controller=controller.replace(anchor,anchor+'    Dim postPeakRecovery As Boolean: postPeakRecovery = (mLocalPeak > 0# And Not mSolver.HasPrincipalBoundaries)\n',1)
    pairs={
      '(mBranchTangentControl Or robustDeparture), recoverSameControl':'(mBranchTangentControl Or robustDeparture Or postPeakRecovery), recoverSameControl',
      'allowHybridRecovery:=allowHybridRecovery, _':'allowHybridRecovery:=(allowHybridRecovery Or postPeakRecovery), _',
      'conditionHybridRecovery:=robustDeparture,':'conditionHybridRecovery:=(robustDeparture Or postPeakRecovery),',
      'allowRefinedRestart:=(IsMaterialGrowthControl() Or robustDeparture),':'allowRefinedRestart:=(IsMaterialGrowthControl() Or robustDeparture Or postPeakRecovery),'}
    for before,after in pairs.items():
        assert controller.count(before)==1,before
        controller=controller.replace(before,after,1)
    values[path]=controller
    path='src/Solver/CCoupledNewtonSolver.cls';newton=(root/path).read_text('utf-8-sig')
    before='''    converged = hybrid.TrySolve(equations, coordinates, ordinaryCalls, ordinarySteps, absoluteTolerance, _
        0.00000000001, 0.2, solution, relativeDifferenceStep:=HYBRID_DIFFERENCE_STEP, _
        automaticScaling:=conditionRecovery, freshJacobian:=conditionRecovery, captureDifferenceBracket:=(reserved > 0), _
        differenceObserver:=observer, completeNewtonLineSearch:=(reserved > 0))'''
    assert newton.count(before)==1
    newton=newton.replace(before,'''    If conditionRecovery And referenceLambda = 0# Then
        converged = hybrid.TrySolveLeastSquaresTrust(equations, coordinates, ordinaryCalls, ordinarySteps, _
            absoluteTolerance, solution, HYBRID_DIFFERENCE_STEP)
    Else
'''+before+'''
    End If''',1)
    values[path]=newton
    path='src/Solver/CCoupledHybridSolver.cls';hybrid=(repo/path).read_text('utf-8-sig')
    method,=re.findall(r'(?ms)^Friend Function TrySolveLeastSquaresTrust\(.*?^End Function',hybrid)
    changed=method.replace('mBestNorm = 1E+300:','mArithmeticDiagnostic = vbNullString\n    mBestNorm = 1E+300:',1)
    changed=changed.replace('Dim originalResidual() As Double','Dim originalResidual() As Double, trialUnavailable As Boolean',1)
    changed=changed.replace('valid = ReadResidual(equations, trial, nextResidual, unused, count)','valid = ReadTrialResidual(equations, trial, nextResidual, unused, count, trialUnavailable)',1)
    changed=changed.replace('ElseIf mStopCode = -1 Then','ElseIf trialUnavailable Then',1)
    assert changed!=method
    values[path]=hybrid.replace(method,changed,1)
    path='tests/CHybridEquationsTest.cls';adapter=(repo/path).read_text('utf-8-sig')
    anchor='    Select Case mMode\n';assert adapter.count(anchor)==1
    adapter=adapter.replace(anchor,anchor+'''        Case "LMUnavailable", "LMOverflow", "LMDivision", "LMCancel", "LMWrongDimensions"
            If coordinates(1) < 0.5 Then
                mUnavailableCallsTest = mUnavailableCallsTest + 1
                If mMode = "LMOverflow" Then Err.Raise 6, "CHybridEquationsTest", "Тестовое переполнение LM-пробы."
                If mMode = "LMDivision" Then Err.Raise 11, "CHybridEquationsTest", "Тестовое деление на ноль LM-пробы."
                If mMode = "LMCancel" Then Err.Raise 18, "CHybridEquationsTest", "Тестовая отмена LM-пробы."
                If mMode = "LMWrongDimensions" Then
                    ReDim residual(1 To count + 1)
                    ICoupledEquationSystem_TryEvaluate = True
                End If
                Exit Function
            End If
            residual(1) = Sqr(coordinates(1)) - 1#
            If withJacobian Then jacobian(1, 1) = 0.5 / Sqr(coordinates(1))
''',1)
    values[path]=adapter
    path='tests/modTestCoupledHybrid.bas'
    values[path]=(repo/path).read_text('utf-8-sig')+'\n'+(here/'TestLeastSquaresTrust.bas.txt').read_text('utf-8-sig')
    imports=sorted(['CCoupledDeformationContinuation','CCoupledNewtonSolver','CCoupledHybridSolver','CHybridEquationsTest','modTestCoupledHybrid','modTestCoupledElasticState'])
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')')
    runner=(root/'Run-TestRealPostCrack.ps1').read_text('utf-8-sig')
    runner=runner.replace('modTestCoupledElasticState.RunTestCoupledPostCrackCapacity','modTestCoupledHybrid.RunTestCoupledLeastSquaresTrust').replace('PostCrackTension.txt','LeastSquaresTrust.txt')
    runner=runner.replace(' -MacroArgument1 $Root','')
    values['Run-TestLeastSquaresTrust.ps1']=runner
    for path,code in values.items():
        raw=code.replace('\n','\r\n').encode('utf-8-sig')
        (root/path).write_bytes(raw);prep['inputs'][path]=sha256(raw).hexdigest().upper()
    for name in (Path(__file__).name,'TestLeastSquaresTrust.bas.txt'):
        raw=(here/name).read_bytes();(root/name).write_bytes(raw);prep['inputs'][name]=sha256(raw).hexdigest().upper()
    prep['modules']=list(dict.fromkeys(prep['modules']+imports))
    prep['directed_scope'].update(actual_import_this_pass=imports,case='PostCrackFullRegularizedTrust',
       LM_uses_existing_complete_equations=True,post_peak_only=True,positive_reference_recovery_unchanged=True,
       no_material_changes=True,original_budgets_and_tolerances=True)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('POSTCRACK FULL LM STAGED; SIX COMPONENT IMPORTS; MATERIAL AND BUDGETS UNCHANGED')


if __name__=='__main__':
    p=ArgumentParser(description=__doc__)
    for name in ('root','parent'):p.add_argument('--'+name,type=Path,required=True)
    p.add_argument('--repo',type=Path,default=Path.cwd());a=p.parse_args()
    TestStage(a.repo.resolve(),a.root.resolve(),a.parent.resolve())
