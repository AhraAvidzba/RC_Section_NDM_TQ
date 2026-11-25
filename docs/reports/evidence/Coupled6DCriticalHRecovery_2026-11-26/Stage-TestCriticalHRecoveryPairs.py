"""ДЛЯ ТЕСТОВ. Ограничивает обязательную H-пару критическим восстановлением и поворотом, сохраняя обычный ход и бюджет известного пика."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ');parent=base/'Cold876VerifiedHCornerPairs20_20261126';root=base/'Cold878CriticalHRecoveryPairs_20261126'
proof=json.loads((parent/'VerifiedHCornerPairsProof.json').read_bytes());assert proof['verified_working_H_pairs']==1
reg=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](base/'Cold877VerifiedHRegressions_20261126','OperationalArcRegressions')[0]
assert reg['passed']==23 and reg['failed']==1 and reg['macros']==1 and reg['own_process_closed']
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
archive=root/'CriticalHCompilationAncestor';archive.mkdir()
for n in ('CompilePoint.log','CompilePointClosure.json'):(root/n).rename(archive/n)
prep=json.loads((root/'Preparation.json').read_bytes());values={}

def replace(code,old,new):
    """Не допускает применения изменения к другому контракту исходников."""
    assert code.count(old)==1,(old[:80],code.count(old));return code.replace(old,new,1)

name='src/Solver/CCoupledNewtonSolver.cls';code=(root/name).read_text('utf-8-sig')
old='        ByVal maxIterations As Long, ByRef probes As Long, Optional ByVal retainEnvelopeControl As Boolean = False) As Boolean'
code=replace(code,old,'''        ByVal maxIterations As Long, ByRef probes As Long, Optional ByVal retainEnvelopeControl As Boolean = False, _
        Optional ByVal verifyCurvePair As Boolean = False) As Boolean''')
start=code.index('    Dim turnRequired As Boolean: turnRequired = Not SpatialIncrementAdvances(');end=code.index('    Dim movement() As Double, coordinate As Double',start)
block=code[start:end];anchor='    Dim fullSteps As Long: fullSteps = mIterations\n';assert block.count(anchor)==1
block=block.replace(anchor,anchor+'    If verifyCurvePair Or turnRequired Then\n',1)+'    End If\n'
code=code[:start]+block+code[end:]
old='''    mEnvelopePairWitnessTest = Array(Array(acceptedLambda, fullLambda, halfLambda, baseline, target, 0.5 * (baseline + target), _
        fullSteps, mIterations - fullSteps, mIterations, tensorError, lambdaError, lambdaBound, jointError, curveScale, Abs(CLng(turnRequired)), _
        mAcceptedRevision, probes), mAccepted, full, half, mTrialResultants, halfActual)'''
new='''    If verifyCurvePair Or turnRequired Then
        mEnvelopePairWitnessTest = Array(Array(acceptedLambda, fullLambda, halfLambda, baseline, target, 0.5 * (baseline + target), _
            fullSteps, mIterations - fullSteps, mIterations, tensorError, lambdaError, lambdaBound, jointError, curveScale, Abs(CLng(turnRequired)), _
            mAcceptedRevision, probes), mAccepted, full, half, mTrialResultants, halfActual)
    End If'''
code=replace(code,old,new)
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledDeformationContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'                trial.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True) Then',
'                trial.Lambda, projection, mAbsTol, mRelTol, mMaxIterations, probes, True, True) Then')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='src/Solver/CCoupledArcContinuation.cls';code=(root/name).read_text('utf-8-sig')
code=replace(code,'        projection, mAbsolute, mRelative, mMaximumIterations \\ 2, probes, False)',
'        projection, mAbsolute, mRelative, mMaximumIterations \\ 2, probes, False, True)')
values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
old='            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False)'
new='            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False, True)'
code=replace(code,old,new);values[name]=code.replace('\n','\r\n').encode('utf-8-sig')
name='tools/build_workbook/Update-IsolatedOptimizerCode.ps1';script=(root/name).read_text('utf-8-sig')
anchor='''        $module=$component.CodeModule
        if ($EntryPointNames.Count -and $file.BaseName -eq 'modTestCoupledElasticState') {'''
patch='''        $module=$component.CodeModule
        if ($file.BaseName -eq 'modTestCoupledElasticState') {
            # ДЛЯ ТЕСТОВ: меняется только один аргумент существующего диагностического вызова; большой модуль не перестраивается.
            $taskOld='            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False)'
            $taskNew='            factor + ratio * (factor - previousFactor), target, returned, projection, absTol, relTol, 30, probes, False, True)'
            if(@(($body -split '\\r?\\n') | Where-Object {$_ -ceq $taskNew}).Count -ne 1){throw 'Prepared critical-H argument line is not unique.'}
            $taskFound=@()
            for($taskLine=1;$taskLine -le $module.CountOfLines;$taskLine++) {
                if(([string]$module.Lines($taskLine,1)).Trim() -ieq $taskOld.Trim()){$taskFound+=,$taskLine}
            }
            if($taskFound.Count -ne 1){throw 'Existing verified-H call line is not unique.'}
            $module.ReplaceLine($taskFound[0],$taskNew)
            Write-Output 'TEST_CRITICAL_H_ARGUMENT_LINE_REPLACED_ONLY'
            $updated += $file.BaseName
            continue
        }
        if ($EntryPointNames.Count -and $file.BaseName -eq 'modTestCoupledElasticState') {'''
script=replace(script,anchor,patch);values[name]=script.replace('\n','\r\n').encode('utf-8-sig')
compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig');compiler=replace(compiler,"-ModuleNames 'modTestCoupledElasticState'","-ModuleNames @('CCoupledNewtonSolver','CCoupledArcContinuation','CCoupledDeformationContinuation','modTestCoupledElasticState')")
values['Compile-TestExternalPoint.ps1']=compiler.replace('\n','\r\n').encode('utf-8-sig')
values.update({'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes(),
 'PriorVerifiedHCornerPairsProof.json':(parent/'VerifiedHCornerPairsProof.json').read_bytes(),
 'NegativeMandatoryAllHRegression.json':(json.dumps(reg,sort_keys=True,indent=2)+'\n').encode(),Path(__file__).name:Path(__file__).read_bytes()})
for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
prep['directed_scope']=dict(case='CriticalHRecoveryPairs',actual_import_this_pass=['CCoupledNewtonSolver','CCoupledArcContinuation','CCoupledDeformationContinuation','modTestCoupledElasticState'],actual_component_count=256,
    H_pair_required_at_recovery_and_turn=True,regular_forward_H_steps_keep_original_control=True,
    known_peak_target_and_global96_primary30_unchanged=True,exact_one_test_argument_line_replaced=True,
    every_critical_pair_keeps_half_midpoint_seed_and_shared_metric=True,force_material_steel_limits_unchanged=True,
    physical_capacity_certified=False,no_release_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print(root,'CRITICAL H-RECOVERY/TURN PAIRS;REGULAR H BUDGET PRESERVED;FOUR IMPORTS;TEST ONLY ONE ARGUMENT LINE PATCH')
