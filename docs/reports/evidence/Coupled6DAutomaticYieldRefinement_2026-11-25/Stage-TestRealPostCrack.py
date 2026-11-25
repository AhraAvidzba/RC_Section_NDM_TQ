"""ДЛЯ ТЕСТОВ. Собирает реальный CDP+сталь с независимой одноосной посттрещинной огибающей."""
from argparse import ArgumentParser
from hashlib import sha256
import json,re
from pathlib import Path
from runpy import run_path


def TestStage(root,parent,adapt=False,singular=False,shoulder=False,capacity=False,balance=False):
    """Использует закрытую частную книгу; импортирует только тест, физический рабочий код сохраняется."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    archive=root/'PostCrackCompilationAncestor';archive.mkdir()
    for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
    prep=json.loads((root/'Preparation.json').read_bytes())
    name='tests/modTestCoupledElasticState.bas';code=(root/name).read_text('utf-8-sig')
    payload=(here/'TestPostCrackTensionControl.bas.txt').read_text('utf-8-sig')
    assert 'Function RunTestCoupledPostCrackTension(' not in code
    code+='\n'+payload
    assert code.count('DETAILS;AcceptedAutomaticPath;capacity_certified=False')==1
    code=code.replace('DETAILS;AcceptedAutomaticPath;capacity_certified=False','DETAILS;AcceptedPostCrackTestPath;capacity_certified=False',1)
    assert code.count('MODE;AcceptedAutomaticPath;NotACapacityResult')==1
    code=code.replace('MODE;AcceptedAutomaticPath;NotACapacityResult','MODE;AcceptedPostCrackTestPath;NotACapacityResult',1)
    assert code.count('META;4;0;Принятая штатным контроллером история. Это не предел.')==1
    code=code.replace('META;4;0;Принятая штатным контроллером история. Это не предел.','META;4;0;Принятая контрольная последовательность. Это не предел.',1)
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    assert compiler.count("-ModuleNames 'modTestCoupledElasticState'")==1
    runner=(here/'Run-TestWorkingGuideRegressions.ps1').read_text('utf-8-sig')
    macros="modTestCoupledEnvelopeControl.RunTestCoupledEnvelopeControl,modTestCoupledElasticState.RunTestCoupledRetainedShoulders,modTestCoupledElasticState.RunTestCoupledDiscardedShoulder"
    assert runner.count(macros)==1
    runner=runner.replace(macros,'modTestCoupledElasticState.RunTestCoupledPostCrackTension').replace('WorkingGuideRegressions.txt','PostCrackTension.txt')
    runner=runner.replace('-Macro $taskMacros -TimeoutSeconds','-Macro $taskMacros -MacroArgument1 $Root -TimeoutSeconds')
    if capacity:runner=runner.replace('modTestCoupledElasticState.RunTestCoupledPostCrackTension','modTestCoupledElasticState.RunTestCoupledPostCrackCapacity')
    values={name:code.replace('\n','\r\n').encode('utf-8-sig'),
      'Compile-TestExternalPoint.ps1':compiler.replace('\n','\r\n').encode('utf-8-sig'),
      'Run-TestRealPostCrack.ps1':runner.replace('\n','\r\n').encode('utf-8-sig'),
      'ParentCodeUpdate.json':(parent/'docs/regression/CodeUpdate.json').read_bytes(),
      'ParentCompileClosure.json':(parent/'CompilePointClosure.json').read_bytes()}
    imports=['modTestCoupledElasticState']
    if adapt:
        name='src/Solver/CCoupledDeformationContinuation.cls';controller=(root/name).read_text('utf-8-sig')
        substitutions={
          'ElseIf mFollowingLocalPeak And Len(mPostPeakControlName) > 0 Then':'ElseIf mLocalPeak > 0# Then',
          'smoothTrial = FitsPostPeakSecant(history, left, current, trial)':'smoothTrial = FitsMaterialGrowthSecant(history, left, current, trial)',
          '(IsMaterialGrowthControl() Or (mFollowingLocalPeak And Len(mPostPeakControlName) > 0))':'(IsMaterialGrowthControl() Or mLocalPeak > 0#)',
          'If smoothTrial And (mFollowingLocalPeak Or IsMaterialGrowthControl()) Then':'If smoothTrial And (mLocalPeak > 0# Or IsMaterialGrowthControl()) Then',
          '            If IsMaterialGrowthControl() Then\n                increment = acceptedIncrement':'            If IsMaterialGrowthControl() Or mLocalPeak > 0# Then\n                increment = acceptedIncrement',
          '            If IsMaterialGrowthControl() And smoothSteps >= 2 Then':'            If (IsMaterialGrowthControl() Or mLocalPeak > 0#) And smoothSteps >= 2 Then'}
        for before,after in substitutions.items():
            assert controller.count(before)==1,before
            controller=controller.replace(before,after,1)
        values[name]=controller.replace('\n','\r\n').encode('utf-8-sig')
        imports=['CCoupledDeformationContinuation','modTestCoupledElasticState']
        values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')').replace('\n','\r\n').encode('utf-8-sig')
    if shoulder:
        assert adapt
        name='src/Solver/CCoupledDeformationContinuation.cls'
        method,=re.findall(r'(?ms)^Private Function PrepareSecantSeed\(.*?^End Function',controller)
        updated=method.replace('ByVal target As Double, ByRef seed As Variant, ByRef factor As Double) As Boolean',
          'ByVal target As Double, ByRef seed As Variant, ByRef factor As Double, Optional ByVal predictorTrace As CCoupledLoadTrace = Nothing) As Boolean')
        assert updated!=method
        anchor='    Dim previousLambda As Double: previousLambda = left.Lambda\n';assert updated.count(anchor)==1
        updated=updated.replace(anchor,'    Dim selectedTrace As CCoupledLoadTrace: Set selectedTrace = predictorTrace\n    If selectedTrace Is Nothing Then Set selectedTrace = mTrace\n'+anchor,1)
        assert updated.count('If Not mTrace Is Nothing Then')==1 and updated.count('mTrace.ReadPoint')==1
        updated=updated.replace('If Not mTrace Is Nothing Then','If Not selectedTrace Is Nothing Then').replace('mTrace.ReadPoint','selectedTrace.ReadPoint')
        controller=controller.replace(method,updated,1)
        method,=re.findall(r'(?ms)^Private Function ReplayShoulders\(.*?^End Function',controller);updated=method
        anchor='    Dim field() As Double, resultants() As Double\n';assert updated.count(anchor)==1
        updated=updated.replace(anchor,anchor+'''    Dim seedFields As New Collection, seedTrace As CCoupledLoadTrace, seedSource As CCoupledLoadTrace
    Dim seedLeft As TBranchPoint, seedCurrent As TBranchPoint, prediction As Variant, predictedLambda As Double
    Dim seedIndex As Long, seedFraction As Double, probed As Boolean
    For seedIndex = 1 To left.HistoryIndex: field = history.Item(seedIndex): seedFields.Add field: Next seedIndex
    Set seedSource = sourceTrace: If seedSource Is Nothing Then Set seedSource = mTrace
    If Not seedSource Is Nothing Then Set seedTrace = seedSource.CopyPrefix(left.HistoryIndex)
''',1)
        anchor='            If Not Probe(point, previousLambda, keepYieldControl:=keepRight, acceptedLambda:=previousLambda) Then Exit Function\n';assert updated.count(anchor)==1
        updated=updated.replace(anchor,'''            seedCurrent = samples(index): seedLeft.HistoryIndex = 0
            If index > 0 Then
                seedLeft = samples(index - 1)
            ElseIf left.HistoryIndex > 1 And Not seedTrace Is Nothing Then
                seedLeft.HistoryIndex = left.HistoryIndex - 1
                seedTrace.ReadPoint seedLeft.HistoryIndex, seedFraction, seedLeft.Lambda
                If seedFraction <> 1# Then seedLeft.HistoryIndex = 0
            End If
            If PrepareSecantSeed(seedFields, seedLeft, seedCurrent, point.Coordinate, prediction, predictedLambda, seedTrace) Then
                probed = Probe(point, predictedLambda, keepYieldControl:=keepRight, initialField:=prediction, acceptedLambda:=previousLambda)
            Else
                probed = Probe(point, previousLambda, keepYieldControl:=keepRight, acceptedLambda:=previousLambda)
            End If
            If Not probed Then Exit Function
''',1)
        anchor='            If saveFields Then\n                mSolver.CopyTrial field, resultants\n';assert updated.count(anchor)==1
        updated=updated.replace(anchor,'            mSolver.CopyTrial field, resultants\n            If saveFields Then\n',1)
        anchor='            mSolver.AcceptStep\n';assert updated.count(anchor)==1
        updated=updated.replace(anchor,anchor+'            seedFields.Add field\n            If Not seedTrace Is Nothing Then seedTrace.Add 1#, point.Lambda\n',1)
        controller=controller.replace(method,updated,1)
        values[name]=controller.replace('\n','\r\n').encode('utf-8-sig')
    if singular:
        name='src/Solver/CCoupledNewtonSolver.cls';newton=(root/name).read_text('utf-8-sig')
        before='            RecordFailure mSystem.FailureCode, "Касательная связанного сечения вырождена или не решается с заданной точностью."\n'
        assert newton.count(before)==1
        after=before+'''            If allowHybridRecovery Then SolveStep = TryHybridControlledField(current, currentLambda, load, _
                absoluteTolerance, relativeTolerance, maxIterations, conditionHybridRecovery, controlPrediction, initialLambda, referenceLambda)
'''
        newton=newton.replace(before,after,1)
        if balance:
            call='''            Dim balanceUnavailable As Boolean, balanceDetails As String
            If TryBalanceControlledLoad(current, currentLambda, load, absoluteTolerance, relativeTolerance, balanceUnavailable, balanceDetails) Then
                SolveStep = True: Exit Function
            End If
            If balanceUnavailable Then Exit Function
'''
            assert newton.count(before)==1
            newton=newton.replace(before,call+before,1)
            newton=newton.replace(before,before+'            If Len(balanceDetails) > 0 Then mMeta.DiagnosticDetails = mMeta.DiagnosticDetails & " | " & balanceDetails\n',1)
            newton+='\n'+(here/'CandidateControlledLoadBalance.cls.txt').read_text('utf-8-sig')
            values['CandidateControlledLoadBalance.cls.txt']=(here/'CandidateControlledLoadBalance.cls.txt').read_bytes()
        anchor='    mHybridEvaluations = mHybridEvaluations + hybrid.FunctionCalls\n'
        assert newton.count(anchor)==1
        newton=newton.replace(anchor,anchor+'''    original.DiagnosticDetails = original.DiagnosticDetails & " | Гибридный резерв: код остановки " & CStr(hybrid.StopCode) & _
        "; принятых поправок " & CStr(hybrid.AcceptedSteps) & "; лучшая невязка " & CStr(hybrid.BestResidual)
''',1)
        if 'Property Get ArithmeticDiagnostic()' in (root/'src/Solver/CCoupledHybridSolver.cls').read_text('utf-8-sig'):
            marker='        "; принятых поправок " & CStr(hybrid.AcceptedSteps) & "; лучшая невязка " & CStr(hybrid.BestResidual)\n'
            assert newton.count(marker)==1
            newton=newton.replace(marker,marker+'    If Len(hybrid.ArithmeticDiagnostic) > 0 Then original.DiagnosticDetails = original.DiagnosticDetails & " | " & hybrid.ArithmeticDiagnostic\n',1)
        values[name]=newton.replace('\n','\r\n').encode('utf-8-sig')
        imports=['CCoupledDeformationContinuation','CCoupledNewtonSolver','modTestCoupledElasticState'] if adapt else ['CCoupledNewtonSolver','modTestCoupledElasticState']
        values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')').replace('\n','\r\n').encode('utf-8-sig')
    for item in (Path(__file__).name,'TestPostCrackTensionControl.bas.txt'):
        values[item]=(here/item).read_bytes()
    for name,raw in values.items():
        (root/name).write_bytes(raw);prep['inputs'][name]=sha256(raw).hexdigest().upper()
    prep['directed_scope'].update(case='RealCDPSteelPostCrackAnalyticalTension',
      actual_import_this_pass=imports,full_cold_from_zero=True,post_peak_adaptation_candidate=adapt,singular_tangent_QR_candidate=singular,
      shoulder_secant_predictor_candidate=shoulder,capacity_only=capacity,full_prescribed_curve_not_run=capacity,
      controlled_load_only_balance_candidate=balance,
      input146_not_read=True,only_axial_case_does_not_replace_HardShear=True,
      actual_working_concrete_and_steel_models=True,analytical_first_peak_N=72144.,
      analytical_later_peak_N=172967.56279474715,all_physical_inputs_inside_test=True,
      ordinary_probe_limit=96,primary_iteration_budget=30,physical_capacity_certified=False)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    assert max(map(len,code.splitlines()))<1024
    print('REAL CDP+STEEL POSTCRACK TEST STAGED; IMPORTS',len(imports),'CAPACITY_ONLY',capacity,'SHOULDER_PREDICTOR',shoulder)


def TestFinalize(repo,root):
    """Проверяет все255 компонентов и254 неимпортированных компонента относительно закрытого родителя."""
    source=Path(__file__).resolve().parent/'Finalize-TestStockYield.py'
    imports=json.loads((root/'Preparation.json').read_bytes())['directed_scope']['actual_import_this_pass']
    code=source.read_text('utf-8-sig').replace('len(new)==253','len(new)==255').replace('len(unchanged)==252','len(unchanged)=='+str(255-len(imports)))
    before="compilation['Modules']==['CCoupledDeformationContinuation']";assert code.count(before)==2
    code=code.replace(before,"compilation['Modules']=="+repr(imports))
    code=code.replace("audit['imported_this_pass']=['CCoupledDeformationContinuation']","audit['imported_this_pass']=compilation['Modules']",1)
    code=code.replace('actual_other_252_components_exact_except_identifier_case','actual_other_'+str(255-len(imports))+'_components_exact_except_identifier_case')
    namespace={'__file__':str(source),'__name__':'real_postcrack_source_test'}
    exec(compile(code,str(source),'exec'),namespace);namespace['TestFinalize'](repo,root)


if __name__=='__main__':
    p=ArgumentParser(description=__doc__);p.add_argument('--root',type=Path,required=True)
    p.add_argument('--parent',type=Path);p.add_argument('--repo',type=Path,default=Path.cwd());p.add_argument('--finalize',action='store_true');p.add_argument('--adapt-post-peak',action='store_true');p.add_argument('--singular-reserve',action='store_true')
    p.add_argument('--shoulder-predictor',action='store_true');p.add_argument('--capacity-only',action='store_true')
    p.add_argument('--load-balance',action='store_true')
    a=p.parse_args()
    if a.finalize:TestFinalize(a.repo.resolve(),a.root.resolve())
    else:TestStage(a.root.resolve(),a.parent.resolve(),a.adapt_post_peak,a.singular_reserve,a.shoulder_predictor,a.capacity_only,a.load_balance)
