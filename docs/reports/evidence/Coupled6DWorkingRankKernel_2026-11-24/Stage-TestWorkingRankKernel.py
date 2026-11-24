"""ДЛЯ ТЕСТОВ. Проверяет фактические четыре текущих исходника вместе с прежними контрольными механизмами продолжения."""
from pathlib import Path
from hashlib import sha256
import json
from runpy import run_path


def TestStage(repo,root,parent):
    """Использует закрытого родителя691 с рабочим контроллером; не переносит постпиковые эксперименты697."""
    here=Path(__file__).resolve().parent
    run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
    archive=root/'WorkingRankCompilationAncestor';archive.mkdir()
    for name in ('CompilePoint.log','CompilePointClosure.json'):(root/name).rename(archive/name)
    prep=json.loads((root/'Preparation.json').read_bytes())
    imports=['CCoupledHybridFactorization','CCoupledHybridSolver','CHybridEquationsTest','modTestCoupledElasticState','modTestCoupledHybrid']
    values={}
    for name in imports:
        path=('src/Solver/' if name.startswith('CCoupled') else 'tests/')+name+('.bas' if name.startswith('mod') else '.cls')
        if name!='modTestCoupledElasticState':values[path]=(repo/path).read_bytes()
    path='tests/modTestCoupledElasticState.bas';code=(root/path).read_text('utf-8-sig')
    assert 'Public Function RunTestTrustStep(' not in code
    payload=(here/'TestLeastSquaresTrustMacro.bas.txt').read_text('utf-8-sig').replace('TryDampedStepTest','TryDampedStep').replace('TryLeastSquaresRadiusTest','TryLeastSquaresRadius')
    values[path]=(code+'\n'+payload).replace('\n','\r\n').encode('utf-8-sig')
    fixture=root.parent/'Cold698RankTrustLinear48_20261124'
    for name in ('Input/TrustSteps.bin','Input/TrustStepCases.json'):values[name]=(fixture/name).read_bytes()
    values['Run-TestRankTrustLinear.ps1']=(fixture/'Run-TestRankTrustLinear.ps1').read_bytes()
    compiler=(here/'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    values['Compile-TestExternalPoint.ps1']=compiler.replace("-ModuleNames 'modTestCoupledElasticState'",'-ModuleNames @('+','.join(repr(x) for x in imports)+')').replace('\n','\r\n').encode('utf-8-sig')
    runner=(here/'Run-TestWorkingGuideRegressions.ps1').read_text('utf-8-sig')
    before='modTestCoupledEnvelopeControl.RunTestCoupledEnvelopeControl,modTestCoupledElasticState.RunTestCoupledRetainedShoulders,modTestCoupledElasticState.RunTestCoupledDiscardedShoulder'
    assert runner.count(before)==1
    after='modTestCoupledHybrid.RunTestCoupledHybrid,modTestCoupledHybrid.RunTestCoupledLeastSquaresTrust,'+before
    runner=runner.replace(before,after).replace('WorkingGuideRegressions.txt','WorkingRank.txt')
    values['Run-TestWorkingRankKernel.ps1']=runner.replace('\n','\r\n').encode('utf-8-sig')
    values['ParentCodeUpdate.json']=(parent/'docs/regression/CodeUpdate.json').read_bytes()
    values['ParentCompileClosure.json']=(parent/'CompilePointClosure.json').read_bytes()
    values[Path(__file__).name]=Path(__file__).read_bytes()
    for name,data in values.items():(root/name).write_bytes(data);prep['inputs'][name]=sha256(data).hexdigest().upper()
    prep['modules']=list(dict.fromkeys(prep['modules']+imports))
    prep['directed_scope'].update(case='WorkingRankKernelTargeted',actual_import_this_pass=imports,
        working_rank_kernel_exact=True,experimental_post_peak_controller_not_used=True,
        macro_count=5,expected_assertions=213,working_square_factorization_without_private_rectangular_extension=True,
        linear48_separate_native_macro=True,physical_capacity_certified=False)
    (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('WORKING KERNEL STAGED; FIVE TARGETED MACROS; ORIGINAL WORKING CONTROLLER')


if __name__=='__main__':
    base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
    TestStage(Path.cwd(),base/'Cold699WorkingRankKernel_20261124',base/'Cold691HybridArithmeticRetry_20261124')
