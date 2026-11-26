"""ДЛЯ ТЕСТОВ. Сравнивает штатную остановку и существующую полевую смену меры до локального пика."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root):
    """Использует фактическую подготовку книги, одни материалы и параметры; меняет только условие активации имеющейся меры."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestClosedBranchPeak.py'))['TestStage'](repo, root, False)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    names = ('src/Solver/CCoupledSectionAnalysis.cls', 'src/Solver/CCoupledDeformationContinuation.cls', 'tests/modTestCoupledElasticState.bas')
    for name in names:
        target = root / 'BaselineSource' / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes((root / name).read_bytes())
    def TestOnce(code, old, new):
        """Не допускает применения инструментирования к неоднозначному месту исходника."""
        assert code.count(old) == 1, old
        return code.replace(old, new)
    name = names[0]
    code = (root / name).read_text('utf-8-sig')
    needle = '''    Dim solver As CCoupledNewtonSolver, branch As New CCoupledLoadContinuation
    Set solver = CreateNewtonSolver(response, fc, ft, models)'''
    replacement = '''    Dim solver As CCoupledNewtonSolver, branch As New CCoupledLoadContinuation
    ' -------------------- ДЛЯ ТЕСТОВ --------------------
    ' Частная копия передаёт неизменённые рабочие данные читающему контрольному сценарию.
    If modTestCoupledElasticState.HardPreparedRequestedTest Then
        Dim testSides As New Collection
        Set solver = CreateNewtonSolver(response, fc, ft, models, testSides)
        modTestCoupledElasticState.RunPreparedHardContinuationTest solver, mOperators, concrete, response, testSides, path, _
            Array(mEb, mPoisson, fc, ft, mYieldRatio * fc, mPeakStrain, mGt, mGc, mLength), _
            INITIAL_STEPS, STEP_ATTEMPTS, initialLambda, maximumLambda, lambdaTolerance, maxIterations, STRESS_TOLERANCE, RELATIVE_TOLERANCE
        capacity.InitializeUnsupportedConfiguration "ДЛЯ ТЕСТОВ: подготовленные данные сохранены в отдельном контрольном файле."
        Exit Function
    End If
    Set solver = CreateNewtonSolver(response, fc, ft, models)'''
    code = TestOnce(code, needle, replacement)
    code = TestOnce(code, '        ByVal models As Object) As CCoupledNewtonSolver',
        '        ByVal models As Object, Optional ByVal testSides As Collection = Nothing) As CCoupledNewtonSolver')
    needle = '    solver.Initialize mOperators, response, mEb, plusResponse, minusResponse'
    code = TestOnce(code, needle, needle + '''
    If Not testSides Is Nothing Then
        testSides.Add minusConcrete: testSides.Add minusResponse
        testSides.Add plusConcrete: testSides.Add plusResponse
    End If''')
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    name = names[1]
    code = (root / name).read_text('utf-8-sig')
    if '    If mUnverifiedDips = 0 Then Exit Function' in code:
        code = TestOnce(code, '    If mUnverifiedDips = 0 Then Exit Function',
            '    If Not modTestCoupledElasticState.EarlyFieldRequestedTest And mUnverifiedDips = 0 Then Exit Function')
        code = TestOnce(code, '    If Not mUnverifiedDecline Or mFollowingLocalPeak Then Exit Function',
            '    If (Not modTestCoupledElasticState.EarlyFieldRequestedTest And Not mUnverifiedDecline) Or mFollowingLocalPeak Then Exit Function')
    else:
        needle = '    If mLastFieldControlRevision = mSolver.AcceptedRevision Then Exit Function'
        code = TestOnce(code, needle, '''    If Not modTestCoupledElasticState.EarlyFieldRequestedTest Then
        If mUnverifiedDips = 0 Or Not mUnverifiedDecline Then Exit Function
    End If
''' + needle)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    name = names[2]
    code = (root / name).read_text('utf-8-sig')
    solver_code = (root / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
    half = re.search(r'Private Const SPATIAL_HALF_STEP As Double = ([\d.E#+-]+)', solver_code)[1]
    globals = '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
Public HardPreparedRequestedTest As Boolean, EarlyFieldRequestedTest As Boolean
Private mHardPreparedPathTest As String, mHardPreparedReportTest As String
Private Const HARD_PREPARED_HALF_TEST As Double = ''' + half + '\n'
    code = TestOnce(code, 'Option Explicit\n', 'Option Explicit\n' + globals)
    code += '\n' + (here / 'TestNativeHardPrepared.bas.txt').read_text('utf-8-sig')
    methods = re.findall(r'(?im)^(?:Public|Private) (?:Function|Sub) (\w+)', code)
    assert len(methods) == len(set(name.casefold() for name in methods))
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    runner = 'Run-TestWorkingASD.ps1'
    code = (root / runner).read_text('utf-8-sig')
    code = TestOnce(code, "'Compile','ClosedPeak'", "'Compile','HardBaseline','HardEarlyField','ClosedPeak'")
    needle = "        if ($Phase -eq 'ClosedPeak') {"
    code = TestOnce(code, needle, '''        if ($Phase -in @('HardBaseline','HardEarlyField')) {
            $caseMode=if ($Phase -eq 'HardBaseline') {'Baseline'} else {'EarlyField'}
            $arguments+=@('-Macro','modTestCoupledElasticState.RunTestHardPrepared','-MacroArgument1',$PSScriptRoot,
                '-MacroArgument2',$caseMode,'-TimeoutSeconds','240')
        } elseif ($Phase -eq 'ClosedPeak') {''')
    (root / runner).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    for relative in (*names, runner):
        prep['inputs'][relative] = sha256((root / relative).read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeHardPrepared.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    for file in (root / 'BaselineSource').rglob('*'):
        if file.is_file(): prep['inputs'][file.relative_to(root).as_posix()] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'] = {'cases': ['Baseline', 'EarlyField'], 'working_sources_changed': False,
        'numerical_tolerances_changed': False, 'global_probe_limit': 96, 'full_acceptance_completed': False}
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('PREPARED HARDSHEAR BASELINE AND EARLY FIELD CONTROL FIXTURE READY', flush=True)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve())
