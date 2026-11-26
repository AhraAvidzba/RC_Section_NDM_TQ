"""ДЛЯ ТЕСТОВ. Готовит быстрые повторные решения из точной штатной истории без повторения длительного поиска."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path


def TestStage(repo, root, source):
    """История считывается как Double; материал, нагрузка и настройки сверяются без округления перед каждым исследованием."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestActualReverseCapture.py'))['TestStage'](repo, root)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    for index in range(1, 5):
        file = source / f'Native/ActualReverse.Reverse{index}.bin'
        name = f'Input/Reverse{index}.bin'
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        (root / name).write_bytes(file.read_bytes())
        prep['inputs'][name] = sha256(file.read_bytes()).hexdigest().upper()
    name = 'Input/SourceOperators.csv'
    file = source / 'Native/ActualReverse.Accepted.Operators.csv'
    (root / name).write_bytes(file.read_bytes())
    prep['inputs'][name] = sha256(file.read_bytes()).hexdigest().upper()
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig') + '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает бюджет, публикацию и фактическую норму после контрольной пробы.
Friend Function ReadPolishEffortTest() As Variant
    ReadPolishEffortTest = Array(mIterations, mEvaluations, mHybridSteps, mHybridMatrices, mHybridEvaluations, _
        IIf(mCandidateReady, 1, 0), mAcceptedRevision, mLastNorm)
End Function
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'src/Solver/CCoupledSectionAnalysis.cls'
    code = (root / name).read_text('utf-8-sig')
    old = '    Set solver = CreateNewtonSolver(response, fc, ft, models, testSides)'
    assert code.count(old) == 1
    code = code.replace(old, old + '''
    If modTestCoupledElasticState.ReversePolishRequestedTest Then
        modTestCoupledElasticState.TestRunReversePolishModel solver, mOperators, concrete, response, testSides, path, _
            Array(mEb, mPoisson, fc, ft, mYieldRatio * fc, mPeakStrain, mGt, mGc, mLength), _
            Array(INITIAL_STEPS, STEP_ATTEMPTS, maxIterations, STRESS_TOLERANCE, RELATIVE_TOLERANCE, lambdaTolerance, maximumLambda)
        capacity.InitializeUnsupportedConfiguration "ДЛЯ ТЕСТОВ: выполнено исследование точной истории без нового поиска предела."
        Exit Function
    End If''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    assert code.count('Option Explicit\n') == 1
    code = code.replace('Option Explicit\n', '''Option Explicit
' -------------------- ДЛЯ ТЕСТОВ --------------------
Public ReversePolishRequestedTest As Boolean
Private mPolishExecutedTest As Boolean, mPolishRootTest As String
''', 1)
    code += '\n' + (here / 'TestNativeReversePolish.bas.txt').read_text('utf-8-sig')
    code = code.replace('MODE;RejectedSpatialEquilibrium;NotAcceptedContinuation;NotACapacityResult',
        'MODE;DiagnosticReSolve;NotAcceptedContinuation;NotACapacityResult')
    code = code.replace('META;6;0;Отклонённое по направлению равновесие. Не принятый шаг и не предел.',
        'META;4;0;Контрольное повторное равновесие. Контроллер его не принимал; это не предел.')
    code = code.replace('DETAILS;RejectedSpatialEquilibrium;capacity_certified=False', 'DETAILS;DiagnosticReSolve;capacity_certified=False')
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'Run-TestWorkingASD.ps1'
    code = (root / name).read_text('utf-8-sig')
    old = "'Compile','ActualReverse'"
    assert code.count(old) == 1
    code = code.replace(old, "'Compile','ReversePolish','ActualReverse'", 1)
    old = "        if ($Phase -eq 'ActualReverse') {"
    assert code.count(old) == 1
    code = code.replace(old, '''        if ($Phase -eq 'ReversePolish') {
            $arguments+=@('-Macro','modTestCoupledElasticState.RunTestReversePolish','-MacroArgument1',$PSScriptRoot,'-TimeoutSeconds','600')
        } elseif ($Phase -eq 'ActualReverse') {''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeReversePolish.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'ReversePolishExactReplay', 'execution': 'Capacity preparation and test-only replay',
        'new_capacity_search_executed': False, 'strict_absolute_tolerance': 1e-10, 'strict_relative_tolerance': 1e-10,
        'variants': ['OriginalTrial', 'StrictTrial', 'StrictForwardSecant'], 'cases': 4,
        'strict_equilibrium_requires_native_and_independent_original_tolerance_checks': True,
        'numerical_outcomes_not_predetermined': True})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('EXACT HISTORY FOUR CASE THREE VARIANT POLISH STUDY READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
