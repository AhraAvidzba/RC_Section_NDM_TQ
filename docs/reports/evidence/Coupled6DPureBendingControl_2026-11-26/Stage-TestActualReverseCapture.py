"""ДЛЯ ТЕСТОВ. Захватывает поля именно штатного Both/Excel, не заменяя вычислительный путь тестовым контроллером."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path


def TestStage(repo, root):
    """Читающие снимки снимаются во время работы; повтор материала и выгрузка выполняются после штатной записи Results."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestReverseCapture.py'))['TestStage'](repo, root)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Solver/CCoupledDeformationContinuation.cls'
    code = (repo / name).read_text('utf-8-sig')
    old = '    If SpatialTrialAdvances Then Exit Function'
    assert code.count(old) == 1
    code = code.replace(old, old + '''
    modTestCoupledElasticState.TestRememberSpatialReverse history, current.HistoryIndex, mTrace, previous, accepted, trial, resultants, _
        mProjection, mSolver.ReadReverseControlTest, signedAdvance, referenceIncrement, mSolver.AcceptedRevision, mFieldControlChanges''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'src/Solver/CCoupledSectionAnalysis.cls'
    code = (repo / name).read_text('utf-8-sig')
    old = '''    Dim solver As CCoupledNewtonSolver, branch As New CCoupledLoadContinuation
    Set solver = CreateNewtonSolver(response, fc, ft, models)'''
    assert code.count(old) == 1
    code = code.replace(old, '''    Dim solver As CCoupledNewtonSolver, branch As New CCoupledLoadContinuation
    Dim testSides As New Collection
    Set solver = CreateNewtonSolver(response, fc, ft, models, testSides)''', 1)
    old = '        ByVal models As Object) As CCoupledNewtonSolver'
    assert code.count(old) == 1
    code = code.replace(old, '        ByVal models As Object, Optional ByVal testSides As Collection = Nothing) As CCoupledNewtonSolver', 1)
    old = '    solver.Initialize mOperators, response, mEb, plusResponse, minusResponse'
    assert code.count(old) == 1
    code = code.replace(old, old + '''
    If Not testSides Is Nothing Then
        testSides.Add minusConcrete: testSides.Add minusResponse
        testSides.Add plusConcrete: testSides.Add plusResponse
    End If''', 1)
    old = '            Set acceptedTrace = controlled.SavedTrace'
    assert code.count(old) == 1
    code = code.replace(old, old + '''
            modTestCoupledElasticState.TestHoldActualReverse solver, mOperators, concrete, response, testSides, path, _
                controlled.AcceptedFields, acceptedTrace, Array(mEb, mPoisson, fc, ft, mYieldRatio * fc, mPeakStrain, mGt, mGc, mLength), _
                Array(INITIAL_STEPS, STEP_ATTEMPTS, maxIterations, STRESS_TOLERANCE, RELATIVE_TOLERANCE, lambdaTolerance, maximumLambda)''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    globals = '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
Private mActualReverseRequestedTest As Boolean
Private mActualSolverTest As CCoupledNewtonSolver, mActualOperatorsTest As CCoupledElasticState
Private mActualConcreteTest As CCoupledConcreteResponse, mActualReinforcedTest As CCoupledReinforcedResponse
Private mActualSidesTest As Collection, mActualFieldsTest As Collection, mActualTraceTest As CCoupledLoadTrace
Private mActualCalibrationTest As Variant, mActualOptionsTest As Variant
Private mActualOffsetTest() As Double, mActualDirectionTest() As Double
'''
    assert code.count('Option Explicit\n') == 1
    code = code.replace('Option Explicit\n', 'Option Explicit\n' + globals, 1)
    old = '        ByVal calibration As Variant, ByVal options As Variant, ByRef offset() As Double, ByRef direction() As Double)'
    assert code.count(old) == 1
    code = code.replace(old, '''        ByVal calibration As Variant, ByVal options As Variant, ByRef offset() As Double, ByRef direction() As Double, _
        Optional ByVal acceptedFields As Collection = Nothing, Optional ByVal acceptedTrace As CCoupledLoadTrace = Nothing)''', 1)
    old = '    For index = 1 To mReverseSnapshotsTest.Count'
    assert code.count(old) == 1
    code = code.replace(old, '''    If Not acceptedFields Is Nothing Then
        Set mReverseFieldsTest = acceptedFields: Set mReverseTraceTest = acceptedTrace.CopyPrefix(acceptedTrace.Count)
        TestExportReverseCandidate 0
    End If
''' + old, 1)
    old = '    Dim path As String: path = mHardPreparedPathTest & ".Reverse" & CStr(caseIndex)'
    assert code.count(old) == 1
    code = code.replace(old, old + '\n    If caseIndex = 0 Then path = mHardPreparedPathTest & ".Accepted"', 1)
    replacements = {
        'MODE;RejectedSpatialEquilibrium;NotAcceptedContinuation;NotACapacityResult': 'MODE;AcceptedAutomaticPath;NotACapacityResult',
        'META;6;0;Отклонённое по направлению равновесие. Не принятый шаг и не предел.': 'META;4;0;Принятая штатным контроллером история. Это не предел.',
        'DETAILS;RejectedSpatialEquilibrium;capacity_certified=False': 'DETAILS;AcceptedAutomaticPath;capacity_certified=False'}
    for old, new in replacements.items():
        line = '    Print #handle, "' + old + '"'
        assert code.count(line) == 1
        code = code.replace(line, '    If caseIndex = 0 Then\n        Print #handle, "' + new + '"\n    Else\n    ' + line + '\n    End If', 1)
    code += '\n' + (here / 'TestNativeActualReverse.bas.txt').read_text('utf-8-sig')
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'Run-TestWorkingASD.ps1'
    code = (root / name).read_text('utf-8-sig')
    old = "'Compile','HardBaseline','HardEarlyField'"
    assert code.count(old) == 1
    code = code.replace(old, "'Compile','ActualReverse','HardBaseline','HardEarlyField'", 1)
    old = "        if ($Phase -in @('HardBaseline','HardEarlyField')) {"
    assert code.count(old) == 1
    code = code.replace(old, '''        if ($Phase -eq 'ActualReverse') {
            $arguments+=@('-Macro','modTestCoupledElasticState.RunTestActualReverseCapture','-MacroArgument1',$PSScriptRoot,
                '-TimeoutSeconds','600','-VerifyResultsReopen')
        } elseif ($Phase -in @('HardBaseline','HardEarlyField')) {''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeActualReverse.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'ActualAutomaticReverseCapture', 'execution': 'Both/Excel',
        'test_controller_reimplementation_used': False, 'native_results_saved_reopened': True,
        'material_replay_only_after_actual_results_written': True})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('ACTUAL BOTH EXCEL ACCEPTED HISTORY AND REJECTED EQUILIBRIA CAPTURE READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve())
