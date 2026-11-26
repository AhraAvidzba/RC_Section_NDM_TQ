"""ДЛЯ ТЕСТОВ. Выгружает весь рабочий путь и точные отклонённые равновесия после 26-го состояния."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root):
    """Сохраняет действующую подготовку, шаг z/L=0,001 и все условия приёмки; добавляет только читающие тестовые вызовы."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestHardPrepared.py'))['TestStage'](repo, root)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    for name in ('src/Interfaces/ICoupledEquationSystem.cls', 'src/Solver/CCoupledHybridFactorization.cls',
                 'src/Solver/CCoupledHybridSolver.cls', 'tests/CHybridEquationsTest.cls', 'tests/modTestCoupledHybrid.bas'):
        file = repo / name
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        (root / name).write_bytes(file.read_bytes())
        prep['inputs'][name] = sha256(file.read_bytes()).hexdigest().upper()
        prep['modules'].append(file.stem)
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig') + '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает управление существующей равновесной пробы без повторного вычисления материала.
Friend Function ReadReverseControlTest() As Variant
    RequireCandidate
    ReadReverseControlTest = Array(mTrialLambda, mTargetControl, mCurrentControlLambda, mArcLoadProjection)
End Function
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'src/Solver/CCoupledDeformationContinuation.cls'
    code = (root / name).read_text('utf-8-sig')
    old = '    If SpatialTrialAdvances Then Exit Function'
    assert code.count(old) == 1
    code = code.replace(old, old + '''
    modTestCoupledElasticState.TestRememberSpatialReverse history, current.HistoryIndex, mTrace, previous, accepted, trial, resultants, _
        mProjection, mSolver.ReadReverseControlTest, signedAdvance, referenceIncrement, mSolver.AcceptedRevision, mFieldControlChanges''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    globals = '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
Private mReverseSnapshotsTest As Collection
Private mReverseSolverTest As CCoupledNewtonSolver, mReverseOperatorsTest As CCoupledElasticState
Private mReverseConcreteTest As CCoupledConcreteResponse, mReverseReinforcedTest As CCoupledReinforcedResponse
Private mReverseSidesTest As Collection, mReverseFieldsTest As Collection, mReverseTraceTest As CCoupledLoadTrace
Private mReverseCalibrationTest As Variant, mReverseOptionsTest As Variant, mReverseModulusTest As Double
Private mReverseBaseTest() As Double, mReverseDirectionTest() As Double
'''
    assert code.count('Option Explicit\n') == 1
    code = code.replace('Option Explicit\n', 'Option Explicit\n' + globals, 1)
    old = '    Dim started As Double: started = Timer\n    Dim branch As New CCoupledLoadContinuation'
    assert code.count(old) == 1
    code = code.replace(old, '    Set mReverseSnapshotsTest = New Collection\n' + old, 1)
    old = '    Close #handle\n    mHardPreparedReportTest = '
    assert code.count(old) == 1
    code = code.replace(old, '''    Close #handle
    TestExportSpatialReverses solver, operators, concrete, reinforced, sides, calibration, _
        Array(initialSteps, attempts, maxIterations, absTol, relTol, lambdaTolerance, maximumLambda), offset, direction
    mHardPreparedReportTest = ''', 1)
    code += '\n' + (here / 'TestNativeReverseCapture.bas.txt').read_text('utf-8-sig')
    stored = (here / 'TestNativeStoredStall.bas.txt').read_text('utf-8-sig')
    exporter, = re.findall(r'(?ms)^Private Sub TestStoredExportCandidate\(\).*?^End Sub', stored)
    exporter = exporter.replace('TestStoredExportCandidate()', 'TestExportReverseCandidate(ByVal caseIndex As Long)')
    exporter = exporter.replace('mStored', 'mReverse')
    exporter = exporter.replace('mHardPreparedPathTest & ".HybridCandidate"', 'mHardPreparedPathTest & ".Reverse" & CStr(caseIndex)')
    exporter = exporter.replace('MODE;ControlledPointWitness;NotACapacityResult', 'MODE;RejectedSpatialEquilibrium;NotAcceptedContinuation;NotACapacityResult')
    exporter = exporter.replace('META;4;0;Подтверждена одна управляемая точка независимого численного метода. Это не предел.',
        'META;6;0;Отклонённое по направлению равновесие. Не принятый шаг и не предел.')
    exporter = exporter.replace('DETAILS;ControlledPointWitness;capacity_certified=False', 'DETAILS;RejectedSpatialEquilibrium;capacity_certified=False')
    exporter = exporter.replace('Контрольный путь внешнего метода', 'Контрольный путь отклонённого равновесия')
    exporter = exporter.replace('Поле внешнего метода', 'Отклонённое поле')
    code += "\n' -------------------- ДЛЯ ТЕСТОВ --------------------\n' Воспроизводит полную контрольную последовательность, включая отклонённое поле, только для выгрузки тензоров.\n" + exporter + '\n'
    assert max(map(len, code.splitlines())) < 1024
    methods = re.findall(r'(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)', code)
    assert len(methods) == len(set(name.lower() for name in methods))
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'Run-TestWorkingASD.ps1'
    code = (root / name).read_text('utf-8-sig')
    old = "'-MacroArgument2',$caseMode,'-TimeoutSeconds','240')"
    assert code.count(old) == 1
    code = code.replace(old, old.replace("'240'", "'600'"), 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeReverseCapture.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'] = {'case': 'ReverseCapture', 'spatial_half_step_dimensionless': .001,
        'material_laws_changed': False, 'working_decisions_changed': False, 'rejected_fields_not_accepted_by_controller': True,
        'post_run_material_replay_only_for_export': True, 'exact_Double_test_binary_format': 'NDMREVTEST000001',
        'outer_watchdog_seconds': 600, 'full_acceptance_completed': False}
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('FULL WORKING PATH AND REJECTED SPATIAL EQUILIBRIA READ ONLY CAPTURE READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve())
