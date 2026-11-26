"""ДЛЯ ТЕСТОВ. Проверяет общий рабочий адаптер на аналитической системе и точной истории связанного сечения."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source):
    """Рабочий интерфейс невязки не изменяется; тестовый наблюдатель сохраняет действительные отображения нового адаптера."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestNewtonYieldPrecision.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    additions = ('src/Interfaces/ICoupledScalarConstraint.cls', 'src/Solver/CCoupledConstraintEquations.cls',
        'src/Solver/CCoupledMaterialConstraint.cls', 'tests/CConstraintEquationsTest.cls', 'tests/modTestCoupledConstraint.bas')
    for name in additions:
        (root / name).parent.mkdir(parents=True, exist_ok=True)
        (root / name).write_bytes((repo / name).read_bytes())
        prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
        prep['modules'].append(Path(name).stem)
    name = 'tests/CProjectedEquationsTest.cls'
    (root / name).write_bytes((here / Path(name).name).read_bytes())
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    prep['modules'].append(Path(name).stem)
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    raw, = re.findall(r'(?ms)^Private Function ICoupledEquationSystem_TryEvaluate\(.*?^End Function', (repo / name).read_text('utf-8-sig'))
    old, = re.findall(r'(?ms)^Private Function ICoupledEquationSystem_TryEvaluate\(.*?^End Function', code)
    code = code.replace(old, raw, 1)
    old = 'mHybridOriginRevision = mAcceptedRevision: mHybridActive = True'
    assert code.count(old) == 1
    code = code.replace(old, 'mHybridOriginRevision = mAcceptedRevision: ActivateHybridContext', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mYieldProjectionHandleTest As Integer\nPrivate mProjectedEquationsTest As ICoupledEquationSystem\n', 1)
    old = '    Dim equations As ICoupledEquationSystem: Set equations = solver'
    assert code.count(old) == 1
    code = code.replace(old, '    Dim equations As ICoupledEquationSystem: Set equations = mProjectedEquationsTest', 1)
    block, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', code)
    old = '        revision = solver.AcceptedRevision'
    assert block.count(old) == 1
    body = block.replace(old, old + '''
        If probeCase = 0 Then target = 0.001 Else target = 0.01
        Print #handle, "EVENT_TARGET;" & CStr(probeCase) & ";" & TestOnsetNumber(target)
''', 1)
    old = '    Print #handle, "CONTROL;" & TestOnsetNumber(target) & TestHardPreparedVector(projection)'
    assert body.count(old) == 1
    body = body.replace(old, '    Print #handle, "CONTROL_EVENT;GenericProjectedTrialYieldMPa;TargetsRecordedPerCase"', 1)
    old = '        Dim scales As Variant: scales = solver.ReadNewtonLineScalesTest'
    assert body.count(old) == 1
    body = body.replace(old, old + '''
        Dim scalarTest As ICoupledScalarConstraint: Set scalarTest = solver.CreateMaterialConstraint(25, 1, 0, target)
        Dim zeroTest(1 To 61) As Double, projectorTest As CCoupledConstraintEquations
        Set projectorTest = New CCoupledConstraintEquations
        projectorTest.Initialize solver, scalarTest, zeroTest, 0.000000000002, 16
        Dim spyTest As CProjectedEquationsTest: Set spyTest = New CProjectedEquationsTest
        spyTest.InitializeTest projectorTest, solver: Set mProjectedEquationsTest = spyTest
        mYieldProjectionHandleTest = FreeFile
        Open mHardPreparedPathTest & ".Projection" & CStr(probeCase) & ".bin" For Binary Access Write As #mYieldProjectionHandleTest
        Dim magicTest As String * 16: magicTest = "NDMQPROJECT00001"
        Dim pivotTest As Long: pivotTest = projectorTest.Pivot
        Put #mYieldProjectionHandleTest, , magicTest: Put #mYieldProjectionHandleTest, , pivotTest
        Print #handle, "ELIMINATION_PIVOT;" & CStr(probeCase) & ";" & CStr(pivotTest)
''', 1)
    old = '        solver.ReadNewtonLineFieldTest coordinates, field, factor'
    assert body.count(old) == 1
    body = body.replace(old, '''        Dim mappedTest() As Double
        If Not projectorTest.TryMap(coordinates, mappedTest) Then Err.Raise 5, , "Не восстановлено поле рабочего адаптера."
        solver.ReadNewtonLineFieldTest mappedTest, field, factor''', 1)
    old = '        solver.EndNewtonLineTest'
    assert body.count(old) == 1
    body = body.replace(old, '''        Close #mYieldProjectionHandleTest: mYieldProjectionHandleTest = 0
''' + old + '''
        If Not TestConstraintContextRejected(scalarTest) Then Err.Raise 5, , "Координата принята после завершения контекста."
        Dim unchangedRevisionTest As Long: unchangedRevisionTest = solver.AcceptedRevision
        If Not solver.BeginNewtonLineTest(mReverseBaseTest, mReverseDirectionTest, projection, target, seed, seedFactor) Then _
            Err.Raise 5, , "Не открыт повторный контекст той же истории."
        If solver.AcceptedRevision <> unchangedRevisionTest Then Err.Raise 5, , "Повторное открытие контекста изменило историю."
        If Not TestConstraintContextRejected(scalarTest) Then Err.Raise 5, , "Прежняя координата принята в новом контексте той же истории."
        Dim renewedTest As ICoupledScalarConstraint, checkValueTest As Double, checkGradientTest() As Double
        Set renewedTest = solver.CreateMaterialConstraint(25, 1, 0, target)
        If Not renewedTest.TryEvaluate(zeroTest, False, checkValueTest, checkGradientTest) Then _
            Err.Raise 5, , "Свежая координата не прочитана в новом контексте."
        solver.EndNewtonLineTest
        Print #handle, "CONTEXT_GATES;" & CStr(probeCase) & ";InactiveRejected;SameRevisionNewContextRejected;FreshReadAccepted"
''', 1)
    old = '    Next probeCase'
    assert body.count(old) == 1
    body = body.replace(old, '''        TestGenericHybridCase solver, probeCase, handle, seed, seedFactor, projection, target, scalarTest
''' + old, 1)
    code = code.replace(block, body, 1)
    for filename in ('TestNativeYieldProjection.bas.txt', 'TestNativeGenericHybrid.bas.txt'):
        code += '\n' + (here / filename).read_text('utf-8-sig')
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'Run-TestWorkingASD.ps1'
    code = (root / name).read_text('utf-8-sig')
    old = "'Compile','ReversePolish','ActualReverse'"
    assert code.count(old) == 1
    code = code.replace(old, "'Compile','ConstraintUnit','ReversePolish','ActualReverse'", 1)
    old = "        if ($Phase -eq 'ReversePolish') {"
    assert code.count(old) == 1
    code = code.replace(old, '''        if ($Phase -eq 'ConstraintUnit') {
            $arguments+=@('-Macro','modTestCoupledConstraint.RunTestCoupledConstraint,modTestCoupledHybrid.RunTestCoupledHybrid','-TimeoutSeconds','165')
        } elseif ($Phase -eq 'ReversePolish') {''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for filename in (Path(__file__).name, 'CProjectedEquationsTest.cls', 'TestNativeYieldProjection.bas.txt', 'TestNativeGenericHybrid.bas.txt'):
        (root / filename).write_bytes((here / filename).read_bytes())
        prep['inputs'][filename] = sha256((root / filename).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'GenericProjectedMaterialEquations', 'nonlinear_event_control': False,
        'nonlinear_event_eliminated': True, 'event_target_MPa': None, 'event_targets_by_case_MPa': {'0': .001, '1': .01},
        'working_raw_equation_interface_used_unchanged': True, 'generic_material_reader_from_working_sources': True,
        'same_two_seeds_and_original_history_for_both_numerical_engines': True,
        'both_complete_Newton_and_existing_hybrid_attempted': True,
        'closed_and_replaced_contexts_rejected': True, 'spy_additional_readonly_map_for_each_successful_function': True,
        'local_event_root_tolerance_MPa': 2e-12, 'local_event_iteration_limit': 16,
        'new_working_classes_not_enabled_in_ordinary_solver': True, 'installed_workbook_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    run_path(str(here / 'Guard-TestExistingOffice.py'))['TestGuard'](root)
    print('GENERIC CONSTRAINT WORKING CLASSES AND MATCHED NATIVE ENGINE TESTS READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
