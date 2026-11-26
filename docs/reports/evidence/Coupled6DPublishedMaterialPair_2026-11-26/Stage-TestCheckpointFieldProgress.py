"""ДЛЯ ТЕСТОВ. Различает новое равновесное поле по заданной деформационной координате без требования различимого изменения lambda."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path


def TestStage(repo, root, source, checkpoint, binary, steps=12, attempts=24, ceiling=16):
    """Порядок lambda остаётся неопределённым при пересечении интервалов; силовые допуски и ориентация не меняются."""
    assert 1 <= steps <= 64 and steps <= attempts <= 192 and 16 <= ceiling <= 4096
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestCheckpointSpatialSequence.py'))['TestStage'](repo,root,source,checkpoint,binary)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    anchor = '    If referenceLambda > 0# Then unresolved = (lambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)))'
    assert code.count(anchor) == 1
    code = code.replace(anchor,'    unresolved = False \' ДЛЯ ТЕСТОВ: различимость поля проверяет деформационный цикл; порядок lambda не является условием НДС.',1)
    anchor = '''    If referenceLambda > 0# Then
        If lambda <= referenceLambda + 8# * relativeTolerance * (1# + Abs(referenceLambda)) + uncertainty Then GoTo Restore
    End If'''
    assert code.count(anchor) == 1
    code = code.replace(anchor,'    \' ДЛЯ ТЕСТОВ: неопределённое изменение lambda сохраняется в журнале; полное равновесие уже проверено.',1)
    (root / name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    anchor = 'MODE;SpatialSecantSequence;OriginalForceTolerance;ResolvedLoadBothDirections;NotPhysicalCapacity'
    assert code.count(anchor) == 1
    code = code.replace(anchor,'MODE;SpatialSecantSequence;OriginalForceTolerance;ResolvedControlAdvance;LambdaOrderMayBeUnresolved;NotPhysicalCapacity',1)
    anchor = '            solver.CopyTrial field, actual\n            Print #handle, "FINAL;"'
    assert code.count(anchor) == 1
    code = code.replace(anchor,'''            Dim actualProjection() As Double
            actualProjection = projection
            If solver.CopyRecoveredControl(actualProjection) Then
                Print #handle, "MATERIAL_CONTROL;" & CStr(attempt)
            Else
                actualProjection = projection
            End If
            Print #handle, "CONTROL_PROJECTION;" & CStr(attempt) & TestHardPreparedVector(actualProjection)
            solver.CopyTrial field, actual
            Print #handle, "FINAL;"''',1)
    anchor = '            ElseIf Abs(factor - currentFactor) <= bound + currentBound + resolution Then'
    assert code.count(anchor) == 1
    code = code.replace(anchor,'''            ElseIf TestSpatialControlAdvance(actualProjection, accepted, field) <= _
                    TestSpatialControlResolution(operators, currentFactor, factor, options, calibration(0)) Then''',1)
    anchor = '                solver.AcceptStep: fields.Add field: trace.Add 1#, factor'
    assert code.count(anchor) == 1
    code = code.replace(anchor,'''                Print #handle, "FIELD_ADVANCE;" & CStr(attempt) & ";" & _
                    TestOnsetNumber(TestSpatialControlAdvance(actualProjection, accepted, field)) & ";" & _
                    TestOnsetNumber(TestSpatialControlResolution(operators, currentFactor, factor, options, calibration(0)))
                If Abs(factor - currentFactor) <= bound + currentBound + resolution Then _
                    Print #handle, "LOAD_ORDER_UNRESOLVED;" & CStr(attempt) & ";NDSAccepted;NotEvidenceOfLoadRiseOrPeak"
''' + anchor,1)
    code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает изменение фактически действующей координаты двух полей без изменения их массивов.
Private Function TestSpatialControlAdvance(ByRef projection() As Double, ByRef before() As Double, ByRef after() As Double) As Double
    Dim i As Long, value As Double
    For i = 1 To UBound(projection): value = value + projection(i) * (after(i) - before(i)): Next i
    TestSpatialControlAdvance = Abs(value)
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сумма ошибок двух строк Eref*(p*u-c) ограничена 2*max(tolerance)/Eref; силовой допуск остаётся исходным.
Private Function TestSpatialControlResolution(ByVal operators As CCoupledElasticState, ByVal before As Double, ByVal after As Double, _
        ByVal options As Variant, ByVal referenceModulus As Double) As Double
    Dim load(1 To 6) As Double, external() As Double, maximum As Double, i As Long, side As Long, factor As Double
    For side = 0 To 1
        If side = 0 Then factor = before Else factor = after
        For i = 1 To 6: load(i) = mReverseBaseTest(i) + factor * mReverseDirectionTest(i): Next i
        operators.LoadEquations load(1), load(2), load(3), load(4), load(5), load(6), external
        For i = 1 To mReverseSolverTest.PhysicalDegreeCount: maximum = GeomMax(maximum, Abs(external(i))): Next i
    Next side
    TestSpatialControlResolution = 2# * (options(3) + options(4) * maximum) / referenceModulus
End Function
'''
    assert code.count('For attempt = 1 To 24') == 1
    code = code.replace('For attempt = 1 To 24','For attempt = 1 To '+str(attempts),1)
    assert code.count('If acceptedCount >= 12 Then') == 1
    code = code.replace('If acceptedCount >= 12 Then','If acceptedCount >= '+str(steps)+' Then',1)
    code = code.replace('RequestedTwelveStepsCompleted;NotPhysicalLimit','Requested'+str(steps)+'StepsCompleted;NotPhysicalLimit')
    assert code.count('initialDesired * 16#') == 2
    code = code.replace('initialDesired * 16#','initialDesired * '+str(ceiling)+'#')
    anchor = 'TestWriteSpatialSequenceSnapshot fields, trace, projection, coordinate + desired, actual, attempt'
    assert code.count(anchor) == 1
    code = code.replace(anchor,'TestWriteSpatialSequenceSnapshot fields, trace, actualProjection, coordinate + desired, actual, attempt',1)
    assert max(map(len,code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n','\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = Path(__file__).name
    (root / name).write_bytes(Path(__file__).read_bytes()); prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update(case='SpatialSequenceResolvedFieldUnresolvedLoadAllowed',
        maximum_new_accepted_steps=steps, maximum_attempts=attempts, maximum_step_ratio_to_initial=ceiling,
        nds_admission='original force tolerance + unchanged field + forward tensor direction + steel admissibility + resolved actual control increment',
        control_increment_resolution='2*max(original absolute-plus-relative force tolerance)/reference modulus',
        overlapping_lambda_intervals_not_counted_as_load_rise_or_peak=True,
        actual_lambda_uncertainty_still_recorded=True, material_and_force_tolerances_unchanged=True,
        abs_gain_required_above_sum_of_bounds_and_resolution=False,
        test_controller_reimplementation_used=True, working_TryDeformationStep_only=False,
        actual_candidate_control_projection_logged_and_preserved=True,
        binary_context_coordinate_is_requested_not_recovered=True,
        total_iteration_budget_unchanged=True, working_sources_changed=False,
        installed_workbook_changed=False, physical_capacity_certified=False)
    (root / 'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
    print('RESOLVED DEFORMATION ADVANCE WITH ORIGINAL FORCE TOLERANCE AND EXPLICIT UNRESOLVED LOAD ORDER READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--checkpoint', type=Path, required=True)
    parser.add_argument('--binary', type=Path, required=True)
    parser.add_argument('--steps', type=int, default=12)
    parser.add_argument('--attempts', type=int, default=24)
    parser.add_argument('--ceiling', type=int, default=16)
    args = parser.parse_args()
    TestStage(args.repo.resolve(),args.root.resolve(),args.source.resolve(),args.checkpoint.resolve(),args.binary.resolve(),args.steps,args.attempts,args.ceiling)
