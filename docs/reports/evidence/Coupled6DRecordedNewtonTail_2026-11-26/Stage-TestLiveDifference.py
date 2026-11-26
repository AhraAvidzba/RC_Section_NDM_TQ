"""ДЛЯ ТЕСТОВ. Проверяет рабочий резерв на реально вычисляемой матрице без внешней подсказки номера точки и уровня."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source):
    """Переиспользует исходную историю и штатные входы; родительская тестовая параметризация не вызывается."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestOriginalAccuracy.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    for name in ('src/Interfaces/ICoupledDifferenceObserver.cls', 'tests/CHybridDifferenceObserverTest.cls'):
        (root / name).write_bytes((repo / name).read_bytes())
        prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
        prep['modules'].append(Path(name).stem)
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает выбранную станцию на фактическом отрезке действующего контекста без изменения истории.
Friend Function ReadLiveDifferenceStationTest(ByRef before() As Double, ByRef after() As Double) As Variant
    If Not mHybridActive Or mAcceptedRevision <> mHybridOriginRevision Then Err.Raise 5, , "Нет действующего контекста наблюдения."
    Dim first() As Double, last() As Double, i As Long, station As TYieldStation, ready As Boolean
    ReDim first(1 To mDegreeCount): ReDim last(1 To mDegreeCount)
    For i = 1 To mDegreeCount
        first(i) = mHybridSeed(i) + mHybridScales(i) * before(i)
        last(i) = mHybridSeed(i) + mHybridScales(i) * after(i)
    Next i
    ready = FirstYieldCrossing(first, last, station)
    ReadLiveDifferenceStationTest = Array(ready, station.Point, station.Side, station.PrincipalIndex, station.FromValue, station.ToValue)
End Function
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'src/Solver/CCoupledHybridSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    anchor = '            If accepted Then\n                mDifferenceBefore = before: mDifferenceAfter = after'
    assert code.count(anchor) == 1
    code = code.replace(anchor, '            If accepted Then\n                modTestCoupledElasticState.TestRecordLiveDifference before, after, score, j, mJacobianCalls, mSteps, mFunctionCalls\n                mDifferenceBefore = before: mDifferenceAfter = after', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    body = (here / 'TestNativeLiveDifference.bas.txt').read_text('utf-8-sig')
    code, count = re.subn(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', lambda _: body, code)
    assert count == 1 and max(map(len, code.splitlines())) < 1024
    code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Сохраняет реально подтверждённые пары; не выбирает управление и не изменяет исходную функцию.
Public Sub TestRecordLiveDifference(ByRef before() As Double, ByRef after() As Double, ByVal score As Double, _
        ByVal column As Long, ByVal matrix As Long, ByVal steps As Long, ByVal calls As Long)
    If mReverseSolverTest Is Nothing Then Exit Sub
    Dim station As Variant: station = mReverseSolverTest.ReadLiveDifferenceStationTest(before, after)
    Dim handle As Integer: handle = FreeFile
    Open mHardPreparedPathTest & ".Brackets.csv" For Append As #handle
    Print #handle, "OBSERVED;" & CStr(matrix) & ";" & CStr(steps) & ";" & CStr(calls) & ";" & CStr(column) & ";" & _
        TestOnsetNumber(score) & ";" & CStr(station(0)) & ";" & CStr(station(1)) & ";" & CStr(station(2)) & ";" & CStr(station(3)) & ";" & _
        TestOnsetNumber(station(4)) & ";" & TestOnsetNumber(station(5)) & TestHardPreparedVector(before) & TestHardPreparedVector(after)
    Close #handle
End Sub
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for filename in (Path(__file__).name, 'TestNativeLiveDifference.bas.txt'):
        (root / filename).write_bytes((here / filename).read_bytes())
        prep['inputs'][filename] = sha256((here / filename).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'WorkingLiveDifferenceMaterialRecovery',
        'working_TryDeformationStep_only': True, 'no_external_matrix_or_material_event_supplied': True,
        'original_material_accuracy': True, 'original_acceptance_and_steel_limits': True,
        'live_normalized_difference_capture_requested': True, 'maximum_total_solver_iterations': 30,
        'recovered_lambda_uncertainty_required': True,
        'installed_workbook_changed': False, 'physical_capacity_certified': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('WORKING LIVE DIFFERENCE RECOVERY READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
