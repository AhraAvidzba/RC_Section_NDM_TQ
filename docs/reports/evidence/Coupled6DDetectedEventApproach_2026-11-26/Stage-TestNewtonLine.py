"""ДЛЯ ТЕСТОВ. Готовит полный Newton с поиском шага на точной истории штатного материала."""

from argparse import ArgumentParser
import csv
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path
import struct


def TestStage(repo, root, source):
    """Меняется только частный тестовый код; действующий интерфейс полной невязки используется без дублирования уравнений."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestReversePolish.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig') + '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Открывает читающий интерфейс полной системы на неизменной истории; не публикует материальную пробу.
Friend Function BeginNewtonLineTest(ByRef baseLoad() As Double, ByRef direction() As Double, _
        ByRef projection() As Double, ByVal target As Double, ByRef seed() As Double, ByVal factor As Double) As Boolean
    If mCandidateReady Or mHybridActive Then Err.Raise 5, , "Не завершена предыдущая тестовая проба."
    If Not PrepareHybridScales() Then Exit Function
    ConfigureControlledLoad baseLoad, direction
    mArcLoadProjection = 0#: mConsistentControl = False
    mControlProjection = projection: mTargetControl = target: mControlActive = True
    ReDim mHybridSeed(1 To mDegreeCount + 1)
    Dim degree As Long
    For degree = 1 To mDegreeCount: mHybridSeed(degree) = seed(degree): Next degree
    mHybridSeed(mDegreeCount + 1) = factor
    mHybridOriginRevision = mAcceptedRevision: mHybridActive = True
    BeginNewtonLineTest = True
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Копирует исходные масштабы без пересчёта закона материала или изменения принятой истории.
Friend Function ReadNewtonLineScalesTest() As Variant
    ReadNewtonLineScalesTest = mHybridScales
End Function

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает физическое поле и множитель из тех же координат, которые использует рабочий интерфейс невязки.
Friend Sub ReadNewtonLineFieldTest(ByRef coordinates() As Double, ByRef field() As Double, ByRef factor As Double)
    If Not mHybridActive Then Err.Raise 5, , "Нет действующей тестовой системы."
    ReDim field(1 To mDegreeCount)
    Dim degree As Long
    For degree = 1 To mDegreeCount: field(degree) = mHybridSeed(degree) + mHybridScales(degree) * coordinates(degree): Next degree
    factor = mHybridSeed(mDegreeCount + 1) + mHybridScales(mDegreeCount + 1) * coordinates(mDegreeCount + 1)
End Sub

' -------------------- ДЛЯ ТЕСТОВ --------------------
' Завершает только читающий тестовый контекст; окончательное поле отдельно проверяется штатной силовой пробой.
Friend Sub EndNewtonLineTest()
    mHybridActive = False: mControlActive = False
End Sub
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    original = (root / name).read_text('utf-8-sig')
    body = (here / 'TestNativeNewtonLine.bas.txt').read_text('utf-8-sig')
    code, count = re.subn(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', lambda _: body, original)
    assert count == 1 and max(map(len, code.splitlines())) < 1024
    methods = re.findall(r'(?im)^(?:Public|Private|Friend) (?:Function|Sub) (\w+)', code)
    assert len(methods) == len(set(item.lower() for item in methods))
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    previous = Path('D:/CodexScratch/RC_Section_NDM_TQ/BestPrediction20261119/Native/BestPrediction.csv')
    rows = list(csv.reader(previous.read_text('cp1251').splitlines(), delimiter=';'))
    best, = [row for row in rows if row[:2] == ['BEST', '0']]
    values = [float(value) for value in best[4:]] + [float(best[2])]
    assert len(values) == 61
    name = 'Input/NearSeed.bin'
    (root / name).write_bytes(struct.pack('<i61d', 61, *values))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    (root / 'Input/NearSeedSource.csv').write_bytes(previous.read_bytes())
    prep['inputs']['Input/NearSeedSource.csv'] = sha256(previous.read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeNewtonLine.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'CompleteNewtonLineSearch', 'new_capacity_search_executed': False,
        'accepted_prefix': 26, 'target_previous_secant_fraction': .5,
        'seeds': ['ForwardSecant', 'PreviousNearEquilibriumCSV'], 'strict_residual_MPa': 1e-10,
        'maximum_corrections_per_seed': 30, 'maximum_function_calls_per_seed': 4500,
        'maximum_line_halvings': 20, 'central_difference_relative_step': 1e-6,
        'linear_solver': 'CCoupledLinearSystem', 'equations': 'existing ICoupledEquationSystem adapter',
        'material_and_equilibrium_limits_unchanged': True, 'working_runtime_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('COMPLETE NEWTON LINE SEARCH ON EXACT HISTORY READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
