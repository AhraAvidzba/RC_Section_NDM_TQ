"""ДЛЯ ТЕСТОВ. Исключает одну координату через точное нелинейное условие пробной поверхности CDP."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source):
    """Глобальные уравнения, материал и история сохранены; фиктивная координата заменяет исключённое неизвестное."""
    here = Path(__file__).resolve().parent
    run_path(str(here / 'Stage-TestNewtonYieldPrecision.py'))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mYieldProjectionPivotTest As Long\n', 1)
    block, = re.findall(r'(?ms)^Friend Function BeginNewtonLineTest\(.*?^End Function', code)
    old = '    BeginNewtonLineTest = True'
    assert block.count(old) == 1
    body = block.replace(old, '''    Dim eventGradientTest() As Double, largestTest As Double
    modTestCoupledElasticState.TestReadNewtonYieldEventGradient seed, eventGradientTest
    mYieldProjectionPivotTest = 0
    For degree = 1 To mDegreeCount
        If Abs(eventGradientTest(degree) * mHybridScales(degree)) > largestTest Then
            largestTest = Abs(eventGradientTest(degree) * mHybridScales(degree)): mYieldProjectionPivotTest = degree
        End If
    Next degree
    If mYieldProjectionPivotTest = 0 Then Err.Raise 5, , "Нет невырожденной координаты исключения события."
''' + old, 1)
    code = code.replace(block, body, 1)
    block, = re.findall(r'(?ms)^Private Function ICoupledEquationSystem_TryEvaluate\(.*?^End Function', code)
    old = '    SetControlledLoad lambda, load'
    assert block.count(old) == 1
    body = block.replace(old, '''    If Not ProjectYieldFieldTest(coordinates, field) Then Exit Function
''' + old, 1)
    old = '    residual(size) = modTestCoupledElasticState.TestReadNewtonYieldEvent(field) - mTargetControl'
    assert body.count(old) == 1
    body = body.replace(old, '    residual(size) = coordinates(mYieldProjectionPivotTest)', 1)
    old = '''        Dim eventGradient() As Double
        modTestCoupledElasticState.TestReadNewtonYieldEventGradient field, eventGradient
        For j = 1 To mDegreeCount: jacobian(j, size) = eventGradient(j) * mHybridScales(j): Next j
        jacobian(size, size) = 0#'''
    assert body.count(old) == 1
    body = body.replace(old, '''        Dim eventGradient() As Double, normalColumnTest() As Double
        modTestCoupledElasticState.TestReadNewtonYieldEventGradient field, eventGradient
        ReDim normalColumnTest(1 To mDegreeCount)
        For i = 1 To mDegreeCount: normalColumnTest(i) = jacobian(mYieldProjectionPivotTest, i) / mHybridScales(mYieldProjectionPivotTest): Next i
        For j = 1 To mDegreeCount
            For i = 1 To mDegreeCount
                jacobian(j, i) = jacobian(j, i) - normalColumnTest(i) * eventGradient(j) * mHybridScales(j) / eventGradient(mYieldProjectionPivotTest)
            Next i
        Next j
        For j = 1 To size: jacobian(j, size) = 0#: Next j
        jacobian(mYieldProjectionPivotTest, size) = 1#''', 1)
    code = code.replace(block, body, 1)
    block, = re.findall(r'(?ms)^Friend Sub ReadNewtonLineFieldTest\(.*?^End Sub', code)
    old = '    factor = mHybridSeed(mDegreeCount + 1) + mHybridScales(mDegreeCount + 1) * coordinates(mDegreeCount + 1)'
    assert block.count(old) == 1
    body = block.replace(old, old + '''
    If Not ProjectYieldFieldTest(coordinates, field) Then Err.Raise 5, , "Не восстановлено поле исключённого события."
''', 1)
    code = code.replace(block, body, 1)
    code += '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Возвращает фиксированный индекс исключаемого неизвестного; он выбирается по ненулевой производной, а не по типу сечения.
Friend Property Get YieldProjectionPivotTest() As Long
    YieldProjectionPivotTest = mYieldProjectionPivotTest
End Property

' -------------------- ДЛЯ ТЕСТОВ --------------------
' При неизменных остальных координатах удерживает пробную поверхность на нуле; фиктивная координата не влияет на физическое поле.
Private Function ProjectYieldFieldTest(ByRef coordinates() As Double, ByRef field() As Double) As Boolean
    Dim pivot As Long: pivot = mYieldProjectionPivotTest
    If pivot < 1 Or pivot > mDegreeCount Then Err.Raise 5, , "Не выбрана координата исключения."
    field(pivot) = mHybridSeed(pivot)
    Dim iteration As Long, value As Double, gradient() As Double, previous As Double
    For iteration = 1 To 16
        value = modTestCoupledElasticState.TestReadNewtonYieldEvent(field)
        If Abs(value - mTargetControl) <= 0.000000000002 Then
            modTestCoupledElasticState.TestRecordYieldProjection coordinates, field, pivot, value
            ProjectYieldFieldTest = True
            Exit Function
        End If
        modTestCoupledElasticState.TestReadNewtonYieldEventGradient field, gradient
        If gradient(pivot) = 0# Then Exit Function
        previous = field(pivot): field(pivot) = previous - (value - mTargetControl) / gradient(pivot)
        If field(pivot) = previous Then Exit Function
    Next iteration
End Function
'''
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mYieldProjectionHandleTest As Integer\n', 1)
    block, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', code)
    old = '        Dim scales As Variant: scales = solver.ReadNewtonLineScalesTest'
    assert block.count(old) == 1
    body = block.replace(old, '''        mYieldProjectionHandleTest = FreeFile
        Open mHardPreparedPathTest & ".Projection" & CStr(probeCase) & ".bin" For Binary Access Write As #mYieldProjectionHandleTest
        Dim projectionMagicTest As String * 16: projectionMagicTest = "NDMQPROJECT00001"
        Dim pivotTest As Long: pivotTest = solver.YieldProjectionPivotTest
        Put #mYieldProjectionHandleTest, , projectionMagicTest: Put #mYieldProjectionHandleTest, , pivotTest
        Print #handle, "ELIMINATION_PIVOT;" & CStr(probeCase) & ";" & CStr(pivotTest)
''' + old, 1)
    old = '        solver.EndNewtonLineTest'
    assert body.count(old) == 1
    body = body.replace(old, '        Close #mYieldProjectionHandleTest: mYieldProjectionHandleTest = 0\n' + old, 1)
    code = code.replace(block, body, 1) + '\n' + (here / 'TestNativeYieldProjection.bas.txt').read_text('utf-8-sig')
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    name = 'Guard-IsolatedOffice.ps1'
    code = (root / name).read_text('utf-8-sig')
    old = 'if (!$instance -or !$item.MainWindowHandle -or'
    assert code.count(old) == 1
    code = code.replace(old, 'if (!$instance -or', 1)
    old = '''        if ($parent -and $parent.ProcessName -ieq 'explorer' -and $parent.SessionId -eq $item.SessionId -and
            $parent.StartTime -lt $item.StartTime -and $instance.CommandLine -notmatch '[/\\-]embedding(?:\\s|$)') {$protected[$item.Id]=$ticks}'''
    assert code.count(old) == 1
    code = code.replace(old, '''        if ($parent -and $parent.SessionId -eq $item.SessionId -and $parent.StartTime -lt $item.StartTime -and
            (($item.MainWindowHandle -and $parent.ProcessName -ieq 'explorer' -and $instance.CommandLine -notmatch '[/\\-]embedding(?:\\s|$)') -or
             ($parent.ProcessName -ieq 'outlook' -and $instance.CommandLine -match '[/\\-]embedding(?:\\s|$)'))) {$protected[$item.Id]=$ticks}''', 1)
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    for file in (Path(__file__), here / 'TestNativeYieldProjection.bas.txt'):
        (root / file.name).write_bytes(file.read_bytes())
        prep['inputs'][file.name] = sha256(file.read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'CompleteNewtonNonlinearYieldElimination',
        'nonlinear_event_control': False, 'nonlinear_event_eliminated': True,
        'elimination_rule': 'one physical field coordinate solved from actual trial yield=0; dummy coordinate fixed at zero',
        'pivot_rule': 'largest abs(analytic event derivative times coordinate scale) at each unchanged seed',
        'local_event_root_tolerance_MPa': 2e-12, 'local_event_iteration_limit': 16,
        'all_raw_coordinates_and_solved_pivot_Double_recorded': True,
        'existing_external_Outlook_Excel_protected_by_PID_and_start_time': True,
        'working_runtime_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    print('NONLINEAR YIELD ELIMINATION STUDY READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
