"""ДЛЯ ТЕСТОВ. Продолжает принятую историю несколькими строго уравновешенными пластическими шагами."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path


def TestStage(repo, root, source, parent='Stage-TestYieldProjection.py'):
    """Каждый новый шаг использует реально принятую историю предшествующего; исходный префикс не подменяется."""
    here = Path(__file__).resolve().parent
    run_path(str(here / parent))['TestStage'](repo, root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    name = 'tests/modTestCoupledElasticState.bas'
    code = (root / name).read_text('utf-8-sig')
    code = code.replace('Option Explicit\n', 'Option Explicit\nPrivate mSequenceOriginalPrefixTest As Long\n', 1)
    block, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', code)
    body = block
    old = '    For probeCase = 0 To 1'
    assert body.count(old) == 1
    body = body.replace(old, '''    ReturnPrecisionFactorTest = 0#
    If Not solver.ReplayAcceptedFields(fields, fields.Count) Then Err.Raise 5, , "Не восстановлена начальная история последовательности."
    mSequenceOriginalPrefixTest = fields.Count
    Set mReverseFieldsTest = New Collection
    For degree = 1 To fields.Count: copy = fields(degree): mReverseFieldsTest.Add copy: Next degree
    Set mReverseTraceTest = trace.CopyPrefix(trace.Count)
    Dim acceptedCountTest As Long
    For probeCase = 0 To 3
        Print #handle, "CASE_BEGIN;" & CStr(probeCase)
''', 1)
    old = '''        ReturnPrecisionFactorTest = 0#
        If Not solver.ReplayAcceptedFields(fields, fields.Count) Then Err.Raise 5, , "Не восстановлена исходная история Newton."
'''
    assert body.count(old) == 1
    body = body.replace(old, '', 1)
    old = '        revision = solver.AcceptedRevision'
    assert body.count(old) == 1
    body = body.replace(old, old + '\n        target = 0.001', 1)
    body, count = re.subn(r'(?ms)^        If probeCase = 1 Then\n.*?^        End If\n', '', body)
    assert count == 1
    start = body.index('        solver.DiscardStep\n        If verified And unchanged Then')
    end = body.index('    Next probeCase', start)
    body = body[:start] + '''        If solved And verified And unchanged And forward And steelUse <= 1# Then
            solver.AcceptStep
            If solver.AcceptedRevision <> revision + 1 Then Err.Raise 5, , "Неверная ревизия принятого шага последовательности."
            mReverseFieldsTest.Add field: mReverseTraceTest.Add 1#, factor
            acceptedCountTest = acceptedCountTest + 1
            Print #handle, "SEQUENCE_ACCEPTED;" & CStr(probeCase) & ";" & TestOnsetNumber(factor) & ";" & CStr(solver.AcceptedRevision)
            previous = accepted: accepted = field
            previousFactor = currentFactor: currentFactor = factor
            For degree = 1 To UBound(movement): movement(degree) = accepted(degree) - previous(degree): Next degree
        Else
            solver.DiscardStep
            Print #handle, "SEQUENCE_STOP;" & CStr(probeCase) & ";NoUnverifiedStateAccepted;NotPhysicalFailure"
            Exit For
        End If
''' + body[end:]
    old = '    ReturnPrecisionFactorTest = 0#\n    Close #handle: mPolishExecutedTest = True'
    assert body.count(old) == 1
    body = body.replace(old, '''    Print #handle, "SEQUENCE_END;" & CStr(acceptedCountTest) & ";" & CStr(probeCase) & ";4;PhysicalCapacityOpen"
    If acceptedCountTest > 0 Then TestExportReverseCandidate 999
''' + old, 1)
    code = code.replace(block, body, 1)
    old = "' Сравнивает полный Newton из секущего и близкого прогноза на одном неизменном материальном префиксе."
    assert code.count(old) == 1
    code = code.replace(old, "' Последовательно принимает до четырёх строгих равновесий, сохраняя пластическую историю каждого принятого шага.", 1)
    block, = re.findall(r'(?ms)^Private Sub TestExportReverseCandidate\(.*?^End Sub', code)
    old = '    If caseIndex = 0 Then path = mHardPreparedPathTest & ".Accepted"'
    assert block.count(old) == 1
    body = block.replace(old, old + '\n    If caseIndex = 999 Then path = mHardPreparedPathTest & ".Sequence"', 1)
    old = 'If index = mReverseFieldsTest.Count Then ReturnPrecisionFactorTest = savedPrecisionTest'
    assert body.count(old) == 1
    body = body.replace(old, 'If index > mSequenceOriginalPrefixTest Then ReturnPrecisionFactorTest = savedPrecisionTest', 1)
    code = code.replace(block, body, 1)
    assert max(map(len, code.splitlines())) < 1024
    (root / name).write_bytes(code.replace('\n', '\r\n').encode('utf-8-sig'))
    prep['inputs'][name] = sha256((root / name).read_bytes()).hexdigest().upper()
    (root / Path(__file__).name).write_bytes(Path(__file__).read_bytes())
    prep['inputs'][Path(__file__).name] = sha256(Path(__file__).read_bytes()).hexdigest().upper()
    prep['directed_scope'].update({'case': 'SequentialProjectedPlasticContinuation',
        'case_ids': list(range(4)), 'step_limit': 4, 'initial_accepted_prefix': 26,
        'initial_prefix_replayed_once_at_original_precision': True,
        'new_accepted_steps_keep_their_own_material_histories_at_precision_factor': .0001,
        'event_target_MPa': .001, 'seeds': ['SequentialForwardSecant'],
        'independent_same_history_pair': False, 'sequence_single_material_export': True,
        'strict_root_original_force_exact_field_forward_and_steel_gates_before_each_accept': True,
        'concrete_physical_capacity_not_certified_by_sequence': True, 'working_runtime_changed': False})
    (root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True)+'\n', encoding='utf-8')
    run_path(str(here / 'Guard-TestExistingOffice.py'))['TestGuard'](root)
    print('SEQUENTIAL PROJECTED PLASTIC CONTINUATION READY')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--repo', type=Path, default=Path.cwd())
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    args = parser.parse_args()
    TestStage(args.repo.resolve(), args.root.resolve(), args.source.resolve())
