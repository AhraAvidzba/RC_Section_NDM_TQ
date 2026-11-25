"""ДЛЯ ТЕСТОВ. Переносит полный холодный путь и последнюю H-пробу в нативное восстановление без десятичного преобразования."""
from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
import re
from runpy import run_path
import struct
import numpy as np


def TestVector(values):
    """Сериализует исходные Double в том же контракте, что VBA Put."""
    values = np.asarray(values, dtype='<f8').ravel()
    return struct.pack('<i', len(values)) + values.tobytes()


def TestTrajectoryBytes(original, frame):
    """Переносит точный двоичный путь в существующий нативный контракт восстановления без изменения физических входов."""
    raw = b'NDMREVTEST000001' + struct.pack('<i', 60)
    for name in ('calibration', 'options', 'offset', 'direction'):
        raw += TestVector(original[name])
    raw += struct.pack('<i', len(frame['fields']))
    for fraction, factor, field in zip(frame['fractions'], frame['lambdas'], frame['fields']):
        raw += struct.pack('<dd', fraction, factor) + TestVector(field)
    for value in (frame['fields'][-2], frame['fields'][-1], frame['fields'][-1],
                  original['offset'] * frame['fractions'][-1] + frame['lambdas'][-1] * original['direction'], np.zeros(60),
                  [frame['lambdas'][-1], 0., frame['lambdas'][-1], 0.], np.zeros(4)):
        raw += TestVector(value)
    return raw


def TestStage(root, source, request_number=None):
    """Проверяет происхождение каждого Double, полный хвост файлов и точную привязку H-запроса к принятому префиксу."""
    here = Path(__file__).resolve().parent
    read = run_path(str(here / 'Read-TestColdRawCapture.py'))
    frames, history = read['TestHistory'](source / 'Native/Stock.RawHistory.bin')
    requests, goals = read['TestRequests'](source / 'Native/Stock.HRequests.bin')
    source_scope = json.loads((source / 'Preparation.json').read_bytes())['directed_scope']
    full_cold = source_scope.get('full_cold_from_zero')
    if full_cold is None:
        full_cold = (source / 'Native/HBackendStock.txt').is_file()
    frame = frames[-1]
    if request_number is None:
        request = requests[-1]
    else:
        assert 1 <= request_number <= len(requests)
        request = requests[request_number - 1]
    matches = np.flatnonzero(np.all(frame['fields'] == request['accepted'], axis=1) & (frame['lambdas'] == request['info'][2]))
    assert len(matches) == 1
    count = int(matches[0]) + 1
    assert count >= 2 and np.array_equal(request['previous'], frame['fields'][count - 2])
    original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source / 'Input/Reverse1.bin')
    assert np.array_equal(original['offset'], request['offset']) and np.array_equal(original['direction'], request['direction'])
    run_path(str(here / 'Stage-TestWorkingControllerResume.py'))['TestStage'](root, source)
    prep = json.loads((root / 'Preparation.json').read_bytes())
    raw = TestTrajectoryBytes(original, frame)
    values = {'Input/Reverse1.bin': raw, 'Input/ColdHPrefix.bin': TestVector([count]),
        'Input/ColdHRequest.bin': b''.join(TestVector(request[name]) for name in ('info', 'offset', 'direction', 'previous', 'seed', 'accepted'))}
    code = (root / 'tests/modTestCoupledElasticState.bas').read_text('utf-8-sig')
    body = (here / 'TestColdRawReplay.bas.txt').read_text('utf-8-sig')
    seed_pattern = r'(?ms)^Private Function TestColdRawHSeed\(.*?^End Function'
    existing_seed = re.findall(seed_pattern, code)
    expected_seed, = re.findall(seed_pattern, body)
    assert len(existing_seed) <= 1
    if existing_seed:
        assert existing_seed[0] == expected_seed
        second_header = body.rfind("' -------------------- ДЛЯ ТЕСТОВ --------------------", 0, body.index('Private Function TestColdRawHSeed('))
        assert second_header > body.index('End Sub')
        body = body[:second_header].rstrip() + '\n'
    code, replacements = re.subn(r'(?ms)^Private Sub TestIntegratedYieldEvent\(\).*?^End Sub', lambda _: body, code)
    assert replacements == 1
    values['tests/modTestCoupledElasticState.bas'] = code.replace('\n', '\r\n').encode('utf-8-sig')
    name = 'src/Solver/CCoupledNewtonSolver.cls'
    code = (root / name).read_text('utf-8-sig')
    addition = '''
' -------------------- ДЛЯ ТЕСТОВ --------------------
' Читает H прогнозного поля без принятия состояния; применяется только для исследования начального приближения.
Friend Function TryReadColdHFieldTest(ByRef field() As Double, ByRef value As Double) As Boolean
    If Not PrepareEnvelopeCoordinate() Then Exit Function
    TryReadColdHFieldTest = mEnvelope.TryCoordinate(field, value)
End Function
'''
    read_pattern = r'(?ms)^Friend Function TryReadColdHFieldTest\(.*?^End Function'
    existing_reader = re.findall(read_pattern, code)
    expected_reader, = re.findall(read_pattern, addition)
    assert len(existing_reader) <= 1
    if existing_reader:
        assert existing_reader[0] == expected_reader
    else:
        code += addition
    values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
    compiler = (root / 'Compile-TestExternalPoint.ps1').read_text('utf-8-sig')
    assert compiler.count("-ModuleNames 'modTestCoupledElasticState'") == 1
    values['Compile-TestExternalPoint.ps1'] = compiler.replace("-ModuleNames 'modTestCoupledElasticState'", "-ModuleNames @('CCoupledNewtonSolver','modTestCoupledElasticState')", 1).encode('utf-8-sig')
    record = dict(source=str(source), native_history_file=history, native_H_file=goals,
        full_fields=len(frame['fields']), H_request_prefix=count, request_number=request['number'],
        native_field_and_lambda_bits_preserved=True, target_and_seed_bits_preserved=True,
        all_history_from_actual_current_cold_run=bool(full_cold), native_recorded_history_preserved=True,
        source_is_full_cold_from_zero=bool(full_cold), no_decimal_field_reconstruction=True,
        H_ray_preparation_has_separate_budget_not_hidden_in30=True)
    values['Input/ColdRawOrigin.json'] = (json.dumps(record, indent=2, sort_keys=True) + '\n').encode()
    for name in (Path(__file__).name, 'TestColdRawReplay.bas.txt'):
        values[name] = (here / name).read_bytes()
    for name, content in values.items():
        (root / name).write_bytes(content)
        prep['inputs'][name] = sha256(content).hexdigest().upper()
    prep['directed_scope'].update(case='ExactColdRawReplayAndH',
        actual_import_this_pass=['CCoupledNewtonSolver', 'modTestCoupledElasticState'],
        initial_fields=len(frame['fields']), exact_H_prefix=count, original_maximum_iterations=30,
        working_source_changed=False, physical_capacity_certified=False)
    (root / 'Preparation.json').write_bytes((json.dumps(prep, indent=2, sort_keys=True) + '\n').encode())
    print('EXACT NATIVE HISTORY REPLAY READY', len(frame['fields']), 'FIELDS; H PREFIX', count, '; NO DECIMAL RECONSTRUCTION')


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--request', type=int)
    args = parser.parse_args()
    TestStage(args.root.resolve(), args.source.resolve(), args.request)
