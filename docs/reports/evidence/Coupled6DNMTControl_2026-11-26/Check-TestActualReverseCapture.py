"""ДЛЯ ТЕСТОВ. Проверяет полную штатную историю, отклонённые равновесия и точные двоичные снимки."""

from argparse import ArgumentParser
from hashlib import sha256
import json
from pathlib import Path
from runpy import run_path
import struct
import numpy as np


def TestReadBinary(path):
    """Читает строго ограниченный тестовый формат little-endian; неверная длина или остаток файла недопустимы."""
    raw, position = path.read_bytes(), 0
    assert raw[:16] == b'NDMREVTEST000001'
    position = 16
    def TestLong():
        """Размеры записаны четырёхбайтовым VBA Long."""
        nonlocal position
        assert position + 4 <= len(raw)
        result, = struct.unpack_from('<i', raw, position)
        position += 4
        return result
    def TestVector():
        """Значения читаются как исходные IEEE-754 Double без десятичного преобразования."""
        nonlocal position
        size = TestLong()
        assert 0 < size <= 10000 and position + 8 * size <= len(raw)
        result = np.frombuffer(raw, dtype='<f8', count=size, offset=position).copy()
        position += 8 * size
        assert np.all(np.isfinite(result))
        return result
    size = TestLong()
    assert size == 60
    result = dict(zip(('calibration', 'options', 'offset', 'direction'), (TestVector() for _ in range(4))))
    count = TestLong()
    assert 2 <= count <= 1000
    fields, fractions, factors = [], [], []
    for _ in range(count):
        assert position + 16 <= len(raw)
        fraction, factor = struct.unpack_from('<dd', raw, position)
        position += 16
        field = TestVector()
        assert field.shape == (size,)
        assert np.isfinite(fraction) and np.isfinite(factor)
        fields.append(field)
        fractions.append(fraction)
        factors.append(factor)
    result.update(fields=np.asarray(fields), fractions=np.asarray(fractions), lambdas=np.asarray(factors))
    result.update(zip(('previous', 'accepted', 'trial', 'actual', 'projection', 'context', 'metrics'), (TestVector() for _ in range(7))))
    assert [result[key].shape for key in ('previous', 'accepted', 'trial', 'actual', 'projection', 'context', 'metrics')] == [(60,), (60,), (60,), (6,), (60,), (4,), (4,)]
    assert position == len(raw)
    result['sha256'] = sha256(raw).hexdigest().upper()
    return result


def TestCheck(root, output):
    """Независимо собирает уравнения по трём станциям; отрицательное направление сохраняет как отклонённое, а не разрушение."""
    here = Path(__file__).resolve().parent
    validate = run_path(str(here / 'Validate-TestHardPrepared.py'))['TestValidate']
    accepted, fields = validate(root, 'ActualReverse.Accepted')
    assert len(fields) == 26 and not accepted['reversed_fields']
    helpers = run_path(str(here / 'Check-TestHardHalfSteps.py'))
    _, _, _, _, coordinates, _, b, slope = helpers['TestOperators'](root / 'Native/ActualReverse.Accepted.Operators.csv', 60, 43)
    weights = coordinates[:, 2]
    tensor = np.asarray((1., 1., 1., .5, .5, .5))
    def TestStrains(field):
        """Три станции соответствуют безразмерной координате z/L; реакции связей в метрику не входят."""
        return np.asarray([np.einsum('pcd,d->pc', b + side * .001 * slope, field) for side in (-1, 0, 1)])
    def TestDot(a, c):
        """Средняя тензорная метрика совпадает по постановке, но вычислена независимо из операторов."""
        return float(np.einsum('spc,spc,p,c->', a, c, weights, tensor) / (3 * sum(weights)))
    records = helpers['TestRows'](root / 'Native/ActualReverse.ReverseIndex.csv')
    index_rows = [row for row in records if row[0] == 'REVERSE']
    assert len(index_rows) > 0
    cases, first_binary = {}, None
    for row in index_rows:
        index, prefix = int(row[1]), int(row[2])
        phase = 'ActualReverse.Reverse' + str(index)
        binary = TestReadBinary(root / 'Native' / (phase + '.bin'))
        proof, rejected_fields = validate(root, phase)
        assert prefix == 26 and len(rejected_fields) == prefix + 1
        assert np.array_equal(rejected_fields[:prefix], fields)
        assert proof['reversed_fields'] == [prefix + 1]
        np.testing.assert_allclose(binary['fields'], fields, rtol=1e-14, atol=1e-300)
        assert np.array_equal(binary['fields'][-1], binary['accepted'])
        assert np.array_equal(binary['fields'][-2], binary['previous'])
        np.testing.assert_allclose(binary['trial'], rejected_fields[-1], rtol=1e-14, atol=1e-300)
        if first_binary is None:
            first_binary = binary
        else:
            assert np.array_equal(first_binary['fields'], binary['fields'])
            assert np.array_equal(first_binary['lambdas'], binary['lambdas'])
        prior = TestStrains(binary['accepted'] - binary['previous'])
        current = TestStrains(binary['trial'] - binary['accepted'])
        prior_length, current_length = np.sqrt(TestDot(prior, prior)), np.sqrt(TestDot(current, current))
        advance = TestDot(prior, current) / prior_length
        assert advance <= 0 and abs(advance - binary['metrics'][0]) < 1e-18
        assert abs(prior_length - binary['metrics'][1]) < 1e-18
        old_load_step = binary['lambdas'][-1] - binary['lambdas'][-2]
        new_load_step = binary['context'][0] - binary['lambdas'][-1]
        load_scale = prior_length / abs(old_load_step)
        augmented = (TestDot(prior, current) + load_scale**2 * old_load_step * new_load_step) / np.sqrt(prior_length**2 + load_scale**2 * old_load_step**2)
        control_error = binary['calibration'][0] * abs(binary['projection'] @ binary['trial'] - binary['context'][1])
        cases[str(index)] = {'equilibrium': proof, 'binary_sha256': binary['sha256'],
            'lambda': float(binary['context'][0]), 'old_load_step': float(old_load_step), 'new_load_step': float(new_load_step),
            'signed_field_advance': float(advance), 'previous_length': float(prior_length), 'trial_length': float(current_length),
            'field_direction_cosine': float(advance / current_length), 'control_residual_MPa': float(control_error),
            'diagnostic_augmented_advance': float(augmented), 'diagnostic_load_scale': float(load_scale),
            'augmented_metric_not_used_to_accept_any_point': True, 'revision_and_control_changes': binary['metrics'][2:].tolist()}
    # Полная история относится именно к штатному запуску, который также сохранил и повторно открыл Results.
    report = (root / 'Native/ActualReverse.txt').read_text('utf-8-sig')
    assert 'TOTAL_COUPLED_WORKBOOK_TIMING: passed=28; failed=0' in report
    assert 'TOTAL_ACTUAL_REVERSE_CAPTURE_TEST: passed=1; failed=0' in report
    assert 'equal=True;' in report and 'WATCHDOG_COMPLETED: exit=0; source unchanged=True' in report
    trace = run_path(str(here / 'Compare-TestSavedCoupledTrace.py'))['TestSavedTrace']
    saved = trace(root, 'ActualReverse')
    saved_lambda = np.asarray([float(saved[f'Coupled.Trace.Capacity.{index}.Lambda']['value']) for index in range(1, 27)])
    saved_fraction = np.asarray([float(saved[f'Coupled.Trace.Capacity.{index}.Offset']['value']) for index in range(1, 27)])
    assert np.array_equal(saved_lambda, first_binary['lambdas'])
    assert np.array_equal(saved_fraction, first_binary['fractions'])
    result = {'actual_accepted_history': accepted, 'rejected_equilibria': cases,
              'native_binary_history_same_in_all_snapshots': True, 'native_binary_loads_equal_saved_Excel_Double': True,
              'numeric_direction_gate_preserved': True, 'spatial_half_step_dimensionless': .001,
              'independent_material_reintegration': False, 'physical_capacity_certified': False}
    output.write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    print('ACTUAL AUTOMATIC 26 FIELD HISTORY AND REJECTED EQUILIBRIA VERIFIED; EXACT BINARY LOADS MATCH SAVED EXCEL; CAPACITY OPEN')
    for index, case in cases.items():
        print('REVERSE', index, 'lambda', case['lambda'], 'advance', case['signed_field_advance'], 'length', case['trial_length'], 'cosine', case['field_direction_cosine'])


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    TestCheck(args.root.resolve(), args.out.resolve())
