"""ДЛЯ ТЕСТОВ. Собирает все уравнения HardShear из сохранённых операторов и напряжений трёх станций."""

from argparse import ArgumentParser
import csv
import json
from pathlib import Path
import numpy as np


def TestRows(file):
    """Читает инвариантные числа и сохраняет смысл человеческой диагностики Windows-1251."""
    with file.open(encoding='cp1251', newline='') as source:
        return [row for row in csv.reader(source, delimiter=';') if row]


def TestValidate(root, phase):
    """Использует фактические поля и тензоры, без вызова численного ядра VBA и изменения допусков решателя."""
    records = TestRows(root / 'Native' / (phase + '.csv'))
    header = {row[0]: row[1:] for row in records if row[0] in ('DIM', 'LAW', 'OPTIONS', 'OFFSET', 'DIRECTION', 'META', 'DETAILS', 'SOLVER_SECONDS')}
    dofs, physical, points, concrete = map(int, header['DIM'][:4])
    half = float(header['DIM'][4])
    law = np.asarray(header['LAW'], float)
    options = np.asarray(header['OPTIONS'], float)
    offset = np.asarray(header['OFFSET'], float)
    direction = np.asarray(header['DIRECTION'], float)
    geometry = np.zeros((dofs, dofs))
    load = np.zeros((dofs, 6))
    b = np.zeros((points, 6, dofs))
    gradient = b.copy()
    w = np.zeros((points, 6, dofs))
    v = w.copy()
    coordinate = np.zeros((points, 3))
    axis = np.zeros((points, 6))
    for kind, *row in TestRows(root / 'Native' / (phase + '.Operators.csv')):
        if kind == 'G': geometry[:, int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'L': load[:, int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'P': coordinate[int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'A': axis[int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'B':
            p, degree = int(row[0]) - 1, int(row[1]) - 1
            b[p, :, degree] = np.asarray(row[2:8], float)
            gradient[p, :, degree] = np.asarray(row[8:], float)
        elif kind in ('W', 'V'):
            (w if kind == 'W' else v)[int(row[0]) - 1, int(row[1]) - 1] = np.asarray(row[2:], float)
        else: raise AssertionError(kind)
    assert np.linalg.matrix_rank(load) == 6
    area = 1 / load[0, 0]
    moment_scale = 1 / load[1, 1]
    assert abs(area - sum(coordinate[:concrete, 2])) < 1e-8
    residual_scales = np.asarray((area, moment_scale, moment_scale, moment_scale, area, area))
    frames = [(int(row[1]), float(row[2]), float(row[3]), np.asarray(row[4:], float)) for row in records if row[0] == 'F']
    count = len(frames)
    stresses = np.zeros((count, 3, points, 6))
    bars = np.zeros((count, 3, points))
    native = {}
    for kind, *row in records:
        if kind == 'S':
            index, side, p = int(row[0]) - 1, int(row[1]) + 1, int(row[2]) - 1
            stresses[index, side, p] = np.asarray(row[3:9], float)
            bars[index, side, p] = float(row[11])
        elif kind == 'R': native[int(row[0])] = np.asarray(row[1:], float)
    net = stresses.copy()
    net[:, :, concrete:] = bars[:, :, concrete:, None] * axis[None, None, concrete:] - stresses[:, :, concrete:]
    weights = coordinate[:, 2]
    center = np.average(coordinate[:concrete, :2], axis=0, weights=weights[:concrete])
    x, y = (coordinate[:, :2] - center).T
    frame_proof, tensors = [], []
    for index, fraction, factor, field in frames:
        assert field.shape == (dofs,) and index == len(frame_proof) + 1
        sigma = net[index - 1, 1]
        target = fraction * offset + factor * direction
        external = load @ target
        residual = geometry @ field + np.einsum('pcd,pc->d', w, sigma)
        residual += np.einsum('pcd,pc->d', v, (net[index - 1, 2] - net[index - 1, 0]) / (2 * half))
        residual -= external
        actual = np.asarray((np.dot(weights, sigma[:, 2]), np.dot(weights * y, sigma[:, 2]),
            np.dot(weights * x, sigma[:, 2]), np.dot(weights, x * sigma[:, 5] - y * sigma[:, 4]),
            np.dot(weights, sigma[:, 4]), np.dot(weights, sigma[:, 5])))
        norm = max(np.max(np.abs(residual[:physical])), law[0] * np.max(np.abs(residual[physical:])))
        force_norm = np.max(np.abs((actual - target) / residual_scales))
        bound = options[3] + options[4] * np.max(np.abs(external[:physical]))
        assert abs(bound - native[index][1]) <= 1e-15
        assert native[index][0] <= native[index][1] and native[index][2] <= 1
        assert norm <= bound + 1e-9 and force_norm <= bound + 1e-9, (phase, index, norm, force_norm, bound)
        center_strain = np.einsum('pcd,d->pc', b, field)
        slope = np.einsum('pcd,d->pc', gradient, field)
        tensors.append(np.asarray([center_strain + side * half * slope for side in (-1, 0, 1)]))
        frame_proof.append({'index': index, 'lambda': factor, 'fraction': fraction,
            'residual_MPa': float(norm), 'force_residual_MPa': float(force_norm), 'bound_MPa': float(bound),
            'steel_utilization': float(native[index][2])})
    weight_tensor = np.asarray((1., 1., 1., .5, .5, .5))
    def TestDot(a, c):
        """Средняя тензорная метрика всех трёх станций без реакций жёстких связей."""
        return float(np.einsum('spc,spc,p,c->', a, c, weights, weight_tensor) / (3 * sum(weights)))
    increments = np.diff(np.asarray(tensors), axis=0)
    orientation = []
    for index in range(1, len(increments)):
        previous, current = increments[index - 1], increments[index]
        square = TestDot(previous, previous)
        orientation.append({'field': index + 2, 'signed_advance': TestDot(previous, current) / np.sqrt(square),
            'previous_length': np.sqrt(square), 'current_length': np.sqrt(TestDot(current, current))})
    proof = {'phase': phase, 'material': law.tolist(), 'options': options.tolist(),
        'states': frame_proof, 'maximum_lambda': max(row['lambda'] for row in frame_proof),
        'orientation': orientation, 'reversed_fields': [row['field'] for row in orientation if row['signed_advance'] <= 0],
        'physical_capacity_certified': False, 'CSV_rounding_allowance_MPa': 1e-9,
        'meta': header['META'], 'details': header['DETAILS']}
    return proof, np.asarray([row[3] for row in frames])


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    results = {}
    arrays = {}
    for phase in ('HardBaseline', 'HardEarlyField'):
        results[phase], arrays[phase] = TestValidate(args.root.resolve(), phase)
    before, after = results.values()
    assert before['material'] == after['material'] and before['options'] == after['options']
    assert np.array_equal(arrays['HardBaseline'], arrays['HardEarlyField'][:len(arrays['HardBaseline'])])
    args.out.write_text(json.dumps(results, indent=2, ensure_ascii=False) + '\n', encoding='utf-8')
    for phase, result in results.items():
        print(phase, 'states', len(result['states']), 'max_lambda', result['maximum_lambda'], 'reversed', result['reversed_fields'])
    print('PREPARED HARDSHEAR ALL SIX FORCES AND SPATIAL EQUATIONS RECONSTRUCTED', flush=True)
