"""ДЛЯ ТЕСТОВ. Независимо измеряет влияние продольного полушага на одни и те же поля, не объявляя их новыми решениями."""

from argparse import ArgumentParser
from hashlib import sha256
import csv
import json
from pathlib import Path
import numpy as np


def TestRows(file):
    """Читает инвариантные числовые записи и сохранённую диагностику Windows-1251."""
    with file.open(encoding='cp1251', newline='') as stream:
        return [row for row in csv.reader(stream, delimiter=';') if row]


def TestData(file):
    """Отделяет поля, напряжения, повреждения и нативные нормы от смысла статуса решателя."""
    rows = TestRows(file)
    headers = {row[0]: row[1:] for row in rows if row[0] in ('DIM', 'LAW', 'OPTIONS', 'OFFSET', 'DIRECTION', 'MODE')}
    dofs, physical, points, concrete = map(int, headers['DIM'][:4])
    frames = [row for row in rows if row[0] == 'F']
    assert [int(row[1]) for row in frames] == list(range(1, len(frames) + 1))
    fields = np.asarray([row[4:] for row in frames], float)
    fractions = np.asarray([row[2] for row in frames], float)
    factors = np.asarray([row[3] for row in frames], float)
    raw = np.empty((len(frames), 3, points, 9))
    seen = np.zeros((len(frames), 3, points), int)
    native = {}
    for row in rows:
        if row[0] == 'S':
            frame, side, point = int(row[1]) - 1, int(row[2]) + 1, int(row[3]) - 1
            seen[frame, side, point] += 1
            raw[frame, side, point] = np.asarray(row[4:], float)
        elif row[0] == 'R': native[int(row[1])] = np.asarray(row[2:], float)
    assert np.all(seen == 1) and fields.shape == (25, dofs)
    return headers, fields, fractions, factors, raw, native


def TestOperators(file, dofs, points):
    """Считывает только геометрические единичные операторы; все материальные суммы собирает заново."""
    g = np.zeros((dofs, dofs))
    load = np.zeros((dofs, 6))
    w = np.zeros((points, 6, dofs))
    v = w.copy()
    b = np.zeros((points, 6, dofs))
    slope = b.copy()
    coordinate = np.zeros((points, 3))
    axis = np.zeros((points, 6))
    for kind, *row in TestRows(file):
        if kind == 'G': g[:, int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'L': load[:, int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'P': coordinate[int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind == 'A': axis[int(row[0]) - 1] = np.asarray(row[1:], float)
        elif kind in ('W', 'V'):
            (w if kind == 'W' else v)[int(row[0]) - 1, int(row[1]) - 1] = np.asarray(row[2:], float)
        elif kind == 'B':
            point, degree = int(row[0]) - 1, int(row[1]) - 1
            b[point, :, degree] = np.asarray(row[2:8], float)
            slope[point, :, degree] = np.asarray(row[8:], float)
        else: raise AssertionError(kind)
    assert np.linalg.matrix_rank(load) == 6
    return g, load, w, v, coordinate, axis, b, slope


def TestCheck(root, output):
    """Проверяет исходный h по исходным допускам; для остальных h показывает измеренную невязку без поиска равновесия."""
    original = TestData(root / 'BaselineHard.csv')
    original_headers, original_fields, original_fraction, original_factor, original_raw, _ = original
    dofs, physical, points, concrete = map(int, original_headers['DIM'][:4])
    operator_file = root / 'Native/HardEarlyField.Operators.csv'
    assert operator_file.read_bytes() == (root / 'BaselineOperators.csv').read_bytes()
    g, load, w, v, coordinate, axis, b, slope = TestOperators(operator_file, dofs, points)
    area = 1 / load[0, 0]
    moment_scale = 1 / load[1, 1]
    length = moment_scale / area
    weights = coordinate[:, 2]
    center = np.average(coordinate[:concrete, :2], axis=0, weights=weights[:concrete])
    x, y = (coordinate[:, :2] - center).T
    residual_scales = np.asarray((area, moment_scale, moment_scale, moment_scale, area, area))
    results, residuals, gradients = {}, {}, {}
    baseline_raw = None
    rounding_allowance = 1e-9
    declared = json.loads((root / 'Preparation.json').read_bytes())['directed_scope']['half_steps']
    if declared == [.001, .0005, .00025, .002]:
        pairs = [('Base', .001), ('Half', .0005), ('Quarter', .00025), ('Double', .002)]
    else:
        assert len(declared) in (8, 13) and declared == [.001 / 2**index for index in range(len(declared))]
        names = ('Base', 'Half', 'Quarter', 'Eighth', 'Sixteenth', 'ThirtySecond', 'SixtyFourth', 'OneTwentyEighth',
            'TwoFiftySix', 'FiveTwelve', 'TenTwentyFour', 'TwentyFortyEight', 'FortyNinetySix')
        pairs = list(zip(names[:len(declared)], declared))
    for key, half in pairs:
        headers, fields, fractions, factors, raw, native = TestData(root / 'Native' / ('HardEarlyField.' + key + '.csv'))
        assert headers['MODE'] == ['FixedFields', 'NotAnEquilibriumSearch', key]
        assert float(headers['DIM'][4]) == half and headers['DIM'][:4] == original_headers['DIM'][:4]
        assert all(headers[name] == original_headers[name] for name in ('LAW', 'OPTIONS', 'OFFSET', 'DIRECTION'))
        assert np.array_equal(fields, original_fields)
        assert np.array_equal(fractions, original_fraction) and np.array_equal(factors, original_factor)
        if key == 'Base':
            baseline_raw = raw.copy()
            assert np.max(np.abs(raw[..., :6] - original_raw[..., :6])) <= 1e-8
            assert np.max(np.abs(raw[..., 8] - original_raw[..., 8])) <= 1e-8
        else:
            assert np.array_equal(raw[:, 1], baseline_raw[:, 1]), 'central material history changed with h'
        net = raw[..., :6].copy()
        net[:, :, concrete:] = raw[:, :, concrete:, 8, None] * axis[None, None, concrete:] - net[:, :, concrete:]
        # Str(Double) сохраняет 15 значащих цифр; оценка округления учитывает усиление делением на 2h.
        net_rounding = 5e-15 * np.abs(raw[..., :6])
        net_rounding[:, :, concrete:] += 5e-15 * np.abs(raw[:, :, concrete:, 8, None] * axis[None, None, concrete:])
        derivative = (net[:, 2] - net[:, 0]) / (2 * half)
        frame_proofs, actual_residuals = [], []
        law, options = np.asarray(headers['LAW'], float), np.asarray(headers['OPTIONS'], float)
        offset, direction = np.asarray(headers['OFFSET'], float), np.asarray(headers['DIRECTION'], float)
        for index, (field, fraction, factor) in enumerate(zip(fields, fractions, factors), 1):
            target = fraction * offset + factor * direction
            external = load @ target
            stress = net[index - 1, 1]
            residual = g @ field + np.einsum('pcd,pc->d', w, stress)
            residual += np.einsum('pcd,pc->d', v, derivative[index - 1])
            residual -= external
            resultants = np.asarray((np.dot(weights, stress[:, 2]), np.dot(weights * y, stress[:, 2]),
                np.dot(weights * x, stress[:, 2]), np.dot(weights, x * stress[:, 5] - y * stress[:, 4]),
                np.dot(weights, stress[:, 4]), np.dot(weights, stress[:, 5])))
            gauge = law[0] * np.max(np.abs(residual[physical:]))
            norm = max(np.max(np.abs(residual[:physical])), gauge)
            force_norm = np.max(np.abs((resultants - target) / residual_scales))
            bound = options[3] + options[4] * np.max(np.abs(external[:physical]))
            sigma_error = net_rounding[index - 1, 1]
            derivative_error = (net_rounding[index - 1, 2] + net_rounding[index - 1, 0]) / (2 * half)
            error = np.einsum('pcd,pc->d', np.abs(w), sigma_error + 5e-15 * np.abs(stress))
            error += np.einsum('pcd,pc->d', np.abs(v), derivative_error + 5e-15 * np.abs(derivative[index - 1]))
            error += 1e-14 * (np.abs(g) @ np.abs(field) + np.abs(external))
            assembly_allowance = max(rounding_allowance, 2 * max(np.max(error[:physical]), law[0] * np.max(error[physical:])))
            assert abs(native[index][1] - bound) < 1e-15
            assert abs(native[index][0] - norm) <= assembly_allowance
            assert abs(native[index][3] - force_norm) <= rounding_allowance
            assert abs(native[index][4] - gauge) <= rounding_allowance
            assert force_norm <= bound + rounding_allowance
            if key == 'Base': assert norm <= bound + rounding_allowance
            frame_proofs.append({'index': index, 'lambda': float(factor), 'residual_MPa': float(norm),
                'force_residual_MPa': float(force_norm), 'bound_MPa': float(bound),
                'assembly_rounding_allowance_MPa': float(assembly_allowance),
                'residual_to_bound_ratio': float(norm / bound),
                'within_original_tolerance': bool(norm <= bound + rounding_allowance)})
            actual_residuals.append(residual)
        residuals[key] = np.asarray(actual_residuals)
        gradients[key] = derivative
        results[key] = {'half_step': half, 'half_station_distance_mm': half * length,
            'states': frame_proofs, 'fields_within_original_tolerance': [row['index'] for row in frame_proofs if row['within_original_tolerance']],
            'maximum_residual_MPa': max(row['residual_MPa'] for row in frame_proofs),
            'maximum_residual_to_bound_ratio': max(row['residual_to_bound_ratio'] for row in frame_proofs),
            'equilibrium_solved': False, 'capacity_searched': False}
    comparison = []
    refined_keys = [key for key, half in pairs if half <= .001]
    for index in range(25):
        full_to_half = float(np.max(np.abs(residuals['Base'][index, :physical] - residuals['Half'][index, :physical])))
        half_to_quarter = float(np.max(np.abs(residuals['Half'][index, :physical] - residuals['Quarter'][index, :physical])))
        changes = [float(np.max(np.abs(residuals[a][index, :physical] - residuals[c][index, :physical])))
            for a, c in zip(refined_keys, refined_keys[1:])]
        comparison.append({'index': index + 1, 'base_to_half_residual_change_MPa': full_to_half,
            'half_to_quarter_residual_change_MPa': half_to_quarter,
            'successive_change_ratio': full_to_half / half_to_quarter if half_to_quarter else None,
            'successive_refinement_changes_MPa': changes,
            'successive_refinement_ratios': [a / c if c else None for a, c in zip(changes, changes[1:])]})
    finest = refined_keys[-1]
    change = residuals[finest][:, :physical] - residuals['Base'][:, :physical]
    frame, row = np.unravel_index(np.argmax(np.abs(change)), change.shape)
    contributions = np.einsum('pc,pc->p', v[:, :, row], gradients[finest][frame] - gradients['Base'][frame])
    assert abs(sum(contributions) - change[frame, row]) < 1e-8
    largest = [{'point': int(point + 1), 'x_mm': float(coordinate[point, 0]), 'y_mm': float(coordinate[point, 1]),
        'kind': 'bulk_concrete' if point < concrete else 'steel_replacement',
        'residual_change_MPa': float(contributions[point])} for point in np.argsort(np.abs(contributions))[-5:][::-1]]
    proof = {'cases': results, 'residual_refinement': comparison,
        'all_fields_exact': True, 'central_material_histories_exact_between_half_steps': True,
        'baseline_concrete_stress_roundtrip_error_MPa': float(np.max(np.abs(baseline_raw[..., :6] - original_raw[..., :6]))),
        'baseline_steel_stress_roundtrip_error_MPa': float(np.max(np.abs(baseline_raw[..., 8] - original_raw[..., 8]))),
        'baseline_stress_roundtrip_allowance_MPa': 1e-8, 'CSV_assembly_rounding_allowance_MPa': rounding_allowance,
        'operator_sha256': sha256(operator_file.read_bytes()).hexdigest().upper(),
        'largest_residual_change': {'field': int(frame + 1), 'row': int(row + 1), 'finest_case': finest,
            'change_MPa': float(change[frame, row]), 'largest_point_contributions': largest},
        'length_scale_mm': length, 'independent_material_reintegration': False,
        'capacity_searched': False, 'working_code_changed': False}
    output.write_text(json.dumps(proof, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    for key, value in results.items():
        print(key, 'h', value['half_step'], 'max R/tol', value['maximum_residual_to_bound_ratio'],
            'within', len(value['fields_within_original_tolerance']), '/25', flush=True)
    print('FIXED FIELDS HALF-STEP DEPENDENCE AND ALL SIX FORCES INDEPENDENTLY RECONSTRUCTED; NO NEW EQUILIBRIUM CLAIM', flush=True)


if __name__ == '__main__':
    parser = ArgumentParser(description=__doc__)
    parser.add_argument('--root', type=Path, required=True)
    parser.add_argument('--out', type=Path, required=True)
    args = parser.parse_args()
    TestCheck(args.root.resolve(), args.out.resolve())
