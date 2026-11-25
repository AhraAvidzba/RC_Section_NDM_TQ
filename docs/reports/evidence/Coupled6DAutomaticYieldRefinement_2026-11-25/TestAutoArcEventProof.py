"""ДЛЯ ТЕСТОВ. Независимая проверка одного автоматически выделенного события на текущей принятой истории."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np


def TestValidateAutoEvent(root, number, source, history, factors, record, pair, operators, metric, resolution_scale=1.):
    """Проверяет точный источник, поверхность, половинный подход, силы и локальность; возвращает только проверенное принятое событие."""
    here = Path(__file__).resolve().parent
    Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
    path = root / f'Native/AutoEvent.{number:03d}.Result.bin'
    reader = Reader(path.read_bytes())
    assert reader.raw[:16] == b'NDMAUTOSPLIT0001'
    info, event, effort, half_effort, restoration, face, half, projection, actual_face, actual_half, parameters = [reader.VectorTest() for _ in range(11)]
    assert reader.position == len(reader.raw)
    solved, half_solved, verified, lf, lh, revision, reported_error, reported_lambda_error, reported_bound, spacing, target, half_target = info
    assert revision == event[-1] == record[8]
    assert effort[0] <= 15 and effort[0] + half_effort[0] + restoration[0] <= 30
    expected_calls = 1 + int(bool(solved)) + int(bool(verified))
    result = dict(attempt=number, source_states=len(history), source_revision=int(revision),
        solved=bool(solved), half_solved=bool(half_solved), accepted=bool(verified), station=event.tolist(),
        nonlinear_calls=expected_calls, primary_steps=int(effort[0] + half_effort[0] + restoration[0]), event_lambda=float(lf))
    geometry, loads, w, v, coordinates, axis, b, gradient = operators
    accepted = history[-1]; reference = factors[-1]
    sf, sh = pair[0][:2]
    bracket = pair[9] if sf else pair[10]
    bracket_lambda = pair[0][2] if sf else pair[0][3]
    assert sf or sh

    def TestLength(delta):
        """Возвращает длину события по физической метрике трёх станций."""
        return float(np.sqrt(max(0., delta @ metric @ delta)))

    if event[0]:
        point, side = map(int, event[1:3])
        assert 1 <= point <= len(coordinates) and side in (-1, 0, 1) and len(parameters) == 10
        alpha, gamma = parameters[:2]; plastic = parameters[2:8]; cc, ct = parameters[8:]
        modulus, poisson = source['calibration'][:2]
        off = modulus * poisson / ((1 + poisson) * (1 - 2 * poisson))
        shear = modulus / (2 * (1 + poisson))
        elastic = np.diag([off + 2 * shear] * 3 + [shear] * 3)
        elastic[:3, :3] += off * (np.ones((3, 3)) - np.eye(3))

        def TestYield(field):
            """Строит пробную поверхность от захваченных до принятия события координат истории."""
            strain = (b[point - 1] + side * .001 * gradient[point - 1]) @ field
            xx, yy, zz, xy, xz, yz = elastic @ (strain - plastic)
            tensor = np.array([[xx, xy, xz], [xy, yy, yz], [xz, yz, zz]])
            trace = xx + yy + zz; high = np.linalg.eigvalsh(tensor)[-1]
            q = np.sqrt(1.5 * np.sum((tensor - np.eye(3) * trace / 3)**2))
            term = (cc / ct * (1 - alpha) - (1 + alpha)) * high if high > 0 else gamma * high
            return float((alpha * trace + q + term) / (1 - alpha) - cc)

        np.testing.assert_allclose([TestYield(accepted), TestYield(bracket)], event[3:5], rtol=1e-6, atol=2e-12)
        assert event[3] < -32e-12 and event[4] > 0 and 0 <= event[5] <= 1
        if solved:
            face_yield = TestYield(face)
            result['independent_event_yield_MPa'] = face_yield
            assert abs(face_yield) <= 2.1e-12
    else:
        assert not solved and not verified and len(parameters) == 1
    if solved and half_solved:
        expected_spacing = TestLength(face - accepted)
        np.testing.assert_allclose(spacing, expected_spacing, rtol=1e-9, atol=1e-25)
        np.testing.assert_allclose(projection, metric @ (face - accepted) / spacing, rtol=1e-8, atol=1e-17)
        np.testing.assert_allclose([target, half_target], [projection @ face, projection @ accepted + .5 * spacing], rtol=0, atol=1e-14)
        error = TestLength(face - (2 * half - accepted)) / spacing
        lambda_error = abs(2 * lh - reference - lf)
        lambda_bound = max(.005 * abs(lf - reference), 8 * source['options'][4] * .0001 * (1 + abs(reference)))
        np.testing.assert_allclose([reported_error, reported_lambda_error, reported_bound], [error, lambda_error, lambda_bound], rtol=1e-6, atol=1e-14)
        bound_original = source['options'][3] + source['options'][4] * max(abs((loads @ (source['offset'] + lf * source['direction']))[:56]))
        half_original = source['options'][3] + source['options'][4] * max(abs((loads @ (source['offset'] + lh * source['direction']))[:56]))
        resolved = .5 * spacing * source['calibration'][0] > 8 * resolution_scale * max(bound_original, half_original)
        arc_progress = float(pair[5] @ (face - accepted) + pair[0][9] * (lf - reference))
        local = TestLength(face - accepted) <= TestLength(bracket - accepted) * (1 + 1e-10) and arc_progress > 0
        result.update(tensor_half_error=float(error), lambda_half_error=float(lambda_error), lambda_half_bound=float(lambda_bound),
            half_resolution_MPa=float(.5 * spacing * source['calibration'][0]), original_force_bound_MPa=float(bound_original),
            resolved_at_selected_tolerance=bool(resolved), resolution_scale=resolution_scale, local_with_positive_arc_progress=bool(local), arc_progress=arc_progress,
            independent_half_yield_MPa=TestYield(half))
        if verified:
            assert error <= .005 and lambda_error <= lambda_bound and lf > 0 and lh > 0 and TestYield(half) < 0
            assert resolved and local, result
            assert restoration[0] == 0
            scales = np.array([1 / loads[0, 0], 1 / loads[1, 1], 1 / loads[2, 2], 1 / loads[3, 3], 1 / loads[0, 0], 1 / loads[0, 0]])
            for label, field, factor, actual, e in (('face', face, lf, actual_face, restoration), ('half', half, lh, actual_half, half_effort)):
                bound = .0001 * (source['options'][3] + source['options'][4] * max(abs((loads @ (source['offset'] + factor * source['direction']))[:56])))
                force = float(max(abs((actual - source['offset'] - factor * source['direction']) / scales)))
                gauge = float(max(abs(geometry[56:, :] @ field)) * source['calibration'][0])
                assert e[-1] <= bound + 1e-14 and force <= bound + 1e-14 and gauge <= bound + 1e-13
                result[label + '_raw_six_residual_MPa'] = force
    if verified:
        assert solved and half_solved and restoration[0] == 0
    (root / f'Native/AutoEvent.{number:03d}.Proof.json').write_text(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
    return result, face if verified else None, float(lf)
