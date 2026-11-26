"""ДЛЯ ТЕСТОВ. Продолжает чистое N независимым TRF/SVD с настоящим ASDConcrete3D и той же принятой историей; рабочий VBA не меняет."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
from time import perf_counter
import json
import numpy as np
import scipy
from scipy.optimize import least_squares
import openseespy.opensees as ops
import openseespywin.opensees as native

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
parser.add_argument('--seconds', type=float, default=120.)
parser.add_argument('--step', type=float, default=.001)
parser.add_argument('--end', type=float, default=.025)
args = parser.parse_args()
root = args.root
here = Path(__file__).resolve().parent
output = root / 'Native/IndependentAxialTRF.json'
assert not output.exists() and args.step > 0 and args.seconds > 0
source = json.loads((root / 'EarlySourceProof.json').read_bytes())
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == source['saved_book_sha256']
assert ops.version() == '3.8.0' and ops.NDTest is native.NDTest
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
history_raw = (root / 'Native/MainCapacityHistory.bin').read_bytes()
r = R(history_raw)
assert history_raw[:16] == b'NDMMAINHIST00003'
header = r.VectorTest()
dof, points, concrete, area, length, half, absolute, relative = header
dof, points, concrete = map(int, (dof, points, concrete))
offset, direction = r.VectorTest(), r.VectorTest()
assert (dof, points, concrete, half) == (60, 43, 25, .001)
assert (absolute, relative) == (5e-7, 5e-7)
assert np.all(offset == 0) and direction[0] > 0 and np.all(direction[1:] == 0)
count, = r.VectorTest()
history = []
for index in range(int(count)):
    fraction, factor = r.VectorTest()
    assert fraction == 1
    history.append((r.VectorTest(), factor))
assert r.position == len(history_raw)
g, L, w, v, coords, axis, B, gradient = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](
    root / 'Native/MainCapacity.Operators.csv', dof, points)
operators = np.array([B + side * half * gradient for side in (-1, 0, 1)])
assert abs(area * L[0, 0] - 1) < 1e-12
weights = coords[:, 2]
center = np.average(coords[:concrete, :2], weights=weights[:concrete], axis=0)
x, y = (coords[:, :2] - center).T
longitudinal = axis[concrete:, 2] == 1
E = 200000.
H = 1500 / (1 - 1500 / E)
plastic = np.zeros((3, points - concrete))
back = plastic.copy()
fy_t = np.where(longitudinal[None, :], 400., 240.)
fy_c = np.where(longitudinal[None, :], 350., 240.)
contract_raw = (root / 'Input/MaterialContract.json').read_bytes()
contract = json.loads(contract_raw)
assert contract['integration_mode'] == 'native_single_update_working'
command = ['ASDConcrete3D', 30000., .2]
for kind, rows in contract['uniaxial_tables'].items():
    for column, name in enumerate(('e', 's', 'd')):
        command.extend(['-' + kind + name, *(row[column] for row in rows)])
command.append('-tangent')
ops.wipe()
ops.model('basic', '-ndm', 3, '-ndf', 3)
for tag in range(1, 3 * points + 1):
    ops.nDMaterial(command[0], tag, *command[1:])
started = perf_counter()
evaluations = 0


def TestResponse(field, factor):
    """Собирает пробу от неизменённой принятой истории; отклонённая проба не коммитит материал или сталь."""
    global evaluations
    evaluations += 1
    if perf_counter() - started > args.seconds:
        raise TimeoutError('Independent axial TRF wall limit')
    strain = np.einsum('spcd,d->spc', operators, field)
    stress = np.zeros_like(strain)
    for side in range(3):
        for point in range(points):
            xx, yy, zz, xy, xz, yz = strain[side, point]
            tag = side * points + point + 1
            ops.NDTest('SetStrain', tag, xx, yy, zz, xy, yz, xz)
            xx, yy, zz, xy, yz, xz = ops.NDTest('GetStress', tag)
            stress[side, point] = [xx, yy, zz, xy, xz, yz]
    eps = np.einsum('spc,pc->sp', strain[:, concrete:], axis[concrete:])
    trial = E * (eps - plastic)
    difference = trial - back
    delta = np.where(difference > fy_t, (difference - fy_t) / (E + H),
                     np.where(difference < -fy_c, (difference + fy_c) / (E + H), 0.))
    steel = trial - E * delta
    net = stress.copy()
    net[:, concrete:] = steel[:, :, None] * axis[None, concrete:] - stress[:, concrete:]
    residual = g @ field - L @ (factor * direction) + np.einsum('pcd,pc->d', w, net[1])
    residual += np.einsum('pcd,pc->d', v, (net[2] - net[0]) / (2 * half))
    residual[-4:] *= 30000.
    actual = np.array([weights @ net[1, :, 2], (weights * y) @ net[1, :, 2],
                       (weights * x) @ net[1, :, 2], weights @ (x * net[1, :, 5] - y * net[1, :, 4]),
                       weights @ net[1, :, 4], weights @ net[1, :, 5]])
    return residual, strain, eps, actual, delta


def TestCommit(delta):
    """Принимает только предварительно проверенное состояние, включая отдельную пластическую историю стали."""
    global plastic, back
    for tag in range(1, 3 * points + 1):
        ops.NDTest('CommitState', tag)
    plastic = plastic + delta
    back = back + H * delta


records = []
termination = None
try:
    for field, factor in history:
        residual, strain, eps, actual, delta = TestResponse(field, factor)
        bound = absolute + relative * max(abs((L @ (factor * direction))[:-4]))
        assert max(abs(residual)) <= bound + 1e-10
        TestCommit(delta)
    current, factor = history[-1]
    scale = np.max(abs(operators), axis=(0, 1, 2))
    scale = np.divide(.001, scale, out=np.full(dof, .001), where=scale > 0)
    target = float(current[0])
    assert 0 < target < args.end
    while target < args.end:
        target = min(args.end, target + args.step)
        initial = np.r_[current[1:] / scale[1:], factor]

        def TestField(value):
            """Подставляет заданную осевую деформацию; остальные 59 неизвестных и lambda остаются свободными."""
            return np.r_[target, value[:-1] * scale[1:]]

        def TestResidual(value):
            """Возвращает все 56 уравнений и четыре связи с исходным размерным масштабом невязок."""
            return TestResponse(TestField(value), value[-1])[0]

        solution = least_squares(TestResidual, initial, method='trf', jac='3-point', x_scale='jac',
                                 diff_step=1e-5, ftol=1e-12, xtol=1e-12, gtol=1e-12, max_nfev=60)
        field = TestField(solution.x)
        factor = float(solution.x[-1])
        residual, strain, eps, actual, delta = TestResponse(field, factor)
        bound = float(absolute + relative * max(abs((L @ (factor * direction))[:-4])))
        scales = np.array([area, area * length, area * length, area * length, area, area])
        force_norm = float(max(abs((actual - factor * direction) / scales)))
        usage = float(np.max(abs(eps) / np.where(longitudinal[None, :], np.where(eps >= 0, .025, .018), .02)))
        accepted = bool(max(max(abs(residual)), force_norm) <= bound and factor > 0)
        records.append(dict(axial_strain=target, lambda_value=factor, force_N=float(actual[0]),
                            physical_residual_MPa=float(max(abs(residual[:-4]))),
                            gauge_residual_MPa=float(max(abs(residual[-4:]))),
                            six_force_residual_MPa=force_norm, bound_MPa=bound,
                            steel_utilization=usage, solver_status=int(solution.status),
                            solver_message=str(solution.message), accepted=accepted,
                            field=field.tolist(), resultants_N_Nmm=actual.tolist(),
                            concrete_min_longitudinal_strain=float(np.min(strain[:, :concrete, 2])),
                            concrete_max_longitudinal_strain=float(np.max(strain[:, :concrete, 2]))))
        print('AXIAL_TRF', json.dumps({k: value for k, value in records[-1].items() if k not in ('field', 'resultants_N_Nmm')}), flush=True)
        if not accepted:
            termination = 'Independent nonlinear solve did not meet original equilibrium tolerances'
            break
        TestCommit(delta)
        current = field
except TimeoutError as error:
    termination = str(error)
finally:
    ops.wipe()
    proof = dict(engine_version='3.8.0', scipy_version=scipy.__version__,
                 driver='SciPy TRF/SVD; actual OpenSees ASD material; VBA geometric operators',
                 source_history_sha256=sha256(history_raw).hexdigest().upper(),
                 material_contract_sha256=sha256(contract_raw).hexdigest().upper(),
                 source_helper_sha256=sha256(Path(__file__).read_bytes()).hexdigest().upper(),
                 starts_from_verified_native_history=True, inherited_states=len(history),
                 original_tolerances_unchanged=True, prescribed_axial_strain_step=args.step,
                 elapsed_seconds=perf_counter() - started, material_evaluations=evaluations,
                 records=records, termination=termination, physical_capacity_certified=False,
                 independent_OpenSees_structural_analysis=False, whole_goal_completed=False)
    output.write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('AXIAL INDEPENDENT RESULT SAVED', output, 'ACCEPTED', sum(record['accepted'] for record in records), flush=True)
