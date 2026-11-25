"""ДЛЯ ТЕСТОВ. Независимо строит производные дуговой невязки и оценивает оставшуюся поправку поля у фактических корней."""
from pathlib import Path
from runpy import run_path
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
root = base / 'Cold774ArcRootDerivative_20261125'
parent = base / 'Cold765AcceptedMapContinuity_20261125'
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
for name in ('src/Materials/CCoupledConcreteResponse.cls', 'src/Materials/CCoupledConcreteReturn.cls', 'src/Solver/CCoupledDeformationContinuation.cls'):
    assert (root / name).read_bytes() == (parent / name).read_bytes(), name
expected = (parent / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
actual = (root / 'src/Solver/CCoupledNewtonSolver.cls').read_text('utf-8-sig')
assert actual.startswith(expected) and actual[len(expected):].count('Friend Function BeginArcRootTest(') == 1
assert 'mArcLoadProjection = loadProjection: mConsistentControl = True' in actual[len(expected):]
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
source = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](root / 'Input/Reverse1.bin')
assert len(source['fields']) == 33
read = Reader((root / 'Input/SourceArcWalk001.bin').read_bytes())
first = [read.VectorTest() for _ in range(13)]
assert np.array_equal(source['fields'][-1], first[9]) and source['lambdas'][-1] == first[0][2]
read = Reader((root / 'Input/SourceArcWalk002.bin').read_bytes())
pair = [read.VectorTest() for _ in range(13)]
validate = run_path(str(here / 'Validate-TestHardPrepared.py'))
mechanics, fields = validate['TestValidate'](root, 'EnvelopeEnergy.Accepted')
np.testing.assert_allclose(fields, source['fields'], rtol=1e-14, atol=1e-300)
rows = validate['TestRows'](root / 'Native/EnvelopeEnergy.Accepted.csv')
dim = next(r for r in rows if r[0] == 'DIM')
dofs, physical, points = map(int, dim[1:4])
dz = float(dim[5])
b = np.zeros((points, 6, dofs)); gradient = b.copy(); areas = np.zeros(points)
for kind, *r in validate['TestRows'](root / 'Native/EnvelopeEnergy.Accepted.Operators.csv'):
    if kind == 'B':
        point, degree = int(r[0]) - 1, int(r[1]) - 1
        b[point, :, degree], gradient[point, :, degree] = np.asarray(r[2:8], float), np.asarray(r[8:], float)
    elif kind == 'P':
        areas[int(r[0]) - 1] = float(r[3])
operators = np.asarray([b + side * dz * gradient for side in (-1, 0, 1)])
metric = np.einsum('spcd,spce,p,c->de', operators, operators, areas, np.array([1., 1., 1., .5, .5, .5])) / (3 * sum(areas))
metric[physical:, :] = 0; metric[:, physical:] = 0


def TestLength(delta):
    """Измеряет поправку в общей физической метрике трёх станций, исключая множители связей."""
    return float(np.sqrt(max(0., delta @ metric @ delta)))


matrices = {}
reader = Reader((root / 'Native/ArcRoots.Matrix.bin').read_bytes())
assert reader.raw[:16] == b'NDMARCROOTS00001'
for index in range(3):
    info, field, projection, scales, residual, matrix = [reader.VectorTest() for _ in range(6)]
    assert info[0] == index and len(field) == dofs and len(scales) == len(residual) == dofs + 1
    assert len(matrix) == (dofs + 1)**2 and info[4] == 35
    expected_field = source['fields'][-1] if index == 0 else pair[8 + index]
    assert np.array_equal(field, expected_field) and np.array_equal(projection, pair[5])
    expected_factor = source['lambdas'][-1] if index == 0 else pair[0][1 + index]
    assert info[1] == expected_factor and info[3] == pair[0][9]
    matrices[index] = dict(info=info, field=field, projection=projection, scales=scales, residual=residual, matrix=matrix.reshape(dofs + 1, dofs + 1))
assert reader.position == len(reader.raw)
reader = Reader((root / 'Native/ArcRoots.Samples.bin').read_bytes())
assert reader.raw[:16] == b'NDMARCROOTDIFF01'
samples = {1: [], 2: []}
while reader.position < len(reader.raw):
    info, coordinates, residual = [reader.VectorTest() for _ in range(3)]
    index, step, column, sign, ready = info
    index, column, sign = int(index), int(column), int(sign)
    assert index in (1, 2) and ready == 1 and len(coordinates) == len(residual) == dofs + 1
    expected = np.zeros(dofs + 1)
    if column:
        expected[column - 1] = sign * step
    assert np.array_equal(coordinates, expected)
    samples[index].append(dict(step=step, column=column, sign=sign, residual=residual))
assert sum(map(len, samples.values())) == 492
proofs = []
corrections = {}
for index, data in matrices.items():
    variants = {'assembled': data['matrix']}
    repeated = None
    if index:
        cases = samples[index]
        assert len(cases) == 246 and cases[0]['column'] == cases[-1]['column'] == 0
        assert np.array_equal(cases[0]['residual'], cases[-1]['residual'])
        assert np.array_equal(cases[0]['residual'], data['residual'])
        repeated = True
        for step in (1e-6, 1e-8):
            matrix = np.zeros((dofs + 1, dofs + 1))
            for column in range(1, dofs + 2):
                plus, = [s['residual'] for s in cases if s['step'] == step and s['column'] == column and s['sign'] == 1]
                minus, = [s['residual'] for s in cases if s['step'] == step and s['column'] == column and s['sign'] == -1]
                matrix[:, column - 1] = (plus - minus) / (2 * step)
            variants[f'central_{step:.0e}'] = matrix
    result = dict(case=index, factor=float(data['info'][1]), norm_MPa=float(max(abs(data['residual']))),
        zero_response_repeated_bitidentical=repeated, corrections={})
    for label, matrix in variants.items():
        norms = np.linalg.norm(matrix, axis=0)
        assert min(norms) > 0
        scaled = matrix / norms
        singular = np.linalg.svd(scaled, compute_uv=False)
        correction = np.linalg.solve(scaled, -data['residual']) / norms
        physical_correction = data['scales'][:-1] * correction[:-1]
        lambda_correction = data['scales'][-1] * correction[-1]
        linear_residual = float(max(abs(matrix @ correction + data['residual'])))
        assert linear_residual <= max(1e-22, 1e-6 * max(abs(data['residual'])))
        increment = TestLength(data['field'] - source['fields'][-1])
        item = dict(column_scaled_condition=float(singular[0] / singular[-1]),
            correction_tensor_length=TestLength(physical_correction), correction_lambda=float(lambda_correction),
            tensor_increment=increment, correction_to_increment=(TestLength(physical_correction) / increment if increment else None),
            linear_residual_MPa=linear_residual,
            relative_frobenius_difference_from_assembled=float(np.linalg.norm(matrix - data['matrix']) / np.linalg.norm(matrix)))
        result['corrections'][label] = item
        corrections[f'{index}:{label}'] = dict(field=physical_correction.tolist(), lambda_value=float(lambda_correction))
    proofs.append(result)
proof = dict(native=native, source_states=33, original_770_source_history_and_roots_exact=True,
    numerical_working_methods_unchanged=True, all492_raw_difference_samples_verified=True,
    accepted_history_independent_equilibrium=mechanics, roots=proofs, physical_corrections=corrections,
    no_new_equilibrium_search=True, physical_capacity_certified=False, whole_goal_completed=False)
(root / 'ArcRootDerivativeProof.json').write_text(json.dumps(proof, ensure_ascii=False, indent=2, sort_keys=True) + '\n', encoding='utf-8')
for r in proofs:
    print('ROOT', r['case'], 'RESIDUAL', r['norm_MPa'], 'CORRECTIONS', r['corrections'])
