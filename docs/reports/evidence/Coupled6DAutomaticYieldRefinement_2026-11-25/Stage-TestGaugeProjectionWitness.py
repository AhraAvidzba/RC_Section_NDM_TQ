"""ДЛЯ ТЕСТОВ. Независимо проектирует кандидата780 на точные линейные связи в нулевых модах деформаций и читает фактический отклик."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import struct
import json
import numpy as np

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold775ArcRootPath_20261125'
failed = base / 'Cold780RejectedYieldEvent_20261125'
root = base / 'Cold781GaugeProjectionWitness_20261125'
Reader = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
reader = Reader((failed / 'Native/ArcYieldSplit.Result.bin').read_bytes())
values = [reader.VectorTest() for _ in range(10)]
assert reader.position == len(reader.raw) and not values[0][0] and values[1][16] == 1
field = values[1][17:]; factor = values[1][15]
g, loads, w, v, coords, axis, b, gradient = run_path(str(here / 'Check-TestHardHalfSteps.py'))['TestOperators'](failed / 'Native/EnvelopeEnergy.Accepted.Operators.csv', 60, 43)
strain_operator = np.concatenate([b + side * .001 * gradient for side in (-1, 0, 1)]).reshape(-1, 60)[:, :56]
scales = np.maximum(np.linalg.norm(strain_operator, axis=0), 1e-300)
_, singular, vt = np.linalg.svd(strain_operator / scales, full_matrices=False)
assert singular[-5] > 1e-5 and singular[-4] < 1e-12
null = vt[-4:].T / scales[:, None]
constraints = g[56:, :56]
coefficients = np.linalg.solve(constraints @ null, -constraints @ field[:56])
delta = null @ coefficients
projected = field.copy(); projected[:56] += delta
assert max(abs(strain_operator @ delta)) < 1e-23
assert max(abs(constraints @ projected[:56])) < 1e-17
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0, failed / 'Input/Reverse1.bin')
reader = Reader((failed / 'Input/ArcRoots.bin').read_bytes()); reader.position = 0
_, _, projection, old_context = [reader.VectorTest() for _ in range(4)]
binary = bytearray()
for vector in (field, projected, projection, [factor, factor, 0., 0., old_context[4], old_context[5]]):
    binary += struct.pack('<i' + 'd' * len(vector), len(vector), *vector)
proof = dict(source_candidate_lambda=float(factor), source_candidate_field=field.tolist(), projected_field=projected.tolist(),
    delta_field=np.r_[delta, np.zeros(4)].tolist(), normalized_strain_singular_values=singular.tolist(),
    all_four_null_modes_verified=True, maximum_incremental_strain=float(max(abs(strain_operator @ delta))),
    gauge_before=(constraints @ field[:56]).tolist(), gauge_after=(constraints @ projected[:56]).tolist(),
    multipliers_unchanged=True, no_material_or_load_change=True, no_new_equilibrium_search=True)
values_to_write = {'Input/ArcRoots.bin': binary, 'Input/ArcPathOptions.bin': struct.pack('<i3d', 3, 0., 1., 0.),
    'Input/Source780Candidate.bin': (failed / 'Native/ArcYieldSplit.Result.bin').read_bytes(),
    'GaugeProjectionPreparation.json': (json.dumps(proof, indent=2, sort_keys=True) + '\n').encode('utf-8'),
    Path(__file__).name: Path(__file__).read_bytes()}
prep = json.loads((root / 'Preparation.json').read_bytes())
for name, data in values_to_write.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
assert (root / 'docs/regression/Diagnostic.xlsm').read_bytes() == (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope'].update(case='GaugeProjectionReadOnlyWitness', immutable_compiled255_reused=True,
    fresh_compilation_claimed=False, source_states=43, original_and_null_projected_candidate_compared=True,
    all_working_methods_exact775=True, physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('781 EXACT KINEMATIC GAUGE PROJECTION; FOUR NULL MODES; STRAIN CHANGE', max(abs(strain_operator @ delta)), 'COMPILED775 REUSED')
