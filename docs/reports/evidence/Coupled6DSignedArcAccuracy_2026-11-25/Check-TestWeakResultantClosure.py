"""ДЛЯ ТЕСТОВ. Независимо восстанавливает тождества шести интегральных усилий из операторов слабых уравнений."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import numpy as np

here = Path(__file__).resolve().parent
root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold753IntegratedColdFrozenReplay_20261125')
mod = run_path(str(here / 'Validate-TestHardPrepared.py'))
rows = mod['TestRows'](root / 'Native/EnvelopeEnergy.Accepted.csv')
dimension = next(row for row in rows if row[0] == 'DIM')
dofs, physical, points, concrete = map(int, dimension[1:5])
file = root / 'Native/EnvelopeEnergy.Accepted.Operators.csv'
helpers = run_path(str(here / 'Check-TestHardHalfSteps.py'))
geometry, load, w, v, coordinates, axis, _, _ = helpers['TestOperators'](file, dofs, points)
area = 1 / load[0, 0]
moment_scale = 1 / load[1, 1]
assert abs(area - coordinates[:concrete, 2].sum()) < 1e-8
center = np.average(coordinates[:concrete, :2], axis=0, weights=coordinates[:concrete, 2])
x, y = (coordinates[:, :2] - center).T
weight = coordinates[:, 2]
raw = np.zeros((6, points, 6))
raw[0, :, 2] = weight / area
raw[1, :, 2] = weight * y / moment_scale
raw[2, :, 2] = weight * x / moment_scale
raw[3, :, 4], raw[3, :, 5] = -weight * y / moment_scale, weight * x / moment_scale
raw[4, :, 4], raw[5, :, 5] = weight / area, weight / area
names = ('central_stress', 'longitudinal_gradient', 'geometry_gauges', 'six_external_loads')
matrices = (w[:, :, :physical].reshape(points * 6, physical),
    v[:, :, :physical].reshape(points * 6, physical), geometry[:physical, :].T, load[:physical, :].T)
targets = (raw.reshape(6, points * 6).T, np.zeros((points * 6, 6)), np.zeros((dofs, 6)),
    np.diag(1 / np.asarray([area, moment_scale, moment_scale, moment_scale, area, area])))
scales = [max(float(np.max(abs(a))), float(np.max(abs(b))), 1e-300) for a, b in zip(matrices, targets)]
assembled = np.concatenate([a / scale for a, scale in zip(matrices, scales)])
right = np.concatenate([b / scale for b, scale in zip(targets, scales)])
projection, _, rank, singular = np.linalg.lstsq(assembled, right, rcond=1e-12)
residuals = {}
for name, a, target, scale in zip(names, matrices, targets, scales):
    error = a @ projection - target
    residuals[name] = dict(maximum_absolute=float(np.max(abs(error))), maximum_scaled=float(np.max(abs(error / scale))))
verified = max(value['maximum_scaled'] for value in residuals.values()) < 1e-10
proof = dict(source_root=str(root), source_operators_sha256=sha256(file.read_bytes()).hexdigest().upper(),
    independent_resultant_projection_from_operators=True, projection_shape=list(projection.shape), linear_rank=int(rank),
    scaled_matrix_singular_values=singular.tolist(), residuals=residuals,
    all6_weak_and_integral_forces_algebraically_equivalent=verified,
    resultant_residual_max_norm_amplification_bound=np.sum(abs(projection), axis=0).tolist(),
    projection=projection.tolist(), no_material_update_or_new_NDS=True, physical_capacity_certified=False)
out = Path('D:/CodexScratch/RC_Section_NDM_TQ/WeakResultantClosure_20261125.json')
out.write_text(json.dumps(proof, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('ALL6 WEAK/RESULTANT IDENTITY', verified, 'RANK', rank, 'RESIDUALS', residuals)
print('MAXNORM AMPLIFICATION', proof['resultant_residual_max_norm_amplification_bound'])
