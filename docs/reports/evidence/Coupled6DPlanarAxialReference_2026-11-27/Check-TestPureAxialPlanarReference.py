"""ДЛЯ ТЕСТОВ. Независимо интегрирует осевое НДС плоской НДМ и предел по фактическим диаграммам и предельным деформациям."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json, struct, numpy as np

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
root = parser.parse_args().root
here = Path(__file__).resolve().parent
native, text = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](root, 'RealColdReplay')
assert native == dict(passed=1, failed=0, macros=1, timeout600=False, own_process_closed=True)
source = json.loads((root / 'PlanarAxialSourceProof.json').read_bytes())
assert source['only_new_axial_test_method'] and source['actual_planar_kernel_matches_current_output']
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == source['saved_book_sha256']
status = (root / 'Native/PlanarAxialReference.txt').read_text('cp1251')
assert status.startswith('CONVERGED;True;LoadMultiplier;')
count, base = struct.unpack('<id', (root / 'Input/PlanarAxialBase.bin').read_bytes())
assert count == 1 and base != 0
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
reader = R((root / 'Native/PlanarAxialReference.bin').read_bytes())
assert reader.raw[:16] == b'NDMPLANEN0000001'
meta, field, limits = reader.VectorTest(), reader.VectorTest(), reader.VectorTest()
factor, mx, force, my, seconds = meta
nc, ns, npc, nps = map(int, limits[4:8])
concrete = np.array([reader.VectorTest() for _ in range(npc)])
steel = np.array([reader.VectorTest() for _ in range(nps)])
geometry = np.array([reader.VectorTest() for _ in range(nc)])
bars = np.array([reader.VectorTest() for _ in range(ns)])
assert reader.position == len(reader.raw) and nc == 400 and ns == 2
assert abs(force - factor * base) <= limits[8] and abs(mx) <= limits[9] and abs(my) <= limits[10]
assert concrete[-1, 0] == 0 and concrete[-1, 1] == 0


def TestPlane(epsilon0, kx, ky):
    """Интегрирует реальные узлы диаграмм и геометрию с вычитанием бетона в точках продольной стали."""
    ec = epsilon0 + kx * geometry[:, 1] + ky * geometry[:, 0]
    es = epsilon0 + kx * bars[:, 1] + ky * bars[:, 0]
    sc = np.interp(ec, concrete[:, 0], concrete[:, 1])
    ss = np.interp(es, steel[:, 0], steel[:, 1]) - np.interp(es, concrete[:, 0], concrete[:, 1])
    forces = np.array([geometry[:, 2] @ sc + bars[:, 2] @ ss,
                       (geometry[:, 2] * geometry[:, 1]) @ sc + (bars[:, 2] * bars[:, 1]) @ ss,
                       (geometry[:, 2] * geometry[:, 0]) @ sc + (bars[:, 2] * bars[:, 0]) @ ss])
    usage = max(float(np.max(-ec)) / abs(limits[0]), float(np.max(-es)) / abs(limits[2]), float(np.max(es)) / limits[3])
    return forces, usage, ec, es


actual, usage, ec, es = TestPlane(*field[:3])
assert np.all(np.isfinite(actual)) and abs(field[1]) < 1e-12 and abs(field[2]) < 1e-12
assert np.max(abs(actual - field[3:6]) / limits[8:11]) <= 1.000001
assert np.max(abs(actual - np.array([force, mx, my])) / limits[8:11]) <= 1.000001
np.testing.assert_allclose([ec.min(), ec.max(), es.min(), es.max()], field[6:10], atol=1e-13, rtol=0)
assert abs(usage - 1) <= 1.1e-7 / min(abs(limits[0]), abs(limits[2]), limits[3])
limit_epsilon = limits[3] if base > 0 else -min(abs(limits[0]), abs(limits[2]))
independent, independent_usage, _, _ = TestPlane(limit_epsilon, 0., 0.)
assert abs(independent_usage - 1) < 1e-12
assert abs(independent[0] - force) <= .1
proof = dict(native=native, base_N_N=base, lambda_capacity=float(factor), capacity_N=float(force),
             elapsed_seconds=float(seconds), solution_method='stock planar LoadMultiplier via pure axial redirect',
             limit_state=status.splitlines()[0], deformation_plane=field[:3].tolist(),
             returned_resultants_N_Nmm=field[3:6].tolist(), independently_integrated_N_Nmm=actual.tolist(),
             independent_limit_N=float(independent[0]), independent_limit_epsilon=float(limit_epsilon),
             physical_strain_utilization=float(usage), concrete_curve=concrete.tolist(), steel_curve=steel.tolist(),
             artificial_diagram_extension=False, concrete_tension_active=False,
             replacement_concrete_subtracted_at_bars=True, planar_reference_physical_limit_verified=True,
             spatial_materials_and_steel_hardening_differ=True, coupled_capacity_not_certified=True,
             whole_goal_completed=False)
(root / 'PureAxialPlanarReferenceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('PLANAR PURE N LIMIT VERIFIED', json.dumps({k: v for k, v in proof.items() if k not in ('concrete_curve', 'steel_curve')}, sort_keys=True))
