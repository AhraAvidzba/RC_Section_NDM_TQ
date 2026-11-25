"""ДЛЯ ТЕСТОВ. Читает фактический материальный отклик вдоль линии к половинному корню без поиска и принятия НДС."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import struct
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold774ArcRootDerivative_20261125'
root = base / 'Cold775ArcRootPath_20261125'
proof = json.loads((parent / 'ArcRootDerivativeProof.json').read_bytes())
assert proof['all492_raw_difference_samples_verified']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
archive = root / 'ArcPathCompilationAncestor'; archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
name = 'tests/modTestCoupledElasticState.bas'; code = (root / name).read_text('utf-8-sig')
assert code.count('    TestRunArcRootDerivative\n') == 1
code = code.replace('    TestRunArcRootDerivative\n', '    TestRunArcRootPath\n', 1)
code += (here / 'TestArcRootPath.bas.txt').read_text('utf-8-sig')
assert max(map(len, code.splitlines())) < 1024
etas = [0., 1e-8, 1e-6, 1e-4, .001, .01, .1, .25, .5, .75, 1., 2., 0.]
values = {name: code.replace('\n', '\r\n').encode('utf-8-sig'),
    'Input/ArcPathOptions.bin': struct.pack('<i' + 'd' * len(etas), len(etas), *etas),
    'Compile-TestExternalPoint.ps1': (here / 'Compile-TestExternalPoint.ps1').read_bytes(),
    'ParentCodeUpdate.json': (parent / 'docs/regression/CodeUpdate.json').read_bytes(),
    'ParentCompileClosure.json': (parent / 'CompilePointClosure.json').read_bytes(),
    'TestArcRootPath.bas.txt': (here / 'TestArcRootPath.bas.txt').read_bytes(), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data); prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'].update(case='ActualArcRootPathMaterialWitness', actual_import_this_pass=['modTestCoupledElasticState'],
    numerical_working_methods_exact761=True, readonly_exact33_history=True,
    actual770_half_root_and_full_root=True, no_new_equilibrium_search=True,
    direct_all43_material_net_tensors_on_three_stations=True, line_samples=etas,
    accepted_revision_unchanged_required=True, physical_capacity_certified=False, no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('775 READONLY ARC ROOT LINE;13 LINE SAMPLES INCLUDING ZERO REPEAT + EXACT FULL ROOT;129 MATERIAL POINT/STATIONS EACH')
