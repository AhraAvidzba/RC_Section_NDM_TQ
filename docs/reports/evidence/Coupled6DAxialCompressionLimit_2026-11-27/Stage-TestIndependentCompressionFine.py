"""ДЛЯ ТЕСТОВ. Уточняет независимую кривую около пика сжатия, используя только проверенный префикс до максимума."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json, shutil

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent = base / 'Cold1054ASDIndependentCompressionPeak_20261127/IndependentCompressionPeakTestReplay'
root = base / 'Cold1057ASDIndependentCompressionFine_20261127'
assert not root.exists()
checked = json.loads((parent / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
assert checked['all56plus4_equations_and_six_resultants_verified'] and checked['accepted_states'] == 58
raw = (parent / 'Native/MainCapacityHistory.bin').read_bytes()
assert sha256(raw).hexdigest().upper() == checked['input_history_sha256']
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
r = R(raw)
for index in range(3):
    r.VectorTest()
count, = r.VectorTest()
prefix = 0
for index in range(int(count)):
    fraction, factor = r.VectorTest()
    field = r.VectorTest()
    assert fraction == 1
    if abs(field[0]) <= .00205 and prefix == index:
        prefix = index + 1
assert 2 <= prefix < count and r.position == len(raw)
for directory in ('Input', 'Native', 'docs/regression'):
    (root / directory).mkdir(parents=True, exist_ok=True)
for name in ['EarlySourceProof.json', 'Input/CapacityLoad.bin', 'Input/MaterialContract.json',
             'Native/MainCapacityHistory.bin', 'Native/MainCapacity.Operators.csv', 'docs/regression/Diagnostic.xlsm']:
    shutil.copyfile(parent / name, root / name)
source = json.loads((root / 'EarlySourceProof.json').read_bytes())
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == source['saved_book_sha256']
control = dict(native_prefix_states=prefix, source_wire_total_states=int(count),
               prefix_origin='Previously independently solved and replayed compression path; future fields excluded',
               source_native_history_sha256=sha256(raw).hexdigest().upper(),
               axial_coordinate_step=.00000625, end_compressive_strain_magnitude=.00215,
               wall_limit_seconds=120., no_native_excel_invoked=True,
               material_and_geometry_unchanged=True, physical_capacity_not_assumed=True)
(root / 'Input/CompressionPeakControl.json').write_text(json.dumps(control, sort_keys=True, indent=2) + '\n', encoding='utf-8')
for name in ['Stage-TestIndependentCompressionFine.py', 'Solve-TestCompressionPeakOpenSees.py',
             'Solve-TestPureAxialOpenSees.py', 'Read-TestColdRawCapture.py', 'Check-TestHardHalfSteps.py']:
    shutil.copyfile(here / name, root / name)
prep = dict(parent=str(parent), scope='Independent finer axial curve around compression maximum',
            inherited_prefix_states=prefix, no_native_macro_execution_claimed=True,
            inputs={p.relative_to(root).as_posix(): sha256(p.read_bytes()).hexdigest().upper()
                    for p in root.rglob('*') if p.is_file()})
(root / 'Preparation.json').write_text(json.dumps(prep, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('INDEPENDENT FINE COMPRESSION PREPARED', root, 'PREFIX', prefix, 'OF', int(count), 'STATES;QUARTER STRAIN STEP')
