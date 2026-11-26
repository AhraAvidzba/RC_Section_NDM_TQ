"""ДЛЯ ТЕСТОВ. Подставляет точные записи всей деформационной ветви прямого ASD в исходный контракт геометрии и нагрузки для независимого повторения."""
from pathlib import Path
from argparse import ArgumentParser
from runpy import run_path
from hashlib import sha256
import json, shutil

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
root = parser.parse_args().root
here = Path(__file__).resolve().parent
out = root / 'FullRecoveryTestReplay'
assert not out.exists()
source = json.loads((root / 'EarlySourceProof.json').read_bytes())
assert source['changed_components_exact']
assert sha256((root / 'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper() == source['saved_book_sha256']
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
raw = (root / 'Native/MainCapacityHistory.bin').read_bytes()
reader = R(raw)
assert raw[:16] == b'NDMMAINHIST00003'
for index in range(3):
    reader.VectorTest()
header_end = reader.position
full = (root / 'Native/FullFollowedHistory.bin').read_bytes()
assert full[:16] == b'NDMFULLTRACE0001'
reader = R(full)
count, = reader.VectorTest()
assert count == int(count) and count > 1
factors = []
for index in range(int(count)):
    fraction, factor = reader.VectorTest()
    field = reader.VectorTest()
    assert fraction == 1 and len(field) == 60
    factors.append(float(factor))
assert reader.position == len(full)
combined = raw[:header_end] + full[16:]
for directory in ('Input', 'Native', 'docs/regression'):
    (out / directory).mkdir(parents=True, exist_ok=True)
for name in ('EarlySourceProof.json', 'Input/CapacityLoad.bin', 'Input/MaterialContract.json',
             'Native/MainCapacity.Operators.csv', 'docs/regression/Diagnostic.xlsm'):
    shutil.copyfile(root / name, out / name)
(out / 'Native/MainCapacityHistory.bin').write_bytes(combined)
proof = dict(original_history_sha256=sha256(raw).hexdigest().upper(),
    full_history_sha256=sha256(full).hexdigest().upper(), combined_history_sha256=sha256(combined).hexdigest().upper(),
    states=int(count), all_original_Double_bytes_preserved=True,
    maximum_lambda=max(factors), final_lambda=factors[-1],
    physical_acceptance_not_assumed=True, no_new_equilibrium_solve_claimed=True)
(out / 'GeneratedTraceProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('FULL DIRECT RECOVERY REPLAY PREPARED', out, 'STATES', int(count), 'EXACT ORIGINAL BYTES')
