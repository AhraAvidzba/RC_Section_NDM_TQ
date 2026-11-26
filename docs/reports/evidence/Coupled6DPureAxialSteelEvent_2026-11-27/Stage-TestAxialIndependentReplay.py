"""ДЛЯ ТЕСТОВ. Создаёт отдельный снимок полного пути TRF для повторной проверки материала, стали и равновесия."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
import json, struct, shutil

parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
root = parser.parse_args().root
out = root / 'IndependentAxialTRFTestReplay'
assert not out.exists()
proof = json.loads((root / 'Native/IndependentAxialTRF.json').read_bytes())
raw = (root / 'Native/MainCapacityHistory.bin').read_bytes()
assert sha256(raw).hexdigest().upper() == proof['source_history_sha256']
assert proof['termination'] is None and proof['original_tolerances_unchanged']
records = proof['records']
assert len(records) == 21 and all(record['accepted'] for record in records)
R = run_path(str(Path(__file__).resolve().parent / 'Read-TestColdRawCapture.py'))['TestReader']
r = R(raw)
assert raw[:16] == b'NDMMAINHIST00003'
for index in range(3):
    r.VectorTest()
before_count = r.position
count, = r.VectorTest()
records_start = r.position
assert int(count) == proof['inherited_states']


def TestVector(values):
    """Записывает исходный бинарный контракт без округления Double при повторной проверке."""
    return struct.pack('<i' + 'd' * len(values), len(values), *values)


combined = raw[:before_count] + TestVector([count + len(records)]) + raw[records_start:]
for record in records:
    combined += TestVector([1., record['lambda_value']]) + TestVector(record['field'])
for directory in ('Input', 'Native', 'docs/regression'):
    (out / directory).mkdir(parents=True, exist_ok=True)
for name in ['EarlySourceProof.json', 'Input/CapacityLoad.bin', 'Input/MaterialContract.json',
             'Native/MainCapacity.Operators.csv', 'docs/regression/Diagnostic.xlsm']:
    shutil.copyfile(root / name, out / name)
(out / 'Native/MainCapacityHistory.bin').write_bytes(combined)
metadata = dict(native_prefix_states=int(count), independent_driver_states=len(records),
                full_path_states=int(count) + len(records), native_prefix_records_kept_bitexact=True,
                source_history_sha256=sha256(raw).hexdigest().upper(),
                new_history_sha256=sha256(combined).hexdigest().upper(), physical_capacity_certified=False)
(out / 'GeneratedTraceProof.json').write_text(json.dumps(metadata, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('INDEPENDENT AXIAL FULL PATH REPLAY PREPARED', out, metadata['full_path_states'], 'STATES')
