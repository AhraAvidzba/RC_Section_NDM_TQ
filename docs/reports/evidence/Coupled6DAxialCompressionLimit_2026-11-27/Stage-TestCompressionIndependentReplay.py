"""ДЛЯ ТЕСТОВ. Сохраняет только использованный докритический префикс и независимые состояния сжатия для полного повторения."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
import json, struct, shutil

root = Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold1054ASDIndependentCompressionPeak_20261127')
out = root / 'IndependentCompressionPeakTestReplay'
assert not out.exists()
driver = json.loads((root / 'Native/IndependentCompressionPeakTRF.json').read_bytes())
raw = (root / 'Native/MainCapacityHistory.bin').read_bytes()
assert sha256(raw).hexdigest().upper() == driver['source_history_sha256']
assert driver['termination'] is None and driver['inherited_states'] == 15
records = driver['records']
assert len(records) == 43 and all(record['accepted'] for record in records)
R = run_path(str(Path(__file__).resolve().parent / 'Read-TestColdRawCapture.py'))['TestReader']
r = R(raw)
assert raw[:16] == b'NDMMAINHIST00003'
for index in range(3):
    r.VectorTest()
before_count = r.position
count, = r.VectorTest()
records_start = r.position
for index in range(int(count)):
    r.VectorTest()
    r.VectorTest()
    if index + 1 == driver['inherited_states']:
        prefix_end = r.position
assert r.position == len(raw)


def TestVector(values):
    """Записывает тот же little-endian контракт Double без десятичного округления и изменения исходных записей."""
    return struct.pack('<i' + 'd' * len(values), len(values), *values)


combined = raw[:before_count] + TestVector([driver['inherited_states'] + len(records)]) + raw[records_start:prefix_end]
for record in records:
    combined += TestVector([1., record['lambda_value']]) + TestVector(record['field'])
for directory in ('Input', 'Native', 'docs/regression'):
    (out / directory).mkdir(parents=True, exist_ok=True)
for name in ['EarlySourceProof.json', 'Input/CapacityLoad.bin', 'Input/MaterialContract.json',
             'Native/MainCapacity.Operators.csv', 'docs/regression/Diagnostic.xlsm']:
    shutil.copyfile(root / name, out / name)
(out / 'Native/MainCapacityHistory.bin').write_bytes(combined)
metadata = dict(source_native_states=int(count), inherited_native_states=driver['inherited_states'],
                independent_states=len(records), total_states=driver['inherited_states'] + len(records),
                unused_future_native_states_excluded=True, inherited_records_kept_bitexact=True,
                source_history_sha256=sha256(raw).hexdigest().upper(),
                new_history_sha256=sha256(combined).hexdigest().upper(), physical_capacity_certified=False)
(out / 'GeneratedTraceProof.json').write_text(json.dumps(metadata, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('COMPRESSION INDEPENDENT REPLAY PREPARED', metadata['total_states'], 'STATES;UNUSED FUTURE NATIVE STATES EXCLUDED')
