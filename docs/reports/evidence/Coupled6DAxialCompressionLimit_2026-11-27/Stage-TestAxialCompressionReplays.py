"""ДЛЯ ТЕСТОВ. Готовит точное независимое повторение полной ветви VBA или внешнего осевого пути без неиспользованных будущих состояний."""
from pathlib import Path
from hashlib import sha256
from runpy import run_path
from argparse import ArgumentParser
import json, struct, shutil


def TestVector(values):
    """Сохраняет исходные Double в том же двоичном контракте без десятичного округления."""
    return struct.pack('<i' + 'd' * len(values), len(values), *values)


def TestReadRecords(raw, magic, header_vectors):
    """Проверяет границы всего пакета и возвращает точные байты записей с параметрами пути."""
    reader = run_path(str(Path(__file__).resolve().parent / 'Read-TestColdRawCapture.py'))['TestReader'](raw)
    assert raw[:16] == magic
    for index in range(header_vectors):
        reader.VectorTest()
    before_count = reader.position
    count, = reader.VectorTest()
    assert count == int(count) and count > 1
    records = []
    for index in range(int(count)):
        start = reader.position
        fraction, factor = reader.VectorTest()
        field = reader.VectorTest()
        assert fraction == 1 and len(field) == 60
        records.append((fraction, factor, field, raw[start:reader.position]))
    assert reader.position == len(raw)
    return raw[:before_count], records


parser = ArgumentParser(description=__doc__)
parser.add_argument('--root', type=Path, required=True)
parser.add_argument('--mode', choices=('full-branch', 'independent'), required=True)
args = parser.parse_args()
root = args.root
raw = (root / 'Native/MainCapacityHistory.bin').read_bytes()
header, original = TestReadRecords(raw, b'NDMMAINHIST00003', 3)
out = root / ('FullCompressionBranchTestReplay' if args.mode == 'full-branch' else 'IndependentCompressionFineTestReplay')
assert not out.exists()
if args.mode == 'full-branch':
    execution = json.loads((root / 'MainCapacityExecutionProof.json').read_bytes())
    assert execution['native']['own_process_closed'] and execution['has_ultimate_point'] == 1
    source = (root / 'Native/FullFollowedHistory.bin').read_bytes()
    _, records = TestReadRecords(source, b'NDMFULLTRACE0001', 0)
    assert len(records) > len(original)
    combined_records = b''.join(record[3] for record in records)
    extra = dict(full_branch_sha256=sha256(source).hexdigest().upper(),
                 saved_peak_prefix_states=len(original), full_branch_states=len(records),
                 source='Read-only export before peak restoration; no material state changed')
else:
    driver_bytes = (root / 'Native/IndependentCompressionPeakTRF.json').read_bytes()
    driver = json.loads(driver_bytes)
    assert driver['source_history_sha256'] == sha256(raw).hexdigest().upper()
    assert driver['termination'] is None
    inherited = driver['inherited_states']
    solved = driver['records']
    assert 1 < inherited < len(original) and solved and all(record['accepted'] for record in solved)
    prefix = original[:inherited]
    assert abs(solved[0]['axial_strain']) > abs(prefix[-1][2][0])
    combined_records = b''.join(record[3] for record in prefix)
    for record in solved:
        combined_records += TestVector([1., record['lambda_value']]) + TestVector(record['field'])
    records = prefix + solved
    extra = dict(driver_sha256=sha256(driver_bytes).hexdigest().upper(),
                 inherited_states=inherited, independent_states=len(solved),
                 unused_future_states_excluded=len(original) - inherited,
                 source='Independent OpenSees material plus SciPy global solve; same exported section operators')
combined = header + TestVector([len(records)]) + combined_records
for directory in ('Input', 'Native', 'docs/regression'):
    (out / directory).mkdir(parents=True, exist_ok=True)
for name in ['EarlySourceProof.json', 'Input/CapacityLoad.bin', 'Input/MaterialContract.json',
             'Native/MainCapacity.Operators.csv', 'docs/regression/Diagnostic.xlsm']:
    shutil.copyfile(root / name, out / name)
(out / 'Native/MainCapacityHistory.bin').write_bytes(combined)
metadata = dict(mode=args.mode, source_history_sha256=sha256(raw).hexdigest().upper(),
                new_history_sha256=sha256(combined).hexdigest().upper(), total_states=len(records),
                all_double_values_kept_bitexact=True, physical_capacity_certified=False, **extra)
(out / 'GeneratedTraceProof.json').write_text(json.dumps(metadata, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('AXIAL REPLAY PREPARED', out, json.dumps(metadata, sort_keys=True))
