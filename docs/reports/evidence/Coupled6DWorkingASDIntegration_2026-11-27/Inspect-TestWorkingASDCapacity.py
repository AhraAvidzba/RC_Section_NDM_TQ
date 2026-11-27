"""ДЛЯ ТЕСТОВ. Показывает точное различие осевого предела без изменения ожиданий проверки."""
from pathlib import Path
from runpy import run_path
import json

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
R = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestReader']
for name in ('Cold1049ASDPureNSeedHighLoad_20261127', 'Cold1070WorkingASDCapacityTension_20261127'):
    root = base / name
    r = R((root / 'Native/MainCapacity.bin').read_bytes())
    values = r.VectorTest(); count, = r.VectorTest()
    trace = [r.VectorTest().tolist() for i in range(int(count))]
    print(name, json.dumps(dict(values=values.tolist(), trace_count=int(count), tail=trace[-3:]), sort_keys=True))
root = base / 'Cold1066ASDPureNDirectColdRepeat_20261127/Case04'
r = R((root / 'Native/MainCapacityHistory.bin').read_bytes())
for i in range(3): r.VectorTest()
count, = r.VectorTest()
trace = []
for i in range(int(count)):
    trace.append(r.VectorTest().tolist()); r.VectorTest()
print('Verified direct overload', json.dumps(dict(trace_count=int(count), tail=trace[-3:]), sort_keys=True))
print((root / 'Native/DirectState.txt').read_text('utf-8-sig')[:1900])
