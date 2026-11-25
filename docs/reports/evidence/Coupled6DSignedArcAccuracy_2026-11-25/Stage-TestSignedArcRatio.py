"""ДЛЯ ТЕСТОВ. Меняет только долю дуговой секущей при неизменном скомпилированном коде и принятой истории."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
from hashlib import sha256
import struct
import json

here = Path(__file__).resolve().parent
parser = ArgumentParser(description=__doc__)
parser.add_argument('--parent', type=Path, required=True)
parser.add_argument('--root', type=Path, required=True)
parser.add_argument('--ratio', type=float, required=True)
args = parser.parse_args()
parent, root = args.parent, args.root
assert 0 < args.ratio <= 1
native, _ = run_path(str(here / 'Check-TestPersistentHEvidence.py'))['TestNative'](parent, 'RealColdReplay')
assert native['own_process_closed'] and not native['timeout600']
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, parent, 1e-6, 1, 1e-6, 0)
prep = json.loads((root / 'Preparation.json').read_bytes())
values = {'Input/SignedArcOptions.bin': struct.pack('<id', 1, args.ratio), Path(__file__).name: Path(__file__).read_bytes()}
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
assert (root / 'docs/regression/Diagnostic.xlsm').read_bytes() == (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope'].update(case='SignedArcPairChangedRatioOnly', signed_arc_ratio=args.ratio,
    unchanged_actual255_compiled_source_reused=True, fresh_compilation_claimed=False,
    original35_history_and_all_loads_unchanged=True, no_working_promotion=True, physical_capacity_certified=False)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('SIGNED ARC RATIO ONLY', args.ratio, 'IMMUTABLE COMPILED BOOK; SAME HISTORY AND30 BUDGET')
