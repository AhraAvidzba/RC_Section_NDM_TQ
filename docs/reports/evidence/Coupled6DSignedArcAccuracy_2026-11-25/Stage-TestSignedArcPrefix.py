"""ДЛЯ ТЕСТОВ. Проверяет дуговую пару от точного префикса холодной истории без изменения собранного VBA."""
from pathlib import Path
from runpy import run_path
from argparse import ArgumentParser
from hashlib import sha256
import json
import struct

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
p = ArgumentParser(description=__doc__)
p.add_argument('--root', type=Path, required=True)
p.add_argument('--count', type=int, required=True)
p.add_argument('--ratio', type=float, required=True)
p.add_argument('--parent', type=Path, default=base / 'Cold761SignedArcHybrid_20261125')
a = p.parse_args()
assert 4 <= a.count <= 36 and 0 < a.ratio <= 1
parent = a.parent
assert parent in (base / 'Cold761SignedArcHybrid_20261125', base / 'Cold764SignedArcTightPair_20261125')
source = base / 'Cold753IntegratedColdFrozenReplay_20261125'
history = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source / 'Input/Reverse1.bin')
assert len(history['fields']) == 36
prefix = {key: history[key][:a.count] for key in ('fields', 'fractions', 'lambdas')}
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](a.root, parent, 1e-6, 1, 1e-6, 0)
prep = json.loads((a.root / 'Preparation.json').read_bytes())
values = {
    'Input/Reverse1.bin': run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](history, prefix),
    'Input/SignedArcOptions.bin': struct.pack('<id', 1, a.ratio),
    Path(__file__).name: Path(__file__).read_bytes(),
}
for name, data in values.items():
    (a.root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
assert (a.root / 'docs/regression/Diagnostic.xlsm').read_bytes() == (parent / 'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope'].update(case='SignedArcExactColdPrefix', source_history_states=a.count,
    signed_arc_ratio=a.ratio, numerical_core_exact761=True, original_force_tolerance=True,
    fresh_compilation_claimed=False, immutable_compiled255_reused=True,
    accepted_point_polished=False, full_and_half_same_history_combined30=True,
    no_working_promotion=True, physical_capacity_certified=False)
prep['directed_scope']['local_solve_tolerance_factor'] = .0001 if parent.name.startswith('Cold764') else 1.
(a.root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('EXACT COLD PREFIX', a.count, 'RATIO', a.ratio, 'COMPILED PARENT', parent.name, 'UNCHANGED30 BUDGET')
