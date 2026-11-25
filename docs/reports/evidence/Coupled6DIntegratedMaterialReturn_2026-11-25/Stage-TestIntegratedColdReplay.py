"""ДЛЯ ТЕСТОВ. Готовит стрессовый повтор точной холодной истории752 с неизменным численным ядром."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json
import re

here = Path(__file__).resolve().parent
base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
source = base / 'Cold752IntegratedGuideDirected_20261125'
root = base / 'Cold753IntegratedColdFrozenReplay_20261125'
proof = json.loads((source / 'IntegratedMaterialColdProof.json').read_bytes())
assert proof['cold']['native']['own_process_closed']
frames, raw = run_path(str(here / 'Read-TestColdRawCapture.py'))['TestHistory'](source / 'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes'] == 0
original = run_path(str(here / 'Check-TestActualReverseCapture.py'))['TestReadBinary'](source / 'Input/Reverse1.bin')
run_path(str(here / 'Stage-TestEnergyNumerics.py'))['TestClone'](root, source, 1e-6, 1, 1e-6, 0)
archive = root / 'FrozenCompilationAncestor'
archive.mkdir()
for name in ('CompilePoint.log', 'CompilePointClosure.json'):
    (root / name).rename(archive / name)
prep = json.loads((root / 'Preparation.json').read_bytes())
values = {'Input/Reverse1.bin': run_path(str(here / 'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original, frames[-1]),
    'Input/Source752History.bin': (source / 'Native/Stock.RawHistory.bin').read_bytes()}
name = 'tests/modTestCoupledElasticState.bas'
code = (root / name).read_text('utf-8-sig')
method, = re.findall(r'(?ms)^Public Sub TestRunReversePolishModel\(.*?^End Sub', code)
assert method.count('    TestRunIntegratedMaterialReturn\n') == 1
code = code.replace(method, method.replace('    TestRunIntegratedMaterialReturn\n', '', 1), 1)
values[name] = code.replace('\n', '\r\n').encode('utf-8-sig')
values['Compile-TestExternalPoint.ps1'] = (here / 'Compile-TestExternalPoint.ps1').read_bytes()
values['ParentCodeUpdate.json'] = (source / 'docs/regression/CodeUpdate.json').read_bytes()
values['ParentCompileClosure.json'] = (source / 'CompilePointClosure.json').read_bytes()
values[Path(__file__).name] = Path(__file__).read_bytes()
for name, data in values.items():
    (root / name).write_bytes(data)
    prep['inputs'][name] = sha256(data).hexdigest().upper()
prep['directed_scope'] = dict(case='ActualCold752FrozenReplay', actual_import_this_pass=['modTestCoupledElasticState'],
    all_numerical_core_exact752=True, source_states=len(frames[-1]['fields']), source_raw_history_exact=True,
    one_stress_only_replay=True, no_new_NDS_or_capacity_search=True, physical_capacity_certified=False,
    no_working_promotion=True)
(root / 'Preparation.json').write_text(json.dumps(prep, indent=2, sort_keys=True) + '\n', encoding='utf-8')
print('753 EXACT', len(frames[-1]['fields']), 'COLD752 FIELDS; NUMERICAL CORE UNCHANGED; TEST MODULE ONLY')
