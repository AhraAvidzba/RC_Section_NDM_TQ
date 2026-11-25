"""ДЛЯ ТЕСТОВ. Повторяет точную принятую холодную историю и одну дуговую попытку на её конце без новой компиляции."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
from argparse import ArgumentParser
import json,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
p=ArgumentParser(description=__doc__)
p.add_argument('--source',default='Cold802NestedMaterialAccuracy_20261125')
p.add_argument('--root',default='Cold803NestedAccuracyEndpoint_20261125')
a=p.parse_args();source=base/a.source;root=base/a.root
outcome=json.loads((source/'OperationalColdProof.json').read_bytes())
assert outcome['cold']['native']['own_process_closed'] and not outcome['physical_capacity_certified']
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](source/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0 and len(frames[-1]['fields'])>=4
original=run_path(str(here/'Check-TestActualReverseCapture.py'))['TestReadBinary'](source/'Input/Reverse1.bin')
run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,source,1e-6,1,1e-6,0)
prep=json.loads((root/'Preparation.json').read_bytes())
values={
 'Input/Reverse1.bin':run_path(str(here/'Stage-TestColdRawReplay.py'))['TestTrajectoryBytes'](original,frames[-1]),
 'Input/SourceColdRawHistory.bin':(source/'Native/Stock.RawHistory.bin').read_bytes(),
 'Input/OperationalArcRatio.bin':struct.pack('<id',1,1.),
 'Input/ArcCallAllowance.bin':struct.pack('<id',1,12.),Path(__file__).name:Path(__file__).read_bytes()}
for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(source/'docs/regression/Diagnostic.xlsm').read_bytes()
prep['directed_scope']=dict(case='ExactColdEndpointArcFromSameAcceptedHistory',
 actual_import_this_pass=json.loads((source/'docs/regression/CodeUpdate.json').read_bytes())['Modules'],
 actual_component_count=256,immutable_parent_code_reused=True,fresh_compilation_claimed=False,
 source_root=source.name,source_states=len(frames[-1]['fields']),all_source_fields_exact=True,
 source_history_replayed_from_zero=True,all_numerical_methods_exact_source=True,
 single_additional_arc_operation=True,arc_ratio=1.,arc_call_allowance=12,
 primary30_unchanged=True,no_new_cold_capacity_search=True,
 physical_capacity_certified=False,no_working_promotion=True)
(root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print(root.name,'EXACT',len(frames[-1]['fields']),'COLD STATES; ONE ENDPOINT ARC; COMPILED PARENT REUSED')
