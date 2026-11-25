"""ДЛЯ ТЕСТОВ. Подготавливает холодный запуск и проверки границы бюджета4/5 на неизменяемой сборке807."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,struct

here=Path(__file__).resolve().parent;base=Path('D:/CodexScratch/RC_Section_NDM_TQ')
parent=base/'Cold807PredictorLoadingEvent_20261125'
proof=json.loads((parent/'ColdEndpointArcProof.json').read_bytes())
assert proof['native']['own_process_closed'] and proof['returned_verified_candidate']
assert proof['event_check']['point']==25 and proof['event_check']['side']==1
for name,allowance,cold in [('Cold808PredictorEventCold_20261125',12,True),
 ('Cold809PredictorEventFourCalls_20261125',4,False),('Cold810PredictorEventFiveCalls_20261125',5,False)]:
 root=base/name
 run_path(str(here/'Stage-TestEnergyNumerics.py'))['TestClone'](root,parent,1e-6,1,1e-6,0)
 prep=json.loads((root/'Preparation.json').read_bytes())
 values={'Input/ArcCallAllowance.bin':struct.pack('<id',1,float(allowance)),Path(__file__).name:Path(__file__).read_bytes()}
 for n,data in values.items():(root/n).write_bytes(data);prep['inputs'][n]=sha256(data).hexdigest().upper()
 assert (root/'docs/regression/Diagnostic.xlsm').read_bytes()==(parent/'docs/regression/Diagnostic.xlsm').read_bytes()
 prep['directed_scope']=dict(case=('ActualColdWithPredictorLoadingEvent' if cold else 'PredictorEventExactCallBudget'),
  actual_import_this_pass=['CCoupledArcContinuation'],actual_component_count=256,
  immutable_compiled807_reused=True,fresh_compilation_claimed=False,
  all_working_code_and_physical_inputs_exact807=True,arc_call_allowance=allowance,
  cold_controller_starts_from_zero=cold,fixed45_endpoint_only=not cold,
  no_physical_rule_or_acceptance_change=True,physical_capacity_certified=False,no_working_promotion=True)
 (root/'Preparation.json').write_text(json.dumps(prep,indent=2,sort_keys=True)+'\n',encoding='utf-8')
 print(name,'ALLOWANCE',allowance,'COLD',cold,'SAME807 COMPILED CODE')
