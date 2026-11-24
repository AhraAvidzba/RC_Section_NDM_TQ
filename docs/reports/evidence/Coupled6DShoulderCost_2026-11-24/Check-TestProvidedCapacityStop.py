"""ДЛЯ ТЕСТОВ. Разделяет численную остановку732 по бюджету и последующий тайм-аут выгрузки."""
from pathlib import Path
from runpy import run_path
from hashlib import sha256
import json,re

here=Path(__file__).resolve().parent;root=Path('D:/CodexScratch/RC_Section_NDM_TQ/Cold732ProvidedJacobianCapacity_20261124')
native,text=run_path(str(here/'Check-TestPersistentHEvidence.py'))['TestNative'](root,'PostCrackTension')
assert native['timeout600'] and native['own_process_closed']
outcome=(root/'Native/AutomaticCapacityOutcome.txt').read_text('cp1251')
assert 'AUTOMATIC_RESULT: hasPeak=False;' in outcome and 'steelEvent=0' in outcome
passed,failed=re.findall(r'PREEXPORT_TESTS: passed=(\d+); failed=(\d+)',outcome)[0]
assert (int(passed),int(failed))==(18,1)
frames,raw=run_path(str(here/'Read-TestColdRawCapture.py'))['TestHistory'](root/'Native/Stock.RawHistory.bin')
assert raw['incomplete_tail_bytes']==0
frame=frames[-1];assert len(frame['fields'])==71
maximum=float(max(frame['lambdas']));known=172967.56279474715/120000
accuracy=float(re.findall(r'погрешность локального максимума: ([0-9.E+-]+)',outcome)[0])
assert 'проб: 96;' in outcome and accuracy<=.001
compilation=json.loads((root/'docs/regression/CodeUpdate.json').read_text('utf-8-sig'))
assert sha256((root/'docs/regression/Diagnostic.xlsm').read_bytes()).hexdigest().upper()==compilation['WorkbookSHA256']
proof=dict(native=native,numerical_search_returned_before_export=True,
    original_capacity_checks=dict(passed=int(passed),failed=int(failed)),
    numerical_stop='96 probe budget exhausted before steel limit',has_certified_capacity=False,
    native_local_peak=maximum,known_uniaxial_maximum=known,absolute_known_maximum_difference=abs(maximum-known),
    native_local_peak_accuracy=accuracy,required_capacity_accuracy=.001,
    steel_event_localized=False,last_control_strain=float(frame['fields'][-1,0]),
    accepted_states=71,final_raw_frame=frame['number'],source_raw_history=raw,
    full_native_tensor_export_incomplete=True,independent_all71_fields_pending=True,
    physical_capacity_certified=False,entire_goal_completed=False)
(root/'ProvidedCapacityStopProof.json').write_text(json.dumps(proof,indent=2,sort_keys=True)+'\n',encoding='utf-8')
print('732 NUMERICAL STOP96;71STATES;PEAK',maximum,'KNOWN ERROR',abs(maximum-known),
    'LOCAL ACCURACY',accuracy,'EPS',proof['last_control_strain'],'NO STEEL EVENT;EXPORT TIMEOUT600')
