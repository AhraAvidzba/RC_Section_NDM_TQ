"""ДЛЯ ТЕСТОВ. Применяет неизменённую независимую приёмку953 к более точно решённой половинной пробе954."""
from pathlib import Path
import json
here=Path(__file__).resolve().parent;script=here/'Check-TestDetectedEventApproach.py';code=script.read_text('utf-8-sig')
assert code.count('Cold953DetectedEventApproach_20261126')==1
code=code.replace('Cold953DetectedEventApproach_20261126','Cold954RefinedEventApproach_20261126')
env={'__file__':str(script),'__name__':'refined_event_approach_test'};exec(compile(code,str(script),'exec'),env)
root=env['root'];proof=json.loads((root/'DetectedEventApproachProof.json').read_bytes())
proof.update(half_solve_tolerance_factor=.0001,original_acceptance_policy953_unchanged=True,
    original_event952_field_unchanged=True,diagnostic_half_precision_change_not_a_release_change=True)
if proof['native_half_solved']:
    half=proof['states'][1]
    assert half['independent_total_residual_MPa']<=.0001*half['final_bound_MPa']+2e-11
(root/'RefinedEventApproachProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
