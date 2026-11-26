"""ДЛЯ ТЕСТОВ. Применяет исходную независимую приёмку955 к более точному второму событию956."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;script=here/'Check-TestSequentialEventHistory.py';code=script.read_text('utf-8-sig')
code=code.replace("root=base/'Cold955SequentialEventHistory_20261126';parent=base/'Cold954RefinedEventApproach_20261126'", "root=base/'Cold956RefinedSequentialHistory_20261126';parent=base/'Cold955SequentialEventHistory_20261126'",1)
env={'__file__':str(script),'__name__':'refined_sequential_history_test'};exec(compile(code,str(script),'exec'),env)
root=env['root'];proof=json.loads((root/'SequentialEventHistoryProof.json').read_bytes())
previous=json.loads((env['parent']/'SequentialEventHistoryProof.json').read_bytes())
np.testing.assert_array_equal(env['records'][1]['field'],env['half']['field'])
assert proof['results']['half']==previous['results']['half'] and proof['results']['one_full']==previous['results']['one_full']
proof.update(second_event_tolerance_factor=.0001,half_and_full_source_unchanged=True,original_acceptance955_preserved=True)
if proof['second_event_solved']:
    result=proof['results']['two_substeps'];assert result['residual_MPa']<=.0001*result['bound_MPa']+2e-11
(root/'RefinedSequentialHistoryProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
