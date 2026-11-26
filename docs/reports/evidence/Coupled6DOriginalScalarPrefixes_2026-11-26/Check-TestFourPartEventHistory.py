"""ДЛЯ ТЕСТОВ. Проверяет все четвертные равновесия957 и сравнивает конечные состояния двух и четырёх материальных подшагов."""
from pathlib import Path
import json,numpy as np
here=Path(__file__).resolve().parent;script=here/'Check-TestSequentialEventHistory.py';code=script.read_text('utf-8-sig')
changes={
 "root=base/'Cold955SequentialEventHistory_20261126';parent=base/'Cold954RefinedEventApproach_20261126'":"root=base/'Cold957FourPartEventHistory_20261126';parent=base/'Cold956RefinedSequentialHistory_20261126'",
 'for stage in range(6):':'for stage in range(8):',
 'assert r.position==len(r.raw)':'assert r.position==len(r.raw)\nraw_records=records\nrecords=[raw_records[i] for i in (0,3,4,5,6,7)]',
 "half['stress'],geometry[0]-.5*geometry[1]":"half['stress'],geometry[0]-.25*geometry[1]",
}
for before,after in changes.items():assert code.count(before)==1;code=code.replace(before,after,1)
env={'__file__':str(script),'__name__':'four_part_event_history_test'};exec(compile(code,str(script),'exec'),env)
root=env['root'];proof=json.loads((root/'SequentialEventHistoryProof.json').read_bytes())
raw=env['raw_records'];results=[]
for index in (1,2,3):
    record=raw[index];trial=record['trial']
    if trial is None:break
    result=env['TestEquilibrium'](trial['field'],trial['lambda_value'],trial['stress'],env['geometry'][0]-(1-.25*index)*env['geometry'][1])
    np.testing.assert_array_equal(record['field'],trial['field'])
    np.testing.assert_allclose(record['stress'],trial['stress'],rtol=0,atol=2e-12)
    result['fraction']=.25*index;result['iterations']=int(record['effort'][0]);results.append(result)
proof['quarter_states']=results;proof['subdivisions']=4
proof['results']['last_intermediate']=proof['results'].pop('half')
proof['results']['last_intermediate_committed']=proof['results'].pop('half_committed')
proof['results']['four_substeps']=proof['results'].pop('two_substeps')
proof['results']['four_committed']=proof['results'].pop('two_committed')
comparison=proof['comparison'];comparison['intermediate_steps']=[int(raw[i]['effort'][0]) for i in (1,2,3)]
comparison['last_intermediate_steps']=comparison.pop('half_steps')
previous=json.loads((env['parent']/'RefinedSequentialHistoryProof.json').read_bytes())
assert proof['results']['one_full']==previous['results']['one_full']
r=env['R']((env['parent']/'Native/SequentialEvent.bin').read_bytes())
for _ in range(3):r.VectorTest()
for index in range(3):
    info,_,_=[r.VectorTest() for _ in range(3)]
    assert info[5]
    field=r.VectorTest();lam=float(r.VectorTest()[0])
    for _ in range(129*2):r.VectorTest()
    r.VectorTest()
    for _ in range(129*4):r.VectorTest()
    if index==2:two_field=field;two_lambda=lam
four=env['two'];length=env['TestNorm'](env['full']['field']-env['source']['fields'][-1])
field_difference=env['TestNorm'](four['field']-two_field)/length;lambda_difference=abs(four['lambda_value']-two_lambda)
same_surface=comparison['same_terminal_yield_surface'] and previous['comparison']['same_terminal_yield_surface']
refinement=dict(two_lambda=two_lambda,four_lambda=four['lambda_value'],endpoint_field_difference_relative_to_original_increment=field_difference,
    endpoint_lambda_difference=lambda_difference,same_terminal_yield_surface=same_surface,
    below_existing_field_and_lambda_noise_thresholds=bool(same_surface and field_difference<=.005 and lambda_difference<=comparison['existing_lambda_noise']),
    original_geometric_guard_not_changed=True,does_not_establish_rigorous_error_bound_or_physical_capacity=True)
proof['two_to_four_refinement']=refinement
(root/'FourPartEventHistoryProof.json').write_text(json.dumps(proof,sort_keys=True,indent=2)+'\n',encoding='utf-8')
print('TWO TO FOUR REFINEMENT',json.dumps(refinement,sort_keys=True))
