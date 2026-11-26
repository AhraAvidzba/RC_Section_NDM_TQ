"""ДЛЯ ТЕСТОВ. Проверяет найденное событие N после прогноза по решённым плечам: N=100 кН и холодный повтор N=10 кН."""
from pathlib import Path

here = Path(__file__).resolve().parent
source = here / 'Stage-TestASDAxialRepeat.py'
code = source.read_text('utf-8-sig')
for before, after in [
    ('Cold1042ASDPureNBudget_20261127', 'Cold1048ASDPureNEventSeed_20261127'),
    ('Cold1044ASDPureNColdRepeat_20261127', 'Cold1049ASDPureNSeedHighLoad_20261127'),
    ('Cold1045ASDPureNScale10kN_20261127', 'Cold1050ASDPureNSeedColdRepeat_20261127'),
    ("proof['controller_change_only'] == 'budget'", "proof['controller_change_only'] == 'seed'"),
    ("case='ASD_pure_N_' + variant", "case=('ASD_seed_100kN_regression' if variant == 'repeat' else 'ASD_seed_10kN_cold_repeat')")
]:
    assert code.count(before) == 1, before
    code = code.replace(before, after, 1)
needle = "assert json.loads((parent / 'MainCapacityExecutionProof.json').read_bytes())['native']['own_process_closed']"
assert code.count(needle) == 1
code = code.replace(needle, needle + "\nassert json.loads((parent / 'MainCapacityExecutionProof.json').read_bytes())['has_steel_limit'] == 1", 1)
exec(compile(code, str(source), 'exec'), dict(__file__=str(Path(__file__).resolve()), __name__='TestStageAxialSeedRepeat'))
