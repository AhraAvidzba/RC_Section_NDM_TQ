"""ДЛЯ ТЕСТОВ. Проверяет арматурный отказ чистого N, холодную повторяемость и масштаб нагрузки, не подменяя бетонный предел."""
from pathlib import Path
from hashlib import sha256
import json, re

base = Path('D:/CodexScratch/RC_Section_NDM_TQ')
roots = {1042: base / 'Cold1042ASDPureNBudget_20261127',
         1044: base / 'Cold1044ASDPureNColdRepeat_20261127',
         1045: base / 'Cold1045ASDPureNScale10kN_20261127'}
events = {}
number = r'[-+]?\d+(?:\.\d+)?(?:[Ee][-+]?\d+)?'
for case, root in roots.items():
    observed = json.loads((root / 'MainCapacityExecutionProof.json').read_bytes())
    external = json.loads((root / 'Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
    assert observed['native']['own_process_closed'] and observed['has_steel_limit'] == 1
    assert external['all56plus4_equations_and_six_resultants_verified'] and external['reference_tolerance_violated_states'] == 0
    diagnostic = (root / 'Native/MainCapacity.txt').read_text('utf-8-sig')
    low, high = map(float, re.search(rf'Арматурное событие: lambda до = ({number}), lambda после = ({number})', diagnostic).groups())
    admissible, exceeded = map(float, re.search(rf'отношения деформаций к пределам: ({number}) и ({number})', diagnostic).groups())
    assert 0 < low <= high and admissible <= 1 < exceeded
    assert abs(low - observed['steel_limit_lambda']) < 1e-12
    assert abs(external['records'][-1]['steel_strain_utilization'] - admissible) < 1e-10
    force = external['direction_N_Nmm'][0]
    if case in (1042, 1044):
        assert force == 100000 and high < 1 and diagnostic.startswith('Status=FAIL;')
        assert observed['has_ultimate_point'] == 1 and observed['lambda_capacity'] == observed['steel_limit_lambda']
    else:
        assert force == 10000 and low > 1
        assert diagnostic.startswith('Status=NumFail;') and observed['has_ultimate_point'] == 0
    events[str(case)] = dict(base_N_N=force, lambda_interval=[low, high], N_interval_N=[force * low, force * high],
                            seconds=observed['elapsed_seconds'], accepted_states=external['accepted_states'],
                            admissible_utilization=admissible, exceeded_utilization=exceeded,
                            status=diagnostic.split(';')[0].split('=')[1],
                            admissible_field_independently_replayed=True, exceeding_field_not_exported=True,
                            native_bracket_from_audited_solver=True,
                            history_sha256=sha256((root / 'Native/MainCapacityHistory.bin').read_bytes()).hexdigest().upper())
assert events['1042']['history_sha256'] == events['1044']['history_sha256']
external_root = base / 'Cold1039ASDPureNTension_20261127'
driver = json.loads((external_root / 'Native/IndependentAxialTRF.json').read_bytes())
checked = json.loads((external_root / 'IndependentAxialTRFTestReplay/Native/OpenSeesMainCapacityEquilibrium.json').read_bytes())
assert driver['termination'] is None and checked['accepted_states'] == 43
assert checked['reference_tolerance_violated_states'] == 0
reference = driver['records'][-1]['force_N']
assert events['1042']['N_interval_N'][0] <= reference <= events['1042']['N_interval_N'][1]
a, b = events['1042']['N_interval_N'], events['1045']['N_interval_N']
proof = dict(events=events, cold_repeat_history_bitexact=True, independent_N_at_steel_limit_N=reference,
             independent_reference_inside_native_bracket=True,
             two_base_load_event_intervals_overlap=max(a[0], b[0]) <= min(a[1], b[1]),
             lower_event_difference_from_independent_percent=100 * (a[0] - reference) / reference,
             steel_failure_before_100kN_design_load_verified=True,
             global_concrete_capacity_unresolved=True, working_controller_not_promoted=True,
             whole_goal_completed=False)
(roots[1042] / 'AxialSteelFailureProof.json').write_text(json.dumps(proof, sort_keys=True, indent=2) + '\n', encoding='utf-8')
print('AXIAL STEEL EVENT VERIFIED', json.dumps(proof, sort_keys=True))
